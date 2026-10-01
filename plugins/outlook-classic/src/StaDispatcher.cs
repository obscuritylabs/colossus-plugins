using System.Runtime.InteropServices;

namespace Colossus.OutlookClassic;

/// <summary>All COM lifetimes stay on one STA with a Windows message pump.</summary>
public sealed class StaDispatcher : IDisposable
{
    private readonly Thread thread;
    private readonly TaskCompletionSource<Control> ready = new(TaskCreationOptions.RunContinuationsAsynchronously);
    private readonly SemaphoreSlim gate = new(1, 1);
    private int poisoned;
    private int disposed;

    public StaDispatcher()
    {
        thread = new Thread(() =>
        {
            try
            {
                using var control = new Control();
                _ = control.Handle;
                ready.TrySetResult(control);
                Application.Run(); // No form or visible window is created.
            }
            catch (Exception ex) { ready.TrySetException(ex); }
        }) { IsBackground = true, Name = "Outlook COM STA" };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
    }

    public async Task<T> InvokeAsync<T>(Func<T> action, CancellationToken cancellationToken)
    {
        ObjectDisposedException.ThrowIf(Volatile.Read(ref disposed) != 0, this);
        if (Volatile.Read(ref poisoned) != 0)
            throw new InvalidOperationException("Outlook connection timed out; restart the MCP process.");
        if (!await gate.WaitAsync(0, cancellationToken))
            throw new InvalidOperationException("Outlook is busy; only one COM operation may run at a time.");

        var finished = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        using var deadline = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        deadline.CancelAfter(TimeSpan.FromSeconds(20));
        var deadlineToken = deadline.Token;
        try
        {
            if (Volatile.Read(ref poisoned) != 0)
                throw new InvalidOperationException("Outlook connection timed out; restart the MCP process.");
            var control = await ready.Task.WaitAsync(deadline.Token);
            control.BeginInvoke((Action)(() =>
            {
                try
                {
                    deadlineToken.ThrowIfCancellationRequested();
                    finished.TrySetResult(action());
                }
                catch (Exception ex) { finished.TrySetException(ex); }
            }));
            return await finished.Task.WaitAsync(deadline.Token);
        }
        catch (OperationCanceledException)
        {
            // Cancellation cannot abort a COM RPC already executing. Do not queue more work.
            Interlocked.Exchange(ref poisoned, 1);
            if (!cancellationToken.IsCancellationRequested)
                throw new TimeoutException("Outlook did not respond within 20 seconds; restart the MCP process.");
            throw;
        }
        finally { gate.Release(); }
    }

    public void Dispose()
    {
        if (Interlocked.Exchange(ref disposed, 1) != 0) return;
        if (ready.Task.IsCompletedSuccessfully)
        {
            try { ready.Task.Result.BeginInvoke((Action)Application.ExitThread); }
            catch (InvalidOperationException) { }
        }
        thread.Join(TimeSpan.FromSeconds(1));
    }
}

internal static class SafeError
{
    public static string Describe(Exception ex) => ex switch
    {
        COMException com => $"Outlook COM unavailable (HRESULT 0x{com.HResult:X8}). Check that classic Outlook is open in the same interactive user session and permitted by the execution policy.",
        ArgumentException => ex.Message,
        InvalidOperationException => ex.Message,
        TimeoutException => ex.Message,
        OperationCanceledException => "Operation cancelled; an in-flight Outlook COM call may still be completing.",
        _ => "Outlook operation failed. Check client availability and policy; no mailbox change was requested."
    };
}
