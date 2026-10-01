using Colossus.OutlookClassic;

static void Check(bool condition, string name)
{
    if (!condition) throw new Exception(name);
    Console.WriteLine($"PASS {name}");
}

static void Reject(Action action, string name)
{
    try { action(); }
    catch (ArgumentException) { Console.WriteLine($"PASS {name}"); return; }
    throw new Exception($"Expected rejection: {name}");
}

var encoded = Handles.Encode("folder", "AABB", "0123");
Check(Handles.Decode(encoded, "folder") == new ItemHandle("folder", "AABB", "0123"), "opaque handle round trip");
Reject(() => Handles.Decode(encoded, "message"), "folder handles cannot address messages");
Reject(() => Handles.Decode("not base64!", "folder"), "malformed handle rejected");
Reject(() => Handles.Decode(new string('a', 24001), "folder"), "oversized handle rejected");
Reject(() => Handles.CheckId("../etc"), "non-Outlook identifiers rejected");
Reject(() => Handles.Page(51, 0), "result count bounded");
Reject(() => Handles.Page(1, -1), "negative page offset rejected");
Reject(() => new OutlookReader().GetMessage("invalid", 5), "bad handles rejected before COM attachment");
Check(Handles.Text("邮件 🔎", 2) == "邮件", "Unicode body bound");
Check(Handles.Text("🔎 sample", 1) == "", "truncation never returns half an emoji");
Check(Handles.Text("a🔎b", 2) == "a", "truncation backs up before a surrogate pair");
Check(Handles.Text("a🔎b", 3) == "a🔎", "complete emoji fits within the UTF-16 limit");

using (var sta = new StaDispatcher())
{
    var first = await sta.InvokeAsync(() => (Environment.CurrentManagedThreadId,
        Thread.CurrentThread.GetApartmentState(), Application.MessageLoop), CancellationToken.None);
    var second = await sta.InvokeAsync(() => Environment.CurrentManagedThreadId, CancellationToken.None);
    Check(first.Item1 == second && first.Item2 == ApartmentState.STA && first.Item3,
        "operations retain one STA thread with active message pump");

    using var started = new ManualResetEventSlim();
    using var release = new ManualResetEventSlim();
    var inFlight = sta.InvokeAsync(() => { started.Set(); release.Wait(TimeSpan.FromSeconds(3)); return true; }, CancellationToken.None);
    Check(started.Wait(TimeSpan.FromSeconds(3)), "test operation starts on STA");
    try
    {
        await sta.InvokeAsync(() => false, CancellationToken.None);
        throw new Exception("Concurrent COM operations were admitted");
    }
    catch (InvalidOperationException) { Console.WriteLine("PASS concurrent COM operation rejected"); }
    finally { release.Set(); }
    await inFlight;
}

using (var sta = new StaDispatcher())
{
    using var started = new ManualResetEventSlim();
    using var release = new ManualResetEventSlim();
    using var cancellation = new CancellationTokenSource();
    var inFlight = sta.InvokeAsync(() => { started.Set(); release.Wait(TimeSpan.FromSeconds(3)); return true; }, cancellation.Token);
    Check(started.Wait(TimeSpan.FromSeconds(3)), "cancellable operation starts");
    cancellation.Cancel();
    try { await inFlight; throw new Exception("Cancellation was ignored"); }
    catch (OperationCanceledException) { Console.WriteLine("PASS in-flight cancellation reported"); }
    try { await sta.InvokeAsync(() => true, CancellationToken.None); throw new Exception("Timed-out STA reused"); }
    catch (InvalidOperationException) { Console.WriteLine("PASS cancelled COM session cannot accept more work"); }
    finally { release.Set(); }
}

Console.WriteLine("All checks passed; no Outlook connection was made.");
