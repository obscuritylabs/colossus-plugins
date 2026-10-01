using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Colossus.OutlookClassic;

/// <summary>Read-only OOM access. No activation, Send, Save, Delete, Move, or Quit calls.</summary>
public sealed class OutlookReader
{
    public object GetStatus()
    {
        using var com = new ComScope();
        dynamic app = com.Attach();
        dynamic session = com.Own((object)app.Session);
        dynamic stores = com.Own((object)session.Stores);
        return new
        {
            connected = true,
            client = "classic-outlook",
            version = Handles.Text((string)app.Version, 64),
            storeCount = (int)stores.Count,
            processArchitecture = RuntimeInformation.ProcessArchitecture.ToString(),
            apartment = Thread.CurrentThread.GetApartmentState().ToString(),
            readOnly = true
        };
    }

    public object ListStores(int limit, int offset)
    {
        Handles.Page(limit, offset);
        using var com = new ComScope();
        dynamic app = com.Attach();
        dynamic session = com.Own((object)app.Session);
        dynamic stores = com.Own((object)session.Stores);
        var rows = new List<object>();
        int count = stores.Count;
        int end = Math.Min(count, offset + limit);
        for (int i = offset + 1; i <= end; i++)
        {
            using var item = new ComScope();
            dynamic store = item.Own((object)stores.Item(i));
            dynamic root = item.Own((object)store.GetRootFolder());
            rows.Add(new
            {
                name = Handles.Text((string)store.DisplayName, 256),
                folderHandle = Handles.Encode("folder", (string)store.StoreID, (string)root.EntryID)
            });
        }
        return new { items = rows, nextOffset = end < count ? (int?)end : null };
    }

    public object ListFolders(string folderHandle, int limit, int offset)
    {
        var handle = Handles.Decode(folderHandle, "folder");
        Handles.Page(limit, offset);
        using var com = new ComScope();
        dynamic app = com.Attach();
        dynamic session = com.Own((object)app.Session);
        string storeId = ResolveStoreId(session, handle.StoreKey);
        dynamic parent = com.Own((object)session.GetFolderFromID(handle.EntryId, storeId));
        dynamic folders = com.Own((object)parent.Folders);
        var rows = new List<object>();
        int count = folders.Count;
        int end = Math.Min(count, offset + limit);
        for (int i = offset + 1; i <= end; i++)
        {
            using var item = new ComScope();
            dynamic folder = item.Own((object)folders.Item(i));
            rows.Add(new
            {
                name = Handles.Text((string)folder.Name, 256),
                folderHandle = Handles.Encode("folder", (string)folder.StoreID, (string)folder.EntryID),
                defaultItemType = (int)folder.DefaultItemType
            });
        }
        return new { items = rows, nextOffset = end < count ? (int?)end : null };
    }

    public object SearchMessages(string folderHandle, string? subjectContains, bool unreadOnly, int limit, int offset)
    {
        var handle = Handles.Decode(folderHandle, "folder");
        Handles.Page(limit, offset);
        if (subjectContains?.Length > 256) throw new ArgumentException("subjectContains is limited to 256 characters.");
        using var com = new ComScope();
        dynamic app = com.Attach();
        dynamic session = com.Own((object)app.Session);
        string storeId = ResolveStoreId(session, handle.StoreKey);
        dynamic folder = com.Own((object)session.GetFolderFromID(handle.EntryId, storeId));
        dynamic items = com.Own((object)folder.Items);
        if (unreadOnly) items = com.Own((object)items.Restrict("[Unread] = true"));
        items.Sort("[ReceivedTime]", true);
        int count = items.Count;
        int position = Math.Min(offset, count);
        int scanned = 0;
        var rows = new List<object>();
        // A bounded scan avoids query injection and unbounded mailbox enumeration.
        while (position < count && scanned < 500 && rows.Count < limit)
        {
            using var item = new ComScope();
            dynamic mail = item.Own((object)items.Item(++position));
            scanned++;
            if ((int)mail.Class != 43) continue; // olMail
            string subject = (string?)mail.Subject ?? "";
            if (!string.IsNullOrEmpty(subjectContains) && !subject.Contains(subjectContains, StringComparison.OrdinalIgnoreCase)) continue;
            rows.Add(new
            {
                messageHandle = Handles.Encode("message", storeId, (string)mail.EntryID),
                subject = Handles.Text(subject, 512),
                senderName = Handles.Text((string?)mail.SenderName, 256),
                receivedAt = ((DateTime)mail.ReceivedTime).ToUniversalTime().ToString("O"),
                unread = (bool)mail.UnRead
            });
        }
        return new
        {
            items = rows,
            nextOffset = position < count ? (int?)position : null,
            scanned,
            ordering = "received-descending; offsets are best-effort while the folder changes"
        };
    }

    public object GetMessage(string messageHandle, int maxBodyChars)
    {
        var handle = Handles.Decode(messageHandle, "message");
        if (maxBodyChars is < 0 or > 24000) throw new ArgumentException("maxBodyChars must be between 0 and 24000.");
        using var com = new ComScope();
        dynamic app = com.Attach();
        dynamic session = com.Own((object)app.Session);
        string storeId = ResolveStoreId(session, handle.StoreKey);
        dynamic mail = com.Own((object)session.GetItemFromID(handle.EntryId, storeId));
        RequireMail(mail);
        string body = maxBodyChars == 0 ? "" : (string?)mail.Body ?? "";
        return new
        {
            messageHandle,
            subject = Handles.Text((string?)mail.Subject, 512),
            senderName = Handles.Text((string?)mail.SenderName, 256),
            senderAddress = Handles.Text((string?)mail.SenderEmailAddress, 1024),
            receivedAt = ((DateTime)mail.ReceivedTime).ToUniversalTime().ToString("O"),
            unread = (bool)mail.UnRead,
            body = Handles.Text(body, maxBodyChars),
            bodyIncluded = maxBodyChars > 0,
            bodyTruncated = body.Length > maxBodyChars,
            contentTrust = "untrusted-email"
        };
    }

    public object ListAttachments(string messageHandle, int limit, int offset)
    {
        var handle = Handles.Decode(messageHandle, "message");
        Handles.Page(limit, offset);
        using var com = new ComScope();
        dynamic app = com.Attach();
        dynamic session = com.Own((object)app.Session);
        string storeId = ResolveStoreId(session, handle.StoreKey);
        dynamic mail = com.Own((object)session.GetItemFromID(handle.EntryId, storeId));
        RequireMail(mail);
        dynamic attachments = com.Own((object)mail.Attachments);
        int count = attachments.Count;
        int end = Math.Min(count, offset + limit);
        var rows = new List<object>();
        for (int i = offset + 1; i <= end; i++)
        {
            using var item = new ComScope();
            dynamic attachment = item.Own((object)attachments.Item(i));
            rows.Add(new { index = i, name = Handles.Text((string?)attachment.FileName, 512), size = (int)attachment.Size });
        }
        return new { items = rows, nextOffset = end < count ? (int?)end : null };
    }

    private static string ResolveStoreId(dynamic session, string key)
    {
        using var com = new ComScope();
        dynamic stores = com.Own((object)session.Stores);
        string? match = null;
        int count = stores.Count;
        for (int i = 1; i <= count; i++)
        {
            using var item = new ComScope();
            dynamic store = item.Own((object)stores.Item(i));
            string id = (string)store.StoreID;
            if (Handles.StoreKey(id) != key) continue;
            if (match is not null) throw new ArgumentException("Outlook store fingerprint is ambiguous; repeat store discovery.");
            match = id;
        }
        return match ?? throw new ArgumentException("Outlook store is no longer available; repeat store discovery.");
    }

    private static void RequireMail(dynamic item)
    {
        if ((int)item.Class != 43) throw new ArgumentException("The handle does not identify an Outlook mail item.");
    }
}

internal sealed class ComScope : IDisposable
{
    private readonly List<object> owned = [];
    // Outlook.Application CLSID: attach only; never instantiate or start Outlook.
    private static readonly Guid OutlookClass = new("0006F03A-0000-0000-C000-000000000046");

    [DllImport("oleaut32.dll", PreserveSig = false)]
    private static extern void GetActiveObject(in Guid clsid, IntPtr reserved, [MarshalAs(UnmanagedType.IUnknown)] out object value);

    public object Own(object value)
    {
        if (!Marshal.IsComObject(value)) throw new InvalidOperationException("Expected an Outlook COM object.");
        owned.Add(value);
        return value;
    }

    public object Attach()
    {
        if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
            throw new InvalidOperationException("Outlook COM must run on the dedicated STA thread.");
        GetActiveObject(in OutlookClass, IntPtr.Zero, out var app);
        return Own(app);
    }

    public void Dispose()
    {
        for (int i = owned.Count - 1; i >= 0; i--)
        {
            try { Marshal.ReleaseComObject(owned[i]); }
            catch (InvalidComObjectException) { }
        }
        owned.Clear();
    }
}
