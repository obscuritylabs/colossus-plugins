using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Colossus.OutlookClassic;

/// <summary>Bounded Outlook Object Model access in the already running user session. Never sends or quits.</summary>
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
            readOnly = false,
            supportsDrafts = true,
            supportsPermanentDelete = false,
            supportsSend = false
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

    public object GetMailFolders(string storeFolderHandle)
    {
        var handle = Handles.Decode(storeFolderHandle, "folder");
        using var com = new ComScope();
        dynamic app = com.Attach();
        dynamic session = com.Own((object)app.Session);
        dynamic store = FindStore(session, handle.StoreKey, com);
        dynamic root = com.Own((object)store.GetRootFolder());
        if (!string.Equals((string)root.EntryID, handle.EntryId, StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("Pass a store root folder handle returned by list_stores.");
        var rows = new List<object>();
        foreach (var (role, number) in new[] { ("inbox", 6), ("drafts", 16), ("sentItems", 5), ("deletedItems", 3) })
        {
            try
            {
                using var item = new ComScope();
                object? folderValue = (object?)store.GetDefaultFolder(number);
                if (folderValue is null) continue;
                dynamic folder = item.Own(folderValue);
                rows.Add(new
                {
                    role,
                    name = Handles.Text((string)folder.Name, 256),
                    folderHandle = Handles.Encode("folder", (string)store.StoreID, (string)folder.EntryID)
                });
            }
            catch (COMException) { /* Some stores do not provide every standard folder. */ }
        }
        return new { items = rows };
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

    public object MarkMessageRead(string messageHandle, string sourceFolderHandle, bool read)
    {
        var message = Handles.Decode(messageHandle, "message");
        var source = Handles.Decode(sourceFolderHandle, "folder");
        RequireSameStore(message, source);
        using var com = new ComScope();
        dynamic app = com.Attach();
        dynamic session = com.Own((object)app.Session);
        string storeId = ResolveStoreId(session, message.StoreKey);
        dynamic mail = com.Own((object)session.GetItemFromID(message.EntryId, storeId));
        RequireMailInFolder(mail, source);
        mail.UnRead = !read;
        mail.Save();
        return new { messageHandle = Handles.Encode("message", storeId, (string)mail.EntryID), read = !(bool)mail.UnRead, changed = true };
    }

    public object MoveMessage(string messageHandle, string sourceFolderHandle, string destinationFolderHandle, string action)
    {
        var message = Handles.Decode(messageHandle, "message");
        var source = Handles.Decode(sourceFolderHandle, "folder");
        var destination = Handles.Decode(destinationFolderHandle, "folder");
        RequireSameStore(message, source);
        RequireSameStore(message, destination);
        if (source.EntryId.Equals(destination.EntryId, StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("Source and destination folders must differ.");
        using var com = new ComScope();
        dynamic app = com.Attach();
        dynamic session = com.Own((object)app.Session);
        string storeId = ResolveStoreId(session, message.StoreKey);
        dynamic mail = com.Own((object)session.GetItemFromID(message.EntryId, storeId));
        RequireMailInFolder(mail, source);
        dynamic folder = com.Own((object)session.GetFolderFromID(destination.EntryId, storeId));
        if (!string.Equals((string)folder.StoreID, storeId, StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("Destination must belong to the same Outlook store.");
        dynamic moved = com.Own((object)mail.Move(folder));
        RequireMailInFolder(moved, destination);
        return new
        {
            action,
            messageHandle = Handles.Encode("message", storeId, (string)moved.EntryID),
            folderHandle = destinationFolderHandle,
            subject = Handles.Text((string?)moved.Subject, 512),
            previousHandleStale = true
        };
    }

    public object DeleteMessage(string messageHandle, string sourceFolderHandle)
    {
        var message = Handles.Decode(messageHandle, "message");
        var source = Handles.Decode(sourceFolderHandle, "folder");
        RequireSameStore(message, source);
        using var com = new ComScope();
        dynamic app = com.Attach();
        dynamic session = com.Own((object)app.Session);
        dynamic store = FindStore(session, message.StoreKey, com);
        string storeId = (string)store.StoreID;
        dynamic deleted = RequireDefaultFolder(store, 3, "Deleted Items", com);
        var deletedHandle = Handles.Encode("folder", storeId, (string)deleted.EntryID);
        if (source.EntryId.Equals((string)deleted.EntryID, StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("The message is already in Deleted Items; permanent deletion is unavailable.");
        dynamic mail = com.Own((object)session.GetItemFromID(message.EntryId, storeId));
        RequireMailInFolder(mail, source);
        dynamic moved = com.Own((object)mail.Move(deleted));
        RequireMailInFolder(moved, Handles.Decode(deletedHandle, "folder"));
        return new
        {
            action = "moved-to-deleted-items",
            messageHandle = Handles.Encode("message", storeId, (string)moved.EntryID),
            folderHandle = deletedHandle,
            previousHandleStale = true,
            permanent = false
        };
    }

    public object CreateDraft(string storeFolderHandle, string? to, string? cc, string? bcc, string? subject, string? body)
    {
        var handle = Handles.Decode(storeFolderHandle, "folder");
        ValidateDraftFields(to, cc, bcc, subject, body);
        if (to is null && cc is null && bcc is null && subject is null && body is null)
            throw new ArgumentException("Supply at least one draft field.");
        using var com = new ComScope();
        dynamic app = com.Attach();
        dynamic session = com.Own((object)app.Session);
        dynamic store = FindStore(session, handle.StoreKey, com);
        string storeId = (string)store.StoreID;
        dynamic root = com.Own((object)store.GetRootFolder());
        if (!string.Equals((string)root.EntryID, handle.EntryId, StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("Pass a store root folder handle returned by list_stores.");
        dynamic drafts = RequireDefaultFolder(store, 16, "Drafts", com);
        var draftsHandle = Handles.Encode("folder", storeId, (string)drafts.EntryID);
        dynamic items = com.Own((object)drafts.Items);
        dynamic newMail = com.Own((object)items.Add("IPM.Note"));
        ApplyDraftFields(newMail, to, cc, bcc, subject, body);
        // Outlook may otherwise save a new item in the profile's default Drafts store.
        dynamic moved = com.Own((object)newMail.Move(drafts));
        RequireMailInFolder(moved, Handles.Decode(draftsHandle, "folder"));
        moved.Save();
        return DraftResult(moved, storeId, draftsHandle, "created");
    }

    public object UpdateDraft(string messageHandle, string? to, string? cc, string? bcc, string? subject, string? body)
    {
        var handle = Handles.Decode(messageHandle, "message");
        ValidateDraftFields(to, cc, bcc, subject, body);
        if (to is null && cc is null && bcc is null && subject is null && body is null)
            throw new ArgumentException("Supply at least one draft field.");
        using var com = new ComScope();
        dynamic app = com.Attach();
        dynamic session = com.Own((object)app.Session);
        dynamic store = FindStore(session, handle.StoreKey, com);
        string storeId = (string)store.StoreID;
        dynamic drafts = RequireDefaultFolder(store, 16, "Drafts", com);
        var draftsHandle = Handles.Encode("folder", storeId, (string)drafts.EntryID);
        dynamic mail = com.Own((object)session.GetItemFromID(handle.EntryId, storeId));
        RequireMailInFolder(mail, Handles.Decode(draftsHandle, "folder"));
        if ((bool)mail.Sent) throw new ArgumentException("Only unsent drafts can be edited.");
        ApplyDraftFields(mail, to, cc, bcc, subject, body);
        mail.Save();
        return DraftResult(mail, storeId, draftsHandle, "updated");
    }

    private static object DraftResult(dynamic mail, string storeId, string draftsHandle, string action) => new
    {
        action,
        messageHandle = Handles.Encode("message", storeId, (string)mail.EntryID),
        folderHandle = draftsHandle,
        to = Handles.Text((string?)mail.To, 2048),
        cc = Handles.Text((string?)mail.CC, 2048),
        bcc = Handles.Text((string?)mail.BCC, 2048),
        subject = Handles.Text((string?)mail.Subject, 512),
        bodyLength = ((string?)mail.Body ?? "").Length,
        sent = (bool)mail.Sent
    };

    private static void ApplyDraftFields(dynamic mail, string? to, string? cc, string? bcc, string? subject, string? body)
    {
        if (to is not null) mail.To = to;
        if (cc is not null) mail.CC = cc;
        if (bcc is not null) mail.BCC = bcc;
        if (subject is not null) mail.Subject = subject;
        if (body is not null) mail.Body = body;
    }

    private static void ValidateDraftFields(string? to, string? cc, string? bcc, string? subject, string? body)
    {
        if (to?.Length > 2048 || cc?.Length > 2048 || bcc?.Length > 2048)
            throw new ArgumentException("Each recipient field is limited to 2048 characters.");
        if (subject?.Length > 512) throw new ArgumentException("subject is limited to 512 characters.");
        if (body?.Length > 24000) throw new ArgumentException("body is limited to 24000 characters.");
        foreach (var value in new[] { to, cc, bcc, subject })
            if (value is not null && value.Any(char.IsControl))
                throw new ArgumentException("Recipient and subject fields cannot contain control characters.");
    }

    private static void RequireSameStore(ItemHandle a, ItemHandle b)
    {
        if (!string.Equals(a.StoreKey, b.StoreKey, StringComparison.Ordinal))
            throw new ArgumentException("Message and folder handles must belong to the same Outlook store.");
    }

    private static dynamic RequireDefaultFolder(dynamic store, int number, string name, ComScope lifetime)
    {
        object? value = (object?)store.GetDefaultFolder(number);
        if (value is null) throw new ArgumentException($"The selected Outlook store has no {name} folder.");
        return lifetime.Own(value);
    }

    private static void RequireMailInFolder(dynamic mail, ItemHandle folder)
    {
        RequireMail(mail);
        using var com = new ComScope();
        dynamic parent = com.Own((object)mail.Parent);
        if (!string.Equals((string)parent.EntryID, folder.EntryId, StringComparison.OrdinalIgnoreCase) ||
            !string.Equals(Handles.StoreKey((string)parent.StoreID), folder.StoreKey, StringComparison.Ordinal))
            throw new ArgumentException("The message is no longer in the supplied source folder; list that folder again.");
    }

    private static dynamic FindStore(dynamic session, string key, ComScope lifetime)
    {
        dynamic stores = lifetime.Own((object)session.Stores);
        object? found = null;
        int count = stores.Count;
        for (int i = 1; i <= count; i++)
        {
            dynamic store = stores.Item(i);
            if (Handles.StoreKey((string)store.StoreID) == key)
            {
                if (found is not null)
                {
                    Marshal.ReleaseComObject((object)store);
                    Marshal.ReleaseComObject(found);
                    throw new ArgumentException("Outlook store fingerprint is ambiguous; repeat store discovery.");
                }
                found = (object)store;
            }
            else Marshal.ReleaseComObject((object)store);
        }
        return lifetime.Own(found ?? throw new ArgumentException("Outlook store is no longer available; repeat store discovery."));
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
