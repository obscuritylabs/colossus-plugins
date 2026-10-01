using System.Text.Json;
using System.Text.RegularExpressions;

namespace Colossus.OutlookClassic;

public sealed record ItemHandle(string Kind, string StoreId, string EntryId);

public static partial class Handles
{
    [GeneratedRegex("\\A[0-9A-Fa-f]{2,8192}\\z", RegexOptions.CultureInvariant)]
    private static partial Regex HexId();

    public static string CheckId(string id)
    {
        if (id is null || id.Length % 2 != 0 || !HexId().IsMatch(id))
            throw new ArgumentException("Expected an Outlook hexadecimal identifier.");
        return id;
    }

    public static string Encode(string kind, string storeId, string entryId) =>
        Convert.ToBase64String(JsonSerializer.SerializeToUtf8Bytes(new ItemHandle(kind, CheckId(storeId), CheckId(entryId))));

    public static ItemHandle Decode(string value, string expectedKind)
    {
        if (string.IsNullOrEmpty(value) || value.Length > 24000)
            throw new ArgumentException("Invalid or oversized Outlook handle.");
        try
        {
            var handle = JsonSerializer.Deserialize<ItemHandle>(Convert.FromBase64String(value));
            if (handle is null || handle.Kind != expectedKind)
                throw new ArgumentException("Outlook handle has the wrong item kind.");
            CheckId(handle.StoreId);
            CheckId(handle.EntryId);
            return handle;
        }
        catch (Exception ex) when (ex is FormatException or JsonException)
        {
            throw new ArgumentException("Invalid Outlook handle encoding.");
        }
    }

    public static void Page(int limit, int offset)
    {
        if (limit is < 1 or > 50) throw new ArgumentException("limit must be between 1 and 50.");
        if (offset is < 0 or > 1000000) throw new ArgumentException("offset must be between 0 and 1000000.");
    }

    public static string Text(string? value, int maximum)
    {
        if (string.IsNullOrEmpty(value)) return "";
        if (value.Length <= maximum) return value;
        // String lengths are UTF-16 code units. Do not leave half of a valid
        // surrogate pair for the JSON serializer to replace with U+FFFD.
        if (maximum > 0 && char.IsHighSurrogate(value[maximum - 1]) && char.IsLowSurrogate(value[maximum]))
            maximum--;
        return value[..maximum];
    }
}
