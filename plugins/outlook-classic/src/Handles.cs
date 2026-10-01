using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace Colossus.OutlookClassic;

public sealed record ItemHandle(string Kind, string StoreKey, string EntryId);

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

    public static string StoreKey(string storeId) =>
        Convert.ToHexString(SHA256.HashData(Convert.FromHexString(CheckId(storeId))).AsSpan(0, 16));

    private static string KindTag(string kind) => kind switch
    {
        "folder" => "f2",
        "message" => "m2",
        _ => throw new ArgumentException("Unknown Outlook handle kind.")
    };

    private static string Checksum(string payload) =>
        Convert.ToHexString(SHA256.HashData(Encoding.ASCII.GetBytes(payload)).AsSpan(0, 4));

    // The store fingerprint avoids repeating very long, path-bearing PST StoreIDs.
    // The checksum detects copying mistakes; it is not an authorization mechanism.
    public static string Encode(string kind, string storeId, string entryId)
    {
        var payload = $"{KindTag(kind)}.{StoreKey(storeId)}.{CheckId(entryId).ToUpperInvariant()}";
        return $"{payload}.{Checksum(payload)}";
    }

    public static ItemHandle Decode(string value, string expectedKind)
    {
        if (string.IsNullOrEmpty(value) || value.Length > 8240)
            throw new ArgumentException("Invalid or oversized Outlook handle.");
        var parts = value.Split('.');
        if (parts.Length != 4 || parts[0] != KindTag(expectedKind) || parts[1].Length != 32)
            throw new ArgumentException("Invalid Outlook handle format or kind. Repeat the lookup and copy the complete returned handle.");
        CheckId(parts[1]);
        CheckId(parts[2]);
        var payload = string.Join('.', parts, 0, 3);
        if (!string.Equals(parts[3], Checksum(payload), StringComparison.Ordinal))
            throw new ArgumentException("Outlook handle was altered. Repeat the lookup and copy the complete returned handle.");
        return new ItemHandle(expectedKind, parts[1], parts[2]);
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
