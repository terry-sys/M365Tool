using Microsoft.Win32;
using System.Text.RegularExpressions;

namespace Office365CleanupTool.Services;

public sealed class NetworkRepairOptions
{
    public bool EnableTls12 { get; set; } = true;
    public bool EnableAutoDetect { get; set; }
    public bool ImportUserProxy { get; set; }
}

public sealed class NetworkRegistryValue
{
    public string Name { get; set; } = string.Empty;
    public bool Exists { get; set; }
    public RegistryValueKind Kind { get; set; }
    public int? Number { get; set; }
    public string? Text { get; set; }
}

public sealed class NetworkRepairBackup
{
    public int FormatVersion { get; set; } = 1;
    public string UserSid { get; set; } = string.Empty;
    public List<NetworkRegistryValue> Values { get; set; } = new();
    public string WinHttpScript { get; set; } = string.Empty;

    internal static void Validate(NetworkRepairBackup backup)
    {
        string[] names = { "SecureProtocols", "AutoDetect", "AutoConfigURL" };
        if (backup.FormatVersion != 1 || backup.Values.Count != 3
            || !backup.Values.Select(item => item.Name).Order().SequenceEqual(names.Order())
            || string.IsNullOrWhiteSpace(backup.WinHttpScript))
            throw new InvalidDataException("Invalid network backup.");
        foreach (var value in backup.Values.Where(item => item.Exists))
        {
            if ((value.Kind == RegistryValueKind.DWord && value.Number.HasValue)
                || ((value.Kind == RegistryValueKind.String || value.Kind == RegistryValueKind.ExpandString) && value.Text != null)) continue;
            throw new InvalidDataException("Unsupported registry value in network backup.");
        }
        string[] commands = backup.WinHttpScript.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(line => line.Trim()).Where(line => line.Length > 0 && !line.StartsWith('#')).ToArray();
        if (commands.Length < 3 || !commands[0].Equals("pushd winhttp", StringComparison.OrdinalIgnoreCase)
            || !commands[^1].Equals("popd", StringComparison.OrdinalIgnoreCase)
            || commands.Skip(1).SkipLast(1).Any(command => !Regex.IsMatch(command,
                @"^(reset\s+proxy|set\s+(proxy|advproxy)\s+.+)$", RegexOptions.IgnoreCase)))
            throw new InvalidDataException("The backup contains unsupported WinHTTP commands.");
    }
}
