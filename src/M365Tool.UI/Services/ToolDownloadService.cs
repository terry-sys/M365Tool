using System.Net.Http;

namespace Office365CleanupTool.Services;

internal static class ToolDownloadService
{
    private static readonly HttpClient Client = new() { Timeout = System.Threading.Timeout.InfiniteTimeSpan };

    public static Task DownloadAsync(string url, string path, CancellationToken cancellationToken) =>
        DownloadAsync(url, path, cancellationToken, TimeSpan.FromMinutes(10));

    internal static async Task DownloadAsync(string url, string path, CancellationToken cancellationToken, TimeSpan timeout)
    {
        using var deadline = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        deadline.CancelAfter(timeout);
        try
        {
            using var response = await Client.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, deadline.Token);
            response.EnsureSuccessStatusCode();
            await using var file = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None);
            await response.Content.CopyToAsync(file, deadline.Token);
        }
        catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
        {
            throw new TimeoutException($"Tool download timed out after {timeout.TotalMinutes:g} minutes.");
        }
    }
}
