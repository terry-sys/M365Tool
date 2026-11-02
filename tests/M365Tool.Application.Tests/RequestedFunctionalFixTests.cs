using System.Reflection;
using System.Net;
using System.Net.Sockets;
using System.Text;
using Office365CleanupTool.Services;
using Xunit;

namespace M365Tool.Application.Tests;

[Collection("Shared maintenance operation")]
public class RequestedFunctionalFixTests
{
    [Theory]
    [InlineData(0)]
    [InlineData(43)]
    public void CalendarFailureCannotBeReportedAsSuccess(int exitCode)
    {
        var result = new OutlookDiagnosticResult { ExitCode = exitCode, StandardOutput = "ExecutionResult: Failed" };
        Assert.False((bool)typeof(OutlookDiagnosticsService).GetMethod("DetermineSuccess", BindingFlags.Static | BindingFlags.NonPublic)!
            .Invoke(null, new object[] { result, "OutlookCalendarScan" })!);
    }

    [Fact]
    public void FailedTeamsResultDoesNotClaimCompletion()
    {
        var result = new TeamsMaintenanceResult { Success = false };
        string summary = (string)typeof(TeamsMaintenanceService).GetMethod("BuildMaintenanceSummary", BindingFlags.Static | BindingFlags.NonPublic)!
            .Invoke(null, new object[] { result, UiLanguage.English })!;
        Assert.DoesNotContain("reset and cache cleanup completed", summary);
    }
    [Theory]
    [InlineData(0, false)]
    [InlineData(43, false)]
    [InlineData(0, true)]
    [InlineData(43, true)]
    public void CalendarRejectsFailureOnEitherOutputStream(int code, bool stderr)
    {
        var result = new OutlookDiagnosticResult
        {
            ExitCode = code,
            StandardOutput = stderr ? "" : "2026-10-30 ExecutionResult : FAILED",
            StandardError = stderr ? "ExecutionResult: Failed" : ""
        };
        Assert.False(DetermineSuccess(result, "OutlookCalendarScan"));
    }

    [Theory]
    [InlineData(0, true)]
    [InlineData(43, true)]
    [InlineData(5, false)]
    public void CalendarKeepsExistingSuccessCodes(int code, bool expected) =>
        Assert.Equal(expected, DetermineSuccess(new OutlookDiagnosticResult { ExitCode = code, StandardOutput = "ExecutionResult: Passed" }, "OutlookCalendarScan"));

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task StalledDownloadTimesOutAndReleasesMaintenanceGate(bool sendHeaders)
    {
        using var directory = new TestDirectory();
        await using var server = new DownloadServer(sendHeaders, stall: true);
        string path = Path.Combine(directory.Path, "stalled.zip");
        using (MaintenanceCoordinator.Shared.Begin(UiLanguage.English))
        {
            Task download = ToolDownloadService.DownloadAsync(server.Url, path, CancellationToken.None, TimeSpan.FromSeconds(1));
            await server.RequestAccepted.Task.WaitAsync(TimeSpan.FromSeconds(5));
            var error = await Assert.ThrowsAsync<TimeoutException>(() => download.WaitAsync(TimeSpan.FromSeconds(5)));
            Assert.Contains("Tool download timed out", error.Message);
            if (sendHeaders) Assert.True(File.Exists(path), "The response body phase was not reached.");
        }
        Assert.False(MaintenanceCoordinator.Shared.IsBusy);
    }

    [Fact]
    public async Task CallerCanCancelAStalledDownloadBody()
    {
        using var directory = new TestDirectory();
        await using var server = new DownloadServer(sendHeaders: true, stall: true);
        using var cancellation = new CancellationTokenSource();
        Task download = ToolDownloadService.DownloadAsync(server.Url, Path.Combine(directory.Path, "canceled.zip"), cancellation.Token, TimeSpan.FromSeconds(10));
        await server.HeadersSent.Task.WaitAsync(TimeSpan.FromSeconds(5));
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => download.WaitAsync(TimeSpan.FromSeconds(5)));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task SharedDownloaderAndCleanupPreserveDownloadedBytes(bool cleanup)
    {
        using var directory = new TestDirectory();
        await using var server = new DownloadServer(sendHeaders: true, stall: false);
        string path = Path.Combine(directory.Path, "complete.zip");
        Task download = cleanup
            ? (Task)typeof(CleanupService).GetMethod("DownloadFileAsync", BindingFlags.Instance | BindingFlags.NonPublic)!
                .Invoke(new CleanupService(new NeverRun()), new object[] { server.Url, path })!
            : ToolDownloadService.DownloadAsync(server.Url, path, CancellationToken.None);
        await download.WaitAsync(TimeSpan.FromSeconds(5));
        Assert.Equal(DownloadServer.Payload, await File.ReadAllBytesAsync(path));
    }

    [Theory]
    [InlineData(UiLanguage.Chinese)]
    [InlineData(UiLanguage.English)]
    public void SuccessfulTeamsResetDoesNotRequireTheCacheToStillExist(UiLanguage language)
    {
        using var directory = new TestDirectory();
        var result = new TeamsMaintenanceResult { ResetSucceeded = true };
        TeamsMaintenanceService.ClearTeamsCache(new TeamsPathTarget("Fixture", System.IO.Path.Combine(directory.Path, "absent")), result, language);
        TeamsMaintenanceService.CompleteMaintenanceResult(result, language);
        Assert.True(result.CacheCleared);
        Assert.Empty(result.CacheTargets);
        Assert.True(result.Success);
        Assert.Contains(LocalizationService.Localize(language, "Teams 重置与缓存清理已执行完成。", "Teams reset and cache cleanup completed."), result.Message);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TeamsCacheCleanupDoesNotClaimAnUnperformedReset(bool resetFailed)
    {
        using var directory = new TestDirectory();
        File.WriteAllText(Path.Combine(directory.Path, "cache.bin"), "fixture");
        var result = new TeamsMaintenanceResult { ResetError = resetFailed ? "Fixture reset failed" : "" };
        TeamsMaintenanceService.ClearTeamsCache(new TeamsPathTarget("Fixture", directory.Path), result, UiLanguage.English);
        TeamsMaintenanceService.CompleteMaintenanceResult(result, UiLanguage.English);
        Assert.Empty(Directory.GetFileSystemEntries(directory.Path));
        Assert.True(result.CacheCleared);
        Assert.Equal(!resetFailed, result.Success);
        Assert.DoesNotContain("reset and cache cleanup completed", result.Message);
        Assert.Contains(resetFailed ? "did not fully succeed" : "app reset was skipped", result.Message);
    }

    [Fact]
    public void NoTeamsPackageAndNoCacheDoesNotPretendToHaveResetTeams()
    {
        using var directory = new TestDirectory();
        var result = new TeamsMaintenanceResult();
        TeamsMaintenanceService.ClearTeamsCache(new TeamsPathTarget("Fixture", Path.Combine(directory.Path, "absent")), result, UiLanguage.English);
        TeamsMaintenanceService.CompleteMaintenanceResult(result, UiLanguage.English);
        Assert.False(result.Success);
        Assert.Contains("did not complete", result.Message);
    }

    [Fact]
    public void LockedCacheFileRemainsAnActualFailure()
    {
        using var directory = new TestDirectory();
        string path = Path.Combine(directory.Path, "locked.bin");
        File.WriteAllText(path, "fixture");
        using var locked = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        var result = new TeamsMaintenanceResult { ResetSucceeded = true };
        Assert.Throws<IOException>(() => TeamsMaintenanceService.ClearTeamsCache(new TeamsPathTarget("Fixture", directory.Path), result, UiLanguage.English));
        Assert.False(result.CacheCleared);
        Assert.False(result.Success);
    }

    [Theory]
    [InlineData(UiLanguage.Chinese)]
    [InlineData(UiLanguage.English)]
    public void MissingOptionalSignInCachesAreAlreadyClear(UiLanguage language)
    {
        using var directory = new TestDirectory();
        string existing = Path.Combine(directory.Path, "existing");
        Directory.CreateDirectory(existing);
        File.WriteAllText(Path.Combine(existing, "token-fixture.txt"), "fixture");
        var targets = new[] { new TeamsPathTarget("Exists", existing), new TeamsPathTarget("Absent", Path.Combine(directory.Path, "absent")) };
        var result = new TeamsSignInRecordCleanupResult();
        TeamsMaintenanceService.ClearSignInRecordTargets(targets, result, language);
        Assert.True(result.Success);
        Assert.Single(result.ProcessedTargets);
        Assert.Equal(1, result.DeletedFiles);
        Assert.True(Directory.Exists(existing));
        Assert.Empty(Directory.GetFileSystemEntries(existing));
    }

    [Theory]
    [InlineData(UiLanguage.Chinese)]
    [InlineData(UiLanguage.English)]
    public void AlreadyAbsentSignInCachesReportNoCleanupNeeded(UiLanguage language)
    {
        using var directory = new TestDirectory();
        var result = new TeamsSignInRecordCleanupResult();
        TeamsMaintenanceService.ClearSignInRecordTargets(new[] { new TeamsPathTarget("Absent", Path.Combine(directory.Path, "absent")) }, result, language);
        Assert.True(result.Success);
        Assert.Empty(result.ProcessedTargets);
        Assert.Contains(language == UiLanguage.Chinese ? "无需清理" : "no cleanup needed", result.Message);
    }

    [Fact]
    public void LockedSignInCacheIsNotTreatedAsAnAbsentDirectory()
    {
        using var directory = new TestDirectory();
        string path = Path.Combine(directory.Path, "locked.bin");
        File.WriteAllText(path, "fixture");
        using var locked = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        var result = new TeamsSignInRecordCleanupResult();
        Assert.Throws<IOException>(() => TeamsMaintenanceService.ClearSignInRecordTargets(new[] { new TeamsPathTarget("Fixture", directory.Path) }, result, UiLanguage.English));
        Assert.False(result.Success);
        Assert.Empty(result.ProcessedTargets);
    }

    private static bool DetermineSuccess(OutlookDiagnosticResult result, string scenario) =>
        (bool)typeof(OutlookDiagnosticsService).GetMethod("DetermineSuccess", BindingFlags.Static | BindingFlags.NonPublic)!
            .Invoke(null, new object[] { result, scenario })!;

    private sealed class TestDirectory : IDisposable
    {
        public string Path { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "M365Tool-345-" + Guid.NewGuid().ToString("N"));
        public TestDirectory() => Directory.CreateDirectory(Path);
        public void Dispose() => Directory.Delete(Path, true);
    }

    private sealed class DownloadServer : IAsyncDisposable
    {
        public static readonly byte[] Payload = Encoding.UTF8.GetBytes("M365Tool-download-fixture");
        private readonly TcpListener _listener = new(IPAddress.Loopback, 0);
        private readonly CancellationTokenSource _stop = new();
        private readonly Task _response;
        public TaskCompletionSource RequestAccepted { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public TaskCompletionSource HeadersSent { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public string Url { get; }
        public DownloadServer(bool sendHeaders, bool stall)
        {
            _listener.Start();
            Url = $"http://127.0.0.1:{((IPEndPoint)_listener.LocalEndpoint).Port}/fixture.zip";
            _response = RespondAsync(sendHeaders, stall);
        }
        private async Task RespondAsync(bool sendHeaders, bool stall)
        {
            using TcpClient connection = await _listener.AcceptTcpClientAsync(_stop.Token);
            await using var stream = connection.GetStream();
            var buffer = new byte[4096];
            await stream.ReadAsync(buffer, _stop.Token);
            RequestAccepted.TrySetResult();
            if (sendHeaders)
            {
                await stream.WriteAsync(Encoding.ASCII.GetBytes($"HTTP/1.1 200 OK\r\nContent-Length: {Payload.Length}\r\nConnection: close\r\n\r\n"), _stop.Token);
                await stream.WriteAsync(stall ? Payload.AsMemory(0, 1) : Payload.AsMemory(), _stop.Token);
                await stream.FlushAsync(_stop.Token);
                HeadersSent.TrySetResult();
            }
            if (stall) await Task.Delay(Timeout.Infinite, _stop.Token);
        }
        public async ValueTask DisposeAsync()
        {
            _stop.Cancel(); _listener.Stop();
            try { await _response; }
            catch (OperationCanceledException) { }
            catch (IOException) when (_stop.IsCancellationRequested) { }
            catch (SocketException) when (_stop.IsCancellationRequested) { }
            _stop.Dispose();
        }
    }

    private sealed class NeverRun : IScriptRunner
    {
        public Task<ScriptExecutionResult> RunAsync(ScriptExecutionRequest request) => throw new InvalidOperationException("No maintenance commands in tests.");
        public void StartDetached(ScriptExecutionRequest request) => throw new InvalidOperationException("No maintenance commands in tests.");
    }

}
