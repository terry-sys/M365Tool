using System.Diagnostics;
using Office365CleanupTool.Services;
using Xunit;

namespace M365Tool.Application.Tests;

[Collection("Shared maintenance operation")]
public class MaintenanceCoordinatorTests
{
    [Fact]
    public void ConcurrentOperationIsRejectedAndNormalCompletionReleasesGate()
    {
        var coordinator = new MaintenanceCoordinator();
        using (coordinator.Begin(UiLanguage.English))
        {
            Assert.True(coordinator.IsBusy);
            Assert.Throws<InvalidOperationException>(() => coordinator.Begin(UiLanguage.English));
        }
        using var next = coordinator.Begin(UiLanguage.English);
        Assert.True(coordinator.IsBusy);
    }

    [Fact]
    public async Task StoppingMonitoringRetainsGateUntilBackgroundCompletion()
    {
        var coordinator = new MaintenanceCoordinator();
        var background = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        using (coordinator.Begin(UiLanguage.English)) coordinator.TrackBackgroundTask(background.Task);
        Assert.Throws<InvalidOperationException>(() => coordinator.Begin(UiLanguage.English));
        background.SetResult();
        await WaitUntilIdle(coordinator);
        using var next = coordinator.Begin(UiLanguage.English);
    }

    [Fact]
    public async Task FailedBackgroundTaskAlsoReleasesGate()
    {
        var coordinator = new MaintenanceCoordinator();
        var background = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        using (coordinator.Begin(UiLanguage.English)) coordinator.TrackBackgroundTask(background.Task);
        background.SetException(new InvalidOperationException("test process has ended"));
        await WaitUntilIdle(coordinator);
    }

    [Fact]
    public async Task RunnerRetainsSharedGateForAnActualDetachedChild()
    {
        ScriptExecutionResult result;
        using (MaintenanceCoordinator.Shared.Begin(UiLanguage.English))
        {
            result = await new ScriptRunner().RunAsync(new ScriptExecutionRequest
            {
                FileName = Path.Combine(Environment.SystemDirectory, "cmd.exe"),
                Arguments = "/d /c ping -n 4 127.0.0.1 >nul",
                Timeout = TimeSpan.FromMilliseconds(200)
            });
            Assert.True(result.IsStillRunning);
        }
        try
        {
            Assert.Throws<InvalidOperationException>(() => MaintenanceCoordinator.Shared.Begin(UiLanguage.English));
            using var process = Process.GetProcessById(result.ProcessId);
            await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(10));
            await WaitUntilIdle(MaintenanceCoordinator.Shared);
        }
        finally
        {
            try
            {
                using var process = Process.GetProcessById(result.ProcessId);
                if (!process.HasExited) { process.Kill(true); await process.WaitForExitAsync(); }
            }
            catch (ArgumentException) { }
        }
    }

    [Fact]
    public async Task ConflictingNetworkRepairCannotReachItsRunner()
    {
        var runner = new NeverRun();
        var repair = new NetworkRepairService(runner, @"Software\M365Tool.Tests\" + Guid.NewGuid().ToString("N"), null);
        using var active = MaintenanceCoordinator.Shared.Begin(UiLanguage.English);
        await Assert.ThrowsAsync<InvalidOperationException>(() => repair.RunAsync(UiLanguage.English));
        Assert.False(runner.Called);
    }

    private static async Task WaitUntilIdle(MaintenanceCoordinator coordinator)
    {
        var timer = Stopwatch.StartNew();
        while (coordinator.IsBusy && timer.Elapsed < TimeSpan.FromSeconds(5)) await Task.Delay(10);
        Assert.False(coordinator.IsBusy);
    }

    private sealed class NeverRun : IScriptRunner
    {
        public bool Called { get; private set; }
        public Task<ScriptExecutionResult> RunAsync(ScriptExecutionRequest request)
        {
            Called = true;
            throw new InvalidOperationException("Maintenance commands are disabled in this test.");
        }
        public void StartDetached(ScriptExecutionRequest request) => throw new NotSupportedException();
    }
}
