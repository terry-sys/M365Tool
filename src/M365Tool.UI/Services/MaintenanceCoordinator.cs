namespace Office365CleanupTool.Services;

internal sealed class MaintenanceCoordinator
{
    public static MaintenanceCoordinator Shared { get; } = new();

    private readonly object _sync = new();
    private readonly AsyncLocal<Operation?> _current = new();
    private Operation? _active;

    internal bool IsBusy { get { lock (_sync) return _active != null; } }

    public Operation Begin(UiLanguage language)
    {
        lock (_sync)
        {
            if (_active != null)
                throw new InvalidOperationException(LocalizationService.Localize(language,
                    "当前仍有维护任务在运行（可能已停止等待），请等待该任务结束后重试。",
                    "A maintenance task is still running (monitoring may have stopped). Wait for it to finish before retrying."));
            var operation = new Operation(this);
            _active = operation;
            _current.Value = operation;
            return operation;
        }
    }

    public void TrackBackgroundTask(Task completion) => _current.Value?.Track(completion);

    private void Release(Operation operation)
    {
        lock (_sync)
        {
            if (ReferenceEquals(_active, operation)) _active = null;
        }
    }

    internal sealed class Operation(MaintenanceCoordinator owner) : IDisposable
    {
        private readonly List<Task> _backgroundTasks = new();
        private bool _disposed;

        internal void Track(Task completion) => _backgroundTasks.Add(completion);

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            if (ReferenceEquals(owner._current.Value, this)) owner._current.Value = null;
            if (_backgroundTasks.All(task => task.IsCompleted)) owner.Release(this);
            else _ = ReleaseAfterCompletionAsync();
        }

        private async Task ReleaseAfterCompletionAsync()
        {
            try { await Task.WhenAll(_backgroundTasks).ConfigureAwait(false); }
            catch { /* The process runner owns reporting execution failures. */ }
            finally { owner.Release(this); }
        }
    }
}
