using System.Collections.Concurrent;

namespace OutlookSearch.Services;

/// <summary>
/// Runs all work on a single dedicated STA thread. Outlook COM objects have
/// thread affinity (they must be created and used on the same STA thread), so
/// every call into Outlook is funneled through here while the UI thread stays free.
/// </summary>
public sealed class StaTaskScheduler : IDisposable
{
    private readonly BlockingCollection<Action> _queue = new();
    private readonly Thread _thread;

    public StaTaskScheduler()
    {
        _thread = new Thread(() =>
        {
            foreach (var action in _queue.GetConsumingEnumerable())
            {
                try { action(); }
                catch { /* per-item exceptions are surfaced through the TaskCompletionSource */ }
            }
        })
        {
            IsBackground = true,
            Name = "Outlook-STA"
        };
        _thread.SetApartmentState(ApartmentState.STA);
        _thread.Start();
    }

    public Task<T> Run<T>(Func<T> func)
    {
        var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        _queue.Add(() =>
        {
            try { tcs.SetResult(func()); }
            catch (Exception ex) { tcs.SetException(ex); }
        });
        return tcs.Task;
    }

    public Task Run(Action action) => Run<object?>(() => { action(); return null; });

    public void Dispose() => _queue.CompleteAdding();
}
