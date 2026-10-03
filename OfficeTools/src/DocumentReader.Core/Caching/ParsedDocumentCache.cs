namespace DocumentReader.Core.Caching;

/// <summary>
/// 解析結果的記憶體快取（LRU）。鍵包含路徑、最後修改時間與檔案大小，檔案一被修改就自然失效；
/// 同一個鍵同時被多個執行緒要求時只會解析一次。
/// </summary>
public sealed class ParsedDocumentCache(int maxEntries)
{
    private readonly object _gate = new();
    private readonly Dictionary<Key, LinkedListNode<Entry>> _entries = [];
    private readonly LinkedList<Entry> _recent = new(); // 最近使用的在前面

    public readonly record struct Key(string Path, long LastWriteTicks, long Length, string Variant);

    private sealed record Entry(Key Key, Lazy<object> Value);

    public int Count
    {
        get
        {
            lock (_gate)
            {
                return _entries.Count;
            }
        }
    }

    public void Clear()
    {
        lock (_gate)
        {
            _entries.Clear();
            _recent.Clear();
        }
    }

    public T GetOrAdd<T>(Key key, Func<T> factory)
        where T : class
    {
        ArgumentNullException.ThrowIfNull(factory);

        Entry entry;
        lock (_gate)
        {
            if (_entries.TryGetValue(key, out var node))
            {
                _recent.Remove(node);
                _recent.AddFirst(node);
                entry = node.Value;
            }
            else
            {
                entry = new Entry(key, new Lazy<object>(() => factory(), LazyThreadSafetyMode.ExecutionAndPublication));
                _entries[key] = _recent.AddFirst(entry);
                while (_entries.Count > Math.Max(1, maxEntries))
                {
                    var oldest = _recent.Last!;
                    _recent.RemoveLast();
                    _entries.Remove(oldest.Value.Key);
                }
            }
        }

        try
        {
            return (T)entry.Value.Value;
        }
        catch
        {
            // 解析失敗不要留在快取裡（Lazy 會把例外記住），下次重新嘗試
            lock (_gate)
            {
                if (_entries.TryGetValue(key, out var node) && ReferenceEquals(node.Value, entry))
                {
                    _recent.Remove(node);
                    _entries.Remove(key);
                }
            }

            throw;
        }
    }
}
