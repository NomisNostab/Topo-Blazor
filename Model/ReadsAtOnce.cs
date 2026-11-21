namespace Topo.Model
{
    // Terrain answers a request in a few hundred milliseconds, so a report that reads one thing per item, one after
    // another, spends nearly all its time waiting. This reads a few at a time instead. The answers come back in the order
    // the items were given, whatever order they finish in, so the report comes out the same as before, only sooner.
    public static class ReadsAtOnce
    {
        // Enough to cut the waiting a lot without sending Terrain a burst of requests
        public const int Default = 6;

        public static async Task<List<TResult>> SelectAsync<T, TResult>(IEnumerable<T> items, Func<T, Task<TResult>> read, int atOnce = Default)
        {
            var list = items.ToList();
            var results = new TResult[list.Count];
            using var gate = new SemaphoreSlim(Math.Max(1, atOnce));
            var failed = new int[1];
            await Task.WhenAll(list.Select(async (item, i) =>
            {
                await gate.WaitAsync();
                try
                {
                    // Once a read has failed the caller is going to get that failure, so don't send Terrain any more requests,
                    // as reading one after another would have stopped at the first failure too.
                    if (Volatile.Read(ref failed[0]) == 1)
                        return;
                    try
                    {
                        results[i] = await read(item);
                    }
                    catch
                    {
                        Volatile.Write(ref failed[0], 1);
                        throw;
                    }
                }
                finally
                {
                    gate.Release();
                }
            }));
            return results.ToList();
        }
    }
}
