using Topo.Model;

namespace TopoReportFunctionTest
{
    [TestClass]
    public class ReadsAtOnceTests
    {
        [TestMethod]
        public async Task TheAnswersComeBackInTheOrderGivenNotTheOrderTheyFinish()
        {
            // the first items take longest, so they finish last
            var results = await ReadsAtOnce.SelectAsync(Enumerable.Range(0, 12), async i =>
            {
                await Task.Delay((12 - i) * 5);
                return i * 10;
            });

            CollectionAssert.AreEqual(Enumerable.Range(0, 12).Select(i => i * 10).ToList(), results);
        }

        [TestMethod]
        public async Task NoMoreThanTheLimitAreReadAtOnceButMoreThanOneAre()
        {
            int running = 0, peak = 0;
            await ReadsAtOnce.SelectAsync(Enumerable.Range(0, 30), async i =>
            {
                var now = Interlocked.Increment(ref running);
                int seen;
                do { seen = peak; } while (now > seen && Interlocked.CompareExchange(ref peak, now, seen) != seen);
                await Task.Delay(20);
                Interlocked.Decrement(ref running);
                return i;
            }, atOnce: 6);

            Assert.AreEqual(6, peak);
        }

        [TestMethod]
        public async Task EveryItemIsReadExactlyOnce()
        {
            var seen = new List<int>();
            await ReadsAtOnce.SelectAsync(Enumerable.Range(0, 25), async i =>
            {
                await Task.Yield();
                lock (seen) seen.Add(i);
                return i;
            });

            CollectionAssert.AreEquivalent(Enumerable.Range(0, 25).ToList(), seen);
        }

        [TestMethod]
        public async Task NothingToReadGivesNothing()
        {
            var results = await ReadsAtOnce.SelectAsync(new List<string>(), s => Task.FromResult(s));

            Assert.AreEqual(0, results.Count);
        }

        [TestMethod]
        public async Task AFailedReadIsThrownToTheCaller()
        {
            await Assert.ThrowsExactlyAsync<InvalidOperationException>(async () =>
                await ReadsAtOnce.SelectAsync(Enumerable.Range(0, 10), async i =>
                {
                    await Task.Delay(5);
                    if (i == 4) throw new InvalidOperationException("boom");
                    return i;
                }));
        }

        [TestMethod]
        public async Task OnceOneReadFailsNoFurtherReadsAreStarted()
        {
            var started = 0;
            await Assert.ThrowsExactlyAsync<InvalidOperationException>(async () =>
                await ReadsAtOnce.SelectAsync(Enumerable.Range(0, 50), async i =>
                {
                    Interlocked.Increment(ref started);
                    await Task.Delay(10);
                    if (i == 0) throw new InvalidOperationException("boom");
                    return i;
                }, atOnce: 2));

            // the two already running finish, and at most a couple more slip in before the failure is seen; the other ~45 never start
            Assert.IsTrue(started <= 6, $"{started} reads were started");
        }

        [TestMethod]
        public async Task ALimitOfOneReadsOneAfterAnother()
        {
            int running = 0, peak = 0;
            await ReadsAtOnce.SelectAsync(Enumerable.Range(0, 8), async i =>
            {
                var now = Interlocked.Increment(ref running);
                peak = Math.Max(peak, now);
                await Task.Delay(5);
                Interlocked.Decrement(ref running);
                return i;
            }, atOnce: 1);

            Assert.AreEqual(1, peak);
        }
    }
}
