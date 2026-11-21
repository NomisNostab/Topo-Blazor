namespace Topo.Model.Approvals
{
    // The state of one pass through the pending approvals: where the leader is, what they have done, and what is left.
    // It holds no Terrain calls. The page hands CompleteAsync the call that does the write, and the session only moves on
    // when that call says it worked.
    //
    // The queue keeps its order. Skipping moves on and leaves the item pending, Back returns to the nearest earlier item
    // that is still pending, and an item that has been actioned is gone for good. Past the last item the session reaches its
    // end: if items were skipped they are still waiting (Restart goes back to the first of them), otherwise it is all done.
    public class PendingApprovalsSession
    {
        private enum ItemState { Pending, Done, Gone }

        private readonly List<PendingApprovalItem> _items;
        private readonly List<ItemState> _states;
        private readonly List<(PendingApprovalItem Item, ApprovalAction Action)> _completed = new();
        private int _index;

        public PendingApprovalsSession(IEnumerable<PendingApprovalItem> items)
        {
            _items = items.ToList();
            _states = _items.Select(_ => ItemState.Pending).ToList();
        }

        // The item being shown, or null once the session has reached its end or been stopped
        public PendingApprovalItem? Current => !IsStopped && _index < _items.Count ? _items[_index] : null;

        public bool IsStopped { get; private set; }

        // Past the last item, whether or not some were skipped on the way
        public bool IsAtEnd => !IsStopped && _index >= _items.Count;

        public bool IsAllDone => RemainingCount == 0;

        // Items still waiting for an answer: the one being shown and anything skipped or not yet reached
        public int RemainingCount => _states.Count(s => s == ItemState.Pending);

        // For "3 of 12". Items another leader dealt with while the page was open are not counted in either number.
        public int Total => _states.Count(s => s != ItemState.Gone);

        public int Position
        {
            get
            {
                var upTo = Math.Min(_index + 1, _items.Count);
                return _states.Take(upTo).Count(s => s != ItemState.Gone);
            }
        }

        public bool CanGoBack => !IsStopped && PreviousPendingIndex() >= 0;

        // Skipped and still waiting: pending items before the one shown. At the end, every pending item.
        public int SkippedCount => _states.Take(Math.Min(_index, _items.Count)).Count(s => s == ItemState.Pending);

        public IReadOnlyList<(PendingApprovalItem Item, ApprovalAction Action)> Completed => _completed;

        public IEnumerable<PendingApprovalItem> Remaining => _items.Where((_, i) => _states[i] == ItemState.Pending);

        // Why the last attempt did not go through; cleared by the next action that does
        public string? LastError { get; private set; }

        public void Skip()
        {
            if (Current == null) return;
            _index = NextPendingIndex(_index + 1);
        }

        public void Back()
        {
            if (IsStopped) return;
            var previous = PreviousPendingIndex();
            if (previous >= 0) _index = previous;
        }

        public void Stop()
        {
            IsStopped = true;
        }

        // From the end with items skipped, go round again from the first of them
        public void Restart()
        {
            if (IsStopped) return;
            _index = NextPendingIndex(0);
        }

        // Does the action on the current item. When the call says it worked the item is done and the session moves to the
        // next one; when it says it did not, or throws, nothing changes and LastError says why.
        // Returns whether it worked.
        public async Task<bool> CompleteAsync(ApprovalAction action, Func<PendingApprovalItem, ApprovalAction, Task<string?>> perform)
        {
            var item = Current;
            if (item == null) return false;

            string? error;
            try
            {
                error = await perform(item, action);
            }
            catch (Exception ex)
            {
                error = ex.Message;
            }

            if (error != null)
            {
                LastError = error;
                return false;
            }

            LastError = null;
            _states[_index] = ItemState.Done;
            _completed.Add((item, action));
            _index = NextPendingIndex(_index + 1);
            return true;
        }

        // After re-reading Terrain's pending list: items that are no longer on it were dealt with elsewhere (another leader,
        // or by hand in Terrain). They are dropped from the session. Returns how many were dropped.
        public int Reconcile(IEnumerable<string> stillPendingSubmissionIds)
        {
            var pending = new HashSet<string>(stillPendingSubmissionIds);
            var dropped = 0;
            var currentId = Current?.SubmissionId;
            for (var i = 0; i < _items.Count; i++)
            {
                if (_states[i] == ItemState.Pending && !pending.Contains(_items[i].SubmissionId))
                {
                    _states[i] = ItemState.Gone;
                    dropped++;
                }
            }

            // If the item on screen was one of those, move to the next that is left
            if (currentId != null && _index < _items.Count && _states[_index] != ItemState.Pending)
                _index = NextPendingIndex(_index + 1);
            return dropped;
        }

        // The first pending index at or after "from", or past the end when there is none
        private int NextPendingIndex(int from)
        {
            for (var i = Math.Max(from, 0); i < _items.Count; i++)
                if (_states[i] == ItemState.Pending) return i;
            return _items.Count;
        }

        private int PreviousPendingIndex()
        {
            for (var i = Math.Min(_index, _items.Count) - 1; i >= 0; i--)
                if (_states[i] == ItemState.Pending) return i;
            return -1;
        }
    }
}
