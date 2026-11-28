using Microsoft.AspNetCore.Components;
using Topo.Model.Approvals;
using Topo.Services;

namespace Topo.Controller
{
    public class PendingApprovalsController : ComponentBase
    {
        [Inject]
        public StorageService _storageService { get; set; } = default!;

        [Inject]
        public IApprovalsService _approvalsService { get; set; } = default!;

        [Inject]
        public NavigationManager NavigationManager { get; set; } = default!;

        public string GroupName = string.Empty;
        public Dictionary<string, string> Units = new();
        public string UnitId = string.Empty;

        // null until the pending approvals have been read
        public PendingApprovalsSession? Session;
        public int NotActionable;
        public bool IsLoading;
        public bool IsBusy;

        // The action the leader has chosen and is being asked to confirm; nothing is sent until they do
        public ApprovalAction? Pending;
        public string Comment = string.Empty;
        public string? ValidationMessage;
        public string? LoadError;

        // What happened to the last item, shown above the next one
        public string? Notice;
        public bool NoticeIsWarning;

        // The submission details of the request on screen: null until read, DetailsFailed when Terrain could not give them
        public SubmissionDetails? Details;
        public bool DetailsLoading;
        public bool DetailsFailed;

        protected override async Task OnInitializedAsync()
        {
            if (!_storageService.IsAuthenticated || _storageService.IsYouthMember)
            {
                NavigationManager.NavigateTo("index");
                return;
            }

            GroupName = _storageService.GroupNameDisplay;
            Units = _storageService.Units;
            UnitId = _storageService.UnitId;
            if (!string.IsNullOrEmpty(UnitId))
            {
                await StartClick();
            }
        }

        internal async Task UnitChange(ChangeEventArgs e)
        {
            UnitId = e.Value?.ToString() ?? "";
            _storageService.UnitId = UnitId;
            Session = null;
            ClearChoice();
            Notice = null;
            LoadError = null;
            if (!string.IsNullOrEmpty(UnitId))
            {
                await StartClick();
            }
        }

        internal async Task StartClick()
        {
            if (string.IsNullOrEmpty(UnitId))
            {
                return;
            }
            IsLoading = true;
            LoadError = null;
            Notice = null;
            ClearChoice();
            var unitId = UnitId;
            try
            {
                var queue = await _approvalsService.GetPendingApprovalQueue(unitId);
                // The leader picked another unit while this was loading; that unit's own read decides what is shown
                if (unitId != UnitId)
                {
                    return;
                }
                Session = new PendingApprovalsSession(queue.Items);
                NotActionable = queue.NotActionable;
            }
            catch (Exception)
            {
                Session = null;
                LoadError = "Could not read the pending approvals from Terrain.";
            }
            finally
            {
                if (unitId == UnitId)
                {
                    IsLoading = false;
                }
            }
            if (unitId == UnitId)
            {
                await LoadDetails();
            }
        }

        // Skip, Back and Stop only move around; they never write anything
        internal async Task SkipClick()
        {
            if (IsBusy)
            {
                return;
            }
            Session?.Skip();
            await AfterMove();
        }

        internal async Task BackClick()
        {
            if (IsBusy)
            {
                return;
            }
            Session?.Back();
            await AfterMove();
        }

        internal async Task StopClick()
        {
            if (IsBusy)
            {
                return;
            }
            Session?.Stop();
            await AfterMove();
        }

        internal async Task RestartClick()
        {
            if (IsBusy)
            {
                return;
            }
            Session?.Restart();
            await AfterMove();
        }

        private async Task AfterMove()
        {
            Notice = null;
            ClearChoice();
            await LoadDetails();
        }

        internal void ApproveClick() => Choose(ApprovalAction.Approve(Comment));

        internal void ImproveClick() => Choose(ApprovalAction.RequestImprovement(Comment));

        // Checks the action and puts up the confirmation; nothing goes to Terrain yet
        private void Choose(ApprovalAction action)
        {
            var item = Session?.Current;
            if (item == null || IsBusy)
            {
                return;
            }
            ValidationMessage = action.Problem(item.SubmissionType);
            Pending = ValidationMessage == null ? action : null;
            Notice = null;
        }

        internal void CancelConfirmClick()
        {
            Pending = null;
        }

        // The leader has confirmed: write to Terrain, and only if that worked check it has left the pending list and move on
        internal async Task ConfirmClick()
        {
            var session = Session;
            var action = Pending;
            var item = session?.Current;
            if (session == null || action == null || item == null || IsBusy)
            {
                return;
            }

            IsBusy = true;
            try
            {
                var worked = await session.CompleteAsync(action, (i, a) => _approvalsService.PerformAction(i.SubmissionId, a));
                if (!worked)
                {
                    // Stay on this item with what was typed still there, so the leader can try again or skip it
                    Pending = null;
                    Notice = $"{session.LastError} Nothing was changed in Terrain.";
                    NoticeIsWarning = true;
                    return;
                }

                ClearChoice();
                Notice = $"{action.PastTense}: {item.MemberName}, {item.AchievementName}.";
                NoticeIsWarning = false;
                await CheckAgainstTerrain(session, item);
            }
            finally
            {
                IsBusy = false;
            }
            await LoadDetails();
        }

        // Reads the pending list again, to drop what others have dealt with and to check Terrain took the write
        private async Task CheckAgainstTerrain(PendingApprovalsSession session, PendingApprovalItem done)
        {
            try
            {
                var pending = await _approvalsService.GetPendingSubmissions(UnitId);
                var stillPending = pending.Select(p => p.SubmissionId).ToList();
                var listed = pending.FirstOrDefault(p => p.SubmissionId == done.SubmissionId);
                if (listed != null && listed.ActionedByLeader)
                {
                    // Some requests (reviews) need a second leader, so they stay pending after the first approval
                    Notice += " Terrain shows your response and the request is waiting for another leader.";
                }
                else if (listed != null)
                {
                    Notice += " Terrain still lists it as pending, so it may take a moment to update.";
                    NoticeIsWarning = true;
                }
                // An empty list while items are still waiting is more likely a failed read than every one having been
                // dealt with elsewhere, so it is not used to drop them; Terrain refuses a write on one that is gone anyway
                var dropped = stillPending.Count > 0 || session.RemainingCount == 0 ? session.Reconcile(stillPending) : 0;
                if (dropped > 0)
                {
                    Notice += $" {dropped} other {(dropped == 1 ? "request was" : "requests were")} dealt with elsewhere and left out.";
                }
            }
            catch (Exception)
            {
                // The write worked; only the check could not be made
                Notice += " Could not re-read Terrain's pending list to check.";
                NoticeIsWarning = true;
            }
        }

        // Reads the submission details of the request now on screen, unless they are already showing
        private string? _detailsFor;

        private async Task LoadDetails()
        {
            var item = Session?.Current;
            if (item == null)
            {
                _detailsFor = null;
                Details = null;
                DetailsFailed = false;
                DetailsLoading = false;
                return;
            }
            if (item.SubmissionId == _detailsFor && (Details != null || DetailsFailed))
            {
                return;
            }

            _detailsFor = item.SubmissionId;
            Details = null;
            DetailsFailed = false;
            DetailsLoading = true;
            StateHasChanged();

            SubmissionDetails? details = null;
            try
            {
                details = await _approvalsService.GetSubmissionDetails(item);
            }
            catch (Exception)
            {
                // shown as "could not be read" below
            }

            // The leader moved on while this was loading; that move reads its own
            if (Session?.Current?.SubmissionId != item.SubmissionId)
            {
                return;
            }
            Details = details;
            DetailsFailed = details == null;
            DetailsLoading = false;
        }

        private void ClearChoice()
        {
            Pending = null;
            Comment = string.Empty;
            ValidationMessage = null;
        }
    }
}
