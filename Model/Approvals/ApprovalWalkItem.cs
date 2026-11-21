namespace Topo.Model.Approvals
{
    // What the walk-through shows for one pending submission
    public class ApprovalWalkItem
    {
        public string SubmissionId { get; init; } = string.Empty;
        public string SubmissionType { get; init; } = string.Empty;
        public DateTime SubmissionDate { get; init; }
        public string SubmissionTypeLabel => SubmissionType.Length == 0 ? string.Empty : char.ToUpper(SubmissionType[0]) + SubmissionType.Substring(1);
        public string MemberId { get; init; } = string.Empty;
        public string MemberName { get; init; } = string.Empty;
        public string AchievementId { get; init; } = string.Empty;
        public string AchievementType { get; init; } = string.Empty;
        public string AchievementName { get; init; } = string.Empty;
        public string SubmittedBy { get; init; } = string.Empty;
        public List<ApprovalWalkPreviousAction> PreviousActions { get; init; } = new();
    }

    public class ApprovalWalkPreviousAction
    {
        public string By { get; init; } = string.Empty;
        public string Outcome { get; init; } = string.Empty;
        public DateTime Time { get; init; }
        public string Comment { get; init; } = string.Empty;
    }

    // The pending submissions the leader can approve or send back, oldest first, and how many were left out
    public class ApprovalWalkQueue
    {
        public List<ApprovalWalkItem> Items { get; init; } = new();
        public int NotActionable { get; init; }

        // Terrain's Approvals screen disables Approve and Improve on a submission that is the leader's own or that
        // the leader has already actioned, so those are left out of the walk and counted instead.
        public static ApprovalWalkQueue FromPending(IEnumerable<Result> pending, string? currentMemberId)
        {
            var items = new List<ApprovalWalkItem>();
            var notActionable = 0;
            foreach (var result in pending)
            {
                // Award requests are not part of the walk-through and are not mentioned
                if (string.Equals(result.submission.type, "award", StringComparison.OrdinalIgnoreCase))
                    continue;

                var isOwn = !string.IsNullOrEmpty(currentMemberId) && result.member.id == currentMemberId;
                var alreadyActioned = !string.IsNullOrEmpty(currentMemberId) && result.submission.actioned_by.Any(a => a.member_id == currentMemberId);
                // With no idea who the leader is nothing can be checked, so nothing is offered
                if (string.IsNullOrEmpty(currentMemberId) || isOwn || alreadyActioned)
                {
                    notActionable++;
                    continue;
                }
                items.Add(FromResult(result));
            }

            return new ApprovalWalkQueue
            {
                Items = items.OrderBy(i => i.SubmissionDate).ThenBy(i => i.MemberName, StringComparer.CurrentCultureIgnoreCase).ThenBy(i => i.SubmissionId, StringComparer.Ordinal).ToList(),
                NotActionable = notActionable
            };
        }

        private static ApprovalWalkItem FromResult(Result result)
        {
            var submittedBy = result.submitted_by == null || string.IsNullOrEmpty(result.submitted_by.id)
                ? string.Empty
                : $"{result.submitted_by.first_name} {result.submitted_by.last_name}".Trim();
            return new ApprovalWalkItem
            {
                SubmissionId = result.submission.id,
                SubmissionType = result.submission.type,
                SubmissionDate = result.submission.date,
                MemberId = result.member.id,
                MemberName = $"{result.member.first_name} {result.member.last_name}".Trim(),
                AchievementId = result.achievement.id,
                AchievementType = result.achievement.type,
                AchievementName = ApprovalNames.AchievementName(result.achievement.type, result.achievement.achievement_meta),
                SubmittedBy = submittedBy,
                PreviousActions = result.submission.actioned_by
                    .OrderBy(a => a.time)
                    .Select(a => new ApprovalWalkPreviousAction
                    {
                        By = $"{a.member_first_name} {a.member_last_name}".Trim(),
                        Outcome = a.outcome,
                        Time = a.time,
                        Comment = a.comment ?? string.Empty
                    })
                    .ToList()
            };
        }
    }

    // A submission still on Terrain's pending list, and whether the signed in leader has already actioned it
    public record PendingSubmission(string SubmissionId, bool ActionedByLeader);
}
