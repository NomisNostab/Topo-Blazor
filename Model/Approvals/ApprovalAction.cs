namespace Topo.Model.Approvals
{
    public enum ApprovalActionKind
    {
        Approve,
        RequestImprovement
    }

    // One thing the leader can do to a pending submission, in the form Terrain's assessments call takes.
    // Terrain's own Approvals screen offers the same two on review and approval submissions: Approve and
    // "Request improvement", each with an optional comment of up to 500 characters. Award submissions are not handled here.
    public class ApprovalAction
    {
        public const int MaxCommentLength = 500;

        public ApprovalActionKind Kind { get; }
        public string Comment { get; }

        private ApprovalAction(ApprovalActionKind kind, string comment)
        {
            Kind = kind;
            Comment = comment;
        }

        public static ApprovalAction Approve(string? comment) => new(ApprovalActionKind.Approve, (comment ?? "").Trim());

        public static ApprovalAction RequestImprovement(string? comment) => new(ApprovalActionKind.RequestImprovement, (comment ?? "").Trim());

        // The value of "outcome" in the request body
        public string Outcome => Kind == ApprovalActionKind.Approve ? "approved" : "rejected";

        // What the leader sees on the button and in the confirmation
        public string Label => Kind == ApprovalActionKind.Approve ? "Approve" : "Request improvement";

        // The same words Terrain shows after the action
        public string PastTense => Kind == ApprovalActionKind.Approve ? "Approved" : "Improvement requested";

        // null when the action can be sent, otherwise what is wrong with it
        public string? Problem(string submissionType)
        {
            if (string.Equals(submissionType, "award", StringComparison.OrdinalIgnoreCase))
            {
                return "Award requests cannot be approved or sent back for improvement.";
            }
            if (Comment.Length > MaxCommentLength)
            {
                return $"The comment can be at most {MaxCommentLength} characters.";
            }
            return null;
        }

        // The JSON body Terrain's own page sends
        public Dictionary<string, string> ToRequestBody() => new() { ["outcome"] = Outcome, ["comment"] = Comment };
    }
}
