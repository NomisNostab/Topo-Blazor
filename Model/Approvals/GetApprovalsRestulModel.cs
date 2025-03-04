﻿namespace Topo.Model.Approvals
{
    public class GetApprovalsResultModel
    {
        public Result[] results { get; set; } = new Result[0];
    }

    public class Result
    {
        public Achievement achievement { get; set; } = new Achievement();
        public Submission submission { get; set; } = new Submission();
        public Unit_Permissions unit_permissions { get; set; } = new Unit_Permissions();
        public Member member { get; set; } = new Member();
        public Submitted_By submitted_by { get; set; } = new Submitted_By();
    }

    public class Achievement
    {
        public string id { get; set; } = string.Empty;
        public string type { get; set; } = string.Empty;
        public Achievement_Meta achievement_meta { get; set; } = new Achievement_Meta();
    }

    public class Achievement_Meta
    {
        public string branch { get; set; } = string.Empty;
        public int stage { get; set; }
        public string stream { get; set; } = string.Empty;
        public string sia_area {  get; set; } = string.Empty;
    }

    public class Submission
    {
        public string id { get; set; } = string.Empty;
        public string type { get; set; } = string.Empty;
        public DateTime date { get; set; }
        public string status { get; set; } = string.Empty;
        public string outcome { get; set; } = string.Empty;
        public Actioned_By[] actioned_by { get; set; } = new Actioned_By[0];
    }

    public class Actioned_By
    {
        public string member_id { get; set; } = string.Empty;
        public string member_first_name { get; set; } = string.Empty;
        public string member_last_name { get; set; } = string.Empty;
        public string outcome { get; set; } = string.Empty;
        public DateTime time{ get; set; }
        public string date_awarded { get; set; }
    }

    public class Unit_Permissions
    {
        public bool write { get; set; }
    }

    public class Member
    {
        public string id { get; set; } = string.Empty;
        public string first_name { get; set; } = string.Empty;
        public string last_name { get; set; } = string.Empty;
    }

    public class Submitted_By
    {
        public string id { get; set; } = string.Empty;
        public string first_name { get; set; } = string.Empty;
        public string last_name { get; set; } = string.Empty;
    }

}
