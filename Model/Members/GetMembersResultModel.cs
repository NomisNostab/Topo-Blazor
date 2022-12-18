namespace Topo.Model.Members
{
    public class GetMembersResultModel
    {
        public Member[] results { get; set; } = new Member[0];
        public int total { get; set; }
        public object limit { get; set; }
        public object offset { get; set; }
    }

    public class Member
    {
        public string id { get; set; } = "";
        public string member_number { get; set; } = "";
        public string first_name { get; set; } = "";
        public string last_name { get; set; } = "";
        public string status { get; set; } = "";
        public string date_of_birth { get; set; } = "";
        public Group[] groups { get; set; } = new Group[0];
        public Unit unit { get; set; } = new Unit();
        public Patrol patrol { get; set; } = new Patrol();
        public Metadata metadata { get; set; } = new Metadata();
    }

    public class Unit
    {
        public string id { get; set; } = "";
        public string section { get; set; } = "";
        public string duty { get; set; } = "";
        public bool unit_council { get; set; }
        public string group_id { get; set; } = "";
    }

    public class Patrol
    {
        public string id { get; set; } = "";
        public string name { get; set; } = "";
        public string duty { get; set; } = "";
    }

    public class Metadata
    {
        public DateTime achievementimport { get; set; }
    }

    public class Group
    {
        public string id { get; set; } = "";
        public string name { get; set; } = "";
    }
}
