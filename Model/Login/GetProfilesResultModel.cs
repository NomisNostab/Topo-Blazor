namespace Topo.Model.Login
{
    public class GetProfilesResultModel
    {
        public string? username { get; set; }
        public Profile[]? profiles { get; set; }
    }

    public class Profile
    {
        public Member? member { get; set; }
        public Unit? unit { get; set; }
        public Group? group { get; set; }
        public Branch? branch { get; set; }
        public bool is_support_group { get; set; } = false;
    }

    public class Member
    {
        public string? id { get; set; }
        public string? name { get; set; }
        public string[]? roles { get; set; }
    }

    public class Unit
    {
        public string? id { get; set; }
        public string? name { get; set; }
        public string[]? roles { get; set; }
        public string? section { get; set; }
    }

    public class Group
    {
        public string? id { get; set; }
        public string? name { get; set; }
        public string[]? roles { get; set; }
    }

    public class Branch
    {
        public string? id { get; set; }
        public string? name { get; set; }
        public string[]? roles { get; set; }
    }

}
