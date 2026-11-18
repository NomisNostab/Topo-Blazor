namespace Topo.Model.PatrolPlanner
{
    public class PatrolPlannerPageViewModel
    {
        public string UnitId { get; set; } = string.Empty;
        public string UnitName { get; set; } = string.Empty;
        public Dictionary<string, string> Units { get; set; } = new Dictionary<string, string>();
        public string GroupName { get; set; } = string.Empty;
        public List<PatrolPlannerMember> Members { get; set; } = new List<PatrolPlannerMember>();
        public List<string> PatrolNames { get; set; } = new List<string>();
        public string NewPatrolName { get; set; } = string.Empty;
    }

    public class PatrolPlannerMember
    {
        public string Id { get; set; } = string.Empty;
        public string FirstName { get; set; } = string.Empty;
        public string LastName { get; set; } = string.Empty;
        public string Age { get; set; } = string.Empty;
        public bool UnitCouncil { get; set; }
        public bool OriginalUnitCouncil { get; set; }
        public string Gender { get; set; } = string.Empty;
        public string OriginalPatrolName { get; set; } = string.Empty;
        public string OriginalPatrolDuty { get; set; } = string.Empty;
        public string PatrolName { get; set; } = string.Empty;
        public string PatrolDuty { get; set; } = string.Empty;

        // Drives both the "changed" card highlight on the board and what shows up in the
        // "Changes to make in Terrain" list/export - anything not reflected here is invisible to the leader.
        public bool IsChanged => PatrolName != OriginalPatrolName || PatrolDuty != OriginalPatrolDuty || UnitCouncil != OriginalUnitCouncil;
    }
}
