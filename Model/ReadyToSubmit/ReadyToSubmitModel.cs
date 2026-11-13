namespace Topo.Model.ReadyToSubmit
{
    // The kinds of badge the report looks at
    public enum ReadyCategory
    {
        Milestone,
        OutdoorAdventureSkill,
        SpecialInterestArea,
        PeakAward
    }

    // One piece of a member's record from Terrain, with its answers flattened to text
    public class AchievementRecord
    {
        public string Id { get; set; } = string.Empty;
        // Whose record it is; a read of the whole unit returns every member's records together
        public string MemberId { get; set; } = string.Empty;
        // Terrain's type: milestone, outdoor_adventure_skill, special_interest_area, adventurous_journey, course_reflection, personal_reflection
        public string Type { get; set; } = string.Empty;
        public string Section { get; set; } = string.Empty;
        public string Template { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty;
        public DateTime StatusUpdated { get; set; }
        // Milestones: Terrain's own "complete" / "incomplete"
        public string RequirementStatus { get; set; } = string.Empty;
        public Dictionary<string, string> Answers { get; set; } = new Dictionary<string, string>();
        // achievement_meta: the stream and stage of an OAS, the stage of a milestone
        public Dictionary<string, string> Meta { get; set; } = new Dictionary<string, string>();
    }

    // Everything the report needs to know about one member
    public class ReadyMemberInput
    {
        public string MemberId { get; set; } = string.Empty;
        // The full name, used when FirstName and LastName are not given
        public string MemberName { get; set; } = string.Empty;
        public string FirstName { get; set; } = string.Empty;
        public string LastName { get; set; } = string.Empty;
        public string Patrol { get; set; } = string.Empty;
        public List<AchievementRecord> Records { get; set; } = new List<AchievementRecord>();
    }

    // A question of a Terrain form (a plan, a review or a reflection) and how it is answered
    public class FormQuestion
    {
        public string Id { get; set; } = string.Empty;
        public string Type { get; set; } = string.Empty;
        public string Label { get; set; } = string.Empty;
    }

    // What each template asks, by the template's path (for a form), or the ids of what must be ticked off (for an OAS stage)
    public class ReadyTemplates
    {
        public Dictionary<string, List<FormQuestion>> Forms { get; set; } = new Dictionary<string, List<FormQuestion>>(StringComparer.OrdinalIgnoreCase);
        // The things to tick off in an OAS stage, with the wording shown to the Scout
        public Dictionary<string, List<FormQuestion>> OasInputs { get; set; } = new Dictionary<string, List<FormQuestion>>(StringComparer.OrdinalIgnoreCase);
    }

    // Which list an item is on
    public enum ReadyKind
    {
        // Everything is done
        Ready,
        // Everything is done but the Scout's own logbook tick
        LogbookTickOnly,
        // An OAS stage a few ticks short
        NearlyThere
    }

    public class ReadyItem
    {
        public ReadyKind Kind { get; set; }
        public string MemberId { get; set; } = string.Empty;
        public string MemberName { get; set; } = string.Empty;
        public string Patrol { get; set; } = string.Empty;
        public ReadyCategory Category { get; set; }
        public string Badge { get; set; } = string.Empty;
        // Where the badge comes in the wall chart, then lowest level to highest (see ReadyToSubmitBuilder.BadgeOrderOf)
        public int BadgeOrder { get; set; }
        // What is waiting to be sent: "Requirements complete", "Plan ready", "Review ready" and so on
        public string Waiting { get; set; } = string.Empty;
        // For an OAS stage that is not finished: what is still to be ticked off
        public List<string> Missing { get; set; } = new List<string>();
        // Where each missing item comes in the stage's template, to keep that order among items missing for as many Scouts
        public List<int> MissingPositions { get; set; } = new List<int>();
        public int TicksShort => Missing.Count;
        // Where the Scout comes in the unit's listing, to keep that order within a badge
        public int MemberOrder { get; set; }
        public DateTime LastUpdated { get; set; }
    }

    // One requirement of a badge and how many of its Scouts still have to tick it off
    public class MissingRequirement
    {
        public string Label { get; set; } = string.Empty;
        public int Scouts { get; set; }
    }

    // The Scouts a few ticks short of one badge (an OAS stream's stage), and the requirements any of them are missing
    public class NearBadge
    {
        public string Badge { get; set; } = string.Empty;
        // In the order the unit is listed
        public List<ReadyItem> Scouts { get; set; } = new List<ReadyItem>();
        // The commonest first
        public List<MissingRequirement> Requirements { get; set; } = new List<MissingRequirement>();

        // Whether this Scout still has to tick this requirement off; for the rest it is blocked out
        public bool IsMissing(ReadyItem scout, MissingRequirement requirement) => scout.Missing.Contains(requirement.Label);
    }

    public class ReadyToSubmitReport
    {
        public string UnitName { get; set; } = string.Empty;
        public string Section { get; set; } = string.Empty;
        public DateTime GeneratedAt { get; set; }
        public List<ReadyCategory> Categories { get; set; } = new List<ReadyCategory>();
        public int MembersChecked { get; set; }
        public List<ReadyItem> Items { get; set; } = new List<ReadyItem>();
        // Stages up to this many ticks short are included (0 for none)
        public int MaxTicksShort { get; set; }
        public bool LogbookTickOptional { get; set; }

        public List<ReadyItem> Ready => Items.Where(i => i.Kind == ReadyKind.Ready).ToList();
        public List<ReadyItem> LogbookOnly => Items.Where(i => i.Kind == ReadyKind.LogbookTickOnly).ToList();
        public List<ReadyItem> NearlyThere => Items.Where(i => i.Kind == ReadyKind.NearlyThere).ToList();

        // The nearly there stages split by badge in the order of the wall chart, lowest level to highest, each with the requirements
        // missing for any of them, the one most Scouts are missing first
        public List<NearBadge> NearlyThereByBadge() =>
            NearlyThere
                .GroupBy(i => i.Badge)
                .Select(g =>
                {
                    var scouts = g.OrderBy(i => i.MemberOrder).ToList();
                    // each requirement's place in the template
                    var position = new Dictionary<string, int>();
                    foreach (var scout in scouts)
                        for (var i = 0; i < scout.Missing.Count; i++)
                            position[scout.Missing[i]] = i < scout.MissingPositions.Count ? scout.MissingPositions[i] : int.MaxValue;
                    return new NearBadge
                    {
                        Badge = g.Key,
                        Scouts = scouts,
                        Requirements = position.Keys
                            .Select(label => new MissingRequirement { Label = label, Scouts = scouts.Count(s => s.Missing.Contains(label)) })
                            .OrderByDescending(r => r.Scouts)
                            .ThenBy(r => position[r.Label])
                            .ToList()
                    };
                })
                .OrderBy(b => b.Scouts[0].BadgeOrder)
                .ThenBy(b => b.Badge, StringComparer.OrdinalIgnoreCase)
                .ToList();
    }
}
