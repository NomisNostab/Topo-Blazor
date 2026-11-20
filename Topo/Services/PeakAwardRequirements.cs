namespace Topo.Services
{
    // Requirements for each section's own top ("Peak") award. Every section has a different award
    // with its own name and its own numeric thresholds, and Scouts Australia has changed these
    // numbers more than once over the years - when they change again, this is the only file that
    // should need editing. Terrain doesn't expose these requirements anywhere, so they're hardcoded
    // here from the group's own records.
    public static class PeakAwardRequirements
    {
        public class SectionRequirement
        {
            public string AwardName { get; set; } = string.Empty;
            public int OASProgressionsRequired { get; set; }
            // Minimum stage required in EACH of the three "core" OAS streams (Bushcraft, Bushwalking, Camping).
            // Which three streams count as "core" is hardcoded in WallchartService.ComputeOASBucket, not
            // configurable here - if that ever changes, that method needs editing too, not just this number.
            public int OASCoreMinLevel { get; set; }
            // How many OAS streams (any of the nine) need to reach stage 4 or higher. 0 = no such rule for this section.
            public int OASLevel4PlusRequired { get; set; }
            public int SIARequired { get; set; }
            public int SIAMinAreasRequired { get; set; }
        }

        // Keyed by Terrain's unit.section value (lowercase). If a section isn't listed here, the
        // OAS and SIA buckets are simply left out of that member's percentage rather than guessing.
        public static readonly Dictionary<string, SectionRequirement> BySection = new(StringComparer.OrdinalIgnoreCase)
        {
            { "joey", new SectionRequirement
                { AwardName = "Joey Scout Challenge Award", OASProgressionsRequired = 3, OASCoreMinLevel = 1, OASLevel4PlusRequired = 0, SIARequired = 6, SIAMinAreasRequired = 2 } },
            { "cub", new SectionRequirement
                { AwardName = "Grey Wolf Award", OASProgressionsRequired = 8, OASCoreMinLevel = 3, OASLevel4PlusRequired = 0, SIARequired = 6, SIAMinAreasRequired = 2 } },
            { "scout", new SectionRequirement
                { AwardName = "Australian Scout Award", OASProgressionsRequired = 10, OASCoreMinLevel = 5, OASLevel4PlusRequired = 0, SIARequired = 6, SIAMinAreasRequired = 3 } },
            { "venturer", new SectionRequirement
                { AwardName = "King Scout Award", OASProgressionsRequired = 12, OASCoreMinLevel = 5, OASLevel4PlusRequired = 4, SIARequired = 6, SIAMinAreasRequired = 3 } },
            { "rover", new SectionRequirement
                { AwardName = "Baden-Powell Scout Award", OASProgressionsRequired = 14, OASCoreMinLevel = 5, OASLevel4PlusRequired = 6, SIARequired = 6, SIAMinAreasRequired = 4 } },
        };
    }
}
