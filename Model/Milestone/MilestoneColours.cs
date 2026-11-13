namespace Topo.Model.Milestone
{
    // The red-to-green shading of the milestone counts, as hex colours so the page and the exported report share it
    public static class MilestoneColours
    {
        public const string White = "#ffffff";

        // Each scale runs from nothing done (red) to the target met (green), one colour per count
        public static readonly string[] Scale1 = { "#ec5252", "#9cf250" };
        public static readonly string[] Scale2 = { "#ec5252", "#fdee00", "#9cf250" };
        public static readonly string[] Scale4 = { "#ec5252", "#f49d2b", "#fdee00", "#cef026", "#9cf250" };
        public static readonly string[] Scale5 = { "#ec5252", "#f18339", "#f8bd1a", "#ddef1a", "#b8f139", "#9cf250" };
        public static readonly string[] Scale6 = { "#ec5252", "#f18139", "#f8c217", "#fdee00", "#daf01d", "#b7f13a", "#9cf250" };

        // The shade of a count for a milestone (1 to 3): participate counts and lead counts have different targets at each
        // level. A skipped milestone has a negative count and is left white.
        public static string For(int level, bool participate, int count)
        {
            if (count < 0)
                return White;

            var scale = level switch
            {
                1 => participate ? Scale6 : Scale2,
                2 => participate ? Scale5 : Scale4,
                3 => participate ? Scale4 : Scale6,
                _ => null
            };
            return scale == null ? White : scale[Math.Min(count, scale.Length - 1)];
        }
    }
}
