using System.Globalization;

namespace Topo.Model.Approvals
{
    // How the approvals pages name an achievement from what Terrain says about its submission
    public static class ApprovalNames
    {
        public static string AchievementName(string achievementType, Achievement_Meta? meta)
        {
            string name;
            switch (achievementType)
            {
                case "special_interest_area":
                    var siaArea = string.IsNullOrEmpty(meta?.sia_area) ? "special_interest_area" : meta.sia_area;
                    name = siaArea.Replace("_", " ");
                    break;
                case "milestone":
                    name = $"milestone {meta?.stage}";
                    break;
                case "outdoor_adventure_skill":
                    name = meta?.stream == meta?.branch
                        ? $"OAS {meta?.stream} {meta?.stage}"
                        : $"OAS {meta?.stream} {meta?.branch} {meta?.stage}";
                    break;
                case "intro_section":
                    name = "intro to section";
                    break;
                case "intro_scouting":
                    name = "intro to scouting";
                    break;
                default:
                    name = achievementType.Replace("_", " ");
                    break;
            }
            return new CultureInfo("en-AU", false).TextInfo.ToTitleCase(name);
        }
    }
}
