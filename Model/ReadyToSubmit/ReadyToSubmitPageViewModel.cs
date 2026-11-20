namespace Topo.Model.ReadyToSubmit
{
    public class ReadyToSubmitPageViewModel
    {
        public string GroupName { get; set; } = string.Empty;
        public Dictionary<string, string> Units { get; set; } = new Dictionary<string, string>();
        public string UnitId { get; set; } = string.Empty;
        public bool Milestones { get; set; } = true;
        public bool OutdoorAdventureSkills { get; set; } = true;
        public bool SpecialInterestAreas { get; set; } = true;
        public bool PeakAwardParts { get; set; } = true;
        // Count an OAS stage that is only waiting on the Scout's logbook tick
        public bool LogbookTickOptional { get; set; }
        // Also list OAS stages up to this many ticks short (0 for none)
        public int MaxTicksShort { get; set; }
        // Otherwise Scouts are named by first name, with enough of the surname to tell them apart
        public bool UseFullNames { get; set; }
        public string ErrorMessage { get; set; } = string.Empty;

        public List<ReadyCategory> SelectedCategories()
        {
            var categories = new List<ReadyCategory>();
            if (Milestones) categories.Add(ReadyCategory.Milestone);
            if (OutdoorAdventureSkills) categories.Add(ReadyCategory.OutdoorAdventureSkill);
            if (SpecialInterestAreas) categories.Add(ReadyCategory.SpecialInterestArea);
            if (PeakAwardParts) categories.Add(ReadyCategory.PeakAward);
            return categories;
        }
    }
}
