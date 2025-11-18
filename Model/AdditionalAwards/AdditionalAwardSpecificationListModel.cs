namespace Topo.Model.AdditionalAwards
{
    public class AdditionalAwardSpecificationListModel
    {
        public string id { get; set; } = string.Empty;
        public string name { get; set; } = string.Empty;
        public int additionalAwardSortIndex { get; set; } = 0;
        public int additionalAwardValue { get; set; } = 0;
    }
}
