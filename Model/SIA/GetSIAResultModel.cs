namespace Topo.Model.SIA
{
    public class GetSIAResultModel
    {
        public string id { get; set; } = string.Empty;
        public string member_id { get; set; } = string.Empty;
        public string template { get; set; } = string.Empty;
        public int version { get; set; }
        public string section { get; set; } = string.Empty;
        public string type { get; set; } = string.Empty;
        public Answers answers { get; set; } = new Answers();
        public string status { get; set; } = string.Empty;
        public DateTime status_updated { get; set; }
    }
}
