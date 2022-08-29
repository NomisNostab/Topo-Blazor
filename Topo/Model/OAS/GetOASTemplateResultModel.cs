namespace Topo.Model.OAS
{
    public class GetOASTemplateResultModel
    {
        public string template { get; set; } = "";
        public Meta meta { get; set; } = new Meta();
        public int version { get; set; }
        public Document[] document { get; set; } = new Document[0];
    }

    public class Meta
    {
        public int stage { get; set; }
        public string stream { get; set; } = "";
    }

    public class Document
    {
        public string title { get; set; } = "";
        public Input_Groups[] input_groups { get; set; } = new Input_Groups[0];
    }

    public class Input_Groups
    {
        public string title { get; set; } = "";
        public Input[] inputs { get; set; } = new Input[0];
    }

    public class Input
    {
        public string id { get; set; } = "";
        public string type { get; set; } = "";
        public string label { get; set; } = "";
        public string dialog_text { get; set; } = "";
        public string alt { get; set; } = "";
    }
}
