using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Topo.Model.Progress

{
    public class GetIntroductionResultsModel
    {
        public Result[] results { get; set; } = new Result[0];
    }

    public class Result
    {
        public string id { get; set; } = string.Empty;
        public string member_id { get; set; } = string.Empty;
        public string section { get; set; } = string.Empty;
        public string type { get; set; } = string.Empty;
        public string status { get; set; } = string.Empty;
        public DateTime status_updated { get; set; }
        public Imported imported { get; set; } = new Imported();
    }

    public class Imported
    {
        public string date_awarded { get; set; } = string.Empty;
        public Awarded_By awarded_by { get; set; } = new Awarded_By();
    }

    public class Awarded_By
    {
        public string id { get; set; } = string.Empty;
        public string name { get; set; } = string.Empty;
    }

}
