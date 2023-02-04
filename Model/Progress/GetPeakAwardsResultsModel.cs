using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Topo.Model.Progress
{
    public class GetPeakAwardsResultsModel
    {
        public PeakAwardResult[] results { get; set; } = new PeakAwardResult[0];
    }

    public class PeakAwardResult
    {
        public string id { get; set; } = string.Empty;
        public string member_id { get; set; } = string.Empty;
        public string template { get; set; } = string.Empty;
        public int version { get; set; }
        public string section { get; set; } = string.Empty;
        public string type { get; set; } = string.Empty;
        public string status { get; set; } = string.Empty;
        public DateTime status_updated { get; set; }
        public DateTime last_updated { get; set; }
    }

}
