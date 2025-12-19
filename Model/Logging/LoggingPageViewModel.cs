using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Topo.Model.Loging
{
    public class LoggingPageViewModel
    {
        [Display(Name = "Enable Debug Loging")]
        public bool StartLogging { get; set; }
        public bool DownloadLog { get; set; }
    }
}
