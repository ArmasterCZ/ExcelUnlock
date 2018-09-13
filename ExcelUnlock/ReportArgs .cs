using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelUnlock
{
    /// <summary>
    /// event for callback
    /// </summary>
    
    class ReportArgs : EventArgs
    {
        public string report { get; set; }
    }
}
