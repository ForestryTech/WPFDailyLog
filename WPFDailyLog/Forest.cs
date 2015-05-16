using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WPFDailyLog
{
    struct Forest
    {
        public string ForestName;
        public string Abbr;
        public string ForestNumber;

        public Forest(string forestName, string abbr, string forestNumber) {
            ForestName = forestName;
            Abbr = abbr;
            ForestNumber = forestNumber;
        }
    }
}
