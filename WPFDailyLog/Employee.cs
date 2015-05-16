using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WPFDailyLog
{
    class Employee
    {
        private string first;
        private string last;

        public string First { get { return first; } set { first = value; } }
        public string Last { get {return last;} set { last = value;}}

        public Employee(string last, string first) {
            this.last = last;
            this.first = first;
        }
    }
}
