using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CDRcompare
{
    public class CDRItem
    {
        public string ani;
        public string dest;
        public DateTime start;
        public double duration;

        public CDRItem()
        {
            ani = "";
            dest = "";
            start = DateTime.MinValue;
            duration = 0.0;
        }

        public CDRItem(string ani, string dest, DateTime start, double duration)
        {
            this.ani = ani;
            this.dest = dest;
            this.start = start;
            this.duration = duration;
        }
    }

}
