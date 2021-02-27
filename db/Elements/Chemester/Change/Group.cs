using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OrgHours.db.Elements.Chemester.Change
{
    class Group
    {
        string groupName;
        int hours;

        public Group(string groupName, int hours)
        {
            this.groupName = groupName;
            this.hours = hours;
        }

        public string GroupName { get => groupName; set => groupName = value; }
        public int Hours { get => hours; set => hours = value; }
        public string ToTextRow() => $"{groupName}|{hours}";
    }
}
