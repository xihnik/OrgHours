using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OrgHours.db.Elements.Chemester.Change
{
    class Person
    {
        string fio, groupName;

        public Person(string fio, string groupName)
        {
            this.fio = fio;
            this.groupName = groupName;
        }

        public string Fio { get => fio; set => fio = value; }
        public string GroupName { get => groupName; set => groupName = value; }
        public string ToTextRow() => $"{fio}|{groupName}";
    }
}
