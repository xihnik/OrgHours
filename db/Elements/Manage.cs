using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OrgHours.db.Elements
{
    class Manage
    {
        List<Chemester.Manage> lChemesters = new List<Chemester.Manage>();

        public Manage(int chemestersCount, int changesCount)
        {
            for (int i = 0; i < chemestersCount; i++)
            {
                lChemesters.Add(new Chemester.Manage(changesCount));
            }
        }

        internal List<Chemester.Manage> LChemesters { get => lChemesters; set => lChemesters = value; }
    }
}
