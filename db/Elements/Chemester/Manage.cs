using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OrgHours.db.Elements.Chemester
{
    class Manage
    {
        List<Change.Control> lChanges = new List<Change.Control>();

        public Manage(int changesCount)
        {
            for (int i = 0; i < changesCount; i++)
            {
                lChanges.Add(new Change.Control());
            }
        }

        public List<Change.Control> LChanges { get => lChanges; set => lChanges = value; }
    }
}
