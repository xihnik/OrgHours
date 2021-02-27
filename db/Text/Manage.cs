using OrgHours.db.Text.ProcessReport;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OrgHours.db.Text
{
    class Manage
    {
        ProcessReport.Errors errors = new ProcessReport.Errors();
        ProcessReport.Complete complete = new ProcessReport.Complete();

        public Errors Errors { get => errors; set => errors = value; }
        internal Complete Complete { get => complete; set => complete = value; }
    }
}
