using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OrgHours.db.Global
{
    class Manage
    {
        int globalYearStart, globalYearEnd, globalStudentsLength;

        public Manage(int globalYearStart, int globalYearEnd, int globalStudentsLength)
        {
            this.globalYearStart = globalYearStart;
            this.globalYearEnd = globalYearEnd;
            this.globalStudentsLength = globalStudentsLength;
        }

        public void GetFromFile()
        {
            if(Directory.Exists("Information/Global"))
            {
                if(File.Exists("Information/Global/info.txt"))
                {
                    var text = File.ReadAllLines("Information/Global/info.txt");
                    globalYearStart = int.Parse(text[0]);
                    globalYearEnd = int.Parse(text[1]);
                    globalStudentsLength = int.Parse(text[2]);
                }
                else
                {
                    File.AppendAllLines("Information/Global/info.txt", new string[3] { "2020", "2021", "12" });
                    globalYearStart = 2020;
                    globalYearEnd = 2021;
                    globalStudentsLength = 12;
                }
            }
            else
            {
                Directory.CreateDirectory("Information/Global");
                File.AppendAllLines("Information/Global/info.txt", new string[3] { "2020", "2021", "12" });
                globalYearStart = 2020;
                globalYearEnd = 2021;
                globalStudentsLength = 12;
            }
        }

        public int GlobalYearStart { get => globalYearStart; set => globalYearStart = value; }
        public int GlobalYearEnd { get => globalYearEnd; set => globalYearEnd = value; }
        public int GlobalStudentsLength { get => globalStudentsLength; set => globalStudentsLength = value; }

        internal void SaveToFile(int globalYearStart, int globalYearEnd, int globalStudentsLength)
        {
            this.globalYearStart = globalYearStart;
            this.globalYearEnd = globalYearEnd;
            this.globalStudentsLength = globalStudentsLength;
            File.Delete("Information/Global/info.txt");
            File.AppendAllLines("Information/Global/info.txt", new string[] { globalYearStart.ToString(), globalYearEnd.ToString(), globalStudentsLength.ToString() });
        }
    }
}
