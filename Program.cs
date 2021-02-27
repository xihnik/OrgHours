using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OrgHours
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {
            db.References.elementsManage = new db.Elements.Manage(2, 2);
            db.References.textManage = new db.Text.Manage();
            db.References.globalManage = new db.Global.Manage(2020, 2021, 12);

            if (!Directory.Exists("Output"))
            {
                CreateDirectoryOutput();
                if (!Directory.Exists("Information"))
                {
                    CreateDirectoryInformation();
                    GetValueFromExistedFiles();
                    db.References.globalManage.GetFromFile();
                }
                else
                {
                    GetValueFromExistedFiles();
                    db.References.globalManage.GetFromFile();
                }
            }
            else
            {
                if (!Directory.Exists("Information"))
                {
                    CreateDirectoryInformation();
                    ///SetValueNull();
                    db.References.globalManage.GetFromFile();
                }
                else
                {
                    GetValueFromExistedFiles();
                    db.References.globalManage.GetFromFile();
                }
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new FGeneral());
        }

        private static void SetValueNull()
        {
            for (int i = 0; i < db.References.elementsManage.LChemesters.Count(); i++)
            {
                db.Elements.Chemester.Manage chemester = db.References.elementsManage.LChemesters[i];

                for (int j = 0; j < chemester.LChanges.Count(); j++)
                {
                    db.Elements.Chemester.Change.Control change = chemester.LChanges[j];

                    change.LGroups = new List<db.Elements.Chemester.Change.Group>();
                    change.LPersons = new List<db.Elements.Chemester.Change.Person>();
                }
            }
        }

        static void GetValueFromExistedFiles()
        {
            for (int i = 0; i < db.References.elementsManage.LChemesters.Count(); i++)
            {
                db.Elements.Chemester.Manage chemester = db.References.elementsManage.LChemesters[i];

                for (int j = 0; j < chemester.LChanges.Count(); j++)
                {
                    db.Elements.Chemester.Change.Control change = chemester.LChanges[j];

                    void GetHoursGroups()
                    {
                        List<db.Elements.Chemester.Change.Group> LGroupsTemp = new List<db.Elements.Chemester.Change.Group>();
                        foreach (var item in File.ReadAllLines($"Information/Chemester{i}/Change{j}/HoursGroups.txt"))
                        {
                            LGroupsTemp.Add(new db.Elements.Chemester.Change.Group(item.Split('|')[0], int.Parse(item.Split('|')[1])));
                        }
                        change.LGroups = LGroupsTemp;
                        change.groupNavigation = new db.Elements.Chemester.Change.GroupNavigation(LGroupsTemp);
                    }

                    void GetPersons()
                    {
                        List<db.Elements.Chemester.Change.Person> LPersonsTemp = new List<db.Elements.Chemester.Change.Person>();
                        foreach (var item in File.ReadAllLines($"Information/Chemester{i}/Change{j}/Persons.txt"))
                        {
                            LPersonsTemp.Add(new db.Elements.Chemester.Change.Person(item.Split('|')[0], item.Split('|')[1]));
                        }
                        change.LPersons = LPersonsTemp;
                        change.personNavigation = new db.Elements.Chemester.Change.PersonNavigation(LPersonsTemp);
                    }

                    change.LGroups = new List<db.Elements.Chemester.Change.Group>();
                    change.LPersons = new List<db.Elements.Chemester.Change.Person>();

                    GetHoursGroups();
                    GetPersons();
                }
            }
        }

        static void CreateDirectoryInformation()
        {
            for (int i = 0; i < db.References.elementsManage.LChemesters.Count(); i++)
            {
                db.Elements.Chemester.Manage chemester = db.References.elementsManage.LChemesters[i];

                Directory.CreateDirectory($"Information/Chemester{i}");

                for (int j = 0; j < chemester.LChanges.Count(); j++)
                {
                    db.Elements.Chemester.Change.Control change = chemester.LChanges[j];

                    Directory.CreateDirectory($"Information/Chemester{i}/Change{j}");
                    File.AppendAllText($"Information/Chemester{i}/Change{j}/HoursGroups.txt", "", System.Text.Encoding.UTF8);
                    File.AppendAllText($"Information/Chemester{i}/Change{j}/Persons.txt", "", System.Text.Encoding.UTF8);
                }
            }
        }
        static void CreateDirectoryOutput()
        {
            Directory.CreateDirectory("Output");
        }
    }
}
