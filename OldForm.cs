using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace OrgHours
{
    public partial class FGeneral : Form
    {
        public FGeneral()
        {
            InitializeComponent();
        }

        private void btnC2AllToNull_Click(object sender, EventArgs e)
        {
            try
            {
                Chemester1.Control.ResetChemesterFolder();
                MessageBox.Show("Выполнено успешно");
            }
            catch (Exception exc)
            {
                MessageBox.Show("Возникла ошибка");
            }
        }

        private void btnC1AllToNull_Click(object sender, EventArgs e)
        {
            try
            {
                Chemester2.Control.ResetChemesterFolder();
                MessageBox.Show("Выполнено успешно");
            }
            catch (Exception exc)
            {
                MessageBox.Show("Возникла ошибка");
            }
        }

        private void btnC1AddPerson_Click(object sender, EventArgs e)
        {
            try
            {
                OrgHours.Chemester1.Person person = new Chemester1.Person(tbC1AddPersonFIO.Text, tbC1AddPersonGroup.Text, int.Parse(tbC1AddPersonHours.Text));
                OrgHours.Chemester1.Control.personNavigation.PersonsAdd(person);
                MessageBox.Show("Добавлен");
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка");
            }
        }

        private void tcC1Actions_Selecting(object sender, TabControlCancelEventArgs e)
        {
            try
            {
                switch (tcC1Actions.SelectedIndex)
                {
                    case 0:
                        OrgHours.Chemester1.Person person = OrgHours.Chemester1.Control.personNavigation.PersonsGet();
                        tbC1ViewPersonFIO.Text = person.Fio;
                        tbC1ViewPersonGroup.Text = person.GroupName;
                        tbC1ViewPersonHours.Text = person.Hours.ToString();
                        break;
                    case 1:
                        tbC1ViewPersonFIO.Text = "";
                        tbC1ViewPersonGroup.Text = "";
                        tbC1ViewPersonHours.Text = "";
                        break;

                    default:
                        break;
                }
            }
            catch (Exception)
            {
                //MessageBox.Show("Ошибка");
            }
        }

        private void tcC2Actions_Selecting(object sender, TabControlCancelEventArgs e)
        {
            try
            {
                switch (tcC1Actions.SelectedIndex)
                {
                    case 0:
                        OrgHours.Chemester2.Person person = OrgHours.Chemester2.Control.personNavigation.PersonsGet();
                        tbC2ViewPersonFIO.Text = person.Fio;
                        tbC2ViewPersonGroup.Text = person.GroupName;
                        tbC2ViewPersonHours.Text = person.Hours.ToString();
                        break;
                    case 1:
                        tbC2ViewPersonFIO.Text = "";
                        tbC2ViewPersonGroup.Text = "";
                        tbC2ViewPersonHours.Text = "";
                        break;

                    default:
                        break;
                }
            }
            catch (Exception)
            {
                //MessageBox.Show("Ошибка");
            }
        }

        private void tcC1ChooseElements_Selecting(object sender, TabControlCancelEventArgs e)
        {
            try
            {
                switch (tcC1ChooseElements.SelectedIndex)
                {
                    case 0:
                        OrgHours.Chemester1.Person person = OrgHours.Chemester1.Control.personNavigation.PersonsGet();
                        tbC1ViewPersonFIO.Text = person.Fio;
                        tbC1ViewPersonGroup.Text = person.GroupName;
                        tbC1ViewPersonHours.Text = person.Hours.ToString();
                        break;
                    case 1:
                        OrgHours.Chemester1.Group group = OrgHours.Chemester1.Control.groupNavigation.GroupsGet();
                        tbC1ViewGroupGroupName.Text = group.GroupName;
                        tbC1ViewGroupHours.Text = group.Hours.ToString();
                        break;

                    default:
                        break;
                }
            }
            catch (Exception)
            {
                //MessageBox.Show("Ошибка");
            }
        }

        private void tcC2ChooseElements_Selecting(object sender, TabControlCancelEventArgs e)
        {
            try
            {
                switch (tcC1ChooseElements.SelectedIndex)
                {
                    case 0:
                        OrgHours.Chemester2.Person person = OrgHours.Chemester2.Control.personNavigation.PersonsGet();
                        tbC2ViewPersonFIO.Text = person.Fio;
                        tbC2ViewPersonGroup.Text = person.GroupName;
                        tbC2ViewPersonHours.Text = person.Hours.ToString();
                        break;
                    case 1:
                        OrgHours.Chemester2.Group group = OrgHours.Chemester2.Control.groupNavigation.GroupsGet();
                        tbC2ViewGroupGroupName.Text = group.GroupName;
                        tbC2ViewGroupHours.Text = group.Hours.ToString();
                        break;

                    default:
                        break;
                }
            }
            catch (Exception)
            {
                // MessageBox.Show("Ошибка");
            }
        }

        private void btnC1PersonNext_Click(object sender, EventArgs e)
        {
            try
            {
                OrgHours.Chemester1.Control.personNavigation.PersonsNext();
                OrgHours.Chemester1.Person person = OrgHours.Chemester1.Control.personNavigation.PersonsGet();
                tbC1ViewPersonFIO.Text = person.Fio;
                tbC1ViewPersonGroup.Text = person.GroupName;
                tbC1ViewPersonHours.Text = person.Hours.ToString();
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка");
            }
        }

        private void btnC1PersonPrevious_Click(object sender, EventArgs e)
        {
            try
            {
                OrgHours.Chemester1.Control.personNavigation.PersonsPrevious();
                OrgHours.Chemester1.Person person = OrgHours.Chemester1.Control.personNavigation.PersonsGet();
                tbC1ViewPersonFIO.Text = person.Fio;
                tbC1ViewPersonGroup.Text = person.GroupName;
                tbC1ViewPersonHours.Text = person.Hours.ToString();
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка");
            }
        }

        private void btnC2PersonPrevious_Click(object sender, EventArgs e)
        {
            try
            {
                OrgHours.Chemester2.Control.personNavigation.PersonsPrevious();
                OrgHours.Chemester2.Person person = OrgHours.Chemester2.Control.personNavigation.PersonsGet();
                tbC2ViewPersonFIO.Text = person.Fio;
                tbC2ViewPersonGroup.Text = person.GroupName;
                tbC2ViewPersonHours.Text = person.Hours.ToString();
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка");
            }
        }

        private void btnC2PersonNext_Click(object sender, EventArgs e)
        {
            try
            {
                OrgHours.Chemester2.Control.personNavigation.PersonsNext();
                OrgHours.Chemester2.Person person = OrgHours.Chemester2.Control.personNavigation.PersonsGet();
                tbC2ViewPersonFIO.Text = person.Fio;
                tbC2ViewPersonGroup.Text = person.GroupName;
                tbC2ViewPersonHours.Text = person.Hours.ToString();
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка");
            }
        }

        private void btnC2AddPerson_Click(object sender, EventArgs e)
        {
            try
            {
                OrgHours.Chemester2.Person person = new Chemester2.Person(tbC2AddPersonFIO.Text, tbC2AddPersonGroup.Text, int.Parse(tbC2AddPersonHours.Text));
                OrgHours.Chemester2.Control.personNavigation.PersonsAdd(person);
                MessageBox.Show("Добавлен");
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка");
            }
        }

        private void btnC1ViewPersonDelete_Click(object sender, EventArgs e)
        {
            try
            {
                OrgHours.Chemester1.Control.personNavigation.PersonsDelete();
                tbC1ViewPersonFIO.Text = "";
                tbC1ViewPersonGroup.Text = "";
                tbC1ViewPersonHours.Text = "";
                MessageBox.Show("Удален");
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка");
            }
        }

        private void btnC2ViewPersonDelete_Click(object sender, EventArgs e)
        {
            try
            {
                OrgHours.Chemester2.Control.personNavigation.PersonsDelete();
                tbC1ViewPersonFIO.Text = "";
                tbC1ViewPersonGroup.Text = "";
                tbC1ViewPersonHours.Text = "";
                MessageBox.Show("Удален");
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка");
            }
        }

        private void btnC1PersonReset_Click(object sender, EventArgs e)
        {
            try
            {
                OrgHours.Chemester1.Person person = OrgHours.Chemester1.Control.personNavigation.PersonsGet();
                tbC1ViewPersonFIO.Text = person.Fio;
                tbC1ViewPersonGroup.Text = person.GroupName;
                tbC1ViewPersonHours.Text = person.Hours.ToString();
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка");
            }
        }

        private void btnC2PersonReset_Click(object sender, EventArgs e)
        {
            try
            {
                OrgHours.Chemester2.Person person = OrgHours.Chemester2.Control.personNavigation.PersonsGet();
                tbC2ViewPersonFIO.Text = person.Fio;
                tbC2ViewPersonGroup.Text = person.GroupName;
                tbC2ViewPersonHours.Text = person.Hours.ToString();
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка");
            }
        }

        private void btnC2GroupNext_Click(object sender, EventArgs e)
        {
            try
            {
                OrgHours.Chemester2.Control.groupNavigation.GroupsNext();
                OrgHours.Chemester2.Group group = OrgHours.Chemester2.Control.groupNavigation.GroupsGet();
                tbC2ViewGroupGroupName.Text = group.GroupName;
                tbC2ViewGroupHours.Text = group.Hours.ToString();
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка");
            }
        }

        private void btnC2GroupPrevious_Click(object sender, EventArgs e)
        {
            try
            {
                OrgHours.Chemester2.Control.groupNavigation.GroupsPrevious();
                OrgHours.Chemester2.Group group = OrgHours.Chemester2.Control.groupNavigation.GroupsGet();
                tbC2ViewGroupGroupName.Text = group.GroupName;
                tbC2ViewGroupHours.Text = group.Hours.ToString();
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка");
            }
        }

        private void btnC1GroupPrevious_Click(object sender, EventArgs e)
        {
            try
            {
                OrgHours.Chemester1.Control.groupNavigation.GroupsPrevious();
                OrgHours.Chemester1.Group group = OrgHours.Chemester1.Control.groupNavigation.GroupsGet();
                tbC1ViewGroupGroupName.Text = group.GroupName;
                tbC1ViewGroupHours.Text = group.Hours.ToString();
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка");
            }
        }

        private void btnC1GroupNext_Click(object sender, EventArgs e)
        {
            try
            {
                OrgHours.Chemester1.Control.groupNavigation.GroupsNext();
                OrgHours.Chemester1.Group group = OrgHours.Chemester1.Control.groupNavigation.GroupsGet();
                tbC1ViewGroupGroupName.Text = group.GroupName;
                tbC1ViewGroupHours.Text = group.Hours.ToString();
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка");
            }
        }

        private void btnC1ViewGroupReset_Click(object sender, EventArgs e)
        {
            try
            {
                OrgHours.Chemester1.Group group = OrgHours.Chemester1.Control.groupNavigation.GroupsGet();
                tbC1ViewGroupGroupName.Text = group.GroupName;
                tbC1ViewGroupHours.Text = group.Hours.ToString();
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка");
            }
        }

        private void btnC2ViewGroupReset_Click(object sender, EventArgs e)
        {
            try
            {
                OrgHours.Chemester2.Group group = OrgHours.Chemester2.Control.groupNavigation.GroupsGet();
                tbC2ViewGroupGroupName.Text = group.GroupName;
                tbC2ViewGroupHours.Text = group.Hours.ToString();
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка");
            }
        }

        private void btnC2AddGroup_Click(object sender, EventArgs e)
        {
            try
            {
                OrgHours.Chemester2.Group group = new Chemester2.Group(tbC2AddGroupGroupName.Text, int.Parse(tbC2AddGroupHours.Text));
                OrgHours.Chemester2.Control.groupNavigation.GroupsAdd(group);
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка");
            }
        }

        private void btnC1AddGroup_Click(object sender, EventArgs e)
        {
            try
            {
                OrgHours.Chemester1.Group group = new Chemester1.Group(tbC1AddGroupGroupName.Text, int.Parse(tbC1AddGroupHours.Text));
                OrgHours.Chemester1.Control.groupNavigation.GroupsAdd(group);
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка");
            }
        }

        private void btnC1ViewGroupDelete_Click(object sender, EventArgs e)
        {
            try
            {
                OrgHours.Chemester1.Control.groupNavigation.GroupsDelete();
                tbC1ViewGroupGroupName.Text = "";
                tbC1ViewGroupHours.Text = "";
                MessageBox.Show("Удален");
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка");
            }
        }

        private void btnC2ViewGroupDelete_Click(object sender, EventArgs e)
        {
            try
            {
                OrgHours.Chemester2.Control.groupNavigation.GroupsDelete();
                tbC2ViewGroupGroupName.Text = "";
                tbC2ViewGroupHours.Text = "";
                MessageBox.Show("Удален");
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка");
            }
        }

        private void btnC1HardReset_Click(object sender, EventArgs e)
        {
            try
            {
                void SetValueNull()
                {
                    void Chemester1()
                    {
                        OrgHours.Chemester1.Control.LGroups = new List<Chemester1.Group>();
                        OrgHours.Chemester1.Control.groupNavigation = new OrgHours.Chemester1.GroupNavigation(new List<Chemester1.Group>());
                        OrgHours.Chemester1.Control.LPersons = new List<Chemester1.Person>();
                        OrgHours.Chemester1.Control.personNavigation = new OrgHours.Chemester1.PersonNavigation(new List<Chemester1.Person>());
                    }
                    Chemester1();
                }

                void ClearTbValues()
                {
                    void Chemester1()
                    {
                        tbC1ViewPersonFIO.Text = "";
                        tbC1ViewPersonGroup.Text = "";
                        tbC1ViewPersonHours.Text = "";
                    }
                    Chemester1();
                }

                OrgHours.Chemester1.Control.ResetChemesterFolder();
                SetValueNull();
                ClearTbValues();
                MessageBox.Show("Выполнена полная очистка информации по семестру 1");
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка");
            }
        }

        private void btnC2HardReset_Click(object sender, EventArgs e)
        {
            try
            {
                void SetValueNull()
                {
                    void Chemester2()
                    {
                        OrgHours.Chemester2.Control.LGroups = new List<Chemester2.Group>();
                        OrgHours.Chemester2.Control.groupNavigation = new OrgHours.Chemester2.GroupNavigation(new List<Chemester2.Group>());
                        OrgHours.Chemester2.Control.LPersons = new List<Chemester2.Person>();
                        OrgHours.Chemester2.Control.personNavigation = new OrgHours.Chemester2.PersonNavigation(new List<Chemester2.Person>());
                    }
                    Chemester2();
                }

                void ClearTbValues()
                {
                    void Chemester2()
                    {
                        tbC2ViewPersonFIO.Text = "";
                        tbC2ViewPersonGroup.Text = "";
                        tbC2ViewPersonHours.Text = "";
                    }
                    Chemester2();
                }

                OrgHours.Chemester2.Control.ResetChemesterFolder();
                SetValueNull();
                ClearTbValues();
                MessageBox.Show("Выполнена полная очистка информации по семестру 2");
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка");
            }
        }

        private void btnC1SaveAllPersonsAndGroups_Click(object sender, EventArgs e)
        {
            try
            {
                void Chemester1()
                {
                    File.Delete("Information/Chemester1/Persons.txt");
                    File.Delete("Information/Chemester1/HoursGroups.txt");
                    File.AppendAllLines("Information/Chemester1/Persons.txt", OrgHours.Chemester1.Control.LPersons.Select(x => x.ToTextRow()));
                    File.AppendAllLines("Information/Chemester1/HoursGroups.txt", OrgHours.Chemester1.Control.LGroups.Select(x => x.ToTextRow()));
                }
                Chemester1();
                MessageBox.Show("Сохранение произошло успешно");
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка");
            }
        }

        private void btnC2SaveAllPersonsAndGroups_Click(object sender, EventArgs e)
        {
            try
            {
                void Chemester2()
                {
                    File.Delete("Information/Chemester2/Persons.txt");
                    File.Delete("Information/Chemester2/HoursGroups.txt");
                    File.AppendAllLines("Information/Chemester2/Persons.txt", OrgHours.Chemester2.Control.LPersons.Select(x => x.ToTextRow()));
                    File.AppendAllLines("Information/Chemester2/HoursGroups.txt", OrgHours.Chemester2.Control.LGroups.Select(x => x.ToTextRow()));
                }
                Chemester2();
                MessageBox.Show("Сохранение произошло успешно");
            }
            catch (Exception exc)
            {
                MessageBox.Show("Ошибка");
            }
        }

        private void btnC2CreateDoc_Click(object sender, EventArgs e)
        {
            try
            {
                void DocSpace(DocX origDocument, int count)
                {
                    for (int i = 1; i <= count; i++)
                    {
                        origDocument.InsertParagraph("").FontSize(12d).SpacingAfter(4d).Alignment = Alignment.left;
                    }
                }
                var document = DocX.Create(Path.Combine(Environment.CurrentDirectory, "\\Output\\Семестр_1.docx"));
                document.InsertParagraph("Список групп").FontSize(20d).SpacingAfter(8d).Alignment = Alignment.center;
                int group_position = 0;

                //GroupBy
                List<Chemester2.Group> LGroupsNow = new List<Chemester2.Group>();
                foreach (var groupItem in Chemester2.Control.LGroups)
                {
                    bool isHas = false;
                    foreach (var groupNow in LGroupsNow)
                    {
                        if (groupNow.GroupName == groupItem.GroupName)
                            isHas = true;
                    }
                    if (!isHas)
                        LGroupsNow.Add(groupItem);
                }
                //End

                //GroupBy
                List<int> LUniqueHoursNow = new List<int>();
                foreach (var group in Chemester2.Control.LGroups)
                {
                    bool isHas = false;
                    foreach (var timeNow in LUniqueHoursNow)
                    {
                        if (group.Hours == timeNow)
                            isHas = true;
                    }
                    if (!isHas)
                        LUniqueHoursNow.Add(group.Hours);
                }
                //End
                int groupNumber = 0;
                for (int i = 0; i < LUniqueHoursNow.Count(); i++)
                {
                    groupNumber++;
                    int timeNow = LUniqueHoursNow[i];
                    document.InsertParagraph($"Группа {groupNumber}     {timeNow} часов").FontSize(12d).SpacingAfter(4d).Alignment = Alignment.left;
                    document.InsertParagraph($"(включая физкультативное занятие)").FontSize(12d).SpacingAfter(4d).Alignment = Alignment.left;

                    int personPosition = 0;
                    foreach (var itemGroup in Chemester2.Control.LGroups)
                    {
                        if (itemGroup.Hours == timeNow)
                        {
                            foreach (var itemPerson in Chemester2.Control.LPersons)
                            {
                                if (itemGroup.GroupName == itemPerson.GroupName)
                                {
                                    ++personPosition;
                                    if (personPosition == 13)
                                    {
                                        DocSpace(document, 2);
                                        personPosition = 1;
                                        ++groupNumber;
                                        document.InsertParagraph($"Группа {groupNumber}     {timeNow} часов").FontSize(12d).SpacingAfter(4d).Alignment = Alignment.left;
                                        document.InsertParagraph($"(включая физкультативное занятие)").FontSize(12d).SpacingAfter(4d).Alignment = Alignment.left;
                                        document.InsertParagraph($"{personPosition}. {itemPerson.Fio} {itemPerson.GroupName}").FontSize(12d).SpacingAfter(4d).Alignment = Alignment.left;
                                    }
                                    else
                                        document.InsertParagraph($"{personPosition}. {itemPerson.Fio} {itemPerson.GroupName}").FontSize(12d).SpacingAfter(4d).Alignment = Alignment.left;
                                }
                            }

                        }
                    }
                    DocSpace(document, 2);
                }
                document.SaveAs(Environment.CurrentDirectory + "\\Output\\Семестр_2.docx");
                MessageBox.Show("Документ успешно создан, находится он в папке Output, а эта папка находится в том же месте где и программа");
                Process.Start(Environment.CurrentDirectory + "\\Output\\Семестр_2.docx");
            }
            catch (Exception exc)
            {
                MessageBox.Show("Ошибка");
            }
        }

        private void btnC1CreateDoc_Click(object sender, EventArgs e)
        {
            try
            {
                void DocSpace(DocX origDocument, int count)
                {
                    for (int i = 1; i <= count; i++)
                    {
                        origDocument.InsertParagraph("").FontSize(12d).SpacingAfter(4d).Alignment = Alignment.left;
                    }
                }
                var document = DocX.Create(Path.Combine(Environment.CurrentDirectory, "\\Output\\Семестр_1.docx"));
                document.InsertParagraph("Список групп").FontSize(20d).SpacingAfter(8d).Alignment = Alignment.center;
                int group_position = 0;

                //GroupBy
                List<Chemester1.Group> LGroupsNow = new List<Chemester1.Group>();
                foreach (var groupItem in Chemester1.Control.LGroups)
                {
                    bool isHas = false;
                    foreach (var groupNow in LGroupsNow)
                    {
                        if (groupNow.GroupName == groupItem.GroupName)
                            isHas = true;
                    }
                    if (!isHas)
                        LGroupsNow.Add(groupItem);
                }
                //End

                //GroupBy
                List<int> LUniqueHoursNow = new List<int>();
                foreach (var group in Chemester1.Control.LGroups)
                {
                    bool isHas = false;
                    foreach (var timeNow in LUniqueHoursNow)
                    {
                        if (group.Hours == timeNow)
                            isHas = true;
                    }
                    if (!isHas)
                        LUniqueHoursNow.Add(group.Hours);
                }
                //End
                int groupNumber = 0;
                for (int i = 0; i < LUniqueHoursNow.Count(); i++)
                {
                    groupNumber++;
                    int timeNow = LUniqueHoursNow[i];
                    document.InsertParagraph($"Группа {groupNumber}     {timeNow} часов").FontSize(12d).SpacingAfter(4d).Alignment = Alignment.left;
                    document.InsertParagraph($"(включая физкультативное занятие)").FontSize(12d).SpacingAfter(4d).Alignment = Alignment.left;

                    int personPosition = 0;
                    foreach (var itemGroup in Chemester1.Control.LGroups)
                    {
                        if (itemGroup.Hours == timeNow)
                        {
                            foreach (var itemPerson in Chemester1.Control.LPersons)
                            {
                                if (itemGroup.GroupName == itemPerson.GroupName)
                                {
                                    ++personPosition;
                                    if (personPosition == 13)
                                    {
                                        DocSpace(document, 2);
                                        personPosition = 1;
                                        ++groupNumber;
                                        document.InsertParagraph($"Группа {groupNumber}     {timeNow} часов").FontSize(12d).SpacingAfter(4d).Alignment = Alignment.left;
                                        document.InsertParagraph($"(включая физкультативное занятие)").FontSize(12d).SpacingAfter(4d).Alignment = Alignment.left;
                                        document.InsertParagraph($"{personPosition}. {itemPerson.Fio} {itemPerson.GroupName}").FontSize(12d).SpacingAfter(4d).Alignment = Alignment.left;
                                    }
                                    else
                                        document.InsertParagraph($"{personPosition}. {itemPerson.Fio} {itemPerson.GroupName}").FontSize(12d).SpacingAfter(4d).Alignment = Alignment.left;
                                }
                            }

                        }
                    }
                    DocSpace(document, 2);
                }

                /*bool isEndGenerateDoc = false;
                group_position = 0;
                Chemester1.GroupNavigation groupNavigation = new Chemester1.GroupNavigation(LGroupsNow);
                while (!isEndGenerateDoc)
                {
                    group_position++;
                    groupNavigation.GroupsGet();
                    isEndGenerateDoc = true;
                }*/

                /*foreach (var groupItem in LGroupsNow)
                {
                    group_position++;
                    List<OrgHours.Chemester1.Person> getTruePersons()
                    {
                        List<OrgHours.Chemester1.Person> LPersonsTemp = new List<Chemester1.Person>();
                        foreach (var item in OrgHours.Chemester1.Control.LPersons)
                        {
                            if (item.GroupName == groupItem.GroupName)
                                LPersonsTemp.Add(item);
                        }

                        return LPersonsTemp;
                    }

                    document.InsertParagraph($"Группа {group_position}     {groupItem.Hours} часов").FontSize(12d).SpacingAfter(4d).Alignment = Alignment.left;
                    document.InsertParagraph($"(включая физкультативное занятие)").FontSize(12d).SpacingAfter(4d).Alignment = Alignment.left;

                    foreach (var itemTruePerson in getTruePersons())
                    {
                        document.InsertParagraph($"{itemTruePerson.Fio}").FontSize(12d).SpacingAfter(4d).Alignment = Alignment.left;
                    }
                    DocSpace(document, 2);
                }*/
                document.SaveAs(Environment.CurrentDirectory + "\\Output\\Семестр_1.docx");
                MessageBox.Show("Документ успешно создан, находится он в папке Output, а эта папка находится в том же месте где и программа");
                Process.Start(Environment.CurrentDirectory + "\\Output\\Семестр_1.docx");
            }
            catch (Exception exc)
            {
                MessageBox.Show("Ошибка");
            }
        }

        private void tcC1Chemesters_Selecting(object sender, TabControlCancelEventArgs e)
        {
            try
            {
                switch (tcC1ChooseElements.SelectedIndex)
                {
                    case 0:
                        OrgHours.Chemester1.Person person1 = OrgHours.Chemester1.Control.personNavigation.PersonsGet();
                        tbC1ViewPersonFIO.Text = person1.Fio;
                        tbC1ViewPersonGroup.Text = person1.GroupName;
                        tbC1ViewPersonHours.Text = person1.Hours.ToString();

                        OrgHours.Chemester1.Group group1 = OrgHours.Chemester1.Control.groupNavigation.GroupsGet();
                        tbC1ViewGroupGroupName.Text = group1.GroupName;
                        tbC1ViewGroupHours.Text = group1.Hours.ToString();
                        break;
                    case 1:
                        OrgHours.Chemester2.Person person2 = OrgHours.Chemester2.Control.personNavigation.PersonsGet();
                        tbC2ViewPersonFIO.Text = person2.Fio;
                        tbC2ViewPersonGroup.Text = person2.GroupName;
                        tbC2ViewPersonHours.Text = person2.Hours.ToString();

                        OrgHours.Chemester2.Group group2 = OrgHours.Chemester2.Control.groupNavigation.GroupsGet();
                        tbC2ViewGroupGroupName.Text = group2.GroupName;
                        tbC2ViewGroupHours.Text = group2.Hours.ToString();
                        break;

                    default:
                        break;
                }
            }
            catch (Exception)
            {
                //MessageBox.Show("Ошибка");
            }
        }

        private void FGeneral_Load(object sender, EventArgs e)
        {
            OrgHours.Chemester1.Person person1 = OrgHours.Chemester1.Control.personNavigation.PersonsGet();
            tbC1ViewPersonFIO.Text = person1.Fio;
            tbC1ViewPersonGroup.Text = person1.GroupName;
            tbC1ViewPersonHours.Text = person1.Hours.ToString();

            OrgHours.Chemester1.Group group1 = OrgHours.Chemester1.Control.groupNavigation.GroupsGet();
            tbC1ViewGroupGroupName.Text = group1.GroupName;
            tbC1ViewGroupHours.Text = group1.Hours.ToString();


            OrgHours.Chemester2.Person person2 = OrgHours.Chemester2.Control.personNavigation.PersonsGet();
            tbC2ViewPersonFIO.Text = person2.Fio;
            tbC2ViewPersonGroup.Text = person2.GroupName;
            tbC2ViewPersonHours.Text = person2.Hours.ToString();

            OrgHours.Chemester2.Group group2 = OrgHours.Chemester2.Control.groupNavigation.GroupsGet();
            tbC2ViewGroupGroupName.Text = group2.GroupName;
            tbC2ViewGroupHours.Text = group2.Hours.ToString();
        }

        private void btnC1ViewPersonChange_Click(object sender, EventArgs e)
        {
            try
            {
                OrgHours.Chemester1.Person person = new Chemester1.Person(tbC1ViewPersonFIO.Text, tbC1ViewPersonGroup.Text, int.Parse(tbC1ViewPersonHours.Text));
                OrgHours.Chemester1.Control.personNavigation.PersonsChange(person);
                MessageBox.Show("Изменено");
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка");
            }
        }

        private void btnC2ViewPersonChange_Click(object sender, EventArgs e)
        {
            try
            {
                OrgHours.Chemester2.Person person = new Chemester2.Person(tbC2ViewPersonFIO.Text, tbC2ViewPersonGroup.Text, int.Parse(tbC2ViewPersonHours.Text));
                OrgHours.Chemester2.Control.personNavigation.PersonsChange(person);
                MessageBox.Show("Изменено");
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка");
            }
        }

        private void btnC1ViewGroupChange_Click(object sender, EventArgs e)
        {
            try
            {
                OrgHours.Chemester1.Group group = new Chemester1.Group(tbC1ViewGroupGroupName.Text, int.Parse(tbC1ViewGroupHours.Text));
                OrgHours.Chemester1.Control.groupNavigation.GroupsChange(group);
                MessageBox.Show("Изменено");
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка");
            }
        }

        private void btnC2ViewGroupChange_Click(object sender, EventArgs e)
        {
            try
            {
                OrgHours.Chemester2.Group group = new Chemester2.Group(tbC2ViewGroupGroupName.Text, int.Parse(tbC2ViewGroupHours.Text));
                OrgHours.Chemester2.Control.groupNavigation.GroupsChange(group);
                MessageBox.Show("Изменено");
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка");
            }
        }
    }
}
