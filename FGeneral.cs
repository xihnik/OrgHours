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
using System.Windows;
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

        private void FGeneral_Load(object sender, EventArgs e)
        {
            tbGlobalYearStart.Text = db.References.globalManage.GlobalYearStart.ToString();
            tbGlobalYearEnd.Text = db.References.globalManage.GlobalYearEnd.ToString();
            tbGlobalStudentsLength.Text = db.References.globalManage.GlobalStudentsLength.ToString();

            if (db.References.elementsManage.LChemesters.Count() > 0)
            {
                for (int numberChemester = 0; numberChemester < db.References.elementsManage.LChemesters.Count(); numberChemester++)
                {
                    db.Elements.Chemester.Manage chemester = db.References.elementsManage.LChemesters[numberChemester];


                    if (chemester.LChanges.Count() > 0)
                    {
                        for (int numberChange = 0; numberChange < chemester.LChanges.Count(); numberChange++)
                        {
                            db.Elements.Chemester.Change.Control change = chemester.LChanges[numberChange];
                            db.Elements.Chemester.Change.Person person = new db.Elements.Chemester.Change.Person("", "");
                            db.Elements.Chemester.Change.Group group = new db.Elements.Chemester.Change.Group("", 0);

                            if (change.LGroups.Count() > 0)
                            {
                                group = change.LGroups[0];
                            }
                            if (change.LPersons.Count() > 0)
                            {
                                person = change.LPersons[0];
                            }

                            void SetLabels()
                            {
                                void SetPerson()
                                {
                                    this.Controls[0].Controls[0].Controls[0].Controls[numberChemester].Controls[0].Controls[numberChange].Controls[0].Controls[0].Controls[0].Controls[0].Controls[0].Controls[$"tb_C{numberChemester}_C{numberChange}_ViewPersonFIO"].Text = person.Fio;
                                    this.Controls[0].Controls[0].Controls[0].Controls[numberChemester].Controls[0].Controls[numberChange].Controls[0].Controls[0].Controls[0].Controls[0].Controls[0].Controls[$"tb_C{numberChemester}_C{numberChange}_ViewPersonGroup"].Text = person.GroupName;
                                }

                                void SetGroup()
                                {
                                    this.Controls[0].Controls[0].Controls[0].Controls[numberChemester].Controls[0].Controls[numberChange].Controls[0].Controls[1].Controls[0].Controls[0].Controls[0].Controls[$"tb_C{numberChemester}_C{numberChange}_ViewGroupGroupName"].Text = group.GroupName;
                                    this.Controls[0].Controls[0].Controls[0].Controls[numberChemester].Controls[0].Controls[numberChange].Controls[0].Controls[1].Controls[0].Controls[0].Controls[0].Controls[$"tb_C{numberChemester}_C{numberChange}_ViewGroupHours"].Text = group.Hours.ToString();
                                }

                                SetPerson();
                                SetGroup();
                            }


                            void SetButtonsClicks()
                            {
                                void SetPersonClicks()
                                {
                                    void Navigation()
                                    {
                                        this.Controls[0].Controls[0].Controls[0].Controls[numberChemester].Controls[0].Controls[numberChange].Controls[0].Controls[0].Controls[0].Controls[0].Controls[$"btn_C{numberChemester}_C{numberChange}_PersonPrevious"].Click += (childSender, childArgs) =>
                                        {
                                            try
                                            {
                                                change.personNavigation.PersonsPrevious();
                                                db.Elements.Chemester.Change.Person _person = change.personNavigation.PersonsGet();
                                                var partsBtn = ((Button)childSender).Name.Split("_C".ToCharArray(), StringSplitOptions.None);
                                                this.Controls[0].Controls[0].Controls[0].Controls[int.Parse(partsBtn[2])].Controls[0].Controls[int.Parse(partsBtn[4])].Controls[0].Controls[0].Controls[0].Controls[0].Controls[0].Controls[$"tb_C{int.Parse(partsBtn[2])}_C{int.Parse(partsBtn[4])}_ViewPersonFIO"].Text = _person.Fio;
                                                this.Controls[0].Controls[0].Controls[0].Controls[int.Parse(partsBtn[2])].Controls[0].Controls[int.Parse(partsBtn[4])].Controls[0].Controls[0].Controls[0].Controls[0].Controls[0].Controls[$"tb_C{int.Parse(partsBtn[2])}_C{int.Parse(partsBtn[4])}_ViewPersonGroup"].Text = _person.GroupName;
                                            }
                                            catch (Exception exp)
                                            {
                                                // MessageBox.Show(db.References.textManage.Errors.END_OF_LIST);
                                            }
                                        };

                                        this.Controls[0].Controls[0].Controls[0].Controls[numberChemester].Controls[0].Controls[numberChange].Controls[0].Controls[0].Controls[0].Controls[0].Controls[$"btn_C{numberChemester}_C{numberChange}_PersonNext"].Click += (childSender, childArgs) =>
                                        {
                                            try
                                            {
                                                change.personNavigation.PersonsNext();
                                                db.Elements.Chemester.Change.Person _person = change.personNavigation.PersonsGet();
                                                var partsBtn = ((Button)childSender).Name.Split("_C".ToCharArray(), StringSplitOptions.None);
                                                this.Controls[0].Controls[0].Controls[0].Controls[int.Parse(partsBtn[2])].Controls[0].Controls[int.Parse(partsBtn[4])].Controls[0].Controls[0].Controls[0].Controls[0].Controls[0].Controls[$"tb_C{int.Parse(partsBtn[2])}_C{int.Parse(partsBtn[4])}_ViewPersonFIO"].Text = _person.Fio;
                                                this.Controls[0].Controls[0].Controls[0].Controls[int.Parse(partsBtn[2])].Controls[0].Controls[int.Parse(partsBtn[4])].Controls[0].Controls[0].Controls[0].Controls[0].Controls[0].Controls[$"tb_C{int.Parse(partsBtn[2])}_C{int.Parse(partsBtn[4])}_ViewPersonGroup"].Text = _person.GroupName;
                                            }
                                            catch (Exception exp)
                                            {
                                                // MessageBox.Show(db.References.textManage.Errors.END_OF_LIST);
                                            }
                                        };
                                    }

                                    void Actions()
                                    {
                                        void PersonChange()
                                        {
                                            this.Controls[0].Controls[0].Controls[0].Controls[numberChemester].Controls[0].Controls[numberChange].Controls[0].Controls[0].Controls[0].Controls[0].Controls[0].Controls[$"btn_C{numberChemester}_C{numberChange}_ViewPersonChange"].Click += (childSender, childArgs) =>
                                            {
                                                try
                                                {
                                                    var partsBtn = ((Button)childSender).Name.Split("_C".ToCharArray(), StringSplitOptions.None);
                                                    db.Elements.Chemester.Change.Person _person = new db.Elements.Chemester.Change.Person(this.Controls[0].Controls[0].Controls[0].Controls[int.Parse(partsBtn[2])].Controls[0].Controls[int.Parse(partsBtn[4])].Controls[0].Controls[0].Controls[0].Controls[0].Controls[0].Controls[$"tb_C{int.Parse(partsBtn[2])}_C{int.Parse(partsBtn[4])}_ViewPersonFIO"].Text, this.Controls[0].Controls[0].Controls[0].Controls[int.Parse(partsBtn[2])].Controls[0].Controls[int.Parse(partsBtn[4])].Controls[0].Controls[0].Controls[0].Controls[0].Controls[0].Controls[$"tb_C{int.Parse(partsBtn[2])}_C{int.Parse(partsBtn[4])}_ViewPersonGroup"].Text);
                                                    change.personNavigation.PersonsChange(_person);
                                                    MessageBox.Show(db.References.textManage.Complete.ACTION_VIEW_CHANGE);
                                                }
                                                catch (Exception exp)
                                                {
                                                    MessageBox.Show(db.References.textManage.Errors.LIST_ERROR);
                                                }
                                            };
                                        }

                                        void PersonDelete()
                                        {
                                            this.Controls[0].Controls[0].Controls[0].Controls[numberChemester].Controls[0].Controls[numberChange].Controls[0].Controls[0].Controls[0].Controls[0].Controls[0].Controls[$"btn_C{numberChemester}_C{numberChange}_ViewPersonDelete"].Click += (childSender, childArgs) =>
                                            {
                                                try
                                                {
                                                    var partsBtn = ((Button)childSender).Name.Split("_C".ToCharArray(), StringSplitOptions.None);
                                                    change.personNavigation.PersonsDelete();
                                                    change.personNavigation.PersonsToStart();
                                                    //db.Elements.Chemester.Change.Person _person = change.personNavigation.PersonsGet();
                                                    this.Controls[0].Controls[0].Controls[0].Controls[int.Parse(partsBtn[2])].Controls[0].Controls[int.Parse(partsBtn[4])].Controls[0].Controls[0].Controls[0].Controls[0].Controls[0].Controls[$"tb_C{int.Parse(partsBtn[2])}_C{int.Parse(partsBtn[4])}_ViewPersonFIO"].Text = "";
                                                    this.Controls[0].Controls[0].Controls[0].Controls[int.Parse(partsBtn[2])].Controls[0].Controls[int.Parse(partsBtn[4])].Controls[0].Controls[0].Controls[0].Controls[0].Controls[0].Controls[$"tb_C{int.Parse(partsBtn[2])}_C{int.Parse(partsBtn[4])}_ViewPersonGroup"].Text = "";
                                                    MessageBox.Show(db.References.textManage.Complete.ACTION_VIEW_DELETE);
                                                }
                                                catch (Exception exp)
                                                {
                                                    MessageBox.Show(db.References.textManage.Errors.LIST_ERROR);
                                                }
                                            };
                                        }

                                        void PersonReset()
                                        {
                                            this.Controls[0].Controls[0].Controls[0].Controls[numberChemester].Controls[0].Controls[numberChange].Controls[0].Controls[0].Controls[0].Controls[0].Controls[0].Controls[$"btn_C{numberChemester}_C{numberChange}_ViewPersonReset"].Click += (childSender, childArgs) =>
                                            {
                                                try
                                                {
                                                    db.Elements.Chemester.Change.Person _person = change.personNavigation.PersonsGet();
                                                    var partsBtn = ((Button)childSender).Name.Split("_C".ToCharArray(), StringSplitOptions.None);
                                                    this.Controls[0].Controls[0].Controls[0].Controls[int.Parse(partsBtn[2])].Controls[0].Controls[int.Parse(partsBtn[4])].Controls[0].Controls[0].Controls[0].Controls[0].Controls[0].Controls[$"tb_C{int.Parse(partsBtn[2])}_C{int.Parse(partsBtn[4])}_ViewPersonFIO"].Text = _person.Fio;
                                                    this.Controls[0].Controls[0].Controls[0].Controls[int.Parse(partsBtn[2])].Controls[0].Controls[int.Parse(partsBtn[4])].Controls[0].Controls[0].Controls[0].Controls[0].Controls[0].Controls[$"tb_C{int.Parse(partsBtn[2])}_C{int.Parse(partsBtn[4])}_ViewPersonGroup"].Text = _person.GroupName;
                                                    MessageBox.Show(db.References.textManage.Complete.ACTION_VIEW_RESET);
                                                }
                                                catch (Exception exp)
                                                {
                                                    MessageBox.Show(db.References.textManage.Errors.LIST_ERROR);
                                                }
                                            };
                                        }

                                        void PersonAdd()
                                        {
                                            this.Controls[0].Controls[0].Controls[0].Controls[numberChemester].Controls[0].Controls[numberChange].Controls[0].Controls[0].Controls[0].Controls[1].Controls[$"btn_C{numberChemester}_C{numberChange}_AddPerson"].Click += (childSender, childArgs) =>
                                            {
                                                var partsBtn = ((Button)childSender).Name.Split("_C".ToCharArray(), StringSplitOptions.None);
                                                db.References.elementsManage.LChemesters[int.Parse(partsBtn[2])].LChanges[int.Parse(partsBtn[4])].LPersons.Add(new db.Elements.Chemester.Change.Person(this.Controls[0].Controls[0].Controls[0].Controls[int.Parse(partsBtn[2])].Controls[0].Controls[int.Parse(partsBtn[4])].Controls[0].Controls[0].Controls[0].Controls[1].Controls[1].Controls[$"tb_C{int.Parse(partsBtn[2])}_C{int.Parse(partsBtn[4])}_AddPersonFIO"].Text, this.Controls[0].Controls[0].Controls[0].Controls[int.Parse(partsBtn[2])].Controls[0].Controls[int.Parse(partsBtn[4])].Controls[0].Controls[0].Controls[0].Controls[1].Controls[1].Controls[$"tb_C{int.Parse(partsBtn[2])}_C{int.Parse(partsBtn[4])}_AddPersonGroup"].Text));
                                                change.personNavigation.PersonsToEnd();
                                                db.Elements.Chemester.Change.Person _person = change.personNavigation.PersonsGet();
                                                this.Controls[0].Controls[0].Controls[0].Controls[int.Parse(partsBtn[2])].Controls[0].Controls[int.Parse(partsBtn[4])].Controls[0].Controls[0].Controls[0].Controls[0].Controls[0].Controls[$"tb_C{int.Parse(partsBtn[2])}_C{int.Parse(partsBtn[4])}_ViewPersonFIO"].Text = _person.Fio;
                                                this.Controls[0].Controls[0].Controls[0].Controls[int.Parse(partsBtn[2])].Controls[0].Controls[int.Parse(partsBtn[4])].Controls[0].Controls[0].Controls[0].Controls[0].Controls[0].Controls[$"tb_C{int.Parse(partsBtn[2])}_C{int.Parse(partsBtn[4])}_ViewPersonGroup"].Text = _person.GroupName;
                                                MessageBox.Show(db.References.textManage.Complete.ACTION_VIEW_ADD);
                                            };
                                        }

                                        PersonChange();
                                        PersonDelete();
                                        PersonReset();
                                        PersonAdd();
                                    }
                                    Navigation();
                                    Actions();
                                }

                                void SetGroupClicks()
                                {
                                    void Navigation()
                                    {
                                        this.Controls[0].Controls[0].Controls[0].Controls[numberChemester].Controls[0].Controls[numberChange].Controls[0].Controls[1].Controls[0].Controls[0].Controls[$"btn_C{numberChemester}_C{numberChange}_GroupPrevious"].Click += (childSender, childArgs) =>
                                        {
                                            try
                                            {
                                                change.groupNavigation.GroupsPrevious();
                                                db.Elements.Chemester.Change.Group _group = change.groupNavigation.GroupsGet();
                                                var partsBtn = ((Button)childSender).Name.Split("_C".ToCharArray(), StringSplitOptions.None);
                                                this.Controls[0].Controls[0].Controls[0].Controls[int.Parse(partsBtn[2])].Controls[0].Controls[int.Parse(partsBtn[4])].Controls[0].Controls[1].Controls[0].Controls[0].Controls[0].Controls[$"tb_C{int.Parse(partsBtn[2])}_C{int.Parse(partsBtn[4])}_ViewGroupGroupName"].Text = _group.GroupName;
                                                this.Controls[0].Controls[0].Controls[0].Controls[int.Parse(partsBtn[2])].Controls[0].Controls[int.Parse(partsBtn[4])].Controls[0].Controls[1].Controls[0].Controls[0].Controls[0].Controls[$"tb_C{int.Parse(partsBtn[2])}_C{int.Parse(partsBtn[4])}_ViewGroupHours"].Text = _group.Hours.ToString();
                                            }
                                            catch (Exception exp)
                                            {
                                                MessageBox.Show(db.References.textManage.Errors.END_OF_LIST);
                                            }
                                        };

                                        this.Controls[0].Controls[0].Controls[0].Controls[numberChemester].Controls[0].Controls[numberChange].Controls[0].Controls[1].Controls[0].Controls[0].Controls[$"btn_C{numberChemester}_C{numberChange}_GroupNext"].Click += (childSender, childArgs) =>
                                        {
                                            try
                                            {
                                                change.groupNavigation.GroupsNext();
                                                db.Elements.Chemester.Change.Group _group = change.groupNavigation.GroupsGet();
                                                var partsBtn = ((Button)childSender).Name.Split("_C".ToCharArray(), StringSplitOptions.None);
                                                this.Controls[0].Controls[0].Controls[0].Controls[int.Parse(partsBtn[2])].Controls[0].Controls[int.Parse(partsBtn[4])].Controls[0].Controls[1].Controls[0].Controls[0].Controls[0].Controls[$"tb_C{int.Parse(partsBtn[2])}_C{int.Parse(partsBtn[4])}_ViewGroupGroupName"].Text = _group.GroupName;
                                                this.Controls[0].Controls[0].Controls[0].Controls[int.Parse(partsBtn[2])].Controls[0].Controls[int.Parse(partsBtn[4])].Controls[0].Controls[1].Controls[0].Controls[0].Controls[0].Controls[$"tb_C{int.Parse(partsBtn[2])}_C{int.Parse(partsBtn[4])}_ViewGroupHours"].Text = _group.Hours.ToString();
                                            }
                                            catch (Exception exp)
                                            {
                                                MessageBox.Show(db.References.textManage.Errors.END_OF_LIST);
                                            }
                                        };
                                    }

                                    void Actions()
                                    {
                                        void GroupChange()
                                        {
                                            this.Controls[0].Controls[0].Controls[0].Controls[numberChemester].Controls[0].Controls[numberChange].Controls[0].Controls[1].Controls[0].Controls[0].Controls[0].Controls[$"btn_C{numberChemester}_C{numberChange}_ViewGroupChange"].Click += (childSender, childArgs) =>
                                            {
                                                try
                                                {
                                                    var partsBtn = ((Button)childSender).Name.Split("_C".ToCharArray(), StringSplitOptions.None);
                                                    db.Elements.Chemester.Change.Group _group = new db.Elements.Chemester.Change.Group(this.Controls[0].Controls[0].Controls[0].Controls[int.Parse(partsBtn[2])].Controls[0].Controls[int.Parse(partsBtn[4])].Controls[0].Controls[1].Controls[0].Controls[0].Controls[0].Controls[$"tb_C{int.Parse(partsBtn[2])}_C{int.Parse(partsBtn[4])}_ViewGroupGroupName"].Text, int.Parse(this.Controls[0].Controls[0].Controls[0].Controls[int.Parse(partsBtn[2])].Controls[0].Controls[int.Parse(partsBtn[4])].Controls[0].Controls[1].Controls[0].Controls[0].Controls[0].Controls[$"tb_C{int.Parse(partsBtn[2])}_C{int.Parse(partsBtn[4])}_ViewGroupHours"].Text));
                                                    change.groupNavigation.GroupsChange(_group);
                                                    MessageBox.Show(db.References.textManage.Complete.ACTION_VIEW_CHANGE);
                                                }
                                                catch (Exception exp)
                                                {
                                                    MessageBox.Show(db.References.textManage.Errors.LIST_ERROR);
                                                }
                                            };
                                        }

                                        void GroupDelete()
                                        {
                                            this.Controls[0].Controls[0].Controls[0].Controls[numberChemester].Controls[0].Controls[numberChange].Controls[0].Controls[1].Controls[0].Controls[0].Controls[0].Controls[$"btn_C{numberChemester}_C{numberChange}_ViewGroupDelete"].Click += (childSender, childArgs) =>
                                            {
                                                try
                                                {
                                                    var partsBtn = ((Button)childSender).Name.Split("_C".ToCharArray(), StringSplitOptions.None);
                                                    change.groupNavigation.GroupsDelete();
                                                    change.groupNavigation.GroupsToStart();
                                                    //db.Elements.Chemester.Change.Group _group = change.groupNavigation.GroupsGet();
                                                    this.Controls[0].Controls[0].Controls[0].Controls[int.Parse(partsBtn[2])].Controls[0].Controls[int.Parse(partsBtn[4])].Controls[0].Controls[1].Controls[0].Controls[0].Controls[0].Controls[$"tb_C{int.Parse(partsBtn[2])}_C{int.Parse(partsBtn[4])}_ViewGroupGroupName"].Text = "";
                                                    this.Controls[0].Controls[0].Controls[0].Controls[int.Parse(partsBtn[2])].Controls[0].Controls[int.Parse(partsBtn[4])].Controls[0].Controls[1].Controls[0].Controls[0].Controls[0].Controls[$"tb_C{int.Parse(partsBtn[2])}_C{int.Parse(partsBtn[4])}_ViewGroupHours"].Text = "";
                                                    MessageBox.Show(db.References.textManage.Complete.ACTION_VIEW_DELETE);
                                                }
                                                catch (Exception exp)
                                                {
                                                    MessageBox.Show(db.References.textManage.Errors.LIST_ERROR);
                                                }
                                            };
                                        }

                                        void GroupReset()
                                        {
                                            this.Controls[0].Controls[0].Controls[0].Controls[numberChemester].Controls[0].Controls[numberChange].Controls[0].Controls[1].Controls[0].Controls[0].Controls[0].Controls[$"btn_C{numberChemester}_C{numberChange}_ViewGroupReset"].Click += (childSender, childArgs) =>
                                            {
                                                try
                                                {
                                                    db.Elements.Chemester.Change.Group _group = change.groupNavigation.GroupsGet();
                                                    var partsBtn = ((Button)childSender).Name.Split("_C".ToCharArray(), StringSplitOptions.None);
                                                    this.Controls[0].Controls[0].Controls[0].Controls[int.Parse(partsBtn[2])].Controls[0].Controls[int.Parse(partsBtn[4])].Controls[0].Controls[1].Controls[0].Controls[0].Controls[0].Controls[$"tb_C{int.Parse(partsBtn[2])}_C{int.Parse(partsBtn[4])}_ViewGroupGroupName"].Text = _group.GroupName;
                                                    this.Controls[0].Controls[0].Controls[0].Controls[int.Parse(partsBtn[2])].Controls[0].Controls[int.Parse(partsBtn[4])].Controls[0].Controls[1].Controls[0].Controls[0].Controls[0].Controls[$"tb_C{int.Parse(partsBtn[2])}_C{int.Parse(partsBtn[4])}_ViewGroupHours"].Text = _group.Hours.ToString();
                                                    MessageBox.Show(db.References.textManage.Complete.ACTION_VIEW_RESET);
                                                }
                                                catch (Exception exp)
                                                {
                                                    MessageBox.Show(db.References.textManage.Errors.LIST_ERROR);
                                                }
                                            };
                                        }

                                        void GroupAdd()
                                        {
                                            this.Controls[0].Controls[0].Controls[0].Controls[numberChemester].Controls[0].Controls[numberChange].Controls[0].Controls[1].Controls[0].Controls[1].Controls[$"btn_C{numberChemester}_C{numberChange}_AddGroup"].Click += (childSender, childArgs) =>
                                            {
                                                var partsBtn = ((Button)childSender).Name.Split("_C".ToCharArray(), StringSplitOptions.None);
                                                db.References.elementsManage.LChemesters[int.Parse(partsBtn[2])].LChanges[int.Parse(partsBtn[4])].LGroups.Add(new db.Elements.Chemester.Change.Group(this.Controls[0].Controls[0].Controls[0].Controls[int.Parse(partsBtn[2])].Controls[0].Controls[int.Parse(partsBtn[4])].Controls[0].Controls[1].Controls[0].Controls[1].Controls[1].Controls[$"tb_C{int.Parse(partsBtn[2])}_C{int.Parse(partsBtn[4])}_AddGroupGroupName"].Text, int.Parse(this.Controls[0].Controls[0].Controls[0].Controls[int.Parse(partsBtn[2])].Controls[0].Controls[int.Parse(partsBtn[4])].Controls[0].Controls[1].Controls[0].Controls[1].Controls[1].Controls[$"tb_C{int.Parse(partsBtn[2])}_C{int.Parse(partsBtn[4])}_AddGroupHours"].Text)));

                                                change.groupNavigation.GroupsToEnd();
                                                db.Elements.Chemester.Change.Group _group = change.groupNavigation.GroupsGet();
                                                this.Controls[0].Controls[0].Controls[0].Controls[int.Parse(partsBtn[2])].Controls[0].Controls[int.Parse(partsBtn[4])].Controls[0].Controls[1].Controls[0].Controls[0].Controls[0].Controls[$"tb_C{int.Parse(partsBtn[2])}_C{int.Parse(partsBtn[4])}_ViewGroupGroupName"].Text = _group.GroupName;
                                                this.Controls[0].Controls[0].Controls[0].Controls[int.Parse(partsBtn[2])].Controls[0].Controls[int.Parse(partsBtn[4])].Controls[0].Controls[1].Controls[0].Controls[0].Controls[0].Controls[$"tb_C{int.Parse(partsBtn[2])}_C{int.Parse(partsBtn[4])}_ViewGroupHours"].Text = _group.Hours.ToString();

                                                MessageBox.Show(db.References.textManage.Complete.ACTION_VIEW_ADD);
                                            };
                                        }
                                        GroupChange();
                                        GroupDelete();
                                        GroupReset();
                                        GroupAdd();
                                    }

                                    Navigation();
                                    Actions();
                                }

                                void SetControlListClicks()
                                {
                                    void CreateDoc()
                                    {
                                        this.Controls[0].Controls[0].Controls[0].Controls[numberChemester].Controls[0].Controls[numberChange].Controls[0].Controls[2].Controls[$"btn_C{numberChemester}_C{numberChange}_CreateDoc"].Click += (childSender, childArgs) =>
                                        {
                                            var partsBtn = ((Button)childSender).Name.Split("_C".ToCharArray(), StringSplitOptions.None);

                                            void DocSpace(DocX origDocument, int count)
                                            {
                                                for (int i = 1; i <= count; i++)
                                                {
                                                    origDocument.InsertParagraph("").Font(new Xceed.Document.NET.Font("Times New Roman")).FontSize(12d).SpacingAfter(4d).Alignment = Alignment.left;
                                                }
                                            }
                                            var document = DocX.Create(Path.Combine(Environment.CurrentDirectory, $"\\Output\\Семестр{1 + int.Parse(partsBtn[2])}_Смена{1 + int.Parse(partsBtn[4])}.docx"));
                                            document.InsertParagraph("Приложение").Font(new Xceed.Document.NET.Font("Times New Roman")).FontSize(16d).SpacingAfter(8d).Alignment = Alignment.right;
                                            document.InsertParagraph("К приказу № _______").Font(new Xceed.Document.NET.Font("Times New Roman")).FontSize(16d).SpacingAfter(8d).Alignment = Alignment.right;
                                            document.InsertParagraph($"от «__»‎ ________ {tbGlobalYearStart.Text}").Font(new Xceed.Document.NET.Font("Times New Roman")).FontSize(16d).SpacingAfter(8d).Alignment = Alignment.right;
                                            DocSpace(document, 1);
                                            document.InsertParagraph("Специальные медицинские группы (СМГ)").Font(new Xceed.Document.NET.Font("Times New Roman")).FontSize(16d).SpacingAfter(8d).Alignment = Alignment.center;
                                            document.InsertParagraph($"на {1 + int.Parse(partsBtn[2])} семестр {tbGlobalYearStart.Text}/{tbGlobalYearEnd.Text} учебного года").Font(new Xceed.Document.NET.Font("Times New Roman")).FontSize(16d).SpacingAfter(8d).Alignment = Alignment.center;
                                            DocSpace(document, 1);

                                            //GroupBy
                                            List<int> LUniqueHoursNow = new List<int>();
                                            foreach (var _group in db.References.elementsManage.LChemesters[int.Parse(partsBtn[2])].LChanges[int.Parse(partsBtn[4])].LGroups)
                                            {
                                                bool isHas = false;
                                                foreach (var timeNow in LUniqueHoursNow)
                                                {
                                                    if (_group.Hours == timeNow)
                                                        isHas = true;
                                                }
                                                if (!isHas)
                                                    LUniqueHoursNow.Add(_group.Hours);
                                            }
                                            //End
                                            int groupNumber = 0;
                                            for (int i = 0; i < LUniqueHoursNow.Count(); i++)
                                            {
                                                groupNumber++;
                                                int timeNow = LUniqueHoursNow[i];
                                                document.InsertParagraph($"Группа {groupNumber}     {timeNow} часов").Font(new Xceed.Document.NET.Font("Times New Roman")).FontSize(12d).SpacingAfter(4d).Alignment = Alignment.left;
                                                document.InsertParagraph($"(включая физкультативное занятие)").Font(new Xceed.Document.NET.Font("Times New Roman")).FontSize(12d).SpacingAfter(4d).Alignment = Alignment.left;

                                                int personPosition = 0;
                                                foreach (var itemGroup in db.References.elementsManage.LChemesters[int.Parse(partsBtn[2])].LChanges[int.Parse(partsBtn[4])].LGroups)
                                                {
                                                    if (itemGroup.Hours == timeNow)
                                                    {
                                                        foreach (var itemPerson in db.References.elementsManage.LChemesters[int.Parse(partsBtn[2])].LChanges[int.Parse(partsBtn[4])].LPersons)
                                                        {
                                                            if (itemGroup.GroupName == itemPerson.GroupName)
                                                            {
                                                                ++personPosition;
                                                                if (personPosition == (int.Parse(tbGlobalStudentsLength.Text)+1))
                                                                {
                                                                    DocSpace(document, 2);
                                                                    personPosition = 1;
                                                                    ++groupNumber;
                                                                    document.InsertParagraph($"Группа {groupNumber}     {timeNow} часов").Font(new Xceed.Document.NET.Font("Times New Roman")).FontSize(12d).SpacingAfter(4d).Alignment = Alignment.left;
                                                                    document.InsertParagraph($"(включая физкультативное занятие)").Font(new Xceed.Document.NET.Font("Times New Roman")).FontSize(12d).SpacingAfter(4d).Alignment = Alignment.left;
                                                                    document.InsertParagraph($"{personPosition}. {itemPerson.Fio} {itemPerson.GroupName}").Font(new Xceed.Document.NET.Font("Times New Roman")).FontSize(12d).SpacingAfter(4d).Alignment = Alignment.left;
                                                                }
                                                                else
                                                                    document.InsertParagraph($"{personPosition}. {itemPerson.Fio} {itemPerson.GroupName}").Font(new Xceed.Document.NET.Font("Times New Roman")).FontSize(12d).SpacingAfter(4d).Alignment = Alignment.left;
                                                            }
                                                        }

                                                    }
                                                }
                                                DocSpace(document, 2);
                                            }
                                            MessageBox.Show(db.References.textManage.Complete.ACTION_CREATE_DOC);
                                            document.SaveAs(Environment.CurrentDirectory + $"\\Output\\Семестр{1 + int.Parse(partsBtn[2])}_Смена{1 + int.Parse(partsBtn[4])}.docx");
                                            Process.Start(Environment.CurrentDirectory + $"\\Output\\Семестр{1 + int.Parse(partsBtn[2])}_Смена{1 + int.Parse(partsBtn[4])}.docx");
                                        };
                                    }

                                    void DeleteDataForChange()
                                    {
                                        this.Controls[0].Controls[0].Controls[0].Controls[numberChemester].Controls[0].Controls[numberChange].Controls[0].Controls[2].Controls[$"btn_C{numberChemester}_C{numberChange}_HardReset"].Click += (childSender, childArgs) =>
                                        {
                                            var partsBtn = ((Button)childSender).Name.Split("_C".ToCharArray(), StringSplitOptions.None);
                                            File.CreateText($"Information/Chemester{int.Parse(partsBtn[2])}/Change{int.Parse(partsBtn[4])}/HoursGroups.txt");
                                            File.CreateText($"Information/Chemester{int.Parse(partsBtn[2])}/Change{int.Parse(partsBtn[4])}/Persons.txt");
                                            db.References.elementsManage.LChemesters[int.Parse(partsBtn[2])].LChanges[int.Parse(partsBtn[4])].LGroups = new List<db.Elements.Chemester.Change.Group>();
                                            db.References.elementsManage.LChemesters[int.Parse(partsBtn[2])].LChanges[int.Parse(partsBtn[4])].groupNavigation = new db.Elements.Chemester.Change.GroupNavigation(db.References.elementsManage.LChemesters[int.Parse(partsBtn[2])].LChanges[int.Parse(partsBtn[4])].LGroups);
                                            db.References.elementsManage.LChemesters[int.Parse(partsBtn[2])].LChanges[int.Parse(partsBtn[4])].LPersons = new List<db.Elements.Chemester.Change.Person>();
                                            db.References.elementsManage.LChemesters[int.Parse(partsBtn[2])].LChanges[int.Parse(partsBtn[4])].personNavigation = new db.Elements.Chemester.Change.PersonNavigation(db.References.elementsManage.LChemesters[int.Parse(partsBtn[2])].LChanges[int.Parse(partsBtn[4])].LPersons);

                                            this.Controls[0].Controls[0].Controls[0].Controls[int.Parse(partsBtn[2])].Controls[0].Controls[int.Parse(partsBtn[4])].Controls[0].Controls[0].Controls[0].Controls[0].Controls[0].Controls[$"tb_C{int.Parse(partsBtn[2])}_C{int.Parse(partsBtn[4])}_ViewPersonFIO"].Text = "";
                                            this.Controls[0].Controls[0].Controls[0].Controls[int.Parse(partsBtn[2])].Controls[0].Controls[int.Parse(partsBtn[4])].Controls[0].Controls[0].Controls[0].Controls[0].Controls[0].Controls[$"tb_C{int.Parse(partsBtn[2])}_C{int.Parse(partsBtn[4])}_ViewPersonGroup"].Text = "";

                                            this.Controls[0].Controls[0].Controls[0].Controls[int.Parse(partsBtn[2])].Controls[0].Controls[int.Parse(partsBtn[4])].Controls[0].Controls[1].Controls[0].Controls[0].Controls[0].Controls[$"tb_C{int.Parse(partsBtn[2])}_C{int.Parse(partsBtn[4])}_ViewGroupGroupName"].Text = "";
                                            this.Controls[0].Controls[0].Controls[0].Controls[int.Parse(partsBtn[2])].Controls[0].Controls[int.Parse(partsBtn[4])].Controls[0].Controls[1].Controls[0].Controls[0].Controls[0].Controls[$"tb_C{int.Parse(partsBtn[2])}_C{int.Parse(partsBtn[4])}_ViewGroupHours"].Text = "";
                                            MessageBox.Show(db.References.textManage.Complete.ACTION_DELETE_DATA_FOR_CHANGE);
                                        };
                                    }

                                    void SaveDataForChange()
                                    {
                                        this.Controls[0].Controls[0].Controls[0].Controls[numberChemester].Controls[0].Controls[numberChange].Controls[0].Controls[2].Controls[$"btn_C{numberChemester}_C{numberChange}_SaveAllPersonsAndGroups"].Click += (childSender, childArgs) =>
                                        {
                                            var partsBtn = ((Button)childSender).Name.Split("_C".ToCharArray(), StringSplitOptions.None);
                                            File.Delete($"Information/Chemester{int.Parse(partsBtn[2])}/Change{int.Parse(partsBtn[4])}/HoursGroups.txt");
                                            File.Delete($"Information/Chemester{int.Parse(partsBtn[2])}/Change{int.Parse(partsBtn[4])}/Persons.txt");
                                            File.AppendAllLines($"Information/Chemester{int.Parse(partsBtn[2])}/Change{int.Parse(partsBtn[4])}/HoursGroups.txt", db.References.elementsManage.LChemesters[int.Parse(partsBtn[2])].LChanges[int.Parse(partsBtn[4])].LGroups.Select(x => x.ToTextRow()));
                                            File.AppendAllLines($"Information/Chemester{int.Parse(partsBtn[2])}/Change{int.Parse(partsBtn[4])}/Persons.txt", db.References.elementsManage.LChemesters[int.Parse(partsBtn[2])].LChanges[int.Parse(partsBtn[4])].LPersons.Select(x => x.ToTextRow()));
                                            MessageBox.Show(db.References.textManage.Complete.ACTION_SAVE_ALL_PERSONS_AND_GROUP);
                                        };
                                    }

                                    CreateDoc();
                                    DeleteDataForChange();
                                    SaveDataForChange();
                                }

                                SetPersonClicks();
                                SetGroupClicks();
                                SetControlListClicks();
                            }

                            SetLabels();
                            SetButtonsClicks();
                        }
                    }
                }
            }
        }

        private void btnSaveGlobal_Click(object sender, EventArgs e)
        {
            db.References.globalManage.SaveToFile(int.Parse(tbGlobalYearStart.Text), int.Parse(tbGlobalYearEnd.Text), int.Parse(tbGlobalStudentsLength.Text));
            MessageBox.Show(db.References.textManage.Complete.ACTION_SAVE_GLOBAL);
        }
    }
}
