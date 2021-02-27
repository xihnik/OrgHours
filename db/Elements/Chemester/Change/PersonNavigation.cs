using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OrgHours.db.Elements.Chemester.Change
{
    class PersonNavigation
    {
        List<Person> LPersons;
        int index;

        public PersonNavigation(List<Person> lPersons)
        {
            LPersons = lPersons;
            index = 0;
        }

        public Person PersonsGet() => LPersons[index];
        public void PersonsNext() 
        {
            if(index < LPersons.Count - 1)
            {
                index = index + 1;
            }
            else
            {
                System.Windows.Forms.MessageBox.Show(References.textManage.Errors.END_OF_LIST);
            }
        }
        public void PersonsPrevious()
        {
            if (index > 0)
            {
                index = index - 1;
            }
            else
            {
                System.Windows.Forms.MessageBox.Show(References.textManage.Errors.START_OF_LIST);
            }
        }
        public void PersonsAdd(Person person) => LPersons.Add(person);
        public void PersonsDelete() => LPersons.RemoveAt(index);
        public void PersonsChange(Person person) => LPersons[index] = person;
        public void PersonsToStart() => index = 0;
        public void PersonsToEnd() => index = LPersons.Count()-1;
    }
}
