using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OrgHours.db.Elements.Chemester.Change
{
    class GroupNavigation
    {
        List<Group> LGroups;
        int index;

        public GroupNavigation(List<Group> lGroups)
        {
            LGroups = lGroups;
            index = 0;
        }

        public Group GroupsGet() => LGroups[index];
        public void GroupsNext()
        {
            if (index < LGroups.Count - 1)
            {
                index = index + 1;
            }
            else
            {
                System.Windows.Forms.MessageBox.Show(References.textManage.Errors.END_OF_LIST);
            }
        }

        public void GroupsPrevious()
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
        
        public void GroupsAdd(Group group) => LGroups.Add(group);
        public void GroupsDelete() => LGroups.RemoveAt(index);
        public void GroupsChange(Group group) => LGroups[index]=group;
        public void GroupsToStart() => index = 0;
        public void GroupsToEnd() => index = LGroups.Count()-1;
    }
}
