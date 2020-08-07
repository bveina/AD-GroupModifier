using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;

namespace AD_GroupAdd
{
    public partial class Form1 : Form
    {
        string MyDomain;
        public Form1()
        {
            InitializeComponent();
            MyDomain = "UMDAR.umassd.edu";
            timer1.Interval = 1000;
            
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtUserName.Focus();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            timer1.Enabled = true;
            updateList();
        }


        public void updateList()
        {
            if (comboBox1.Text == "") return;
            PopulateUsers(this.MyDomain, comboBox1.Text);
            
        }
        /*

        //https://stackoverflow.com/questions/7915145/get-all-users-from-a-group-in-active-directory
        public static SearchResultCollection GetListOfAdUsersByGroup(string domainName, string groupName)
        {
            DirectoryEntry entry = new DirectoryEntry("LDAP://UMDAR.umassd.edu");
            DirectorySearcher search = new DirectorySearcher(entry);
            string query = "(&(objectCategory=person)(objectClass=user)(memberOf=\"ii-214-users\"))";
            search.Filter = query;
            search.PropertiesToLoad.Add("memberOf");
            search.PropertiesToLoad.Add("name");

            return search.FindAll();
        }
        */
        //https://stackoverflow.com/questions/2143052/adding-and-removing-users-from-active-directory-groups-in-net#2143742
        public void AddUserToGroup(string domainName, string userId, string groupName)
        {
            try
            {
                using (PrincipalContext pc = new PrincipalContext(ContextType.Domain, domainName))
                {
                    GroupPrincipal group = GroupPrincipal.FindByIdentity(pc, groupName);
                    group.Members.Add(pc, IdentityType.SamAccountName, userId);
                    group.Save();
                }
            }
            catch (UnauthorizedAccessException e)
            {
                MessageBox.Show("Im sorry dave im afaid you cant do that",e.Message);
            }
            catch (System.DirectoryServices.DirectoryServicesCOMException E)
            {
                //doSomething with E.Message.ToString(); 
                MessageBox.Show("Im sorry dave im afaid you cant do that", E.Message);

            }
        }

        public void RemoveUserFromGroup(string domainName, string userId, string groupName)
        {
            if (userId=="")
            {
                MessageBox.Show("BlankUsername");
            }
            try
            {
                using (PrincipalContext pc = new PrincipalContext(ContextType.Domain, domainName))
                {
                    GroupPrincipal group = GroupPrincipal.FindByIdentity(pc, groupName);
                    group.Members.Remove(pc, IdentityType.SamAccountName, userId);
                    group.Save();
                }
            }
            catch (UnauthorizedAccessException e)
            {
                MessageBox.Show("Im sorry dave im afaid you cant do that", e.Message);
            }
            catch (System.DirectoryServices.DirectoryServicesCOMException E)
            {
                //doSomething with E.Message.ToString(); 

            }
        }

        public void PopulateUsers(string domainName, string groupName)
        {
            using (var context = new PrincipalContext(ContextType.Domain, domainName))
            {
                using (var group = GroupPrincipal.FindByIdentity(context, groupName))
                {
                    
                    if (group == null)
                    {
                        MessageBox.Show("Group does not exist");
                        timer1.Enabled = false;
                        return;
                    }
                    else
                    {
                        PrincipalSearchResult<Principal> users = group.GetMembers(false);
                        listBox1.SuspendLayout();
                        int sel = listBox1.SelectedIndex;
                        listBox1.Items.Clear();
                        foreach (Principal v in users)
                        {
                            //if (v.Name.Contains("ECE-")) continue;
                            //if (listBox1.Items.Contains(v)) continue;
                            listBox1.Items.Add(v);
                        }
                        listBox1.SelectedIndex = sel;
                        listBox1.ResumeLayout();

                    }
                }
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (comboBox1.Focused) return;
            updateList();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            if (txtUserName.Text=="")
            {
                MessageBox.Show("Blank username?");
                return;
            }
            AddUserToGroup(MyDomain, txtUserName.Text, comboBox1.Text);
            txtUserName.Text = "";
        }

        private void btnRemove_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex==-1)
            {
                MessageBox.Show("select a user to remove");
                return;
            }
            RemoveUserFromGroup(MyDomain, (listBox1.SelectedItem as Principal).SamAccountName, comboBox1.Text);

        }

        private void comboBox1_KeyDown(object sender, KeyEventArgs e)
        {
            
            if (e.KeyCode == Keys.Enter)
            {
                //todo: check if group exists
            }
        }

        private void txtUserName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (txtUserName.Text == "")
                {
                    MessageBox.Show("Blank username?");
                    return;
                }
                AddUserToGroup(MyDomain, txtUserName.Text, comboBox1.Text);
                txtUserName.Text = "";
            }
        }
    }
}
    