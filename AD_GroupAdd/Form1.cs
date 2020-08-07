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
        string MyGroupName;
        public Form1()
        {
            InitializeComponent();
            MyDomain = Properties.Settings.Default.defaultDomain;                ;
            timer1.Interval = 1000;
            
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtUserName.Focus();
        }

        private void comboBox1_KeyDown(object sender, KeyEventArgs e)
        {

            if (e.KeyCode == Keys.Enter)
            {
                if (GroupExists(this.MyDomain, comboBox1.Text))
                {
                    this.MyGroupName = comboBox1.Text;
                    txtUserName.Focus();
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                }
            }
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            this.MyGroupName = Properties.Settings.Default.defaultADGroup;
            comboBox1.Text= Properties.Settings.Default.defaultADGroup;
            timer1.Enabled = true;
            updateList();
        }


        public void updateList()
        {
            PopulateUsers();            
        }
       

        //https://stackoverflow.com/questions/2143052/adding-and-removing-users-from-active-directory-groups-in-net#2143742
        private void AddUserToGroup(string domainName, string userId, string groupName)
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
            catch (System.DirectoryServices.AccountManagement.NoMatchingPrincipalException e)
            {
                MessageBox.Show("cant find that user");
            }
            catch (UnauthorizedAccessException e)
            {
                MessageBox.Show(Properties.Resources.ErrorMsg,e.Message);
            }
            catch (System.DirectoryServices.DirectoryServicesCOMException E)
            {
                MessageBox.Show(Properties.Resources.ErrorMsg, E.Message);

            }
        }

        private void RemoveUserFromGroup(string domainName, string userId, string groupName)
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
            catch (System.DirectoryServices.AccountManagement.NoMatchingPrincipalException e)
            {
                MessageBox.Show("cant find that user");
            }
            catch (UnauthorizedAccessException e)
            {
                MessageBox.Show(Properties.Resources.ErrorMsg, e.Message);
            }
            catch (System.DirectoryServices.DirectoryServicesCOMException E)
            {
                MessageBox.Show(Properties.Resources.ErrorMsg, E.Message);
            }
        }

        public void PopulateUsers()
        {
            using (var context = new PrincipalContext(ContextType.Domain, this.MyDomain))
            {
                using (var group = GroupPrincipal.FindByIdentity(context, this.MyGroupName))
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
                        if (sel <= (listBox1.Items.Count-1))
                            listBox1.SelectedIndex = sel;
                        else
                            listBox1.SelectedIndex = listBox1.Items.Count-1;
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
            AddUserToGroup(MyDomain, txtUserName.Text, this.MyGroupName);
            txtUserName.Text = "";
        }

        private void btnRemove_Click(object sender, EventArgs e)
        {
            removeSelected();

        }

        private void removeSelected()
        {
            if (listBox1.SelectedIndex == -1)
            {
                MessageBox.Show("select a user to remove");
                return;
            }
            RemoveUserFromGroup(MyDomain, (listBox1.SelectedItem as Principal).SamAccountName, this.MyGroupName);
        }


        


        private bool GroupExists(string domainName, string groupName)
        {
            using (var context = new PrincipalContext(ContextType.Domain, domainName))
            {
                using (var group = GroupPrincipal.FindByIdentity(context, groupName))
                {

                    if (group == null)
                    {
                        return false;
                    }
                    else
                    {
                        return true;
                    }
                }
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
                e.Handled = true;
                e.SuppressKeyPress = true;

            }
        }

        private void listBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode==Keys.Delete)
            {
                removeSelected();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }
    }
}
    