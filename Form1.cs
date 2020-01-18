using System;
using System.Windows.Forms;
using System.Net.NetworkInformation;
using System.Text;
using System.Diagnostics;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Management;
using System.Net;

//Developed by: Byanka - Service Desk Engineer
//Feb 2019

namespace SDTool
{
    public partial class SDToolForm : Form
    {
        public SDToolForm()
        {
            InitializeComponent();
            label4.Text = UserPrincipal.Current.SamAccountName;            
        }
        
        string loginUser = UserPrincipal.Current.SamAccountName;
        string domain = "(AD server name)";
       
        //Viewing User Information
        private void AccountInfoBtn_Click(object sender, EventArgs e)
        {
            if(String.IsNullOrEmpty(UserID.Text))
            {
                MessageBox.Show("Please enter User ID!");
            }
            else
            {
                string username = UserID.Text;
                string loginPwd = textBox1.Text;
                PrincipalContext ctx = new PrincipalContext(ContextType.Domain, domain, loginUser, loginPwd);
                UserPrincipal user = UserPrincipal.FindByIdentity(ctx, IdentityType.SamAccountName, username);

                richTextBox2.Text = "";
                richTextBox3.Text = "";
                richTextBox4.Text = "";
                UserGroupListBox.Items.Clear();

                if (user != null)
                {
                    if (user.Enabled == false)
                    {
                        ViewAccountInfo();
                        richTextBox3.Text = "Account is DISABLED.";
                    }
                    if (user.Enabled == true)
                    {
                        ViewAccountInfo();

                        richTextBox4.Text = "Phone Number: " + UserPhoneNumber() + "\n"
                            + "Email Address: " + UserEmailAddress();

                        richTextBox3.Text = "Account is ACTIVE." + "\n"
                            + AccountLockedOutStatus().ToString() + "\n"
                            + "Account Expiry Date: " + AccountExpiryDate() + "\n" + "\n"
                            + "Password Last Changed: " + LastPasswordChanged() + "\n"
                            + "Password Expiry Date: " + PasswordExpiryDate() + "\n" + "\n"
                            + "Last Logon: " + LastLogon() + "\n"
                            + "Last Incorrect Password: " + LastBadPassword();

                        UserGroupMembership();
                    }
                }
                else
                {
                    MessageBox.Show("User not found");
                }
            }
        }

        //Viewing Account Info
        public void ViewAccountInfo()
        {
            string username = UserID.Text;
            string loginPwd = textBox1.Text;
            DirectoryEntry myLdapConnection = new DirectoryEntry("LDAP://" + domain, loginUser, loginPwd);
            DirectorySearcher search = new DirectorySearcher(myLdapConnection);
            search.Filter = "(sAMAccountName=" + username + ")";
            SearchResult result = search.FindOne();

            if (result.GetDirectoryEntry().Properties["givenName"].Value == null)
            {
                richTextBox2.Text = "Name :" + "n/a" + "\n"
              + "Job Title :" + "n/a" + "\n"
              + "Location :" + "n/a" + "\n"
              + "Description :" + "n/a" + "\n";
            }
            else
            {
                if (result.GetDirectoryEntry().Properties["title"].Value != null && result.GetDirectoryEntry().Properties["physicalDeliveryOfficeName"].Value != null && result.GetDirectoryEntry().Properties["description"].Value != null)
                {
                    richTextBox2.Text = "Name :" + result.GetDirectoryEntry().Properties["givenName"].Value.ToString() + " "
                       + result.GetDirectoryEntry().Properties["sn"].Value.ToString() + "\n"
                   + "Job Title :" + result.GetDirectoryEntry().Properties["title"].Value.ToString() + "\n"
                   + "Location :" + result.GetDirectoryEntry().Properties["physicalDeliveryOfficeName"].Value.ToString() + "\n"
                   + "Description :" + result.GetDirectoryEntry().Properties["description"].Value.ToString();
                }
                if (result.GetDirectoryEntry().Properties["title"].Value == null && result.GetDirectoryEntry().Properties["physicalDeliveryOfficeName"].Value != null && result.GetDirectoryEntry().Properties["description"].Value != null)
                {
                    richTextBox2.Text = "Name :" + result.GetDirectoryEntry().Properties["givenName"].Value.ToString() + " "
                       + result.GetDirectoryEntry().Properties["sn"].Value.ToString() + "\n"
                   + "Job Title :" + "n/a" + "\n"
                   + "Location :" + result.GetDirectoryEntry().Properties["physicalDeliveryOfficeName"].Value.ToString() + "\n"
                   + "Description :" + result.GetDirectoryEntry().Properties["description"].Value.ToString();
                }
                if (result.GetDirectoryEntry().Properties["title"].Value != null && result.GetDirectoryEntry().Properties["physicalDeliveryOfficeName"].Value == null && result.GetDirectoryEntry().Properties["description"].Value != null)
                {
                    richTextBox2.Text = "Name :" + result.GetDirectoryEntry().Properties["givenName"].Value.ToString() + " "
                       + result.GetDirectoryEntry().Properties["sn"].Value.ToString() + "\n"
                   + "Job Title :" + result.GetDirectoryEntry().Properties["title"].Value.ToString() + "\n"
                   + "Location :" + "n/a" + "\n"
                   + "Description :" + result.GetDirectoryEntry().Properties["description"].Value.ToString();
                }
                if (result.GetDirectoryEntry().Properties["title"].Value == null && result.GetDirectoryEntry().Properties["physicalDeliveryOfficeName"].Value == null && result.GetDirectoryEntry().Properties["description"].Value != null)
                {
                    richTextBox2.Text = "Name :" + result.GetDirectoryEntry().Properties["givenName"].Value.ToString() + " "
                       + result.GetDirectoryEntry().Properties["sn"].Value.ToString() + "\n"
                   + "Job Title :" + "n/a" + "\n"
                   + "Location :" + "n/a" + "\n"
                   + "Description :" + result.GetDirectoryEntry().Properties["description"].Value.ToString();
                }
                if (result.GetDirectoryEntry().Properties["title"].Value == null && result.GetDirectoryEntry().Properties["physicalDeliveryOfficeName"].Value == null && result.GetDirectoryEntry().Properties["description"].Value == null)
                {
                    richTextBox2.Text = "Name :" + result.GetDirectoryEntry().Properties["givenName"].Value.ToString() + " "
                       + result.GetDirectoryEntry().Properties["sn"].Value.ToString() + "\n"
                   + "Job Title :" + "n/a" + "\n"
                   + "Location :" + "n/a" + "\n"
                   + "Description :" + "n/a" + "\n";
                }
                if (result.GetDirectoryEntry().Properties["title"].Value == null && result.GetDirectoryEntry().Properties["physicalDeliveryOfficeName"].Value != null && result.GetDirectoryEntry().Properties["description"].Value == null)
                {
                    richTextBox2.Text = "Name :" + result.GetDirectoryEntry().Properties["givenName"].Value.ToString() + " "
                       + result.GetDirectoryEntry().Properties["sn"].Value.ToString() + "\n"
                   + "Job Title :" + "n/a" + "\n"
                   + "Location :" + result.GetDirectoryEntry().Properties["physicalDeliveryOfficeName"].Value.ToString() + "\n"
                   + "Description :" + "n/a" + "\n";
                }
                if (result.GetDirectoryEntry().Properties["title"].Value != null && result.GetDirectoryEntry().Properties["physicalDeliveryOfficeName"].Value == null && result.GetDirectoryEntry().Properties["description"].Value == null)
                {
                    richTextBox2.Text = "Name :" + result.GetDirectoryEntry().Properties["givenName"].Value.ToString() + " "
                       + result.GetDirectoryEntry().Properties["sn"].Value.ToString() + "\n"
                   + "Job Title :" + result.GetDirectoryEntry().Properties["title"].Value.ToString() + "\n"
                   + "Location :" + "n/a" + "\n"
                   + "Description :" + "n/a" + "\n";
                }
                if (result.GetDirectoryEntry().Properties["title"].Value != null && result.GetDirectoryEntry().Properties["physicalDeliveryOfficeName"].Value != null && result.GetDirectoryEntry().Properties["description"].Value == null)
                {
                    richTextBox2.Text = "Name :" + result.GetDirectoryEntry().Properties["givenName"].Value.ToString() + " "
                       + result.GetDirectoryEntry().Properties["sn"].Value.ToString() + "\n"
                   + "Job Title :" + result.GetDirectoryEntry().Properties["title"].Value.ToString() + "\n"
                   + "Location :" + result.GetDirectoryEntry().Properties["physicalDeliveryOfficeName"].Value.ToString() + "\n"
                   + "Description :" + "n/a" + "\n";
                }
            }
        }

        //Viewing user phone number
        public string UserPhoneNumber()
        {
            string username = UserID.Text;
            string loginPwd = textBox1.Text;
            PrincipalContext ctx = new PrincipalContext(ContextType.Domain, domain, loginUser, loginPwd);
            UserPrincipal user = UserPrincipal.FindByIdentity(ctx, IdentityType.SamAccountName, username);

            if (user.VoiceTelephoneNumber != null)
            {
                string phoneNo = user.VoiceTelephoneNumber.ToString();
                return phoneNo;
            }
            else
            {
                string phoneNo = "n/a";
                return phoneNo;
            }
        }

        //Viewing user's email address
        public string UserEmailAddress()
        {
            string username = UserID.Text;
            string loginPwd = textBox1.Text;
            PrincipalContext ctx = new PrincipalContext(ContextType.Domain, domain, loginUser, loginPwd);
            UserPrincipal user = UserPrincipal.FindByIdentity(ctx, IdentityType.SamAccountName, username);

            if (user.EmailAddress != null)
            {
                string emailAddr = user.EmailAddress.ToString();
                return emailAddr;
            }
            else
            {
                string emailAddr = "n/a";
                return emailAddr;
            }
        }

        //Account Locked out status
        public string AccountLockedOutStatus()
        {
            string username = UserID.Text;
            string loginPwd = textBox1.Text;
            PrincipalContext ctx = new PrincipalContext(ContextType.Domain, domain, loginUser, loginPwd);
            UserPrincipal user = UserPrincipal.FindByIdentity(ctx, IdentityType.SamAccountName, username);

            if (user.IsAccountLockedOut())
            {
                string lockedOutStatus = "Account is currently LOCKED OUT";
                return lockedOutStatus;
            }
            else
            {
                string lockedOutStatus = "Account is NOT locked out";
                return lockedOutStatus;
            }
        }

        //Account Expiry Date
        public string AccountExpiryDate()
        {
            string username = UserID.Text;
            string loginPwd = textBox1.Text;
            PrincipalContext ctx = new PrincipalContext(ContextType.Domain, domain, loginUser, loginPwd);
            UserPrincipal user = UserPrincipal.FindByIdentity(ctx, IdentityType.SamAccountName, username);

            if (user.AccountExpirationDate.HasValue)
            {
                DateTime expiration = user.AccountExpirationDate.Value.ToLocalTime();
                string accountExpiryDate = expiration.ToShortDateString();
                return accountExpiryDate;
            }
            else
            {
                string accountExpiryDate = "Never";
                return accountExpiryDate;
            }
        }

        //Password Expiry Date
        public string PasswordExpiryDate()
        {
            string username = UserID.Text;
            string loginPwd = textBox1.Text;
            PrincipalContext ctx = new PrincipalContext(ContextType.Domain, domain, loginUser, loginPwd);
            UserPrincipal user = UserPrincipal.FindByIdentity(ctx, IdentityType.SamAccountName, username);

            if (user.LastPasswordSet.HasValue)
            {
                DateTime exp = user.LastPasswordSet.Value.AddDays(90);
                string pwdExpiryDate = exp.ToShortDateString();
                return pwdExpiryDate;
            }
            else
            {
                string pwdExpiryDate = "n/a";
                return pwdExpiryDate;
            }
        }

        //Last Password Changed
        public string LastPasswordChanged()
        {
            string username = UserID.Text;
            string loginPwd = textBox1.Text;
            PrincipalContext ctx = new PrincipalContext(ContextType.Domain, domain, loginUser, loginPwd);
            UserPrincipal user = UserPrincipal.FindByIdentity(ctx, IdentityType.SamAccountName, username);

            if (user.LastPasswordSet.HasValue)
            {
                DateTime pwdChanged = user.LastPasswordSet.Value;
                return pwdChanged.ToShortDateString();
            }
            else
            {
                string pwdChanged = "n/a";
                return pwdChanged;
            }
        }

        //Last Incorrect Password
        public string LastBadPassword()
        {
            string username = UserID.Text;
            string loginPwd = textBox1.Text;
            PrincipalContext ctx = new PrincipalContext(ContextType.Domain, domain, loginUser, loginPwd);
            UserPrincipal user = UserPrincipal.FindByIdentity(ctx, IdentityType.SamAccountName, username);

            if(user.LastBadPasswordAttempt.HasValue)
            {
                DateTime badPwd = user.LastBadPasswordAttempt.Value;
                return badPwd.ToLocalTime().ToString();
            }
            else
            {
                string badPwd = "n/a";
                return badPwd;
            }
        }

        // Last Logon
        public string LastLogon()
        {
            string username = UserID.Text;
            string loginPwd = textBox1.Text;
            PrincipalContext ctx = new PrincipalContext(ContextType.Domain, domain, loginUser, loginPwd);
            UserPrincipal user = UserPrincipal.FindByIdentity(ctx, IdentityType.SamAccountName, username);

            if (user.LastLogon.HasValue)
            {
                DateTime lastLogon = user.LastLogon.Value;
                return lastLogon.ToLocalTime().ToString();
            }
            else
            {
                string lastLogon = "n/a";
                return lastLogon;
            }
        }

        //User's group membership
        public void UserGroupMembership()
        {
            string username = UserID.Text;
            string loginPwd = textBox1.Text;
            PrincipalContext ctx = new PrincipalContext(ContextType.Domain, domain, loginUser, loginPwd);
            UserPrincipal user = UserPrincipal.FindByIdentity(ctx, IdentityType.SamAccountName, username);

            UserGroupListBox.Items.Clear();
           foreach (var group in user.GetGroups())
            {
                UserGroupListBox.Items.Add(group.Name);
            }
        }

        //Device's group membership
        public void DeviceGroupMembership()
        {
            string deviceName = DeviceAssetNo.Text;
            string loginPwd = textBox1.Text;
            PrincipalContext ctx = new PrincipalContext(ContextType.Domain, domain, loginUser, loginPwd);
            ComputerPrincipal device = ComputerPrincipal.FindByIdentity(ctx, deviceName);

            ComputerGroupListBox.Items.Clear();
            foreach (var group in device.GetGroups())
            {
                ComputerGroupListBox.Items.Add(group.Name);
            }
        }
        
        //Unlock account button
        private void unlockAcctBtn_CLick(object sender, EventArgs e)
        {
            string username = UserID.Text;
            string loginPwd = textBox1.Text;

            if (String.IsNullOrEmpty(UserID.Text))
            {
                MessageBox.Show("Please enter User ID!");
            }
            else
            {
                PrincipalContext ctx = new PrincipalContext(ContextType.Domain, domain, loginUser, loginPwd);
                UserPrincipal user = UserPrincipal.FindByIdentity(ctx, IdentityType.SamAccountName, username);

                user.UnlockAccount();

                MessageBox.Show("Account in unlocked.");

                //Reloading the information
                richTextBox2.Text = "";
                richTextBox3.Text = "";
                richTextBox4.Text = "";
                UserGroupListBox.Items.Clear();

                ViewAccountInfo();

                richTextBox4.Text = "Phone Number: " + UserPhoneNumber() + "\n"
                + "Email Address: " + UserEmailAddress();

                richTextBox3.Text = "Account is ACTIVE." + "\n"
                 + AccountLockedOutStatus().ToString() + "\n"
                 + "Account Expiry Date: " + AccountExpiryDate() + "\n" + "\n"
                 + "Password Last Changed: " + LastPasswordChanged() + "\n"
                 + "Password Expiry Date: " + PasswordExpiryDate() + "\n" + "\n"
                 + "Last Logon: " + LastLogon() + "\n"
                 + "Last Incorrect Password: " + LastBadPassword();

                UserGroupMembership();
            } 
        }

        //Reset password button
        private void resetPwdBtn_Click(object sender, EventArgs e)
        {
            string username = UserID.Text;
            string loginPwd = textBox1.Text;

            if (String.IsNullOrEmpty(UserID.Text))
            {
                MessageBox.Show("Please enter User ID!");
            }
            else
            {
                PrincipalContext ctx = new PrincipalContext(ContextType.Domain, domain, loginUser, loginPwd);
                UserPrincipal user = UserPrincipal.FindByIdentity(ctx, IdentityType.SamAccountName, username);

                DialogResult result = MessageBox.Show("Are you sure you want to reset user's password?", "Confirmation", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    string newPwd = DateTime.Now.DayOfWeek.ToString() + "123456789";

                    user.SetPassword(newPwd);
                    MessageBox.Show("Password has been reset to " + newPwd);
                }
            }
        }

        //Clear button - Device Info
        private void button9_Click(object sender, EventArgs e)
        {
            DeviceAssetNo.Text = "";
            PingStatusTextBox.Text = "";
            PingResultsTextBox.Text = "";
            ComputerGroupListBox.Items.Clear();
        }
  
        //Clear button - User Info
        private void Clearbtn1_Click(object sender, EventArgs e)
        {
            UserID.Text = "";
            richTextBox2.Text = "";
            richTextBox3.Text = "";
            richTextBox4.Text = "";
            UserGroupListBox.Items.Clear();
        }

        //Open Active Directory
        private void openADBtn_Click(object sender, EventArgs e)
        {
           
                ProcessStartInfo psi = new ProcessStartInfo("cmd");
                psi.UseShellExecute = false;
                psi.RedirectStandardInput = true;
                psi.RedirectStandardOutput = true;
                psi.CreateNoWindow = true; var proc = Process.Start(psi);
                proc.StandardInput.WriteLine("dsa.msc");
        }

        //Viewing Device Info
        private void DeviceInfoBtn_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(DeviceAssetNo.Text))
            {
                MessageBox.Show("Please enter Device Asset Number!");
            }
            else
            {
                try
                {
                    PingResultsTextBox.Text = "";
                    ComputerGroupListBox.Items.Clear();

                    Ping pingSender = new Ping();
                    PingOptions options = new PingOptions();
                    options.DontFragment = true;
                    string data = DeviceAssetNo.Text;
                    byte[] buffer = Encoding.ASCII.GetBytes(data);
                    int timeout = 120;
                    PingReply reply = pingSender.Send(data, timeout, buffer, options);
                    if (reply.Status == IPStatus.Success)
                    {
                        PingResultsTextBox.Clear();
                        PingResultsTextBox.Text = "Asset Information: " + "\n"
                            + "Operating System: " + GetDeviceOSInfo() + "\n"
                            + "Manufacturer: " + GetDeviceManufacturerInfo() + "\n"
                            + "Model: " + GetDeviceModelInfo() + "\n"
                            + "Last Logged On User: " + GetDeviceUserInfo();
                    }
                    else
                    {
                        PingResultsTextBox.Text = "Asset Information: " + "\n"
                            + "Cannot retrieve data.  Computer may be offline.  Please use Lansweeper instead.";
                    }
                    DeviceGroupMembership();
                }
                catch (Exception)
                {
                    PingResultsTextBox.Text = "Asset Information: " + "\n"
                        + "Some unknown error. Please use Lansweeper instead.";
                    DeviceGroupMembership();
                }
            }
        }

        //Get Device OS info
        public string GetDeviceOSInfo()
        {
            string deviceName = DeviceAssetNo.Text;
            string loginPwd = textBox1.Text;
            if (deviceName != null)
            {
                ConnectionOptions options = new ConnectionOptions();
                options.Username = loginUser;
                options.Password = loginPwd;
                options.Impersonation = ImpersonationLevel.Impersonate;
                options.EnablePrivileges = true;

                ManagementScope scope = new ManagementScope("\\\\" + deviceName + "\\root\\cimv2", options);
                scope.Connect();

                ObjectQuery query = new ObjectQuery("SELECT * FROM Win32_OperatingSystem");
                ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);

                ManagementObjectCollection queryCollection = searcher.Get();
                foreach (ManagementObject m in queryCollection)
                {
                    string OSInfo = m["Caption"].ToString();
                    return OSInfo;
                }
            }
            return null;
        }

        //Get Device Manufacturer Info
        public string GetDeviceManufacturerInfo()
        {
            string deviceName = DeviceAssetNo.Text;
            string loginPwd = textBox1.Text;
            if (deviceName != null)
            {
                ConnectionOptions options = new ConnectionOptions();
                options.Username = loginUser;
                options.Password = loginPwd;
                options.Impersonation = ImpersonationLevel.Impersonate;
                options.EnablePrivileges = true;

                ManagementScope scope = new ManagementScope("\\\\" + deviceName + "\\root\\cimv2", options);
                scope.Connect();

                ObjectQuery query = new ObjectQuery("SELECT * FROM Win32_ComputerSystem");
                ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);

                ManagementObjectCollection queryCollection = searcher.Get();
                foreach (ManagementObject m in queryCollection)
                {
                    string ManufacturerInfo = m["Manufacturer"].ToString();
                    return ManufacturerInfo;
                }
            }
            return null;
        }

        //Get Device Model Info
        public string GetDeviceModelInfo()
        {
            string deviceName = DeviceAssetNo.Text;
            string loginPwd = textBox1.Text;
            if (deviceName != null)
            {
                ConnectionOptions options = new ConnectionOptions();
                options.Username = loginUser;
                options.Password = loginPwd;
                options.Impersonation = ImpersonationLevel.Impersonate;
                options.EnablePrivileges = true;

                ManagementScope scope = new ManagementScope("\\\\" + deviceName + "\\root\\cimv2", options);
                scope.Connect();

                ObjectQuery query = new ObjectQuery("SELECT * FROM Win32_ComputerSystem");
                ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);

                ManagementObjectCollection queryCollection = searcher.Get();
                foreach (ManagementObject m in queryCollection)
                {
                    string ModelInfo = m["Model"].ToString();
                    return ModelInfo;
                }
            }
            return null;
        }

        //Get Device Last User Info
        public string GetDeviceUserInfo()
        {
            string deviceName = DeviceAssetNo.Text;
            string loginPwd = textBox1.Text;
            if (deviceName != null)
            {
                ConnectionOptions options = new ConnectionOptions();
                options.Username = loginUser;
                options.Password = loginPwd;
                options.Impersonation = ImpersonationLevel.Impersonate;
                options.EnablePrivileges = true;

                ManagementScope scope = new ManagementScope("\\\\" + deviceName + "\\root\\cimv2", options);
                scope.Connect();

                ObjectQuery query = new ObjectQuery("SELECT * FROM Win32_ComputerSystem");
                ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);

                ManagementObjectCollection queryCollection = searcher.Get();
                foreach (ManagementObject m in queryCollection)
                {
                    string UserInfo = m["Username"].ToString();
                    return UserInfo;
                }
            }
            return null;
        }

        //Quick Pinging Status
        private void PingBtn_Click(object sender, EventArgs e)
        {
            Ping pingSender = new Ping();
            PingOptions options = new PingOptions();
            options.DontFragment = true;
            string data = DeviceAssetNo.Text;
            byte[] buffer = Encoding.ASCII.GetBytes(data);
            int timeout = 120;
            try
            {
                PingReply reply = pingSender.Send(data, timeout, buffer, options);
                if (reply.Status == IPStatus.Success)
                {
                    PingStatusTextBox.Text = "Pinging success.  Device is ONLINE. ";
                }
                else
                {
                    PingStatusTextBox.Text = "Pinging failed.  Device maybe OFFLINE.";
                }
            }
            catch (Exception)
            {

            }
        }

        //Output CMD result
        private void PingResultBtn_Click(object sender, EventArgs e)
        {
            PingResultsTextBox.Clear();
            ProcessStartInfo psi = new ProcessStartInfo("cmd");
            psi.UseShellExecute = false;
            psi.RedirectStandardInput = true;
            psi.RedirectStandardOutput = true;
            psi.CreateNoWindow = true;
            var proc = Process.Start(psi);
            proc.StandardInput.WriteLine("ping " + DeviceAssetNo.Text);
            proc.StandardInput.WriteLine("exit");
            string s = proc.StandardOutput.ReadToEnd();
            PingResultsTextBox.Text = s;
        }

        //Open Remote Control Viewer
        private void openRCVBtn_Click(object sender, EventArgs e)
        {
                Process.Start(@"C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\i386\CmRcViewer.exe");
        }

        //Open MSRA
        private void openMSRABtn_Click(object sender, EventArgs e)
        {
                  Process.Start(@"C:\Windows\System32\msra.exe");
        }

        //Clear button
        private void button6_Click(object sender, EventArgs e)
        {
            richTextBox1.Text = "";
        }

        //Copy button - Notes
        private void button5_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(richTextBox1.Text))
            {
                MessageBox.Show("Field cannot be empty!");
            }
            else
            {
                Clipboard.SetText(richTextBox1.Text);
            }
        }

        //Copy button - user ID
        private void button2_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(UserID.Text))
            {
                MessageBox.Show("Field cannot be empty!");
            }
            else
            {
                Clipboard.SetText(UserID.Text);
            }
        }

        //Copy button - Device
        private void button4_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(DeviceAssetNo.Text))
            {
                MessageBox.Show("Field cannot be empty!");
            }
            else
            {
                Clipboard.SetText(DeviceAssetNo.Text);
            }
        }

        //Open more Service Desk Links
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start(@"\\perdata\Public\_Service Desk Team\HZP Service Desk Notes\HZP Service Desk notes.htm");
        }

        //Open LANSWEEPER
        private void button7_Click(object sender, EventArgs e)
        {

                Process.Start(@"http://ea:85/");     
        }

        //Tooltip
        private void SDToolForm_Load(object sender, EventArgs e)
        {
            toolTip1.SetToolTip(resetPwdBtn, "Reset password to [today's day]123456789");
            toolTip1.SetToolTip(AccountInfoBtn, "Get user account information");
            toolTip1.SetToolTip(unlockAcctBtn, "Unlock the account");
            toolTip1.SetToolTip(openADBtn, "Open Active Directory");
            toolTip1.SetToolTip(DeviceInfoBtn, "Get Device Information");
            toolTip1.SetToolTip(openRCVBtn, "Open Remote Control Viewer");
            toolTip1.SetToolTip(openMSRABtn, "Open MSRA");
            toolTip1.SetToolTip(FullNameTextBox, "Enter full or partial name of the user");
            toolTip1.SetToolTip(UserSearchBtn, "Search User ID related to the name");
            toolTip1.SetToolTip(UserIDListBox, "List of User ID related to the name");
            toolTip1.SetToolTip(GetAccountInfoBtn2, "Get user account information");
            toolTip1.SetToolTip(DeviceAssetNo, "Enter Device Asset Number");
        }

        //Search User ID by Name
        private void UserSearchBtn_Click(object sender, EventArgs e)
        {
            string username = FullNameTextBox.Text;
            string loginPwd = textBox1.Text;

            if (String.IsNullOrEmpty(FullNameTextBox.Text))
            {
                MessageBox.Show("Please enter the name!");
            }
            else
            {
                PrincipalContext ctx = new PrincipalContext(ContextType.Domain, domain, loginUser, loginPwd);
                UserPrincipal user = new UserPrincipal(ctx);
                user.DisplayName = "*" + username + "*";
                PrincipalSearcher srch = new PrincipalSearcher(user);

                UserIDListBox.Items.Clear();
                foreach (var found in srch.FindAll())
                {
                    if (found != null)
                    {
                        UserIDListBox.Items.Add(found.SamAccountName);
                    }
                    else
                    {
                        MessageBox.Show("User not found.");
                    }
                }
            }
        }

        //Get Account Info from Selected User ID
        private void GetAccountInfoBtn2_Click(object sender, EventArgs e)
        {
            if (UserIDListBox.SelectedIndex == -1)
            {
                MessageBox.Show("Please select a User ID.");
            }
            else
            {
                string curItem = UserIDListBox.SelectedItem.ToString();
                UserID.Text = curItem;

                string username = UserID.Text;
                string loginPwd = textBox1.Text;
                PrincipalContext ctx = new PrincipalContext(ContextType.Domain, domain, loginUser, loginPwd);
                UserPrincipal user = UserPrincipal.FindByIdentity(ctx, IdentityType.SamAccountName, username);

                richTextBox2.Text = "";
                richTextBox3.Text = "";
                richTextBox4.Text = "";
                UserGroupListBox.Items.Clear();

                if (user != null)
                {
                    if (user.Enabled == false)
                    {
                        ViewAccountInfo();
                        richTextBox3.Text = "Account is DISABLED.";
                    }
                    if (user.Enabled == true)
                    {
                        ViewAccountInfo();

                        richTextBox4.Text = "Phone Number: " + UserPhoneNumber() + "\n"
                            + "Email Address: " + UserEmailAddress();

                        richTextBox3.Text = "Account is ACTIVE." + "\n"
                            + AccountLockedOutStatus().ToString() + "\n"
                            + "Account Expiry Date: " + AccountExpiryDate() + "\n" + "\n"
                            + "Password Last Changed: " + LastPasswordChanged() + "\n"
                            + "Password Expiry Date: " + PasswordExpiryDate() + "\n" + "\n"
                            + "Last Logon: " + LastLogon() + "\n"
                            + "Last Incorrect Password: " + LastBadPassword();

                        UserGroupMembership();
                    }
                }
                else
                {
                    MessageBox.Show("User not found");
                }
            }
        }

        //Clear User ID Listbox
        private void ClearBtn3_Click(object sender, EventArgs e)
        {
            FullNameTextBox.Text = "";
            UserIDListBox.Items.Clear();
        }

        //Copy User Name
        private void button1_Click(object sender, EventArgs e)
        {
            if(String.IsNullOrEmpty(FullNameTextBox.Text))
            {
                MessageBox.Show("Field cannot be empty!");
            }
            else
            {
                Clipboard.SetText(FullNameTextBox.Text);
            }
        }

        //About
        private void AboutBtn_Click(object sender, EventArgs e)
        {
            AboutBox1 about = new AboutBox1();
            about.Show();
        }
    }

}




