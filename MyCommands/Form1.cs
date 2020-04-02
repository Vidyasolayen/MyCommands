using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Data.SqlClient;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Drawing.Imaging;
using System.Xml;
using System.Management.Automation;

namespace MyCommands
{
    public partial class Form1 : Form
    {
        #region databaseConnection
        SqlConnection con = new SqlConnection("Integrated Security=SSPI;Persist Security Info=False;User ID=wbpoc;Initial Catalog=DFCommands;Data Source=.");
        SqlCommand cmd;
        SqlDataAdapter adapt;
        SqlDataReader reader;
        Dictionary<string, System.Diagnostics.Process> procList;
        MyCommandsDAL mcd = new MyCommandsDAL();
        #endregion
        int count = 0;
        bool alreadyPopulatedcbVersion = false;
        List<int> versionNumbers = null;
        List<string> alreadyOpenedVersions = new List<string>();
        Dictionary<string, Process> processesOpen = new Dictionary<string, Process>();
        //string xmlFilePath = "D:\\commands.xml";
        string xmlFilePath = @"https://mycworks.000webhostapp.com/commands.xml";
        string utilsFolderPath = @"C:\Dayforce\Utils";
        frmAddShortcut frmShortcut;

        #region configurations
        string[] openSolutions = new string[] { "Main.sln", "Dsvc.sln" };
        #endregion
        public Form1()
        {
            InitializeComponent();
            #region shortcutkeys
            int zHotKey = (int)Keys.Z;
            int kill_iis_hotkey = (int)Keys.I;
            int mHotKey = (int)Keys.M;
            int kHotKey = (int)Keys.K;
            int lHotKey = (int)Keys.L;
            int vHotKey = (int)Keys.V;
            int xHotKey = (int)Keys.X;
            int oHotKey = (int)Keys.O;
            int cHotKey = (int)Keys.C;
            int bHotKey = (int)Keys.B;
            int gHotKey = (int)Keys.G;
            int rHotKey = (int)Keys.R;
            int qHotKey = (int)Keys.Q;
            #endregion
            #region registering_shortcuts
            Boolean success = Form1.RegisterHotKey(this.Handle, this.GetType().GetHashCode(), 0x0001, zHotKey);//Set hotkey as 'Alt + z'
            Boolean kill_iis = Form1.RegisterHotKey(this.Handle, this.GetType().GetHashCode(), 0x0001, kHotKey);//Set hotkey as 'Alt + k'
            Boolean moveDllKey = Form1.RegisterHotKey(this.Handle, this.GetType().GetHashCode(), 0x0001, mHotKey);//Set hotkey as 'Alt + m
            Boolean move_iis = Form1.RegisterHotKey(this.Handle, this.GetType().GetHashCode(), 0x0001, kill_iis_hotkey);//Set hotkey as 'Alt + i'
            Boolean launchSolutions = Form1.RegisterHotKey(this.Handle, this.GetType().GetHashCode(), 0x0001, lHotKey);//Set hotkey as 'Alt + l'
            Boolean changeVersion = Form1.RegisterHotKey(this.Handle, this.GetType().GetHashCode(), 0x0001, vHotKey);//Set hotkey as 'Alt + v'
            Boolean killvs = Form1.RegisterHotKey(this.Handle, this.GetType().GetHashCode(), 0x0001, xHotKey);//Set hotkey as 'Alt + x'
            Boolean openLog = Form1.RegisterHotKey(this.Handle, this.GetType().GetHashCode(), 0x0001, oHotKey);//Set hotkey as 'Alt + o'
            Boolean clearLog = Form1.RegisterHotKey(this.Handle, this.GetType().GetHashCode(), 0x0001, cHotKey);//Set hotkey as 'Alt + c'
            Boolean launchBJE = Form1.RegisterHotKey(this.Handle, this.GetType().GetHashCode(), 0x0001, bHotKey);//Set hotkey as 'Alt + b'
            Boolean generateComment = Form1.RegisterHotKey(this.Handle, this.GetType().GetHashCode(), 0x0001, gHotKey);//Set hotkey as 'Alt + g'
            Boolean reviewersList = Form1.RegisterHotKey(this.Handle, this.GetType().GetHashCode(), 0x0001, rHotKey);//Set hotkey as 'Alt + r'
            Boolean branchName = Form1.RegisterHotKey(this.Handle, this.GetType().GetHashCode(), 0x0001, qHotKey);//Set hotkey as 'Alt + q'
            #endregion
        }
        [DllImport("user32.dll")]
        public static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vlc);

        #region ASSIGNING ACTIONS TO SHORTCUTS
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x0312)
            {
                int id = m.LParam.ToInt32();
                bool shouldExecute = true;

                if (!IsFormOpen("frmAddShortcut"))
                {
                    frmShortcut = new frmAddShortcut();
                }
                else
                {
                    shouldExecute = false;
                    frmShortcut.showCodeOnScreen(id.ToString());
                }

                if (shouldExecute)
                {
                    //You can replace this statement with your desired response to the Hotkey.
                    //Alt + z
                    if (id == 5898241)
                    {
                        //checkIfSolutionRunning();
                        sendEmail();
                    }
                    //Alt + i
                    if (id == 4784129)
                    {
                        changeIISVersion();
                    }
                    //Alt + m
                    if (id == 5046273)
                    {
                        moveDLLs();
                    }
                    //Alt + k
                    if (id == 4915201)
                    {
                        btnKill_IISExp_Click(null, null);
                    }
                    if (id == 4980737)
                    {
                        btncbLaunch_Click(null, null);
                    }
                    //Alt + v
                    if (id == 5636097)
                    {
                        changeDFVersion();
                    }
                    if (id == 5767169) // alt + x
                    {
                        closeCurrentVersion();
                    }
                    //Alt + o
                    if (id == 5177345)
                    {
                        btnOpenLog_Click(null, null);
                    }
                    //Alt + c
                    if (id == 4390913)
                    {
                        btnClearLog_Click(null, null);

                    }
                    //Alt + b
                    if (id == 4325377)
                    {
                        //btnLaunchBJE_Click(null,null);
                    }
                    //Alt + g
                    if (id == 4653057)
                    {
                        generateTFSComment();
                    }
                    //Alt + r
                    if (id == 5373953)
                    {
                        //do something
                    }
                    if (id == 5308417)
                    {
                        copyBranchName();
                    }
                    //Console.WriteLine(id);
                }

            }
            //Console.WriteLine(m.Msg);

            base.WndProc(ref m);
        }

        private void moveDLLs()
        {
            string versionPath = mcd.getVersionPath(cbVersion.Text);
            string fileName = @"C:\Temp\tempBat.bat";

            string command = $@"@echo off
 set src_recruitingcommon='{versionPath}\Services\Platform\WBDataSvc\RecruitingCommon\bin\Debug\RecruitingCommon.* '
 set src_businessapi='{versionPath}\Services\Platform\BusinessAPI\bin\Debug\Dayforce.BusinessAPI.*'
 set dst_folder={versionPath}\UI\MyWORKBits\bin\


 xcopy /y %src_recruitingcommon% %dst_folder%
 xcopy /y %src_businessapi% %dst_folder%
 pause
";
            command = command.Replace("'", "\"");

            executeCommand(command, fileName);
        }

        private void executeCommand(string command, string fileName)
        {
            try
            {
                //This will create a new .bat file in the bat directory.
                System.IO.StreamWriter file = new System.IO.StreamWriter(fileName);
                //The code below will write lines to the .bat file.
                //The dir command is used in a Command to list files in the specified directory.
                //The > command in this case sends the list to a text file.
                file.WriteLine(command);
                file.Close();

                //The System.Dignostics.Process Process Class allows a Visual Studio Program
                //to execute another application
                System.Diagnostics.Process.Start(fileName);

                //Process process = new Process();

                //process.StartInfo.RedirectStandardOutput = true;
                //process.StartInfo.RedirectStandardError = true;
                //process.StartInfo.UseShellExecute = false;
                //process.StartInfo.CreateNoWindow = true;
                //process.StartInfo.UseShellExecute = false;

                //var s = process.StandardInput;

                //s.WriteLine();
                //s.WriteLine();
                //s.WriteLine();
                //s.WriteLine();
                //s.WriteLine("\n");

            }
            catch (System.Exception err)
            {
                System.Windows.Forms.MessageBox.Show(err.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void changeIISVersion()
        {
            string fileName = @"C:\Temp\tempBat.bat";

            string folderPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.SystemX86);

            string versionPath = mcd.getVersionPath(cbVersion.Text);

            string strCmdText;
            strCmdText = $@"@echo off

SET DRIVEPATH='{versionPath}'
   
{folderPath}/inetsrv/appcmd.exe set vdir 'AdminService/' -physicalPath:%DRIVEPATH%\bin\_PublishedWebsites\AdminService
{folderPath}/inetsrv/appcmd.exe set vdir 'DataSvc/' -physicalPath:%DRIVEPATH%\bin\_PublishedWebsites\DataSvc
{folderPath}/inetsrv/appcmd.exe set vdir 'MyDayforce/' -physicalPath:%DRIVEPATH%\bin\_PublishedWebsites\MyDayforce
{folderPath}/inetsrv/appcmd.exe set vdir 'HCMAnywhere/' -physicalPath:%DRIVEPATH%\bin\Api
{folderPath}/inetsrv/appcmd.exe set vdir 'CandidatePortal/' -physicalPath:%DRIVEPATH%\bin\_PublishedWebsites\CandidatePortal
::pause
";
            if (!string.IsNullOrEmpty(cbVersion.Text) && int.Parse(cbVersion.Text) < 858)
            {
                strCmdText = $@"@echo off
SET DRIVEPATH='{versionPath}'
   
{folderPath}/inetsrv/appcmd.exe set vdir 'AdminService/' -physicalPath:%DRIVEPATH%\Services\Platform\AdminService
{folderPath}/inetsrv/appcmd.exe set vdir 'DataSvc/' -physicalPath:%DRIVEPATH%\Services\Platform\WBDataSvc\DataSvc
{folderPath}/inetsrv/appcmd.exe set vdir 'MyDayforce/' -physicalPath:%DRIVEPATH%\UI\MyWorkBits
{folderPath}/inetsrv/appcmd.exe set vdir 'CandidatePortal/' -physicalPath:%DRIVEPATH%\UI\CandidatePortal
{folderPath}/inetsrv/appcmd.exe set vdir 'HCMAnywhere/' -physicalPath:%DRIVEPATH%\UI\HCMAnywhere
::pause
";
            }

            strCmdText = strCmdText.Replace("'", "\"");

            executeCommand(strCmdText, fileName);
        }
        #endregion

        private void copyBranchName()
        {
            string tfsItemTitle = Clipboard.GetText();

            string[] strArr = null;

            strArr = tfsItemTitle.Split(':');


            string tfsnum = "";

            string result = "";

            if (strArr.Length > 1)
            {
                for (int count = 0; count <= strArr.Length - 1; count++)
                {
                    if (count == 0)
                    {
                        tfsnum = strArr[count].ToLower().Trim();
                        tfsnum = tfsnum.ToLower().Replace("task", "").Replace("issue", "").Replace("bug", "").Replace("pbi", "").Replace("dev", "").Replace(" ", "");
                    }
                    else
                    {
                        result += strArr[count];
                    }
                }

                string finalResult = result.ToLower().Replace("- ", "-").Replace(" -", "").TrimStart().Replace(' ', '-');

                string teamName = "recruiting/";

                string copyBranchName = $"{teamName}{tfsnum}-{finalResult}";

                Clipboard.SetText(copyBranchName);
            }
        }

        private void populateCbVersion()
        {
            using (SqlConnection conn = new SqlConnection("Integrated Security=SSPI;Persist Security Info=False;User ID=wbpoc;Initial Catalog=DFCommands;Data Source=."))
            {
                try
                {
                    string query = "select versionNumber, versionID from tblVersion";
                    SqlDataAdapter da = new SqlDataAdapter(query, conn);
                    conn.Open();
                    DataSet ds = new DataSet();
                    da.Fill(ds, "tblVersion");
                    cbVersion.DisplayMember = "versionNumber";
                    cbVersion.ValueMember = "versionID";
                    cbVersion.DataSource = ds.Tables["tblVersion"];
                    conn.Close();
                }
                catch (Exception ex)
                {
                    // write exception info to log or anything else
                    MessageBox.Show("Error occured!", ex.Message);
                }
            }
        }

        private void changeDFVersion()
        {
            if (!alreadyPopulatedcbVersion)
            {
                populateCbVersion();
                versionNumbers = mcd.getAllVersionNumbers();
                alreadyPopulatedcbVersion = true;
            }

            if (count == 0)
            {
                cbVersion.SelectedIndex = count;
            }
            else
            {
                cbVersion.SelectedIndex = count;
            }

            count++;

            if (versionNumbers != null && versionNumbers.Count == count)
            {
                count = 0;
            }
            notifyIcon1.BalloonTipText = cbVersion.Text;
            notifyIcon1.BalloonTipIcon = ToolTipIcon.Info;
            notifyIcon1.BalloonTipTitle = "Version Chosen";
            notifyIcon1.ShowBalloonTip(100);
        }

        private void generateTFSComment()
        {
            string taskName = Clipboard.GetText();

            if (taskName != null)
            {
                string release = cbVersion.Text;

                string[] strArr = null;

                strArr = taskName.Split(':');

                string result = "";

                if (strArr.Length > 1)
                {
                    for (int count = 0; count <= strArr.Length - 1; count++)
                    {
                        if (count == 0)
                        {
                            var tfsnum = strArr[count].ToLower().Trim();
                            tfsnum = tfsnum.ToLower().Replace("task", "").Replace("issue", "").Replace("bug", "").Replace("pbi", "").Replace("dev", "");
                            result += "<TFS Number> " + tfsnum + "\n";
                            result += "<Release>" + release + "\n";
                            result += "<Module>" + "Recruiting" + "\n";
                            result += "<Problem or Task>";
                        }
                        else
                        {
                            result += strArr[count];
                        }
                    }

                    result += "\n<Solution>";

                    Console.WriteLine(result);

                    Clipboard.SetText(result);
                }

            }
        }

        private void sendEmail()
        {
            #region saving screenshot
            Graphics graphics;
            Bitmap myScreen = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height, PixelFormat.Format32bppPArgb);
            graphics = Graphics.FromImage(myScreen);
            graphics.CopyFromScreen(Screen.PrimaryScreen.Bounds.X, Screen.PrimaryScreen.Bounds.Y, 0, 0, Screen.PrimaryScreen.Bounds.Size, CopyPixelOperation.SourceCopy);
            string imageName = "MyScreen " + DateTime.Now.ToString();
            imageName = imageName.Replace("/", "_");
            imageName = imageName.Replace(":", "_");
            imageName += ".jpg";
            myScreen.Save(imageName, System.Drawing.Imaging.ImageFormat.Jpeg);
            string logpath = System.IO.Path.GetDirectoryName(Application.ExecutablePath) + @"\" + imageName;
            #endregion

            try
            {
                Outlook.Application oApp = new Outlook.Application();
                Outlook.MailItem mailItem = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
                mailItem.HTMLBody = "Sending Screenshot";
                mailItem.Subject = "Screenshot test";
                mailItem.Attachments.Add(logpath); //logPath is a string holding path to the log.txt file
                Outlook.Recipients oRecips = (Outlook.Recipients)mailItem.Recipients;
                Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add("ashley.mohun@ceridian.com");
                oRecip.Resolve();
                mailItem.Send();
                oRecip = null;
                oRecips = null;
                mailItem = null;
                oApp = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
    }
        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            //Show();
            //this.WindowState = FormWindowState.Normal;
            //notifyIcon1.Visible = false;
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            //        //if the form is minimized  
            //        //hide it from the task bar  
            //        //and show the system tray icon (represented by the NotifyIcon control)  
            //if (this.WindowState == FormWindowState.Minimized)
            //{
            //    Hide();
            //    notifyIcon1.Visible = true;
            //}
        }

        private void btnClearLog_Click(object sender, EventArgs e)
        {
            if (File.Exists(@"C:\Users\gmohun\Desktop\Commands\ClearLog.bat"))
            {
                Process proc = new System.Diagnostics.Process();
                proc.EnableRaisingEvents = false;
                proc.StartInfo.FileName = @"C:\Users\gmohun\Desktop\Commands\ClearLog.bat";
                proc.Start();
            }
        }

        private void btnOpenLog_Click(object sender, EventArgs e)
        {
            // opens the folder in explorer
            Process.Start(@"c:\Log");
        }

        private void btnCloseALLVS_Click(object sender, EventArgs e)
        {
            if (File.Exists(@"C:\Users\gmohun\Desktop\Commands\killvs.bat"))
            {
                Process proc = new System.Diagnostics.Process();
                proc.EnableRaisingEvents = false;
                proc.StartInfo.FileName = @"C:\Users\gmohun\Desktop\Commands\killvs.bat";
                proc.Start();
            }
        }

        private void btnKill_IISExp_Click(object sender, EventArgs e)
        {
            string versionPath = mcd.getVersionPath(cbVersion.Text);
            string fileName = @"C:\Temp\tempBat.bat";

            string command = $@"pskill iisexpress";

            executeCommand(command, fileName);
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Home))
            {
                MessageBox.Show("What the Ctrl+F?");
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void btncbLaunch_Click(object sender, EventArgs e)
        {
            if (cbVersion.Text != "")
            {
                genericLaunchSolution();
            }
        }

        private void closeCurrentVersion()
        {
            var prs = Process.GetProcesses();
            string versionPath = mcd.getVersionPath(cbVersion.Text);
            var mainProcessPath = $@"{versionPath}\Main";
            var dsvcProcessPath = $@"{versionPath}\DataSvc";

            foreach (Process pr in prs)
            {
                if (pr.ProcessName.ToString().ToLower().Contains("devenv"))
                {
                    if (pr.MainWindowTitle.ToLower().Contains(mainProcessPath.ToLower()) || pr.MainWindowTitle.ToLower().Contains(dsvcProcessPath.ToLower()))
                    {
                        pr.Kill();
                    }
                }
            }
        }
        //Generic way of lauching solutions
        private void genericLaunchSolution()
        {
            string versionPath = mcd.getVersionPath(cbVersion.Text);

            bool isSelectedVersionMainOpen = false;
            bool isSelectedVersionDsvcOpen = false;

            var mainProcessPath = $@"{versionPath}\Main";
            var dsvcProcessPath = $@"{versionPath}\DataSvc";

            var prs = Process.GetProcesses();

            foreach (Process pr in prs)
            {
                if (pr.ProcessName.ToString().ToLower().Contains("devenv"))
                {
                    if (pr.MainWindowTitle.ToLower().Contains(mainProcessPath.ToLower()))
                    {
                        isSelectedVersionMainOpen = true;
                    }
                    if (pr.MainWindowTitle.ToLower().Contains(dsvcProcessPath.ToLower()))
                    {
                        isSelectedVersionDsvcOpen = true;
                    }
                }
            }

            if (!isSelectedVersionMainOpen)
            {
                Process procMain = new System.Diagnostics.Process();
                procMain.EnableRaisingEvents = false;
                procMain.StartInfo.FileName = $@"{versionPath}\Main.sln";
                procMain.Start();
            }

            if (!isSelectedVersionDsvcOpen)
            {
                Process procDSvc = new System.Diagnostics.Process();
                procDSvc.EnableRaisingEvents = false;
                procDSvc.StartInfo.FileName = $@"{versionPath}\DataSvc.sln";
                procDSvc.Start();
            }
        }

        private void btnConfig_Click(object sender, EventArgs e)
        {
            Configuration frmConfig = new Configuration();
            frmConfig.Show();
            this.Hide();
        }

        private void btnScreenshot_Click(object sender, EventArgs e)
        {
            sendEmail();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //on form load
            //add listener to xml file
            Timer MyTimer = new Timer();
            MyTimer.Interval = (1 * 20 * 1000); // 45 mins
            MyTimer.Tick += new EventHandler(TimerCallback);
            MyTimer.Start();
        }

        private void TimerCallback(object sender, EventArgs e)
        {
            ////read xml file
            //XmlDocument xmlDocument = new XmlDocument();
            //xmlDocument.Load(xmlFilePath);

            //XmlNodeList commands = xmlDocument.SelectNodes("/CommandList/Command");

            //if (commands.Count > 0)
            //{
            //    foreach (XmlNode command in commands)
            //    {
            //        XmlNode commandText = command.SelectSingleNode("CommandText");
            //        XmlNode commandDate = command.SelectSingleNode("CommandDate");
            //        XmlNode commandTime = command.SelectSingleNode("CommandTime");
            //        XmlNode commandNumber = command.SelectSingleNode("CommandId");

            //        if (!string.IsNullOrEmpty(commandText?.InnerText))
            //        {
            //            if (!mcd.IsCommandAlreadyExecuted(commandNumber?.InnerText))
            //            {
            //                Process.Start(@"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe", $"{commandText.InnerText}");
            //                mcd.SaveExecutedCommand(commandText?.InnerText, commandNumber?.InnerText);
            //            }
            //        }
            //    }
            //}
        }

        private void btnShortcuts_Click(object sender, EventArgs e)
        {
            if (!IsFormOpen("frmAddShortcut")) {
                frmShortcut = new frmAddShortcut();
            }
            this.Hide();
            frmShortcut.Show();
        }

        private bool IsFormOpen(string formName)
        {
            FormCollection listOfForms = Application.OpenForms;

            bool result = false;

            foreach (Form frm in listOfForms)
            {
                if (frm.Name == formName)
                {
                    result = true;
                }
            }
            return result;
        }
    }
}
