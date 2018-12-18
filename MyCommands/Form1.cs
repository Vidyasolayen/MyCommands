using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Speech.Recognition;
using System.Runtime.InteropServices;
using System.Data.SqlClient;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Drawing.Imaging;

namespace MyCommands
{
    public partial class Form1 : Form
    {

        SqlConnection con = new SqlConnection("Integrated Security=SSPI;Persist Security Info=False;User ID=wbpoc;Initial Catalog=DFCommands;Data Source=.");
        SqlCommand cmd;
        SqlDataAdapter adapt;
        SqlDataReader reader;
        Dictionary<string, System.Diagnostics.Process> procList;

        SpeechRecognitionEngine recEngine = new SpeechRecognitionEngine();
        int count = 0;
        #region configurations
        string versionTip = @"C:\Dayforce\SharpTop";
        string version855 = @"D:\Dayforce855\SharpTop";
        string version854 = @"D:\Dayforce854\SharpTop";
        string version853 = @"D:\Dayforce853\SharpTop";

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
            #endregion
        }
        [DllImport("user32.dll")]
        public static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vlc);

        private void lauchSolutions(string version) {
            //launch Main And DataSvc
            Process[] prs = Process.GetProcesses();
            foreach (Process pr in prs)
            {
                if (pr.ProcessName.ToString().ToLower().Contains("devenv"))
                {
                    if (pr.MainWindowTitle == $"{versionTip} - Microsoft Visual Studio *" || pr.MainWindowTitle == "Dayforce\\SharpTop\\Main - Microsoft Visual Studio (Administrator) *")
                    {
                        pr.Kill();
                    }
                    if (pr.MainWindowTitle == "Dayforce\\SharpTop\\DataSvc - Microsoft Visual Studio *" || pr.MainWindowTitle == "Dayforce\\SharpTop\\DataSvc - Microsoft Visual Studio (Administrator) *")
                    {
                        pr.Kill();
                    }
                }
            }

            Process procMain = new System.Diagnostics.Process();
            procMain.EnableRaisingEvents = false;
            procMain.StartInfo.FileName = @"C:\Dayforce\SharpTop\Main.sln";
            procMain.Start();

            Process procDSvc = new System.Diagnostics.Process();
            procDSvc.EnableRaisingEvents = false;
            procDSvc.StartInfo.FileName = @"C:\Dayforce\SharpTop\DataSvc.sln";
            procDSvc.Start();
        }

        private void btnLauchSolutions_Click(object sender, EventArgs e)
        {
            //added comment to test
            var versionNumber = cbVersion.SelectedItem.ToString().ToLower();

            if (versionNumber == "tip")
            {
                versionNumber = "856";
            }

            if (procList != null && !procList.ContainsKey(versionNumber))
            {
                startSolutionsProcess(versionNumber);
            }
            else if(procList == null)
            {
                startSolutionsProcess(versionNumber);
            }
        }

        private void startSolutionsProcess(string versionNumber)
        {
            Process procMain = new System.Diagnostics.Process();
            procMain.EnableRaisingEvents = false;
            procMain.StartInfo.FileName = $@"{getVersionPath(versionNumber)}\Main.sln";
            procMain.Start();

            Process procDSvc = new System.Diagnostics.Process();
            procDSvc.EnableRaisingEvents = false;
            procDSvc.StartInfo.FileName = $@"{getVersionPath(versionNumber)}\DataSvc.sln";
            procDSvc.Start();

            //procList.Add(versionNumber, procMain);
            //procList.Add(versionNumber, procDSvc);
        }
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x0312)
            {
                //s = 5439489
                //You can replace this statement with your desired response to the Hotkey.
                int id = m.LParam.ToInt32();
                //MessageBox.Show(string.Format("Hotkey #{0} pressed", id));
                if (id == 5898241)
                {
                    sendEmail();
                }
                if (id == 4784129)
                {
                    btncbIIS_Click(null, null);
                }
                if (id == 5046273)
                {
                    btncbMoveDLL_Click(null, null);
                }
                if (id == 4915201)
                {
                    btnKill_IISExp_Click(null, null);
                }
                if (id == 4980737)
                {
                    btncbLaunch_Click(null, null);
                }
                if (id == 5636097)
                {
                    if (count == 0)
                    {
                        cbVersion.SelectedIndex = cbVersion.FindStringExact("Tip");
                    }
                    if (count == 1)
                    {
                        cbVersion.SelectedIndex = cbVersion.FindStringExact("855");
                    }
                    if (count == 2)
                    {
                        cbVersion.SelectedIndex = cbVersion.FindStringExact("854");
                    }
                    if (count == 3) {
                        cbVersion.SelectedIndex = cbVersion.FindStringExact("853");
                    }

                    count++;

                    if (count > 3)
                    {
                        count = 0;
                    }
                    notifyIcon1.BalloonTipText = "Version Chosen "+ cbVersion.SelectedItem.ToString() ;
                    notifyIcon1.BalloonTipIcon = ToolTipIcon.Info;
                    notifyIcon1.BalloonTipTitle = "Alert!";
                    notifyIcon1.ShowBalloonTip(500);
                }
                if (id == 5767169) // alt + x
                {
                    btnCloseVS_Click(null, null);
                }

                if (id == 5177345)
                {
                    btnOpenLog_Click(null, null);
                }
                if (id == 4390913)
                {
                    btnClearLog_Click(null, null);

                }
                if (id == 4325377)
                {
                    btnLaunchBJE_Click(null,null);
                }
                if (id == 4653057) {
                    generateTFSComment();
                }
                Console.WriteLine(id);

            }
            Console.WriteLine(m.Msg);

            base.WndProc(ref m);
        }
        //move this to a different class
        private void generateTFSComment()
        {
            string taskName = Clipboard.GetText();

            if (taskName != null)
            {
                string release = "856";

                if (cbVersion.SelectedItem != null && cbVersion.SelectedItem.ToString().ToLower() != "tip") {
                    release = cbVersion.SelectedItem.ToString();
                }

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
                            tfsnum = tfsnum.Replace("task", "").Replace("issue", "").Replace("bug", "").Replace("pbi", "");
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
            string imageName = "MyScreen "+ DateTime.Now.ToString();
            imageName = imageName.Replace("/", "_");
            imageName = imageName.Replace(":", "_");
            imageName += ".jpg";
            myScreen.Save(imageName, System.Drawing.Imaging.ImageFormat.Jpeg);
            string logpath = System.IO.Path.GetDirectoryName(Application.ExecutablePath) + @"\"+imageName;
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


        private void btnLaunch853_Click(object sender, EventArgs e)
        {
            if (File.Exists(@"C:\Users\gmohun\Desktop\Commands\Main853.bat"))
            {
                Process proc = new System.Diagnostics.Process();
                proc.EnableRaisingEvents = false;
                proc.StartInfo.FileName = @"C:\Users\gmohun\Desktop\Commands\Main853.bat";
                proc.Start();
            }
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

        private void btnIISSharptop_Click(object sender, EventArgs e)
        {
            if (File.Exists(@"C:\Users\gmohun\Desktop\Commands\IIS.bat"))
            {
                Process proc = new System.Diagnostics.Process();
                proc.EnableRaisingEvents = false;
                proc.StartInfo.FileName = @"C:\Users\gmohun\Desktop\Commands\IIS.bat";
                proc.Start();
            }
        }

        private void btnIIS853_Click(object sender, EventArgs e)
        {
            if (File.Exists(@"C:\Users\gmohun\Desktop\Commands\IIS853.bat"))
            {
                Process proc = new System.Diagnostics.Process();
                proc.EnableRaisingEvents = false;
                proc.StartInfo.FileName = @"C:\Users\gmohun\Desktop\Commands\IIS853.bat";
                proc.Start();
            }
        }

        private void btnUpgradeSharptop_Click(object sender, EventArgs e)
        {
            if (File.Exists(@"C:\Users\gmohun\Desktop\Commands\upgradeSharptop.bat"))
            {
                Process proc = new System.Diagnostics.Process();
                proc.EnableRaisingEvents = false;
                proc.StartInfo.FileName = @"C:\Users\gmohun\Desktop\Commands\upgradeSharptop.bat";
                proc.Start();
            }
        }

        private void btnUpgrade853_Click(object sender, EventArgs e)
        {
            if (File.Exists(@"C:\Users\gmohun\Desktop\Commands\upgrade853db.bat"))
            {
                Process proc = new System.Diagnostics.Process();
                proc.EnableRaisingEvents = false;
                proc.StartInfo.FileName = @"C:\Users\gmohun\Desktop\Commands\upgrade853db.bat";
                proc.Start();
            }
        }

        private void btnMoveDLLSharptop_Click(object sender, EventArgs e)
        {
            if (File.Exists(@"C:\Users\gmohun\Desktop\Commands\MoveDlls.bat"))
            {
                Process proc = new System.Diagnostics.Process();
                proc.EnableRaisingEvents = false;
                proc.StartInfo.FileName = @"C:\Users\gmohun\Desktop\Commands\MoveDlls.bat";
                proc.Start();
            }
        }

        private void btnMoveDLL855_Click(object sender, EventArgs e)
        {
            if (File.Exists(@"C:\Users\gmohun\Desktop\Commands\MoveDlls855.bat"))
            {
                Process proc = new System.Diagnostics.Process();
                proc.EnableRaisingEvents = false;
                proc.StartInfo.FileName = @"C:\Users\gmohun\Desktop\Commands\MoveDlls855.bat";
                proc.Start();
            }
        }

        private void btnMoveDLL854_Click(object sender, EventArgs e)
        {
            if (File.Exists(@"C:\Users\gmohun\Desktop\Commands\MoveDlls854.bat"))
            {
                Process proc = new System.Diagnostics.Process();
                proc.EnableRaisingEvents = false;
                proc.StartInfo.FileName = @"C:\Users\gmohun\Desktop\Commands\MoveDlls854.bat";
                proc.Start();
            }
        }

        private void btnMoveDLL853_Click(object sender, EventArgs e)
        {
            if (File.Exists(@"C:\Users\gmohun\Desktop\Commands\MoveDlls853.bat"))
            {
                Process proc = new System.Diagnostics.Process();
                proc.EnableRaisingEvents = false;
                proc.StartInfo.FileName = @"C:\Users\gmohun\Desktop\Commands\MoveDlls853.bat";
                proc.Start();
            }
        }

        private void btnOpenLog_Click(object sender, EventArgs e)
        {
            // opens the folder in explorer
            Process.Start(@"c:\Log");
        }

        private void btnCloseVS_Click(object sender, EventArgs e)
        {
            if (cbVersion.SelectedItem.ToString().ToLower() == "tip")
            {
                closevs_sharptop(sender, e);
            }
            if (cbVersion.SelectedItem.ToString().ToLower() == "855")
            {
                closevs_855(sender, e);
            }
            if (cbVersion.SelectedItem.ToString().ToLower() == "854")
            {
                closevs_854(sender, e);
            }
            if (cbVersion.SelectedItem.ToString().ToLower() == "853")
            {
                closevs_853(sender, e);
            }

        }

        private void closevs_sharptop(object sender, EventArgs e)
        {
            Process[] prs = Process.GetProcesses();

            foreach (Process pr in prs)
            {
                if (pr.ProcessName.ToString().ToLower().Contains("devenv"))
                {
                    if (pr.MainWindowTitle == "Dayforce\\SharpTop\\Main - Microsoft Visual Studio *" || pr.MainWindowTitle == "Dayforce\\SharpTop\\Main - Microsoft Visual Studio (Administrator) *")
                    {
                        pr.Kill();
                    }
                    if (pr.MainWindowTitle == "Dayforce\\SharpTop\\DataSvc - Microsoft Visual Studio *" || pr.MainWindowTitle == "Dayforce\\SharpTop\\DataSvc - Microsoft Visual Studio (Administrator) *")
                    {
                        pr.Kill();
                    }
                }
            }

        }

        private void closevs_855(object sender, EventArgs e)
        {
            Process[] prs = Process.GetProcesses();

            foreach (Process pr in prs)
            {
                if (pr.ProcessName.ToString().ToLower().Contains("devenv"))
                {
                    if (pr.MainWindowTitle == "Dayforce855\\SharpTop\\Main - Microsoft Visual Studio *" || pr.MainWindowTitle == "Dayforce855\\SharpTop\\Main - Microsoft Visual Studio (Administrator) *")
                    {
                        pr.Kill();
                    }
                    if (pr.MainWindowTitle == "Dayforce855\\SharpTop\\DataSvc - Microsoft Visual Studio *" || pr.MainWindowTitle == "Dayforce855\\SharpTop\\DataSvc - Microsoft Visual Studio (Administrator) *")
                    {
                        pr.Kill();
                    }
                }
            }

        }

        private void closevs_854(object sender, EventArgs e)
        {
            Process[] prs = Process.GetProcesses();

            foreach (Process pr in prs)
            {
                if (pr.ProcessName.ToString().ToLower().Contains("devenv"))
                {
                    if (pr.MainWindowTitle == "Dayforce854\\SharpTop\\Main - Microsoft Visual Studio *" || pr.MainWindowTitle == "Dayforce854\\SharpTop\\Main - Microsoft Visual Studio (Administrator) *")
                    {
                        pr.Kill();
                    }
                    if (pr.MainWindowTitle == "Dayforce854\\SharpTop\\DataSvc - Microsoft Visual Studio *" || pr.MainWindowTitle == "Dayforce854\\SharpTop\\DataSvc - Microsoft Visual Studio (Administrator) *")
                    {
                        pr.Kill();
                    }
                }
            }

        }
        private void closevs_853(object sender, EventArgs e)
        {
            Process[] prs = Process.GetProcesses();

            foreach (Process pr in prs)
            {
                if (pr.ProcessName.ToString().ToLower().Contains("devenv"))
                {
                    if (pr.MainWindowTitle == "Dayforce853\\SharpTop\\Main - Microsoft Visual Studio *" || pr.MainWindowTitle == "Dayforce853\\SharpTop\\Main - Microsoft Visual Studio (Administrator) *")
                    {
                        pr.Kill();
                    }
                    if (pr.MainWindowTitle == "Dayforce853\\SharpTop\\DataSvc - Microsoft Visual Studio *" || pr.MainWindowTitle == "Dayforce853\\SharpTop\\DataSvc - Microsoft Visual Studio (Administrator) *")
                    {
                        pr.Kill();
                    }
                }
            }

        }

        private void btnLaunchBJE_Click(object sender, EventArgs e)
        {
            if (cbVersion.SelectedItem.ToString().ToLower() == "tip")
            {
                launchBJE("tip");
            }
            if (cbVersion.SelectedItem.ToString().ToLower() == "855")
            {
                launchBJE("855");
            }
            if (cbVersion.SelectedItem.ToString().ToLower() == "854")
            {
                launchBJE("854");
            }
            if (cbVersion.SelectedItem.ToString().ToLower() == "853")
            {
                launchBJE("853");
            }
        }

        private void launchBJE(string version) {
            if (version == "tip") {
                if (File.Exists(@"C:\Users\gmohun\Desktop\Commands\BJE.bat"))
                {
                    Process proc = new System.Diagnostics.Process();
                    proc.EnableRaisingEvents = false;
                    proc.StartInfo.FileName = @"C:\Users\gmohun\Desktop\Commands\BJE.bat";
                    proc.Start();
                }
            }
            if (version == "855")
            {
                if (File.Exists(@"C:\Users\gmohun\Desktop\Commands\BJE855.bat"))
                {
                    Process proc = new System.Diagnostics.Process();
                    proc.EnableRaisingEvents = false;
                    proc.StartInfo.FileName = @"C:\Users\gmohun\Desktop\Commands\BJE855.bat";
                    proc.Start();
                }
            }
            if (version == "854")
            {
                if (File.Exists(@"C:\Users\gmohun\Desktop\Commands\BJE854.bat"))
                {
                    Process proc = new System.Diagnostics.Process();
                    proc.EnableRaisingEvents = false;
                    proc.StartInfo.FileName = @"C:\Users\gmohun\Desktop\Commands\BJE854.bat";
                    proc.Start();
                }
            }
            if (version == "853")
            {
                if (File.Exists(@"C:\Users\gmohun\Desktop\Commands\BJE853.bat"))
                {
                    Process proc = new System.Diagnostics.Process();
                    proc.EnableRaisingEvents = false;
                    proc.StartInfo.FileName = @"C:\Users\gmohun\Desktop\Commands\BJE853.bat";
                    proc.Start();
                }
            }
        }

        private void btnGetLatest_Click(object sender, EventArgs e)
        {
            if (File.Exists(@"C:\Users\gmohun\Desktop\Commands\getLatest.ps1"))
            {
                Process proc = new System.Diagnostics.Process();
                proc.EnableRaisingEvents = false;
                proc.StartInfo.FileName = @"C:\Users\gmohun\Desktop\Commands\getLatest.ps1";
                proc.Start();
            }

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
            if (File.Exists(@"C:\Users\gmohun\Desktop\Commands\killiis.bat"))
            {
                Process proc = new System.Diagnostics.Process();
                proc.EnableRaisingEvents = false;
                proc.StartInfo.FileName = @"C:\Users\gmohun\Desktop\Commands\killiis.bat";
                proc.Start();
            }
        }

        //SPEECH RECOGNITION CODE BELOW
        private void Form1_Load(object sender, EventArgs e)
        {
            this.KeyPreview = true;

            Choices commands = new Choices();

            //put commands here
            //TODO put in database
            commands.Add(new string[] { "say hello", "print my name", "get latest", "latest", "clear log", "show numbers", "hide numbers" });

            GrammarBuilder gBuilder = new GrammarBuilder();
            gBuilder.Append(commands);

            Grammar grammar = new Grammar(gBuilder);

            recEngine.LoadGrammarAsync(grammar);

            recEngine.SetInputToDefaultAudioDevice();

            recEngine.SpeechRecognized += recEngine_SpeechRecognized;

        }

        private void recEngine_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            switch (e.Result.Text.ToString().ToLower())
            {
                case "say hello":
                    MessageBox.Show("Hello Ashley, How are you?");
                    break;
                case "print my name":
                    txtLog.Text += "\nAshley";
                    break;
                case "get latest":
                    txtLog.Text += "\nGetting latest";
                    break;
                case "latest":
                    txtLog.Text += "\nlatest...";
                    break;
                case "clear log":
                    btnClearLog_Click(sender, e);
                    break;
                case "show numbers":
                    //showAllLabels();
                    break;
                case "hide numbers":
                    //hideAllLabels();
                    break;
                default:
                    txtLog.Text += "\nmisunderstood " + e.Result.ToString();
                    break;
            }
        }
        
        private void btnEnable_Click(object sender, EventArgs e)
        {
            recEngine.RecognizeAsync(RecognizeMode.Multiple);
            btnDisable.Enabled = true;
            btnEnable.Enabled = false;
        }

        private void btnDisable_Click(object sender, EventArgs e)
        {
            recEngine.RecognizeAsyncStop();
            btnEnable.Enabled = true;
            btnDisable.Enabled = false;
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

        private void btnLauch854_Click(object sender, EventArgs e)
        {
            if (File.Exists(@"C:\Users\gmohun\Desktop\Commands\MainSoln854.bat"))
            {
                Process proc = new System.Diagnostics.Process();
                proc.EnableRaisingEvents = false;
                proc.StartInfo.FileName = @"C:\Users\gmohun\Desktop\Commands\MainSoln854.bat";
                proc.Start();
            }
        }

        private void btnLauch855_Click(object sender, EventArgs e)
        {
            if (File.Exists(@"C:\Users\gmohun\Desktop\Commands\MainSoln855.bat"))
            {
                Process proc = new System.Diagnostics.Process();
                proc.EnableRaisingEvents = false;
                proc.StartInfo.FileName = @"C:\Users\gmohun\Desktop\Commands\MainSoln855.bat";
                proc.Start();
            }
        }

        private void btnIIS855_Click(object sender, EventArgs e)
        {
            if (File.Exists(@"C:\Users\gmohun\Desktop\Commands\iis855.bat"))
            {
                Process proc = new System.Diagnostics.Process();
                proc.EnableRaisingEvents = false;
                proc.StartInfo.FileName = @"C:\Users\gmohun\Desktop\Commands\iis855.bat";
                proc.Start();
            }

        }

        private void btnIIS854_Click(object sender, EventArgs e)
        {
            if (File.Exists(@"C:\Users\gmohun\Desktop\Commands\iis854.bat"))
            {
                Process proc = new System.Diagnostics.Process();
                proc.EnableRaisingEvents = false;
                proc.StartInfo.FileName = @"C:\Users\gmohun\Desktop\Commands\iis854.bat";
                proc.Start();
            }

        }

        private void cbVersion_SelectedIndexChanged(object sender, EventArgs e)
        {
            //MessageBox.Show(cbVersion.SelectedItem.ToString());
        }

        private void btncbLaunch_Click(object sender, EventArgs e)
        {
            btnLauchSolutions_Click(sender, e);
        }

        private void btncbIIS_Click(object sender, EventArgs e)
        {
            if (cbVersion.SelectedItem.ToString().ToLower() == "tip")
            {
                btnIISSharptop_Click(sender, e);
            }
            if (cbVersion.SelectedItem.ToString().ToLower() == "855")
            {
                btnIIS855_Click(sender, e);
            }
            if (cbVersion.SelectedItem.ToString().ToLower() == "854")
            {
                btnIIS854_Click(sender, e);
            }
            if (cbVersion.SelectedItem.ToString().ToLower() == "853")
            {
                btnIIS853_Click(sender, e);
            }

        }

        private void btncbGetLatest_Click(object sender, EventArgs e)
        {
            if (cbVersion.SelectedItem.ToString().ToLower() == "tip")
            {
                btnGetLatest_Click(sender, e);
            }
            if (cbVersion.SelectedItem.ToString().ToLower() == "854")
            {
                //btnIIS854_Click(sender, e);
            }
            if (cbVersion.SelectedItem.ToString().ToLower() == "853")
            {
                //btnIIS853_Click(sender, e);
            }
        }

        private void btncbMoveDLL_Click(object sender, EventArgs e)
        {
            if (cbVersion.SelectedItem.ToString().ToLower() == "tip")
            {
                btnMoveDLLSharptop_Click(sender, e);
            }
            if (cbVersion.SelectedItem.ToString().ToLower() == "855")
            {
                btnMoveDLL855_Click(sender, e);
            }
            if (cbVersion.SelectedItem.ToString().ToLower() == "854")
            {
                btnMoveDLL854_Click(sender, e);
            }
            if (cbVersion.SelectedItem.ToString().ToLower() == "853")
            {
                btnMoveDLL853_Click(sender, e);
            }

        }

        private void btnConfig_Click(object sender, EventArgs e)
        {
            Configuration frmConfig = new Configuration();
            frmConfig.Show();
            this.Hide();
        }

        private void btnTest_Click(object sender, EventArgs e)
        {
            MessageBox.Show(getVersionPath("856"));
        }

        private string getVersionPath(string versionNumber)
        {
            SqlConnection sqlConnection1 = new SqlConnection("Integrated Security=SSPI;Persist Security Info=False;User ID=wbpoc;Initial Catalog=DFCommands;Data Source=.");
            SqlCommand cmd = new SqlCommand();
            SqlDataReader reader;

            cmd.CommandText = $"SELECT VersionPath FROM tblVersion where versionNumber='{versionNumber}'";
            cmd.CommandType = CommandType.Text;
            cmd.Connection = sqlConnection1;

            sqlConnection1.Open();

            reader = cmd.ExecuteReader();

            string results = "";

            // Data is accessible through the DataReader object here..
            if (reader.Read())
            {
                results = reader["VersionPath"].ToString();
            }

            sqlConnection1.Close();
            return results;
        }

        private void btnScreenshot_Click(object sender, EventArgs e)
        {
            sendEmail();
        }
    }
}
