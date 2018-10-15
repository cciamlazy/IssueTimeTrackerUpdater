using IWshRuntimeLibrary;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace IssueTimeTrackerUpdater
{
    public partial class Updater : Form
    {
        Update _Old;
        Update _Update;
        bool isUpdated = false;
        bool forceUpdate = false;

        public Updater()
        {
            InitializeComponent();
            Program.DataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\IssueTimeTracker";
            Program.webClient = new System.Net.WebClient();

            
            //Load Current version update file
            string curPath = Program.DataPath + "\\Data\\CurrentVersion.json";
            Update currentVersion;
            if (System.IO.File.Exists(curPath))
                currentVersion = Program.GetUpdateFile(curPath);
            else
                currentVersion = Program.GetUpdateFile(curPath.Replace("json", "xml"));

            string oldVersionFile = Program.DataPath + "\\Data\\OldVersion.xml";
            string currentVersionFile = Program.DataPath + "\\Data\\CurrentVersion.xml";
            string newVersionFile = Program.DataPath + "\\Data\\NewVersion.xml";

            if (System.IO.File.Exists(oldVersionFile))
                System.IO.File.Delete(oldVersionFile);
            if (System.IO.File.Exists(currentVersionFile))
                System.IO.File.Delete(currentVersionFile);
            if (System.IO.File.Exists(newVersionFile))
                System.IO.File.Delete(newVersionFile);

            //Download Update File
            if (!System.IO.File.Exists(Program.DataPath + "\\Data\\NewVersion.json"))
            {
                Program.webClient.DownloadFile(Program.BaseSite + Program.getLatestVersion() + "/Data/CurrentVersion.json", Program.DataPath + "\\Data\\NewVersion.json");
            }
            else
            {
                if (Program.isNewer(Program.GetUpdateFile(Program.DataPath + "\\Data\\NewVersion.json").Version, Program.getLatestVersion()))
                    Program.webClient.DownloadFile(Program.BaseSite + Program.getLatestVersion() + "/Data/CurrentVersion.json", Program.DataPath + "\\Data\\NewVersion.json");
            }
            Update update = Program.GetUpdateFile(Program.DataPath + "\\Data\\NewVersion.json");
            /*if (update.RequiredUpdate)
            {
                Process issueTimeTracker = Process.GetProcessesByName("IssueTimeTracker").FirstOrDefault(p => p.MainModule.FileName.StartsWith(Program.DataPath + "\\IssueTimeTracker.exe"));
                bool isRunning = issueTimeTracker != default(Process);
                if (isRunning)
                    issueTimeTracker.Kill();
            }*/
            if (currentVersion != null && !Program.isNewer(currentVersion.Version, update.Version))
                System.IO.File.Delete(Program.DataPath + "\\Data\\NewVersion.json");

            LoadUpdater(currentVersion, update);
        }

        private void LoadUpdater(Update oldFile, Update update)
        {
            _Old = oldFile;
            _Update = update;

            //Load update data into program
            if (update.RequiredUpdate)
            {
                UpdateAvailable.Text = "A Required Update Is Available";
                RemindButton.Enabled = false;
                SkipButton.Enabled = false;
            }
            else
            {
                UpdateAvailable.Text = "An Optional Update Is Available";
                RemindButton.Enabled = true;
                SkipButton.Enabled = true;
            }
            ReleaseNotes.Lines = update.ReleaseNotes.Split('\n');

            
            if (!Program.isNewer((oldFile != null ? oldFile.Version : ""), update.Version))
            {
                UpdateAvailable.Text = "No Update is Available";
                RemindButton.Text = "Force Update";
                UpdateButton.Enabled = false;
                RemindButton.Enabled = true;
                SkipButton.Enabled = false;
            }
        }

        private bool EndProcess(bool kill = false)
        {
            bool killed = false;
            Process[] issueTimeTracker = Process.GetProcessesByName("IssueTimeTracker")/*.FirstOrDefault(p => p.MainModule.FileName.StartsWith(Program.DataPath + "\\IssueTimeTracker.exe"))*/;
            bool[] isRunning = new bool[issueTimeTracker.Length];
            for (int i = 0; i < issueTimeTracker.Length; i++)
            {
                isRunning[i] = issueTimeTracker[i] != default(Process);
                if (isRunning[i])
                {
                    if(!kill)
                        issueTimeTracker[i].CloseMainWindow();
                    else
                        issueTimeTracker[i].Kill();
                    killed = true;
                }
            }
            return killed;
        }

        private void Update(Update oldFile, Update updateFile)
        {
            if (EndProcess())
                Thread.Sleep(500);

            progressBar1.Maximum = updateFile.Files.Count * 2;
            string failedList = "";
            string failedStack = "";
            foreach (DownloadFile file in updateFile.Files)
            {
                string path = Program.DataPath + file.FileName;
                this.Enabled = false;
                string version = "";
                if (oldFile != null)
                    version = FindVersionByFile(oldFile, file.FileName);
                bool download = false;
                if (forceUpdate || version == "" || Program.isNewer(version, file.FileVersion) || !System.IO.File.Exists(path))
                    download = true;
                progressBar1.Value++;
                if(download)
                {
                    FileInfo fi = new FileInfo(path);
                    if (!Directory.Exists(fi.Directory.Parent.FullName))
                        Directory.CreateDirectory(fi.Directory.Parent.FullName);
                    if (!Directory.Exists(fi.Directory.FullName))
                        Directory.CreateDirectory(fi.Directory.FullName);
                    try
                    {
                        if(file.FileName.ToLower().Contains("issuetimetracker.exe") || file.FileName.ToLower().Contains("atlassian.jira.dll") || file.FileName.ToLower().Contains("circularprogressbar.dll"))
                        {
                            EndProcess(true);
                        }
                        Program.webClient.DownloadFile(Program.BaseSite + file.FileVersion + file.DownloadLink, path);
                    }
                    catch (Exception e)
                    {
                        failedList += file.FileName + " - " + e.Message + " - " + e.StackTrace + "\n";
                        failedStack += e.StackTrace + " \n" + e.InnerException + "\n\n\n";
                    }
                }
                progressBar1.Value++;
            }

            if (failedList != "")
            {
                MessageBox.Show("Failed to update\n" + failedList, "Failed Update");
                SendMail("csmith@eccoviasolutions.com", "Updater errors", failedStack);
            }
            string oldVersionFile = Program.DataPath + "\\Data\\OldVersion.json";
            string currentVersionFile = Program.DataPath + "\\Data\\CurrentVersion.json";
            string newVersionFile = Program.DataPath + "\\Data\\NewVersion.json";

            if (System.IO.File.Exists(currentVersionFile))
            {
                if (System.IO.File.Exists(oldVersionFile))
                    System.IO.File.Delete(oldVersionFile);
                System.IO.File.Copy(currentVersionFile, Program.DataPath + "\\Data\\OldVersion.json");
            }

            if (System.IO.File.Exists(Program.DataPath + "\\Data\\NewVersion.json"))
            {
                if (System.IO.File.Exists(currentVersionFile))
                    System.IO.File.Delete(currentVersionFile);
                System.IO.File.Copy(Program.DataPath + "\\Data\\NewVersion.json", currentVersionFile);
                System.IO.File.Delete(Program.DataPath + "\\Data\\NewVersion.json");
            }

            if (updateFile.VerifyUpdate != null && updateFile.VerifyUpdate)
            {
                if(System.IO.File.Exists(oldVersionFile))
                    System.IO.File.Delete(oldVersionFile);
                if (System.IO.File.Exists(Program.DataPath + "\\IssueTimeTracker.exe"))
                    Process.Start(Program.DataPath + "\\IssueTimeTracker.exe");
                this.Close();
            }
            else
            {
                this.Enabled = true;
                isUpdated = true;
                UpdateButton.Text = "Finish";
                UpdateButton.Enabled = true;
            }
        }

        private string FindVersionByFile(Update update, string fileName)
        {
            string version = "";

            if (update != null && update.Files.Count > 0)
            {
                foreach (DownloadFile file in update.Files)
                {
                    if (file.FileName == fileName)
                    {
                        version = file.FileVersion;
                        break;
                    }
                }
            }

            return version;
        }

        private void UpdateButton_Click(object sender, EventArgs e)
        {
            if (!isUpdated)
            {
                UpdateButton.Enabled = false;
                RemindButton.Enabled = false;
                SkipButton.Enabled = false;
                Update(_Old, _Update);

                string commonStartMenuPath = Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu);
                string appStartMenuPath = Path.Combine(commonStartMenuPath, "Programs", "Issue Time Tracker");
                string shortcutLocation = Path.Combine(appStartMenuPath, "IssueTimeTracker" + ".lnk");
                if (!System.IO.File.Exists(shortcutLocation))
                {
                    DialogResult dialogResult = MessageBox.Show("Would you like to create a shortcut to your start menu?", "Start Menu Shortcut", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        try
                        {
                            AddShortcut();
                        }
                        catch
                        {
                            MessageBox.Show("Failed to create shortcut. File Location: " + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\IssueTimeTracker\\IssueTimeTracker.exe", "Failed");
                        }
                    }
                }
                
            }
            else
            {
                if (System.IO.File.Exists(Program.DataPath + "\\IssueTimeTracker.exe"))
                    Process.Start(Program.DataPath + "\\IssueTimeTracker.exe");
                this.Close();
            }
        }

        private static void AddShortcut()
        {
            string pathToExe = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\IssueTimeTracker\\IssueTimeTracker.exe";
            string commonStartMenuPath = Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu);
            string appStartMenuPath = Path.Combine(commonStartMenuPath, "Programs", "Issue Time Tracker");

            if (!Directory.Exists(appStartMenuPath))
                Directory.CreateDirectory(appStartMenuPath);

            string shortcutLocation = Path.Combine(appStartMenuPath, "IssueTimeTracker" + ".lnk");

            if (System.IO.File.Exists(shortcutLocation))
            {
                return;
            }

            WshShell shell = new WshShell();
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutLocation);

            shortcut.Description = "IssueTimeTracker";
            shortcut.IconLocation = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\IssueTimeTracker\\ITT.ico";
            shortcut.TargetPath = pathToExe;
            shortcut.Save();
        }

        private void RemindButton_Click(object sender, EventArgs e)
        {
            if (RemindButton.Text == "Force Update")
            {
                UpdateButton.Enabled = false;
                RemindButton.Enabled = false;
                SkipButton.Enabled = false;
                forceUpdate = true;
                Update(_Old, _Update);
            }
            else
            {
                if (System.IO.File.Exists(Program.DataPath + "\\IssueTimeTracker.exe"))
                    Process.Start(Program.DataPath + "\\IssueTimeTracker.exe");
                this.Close();
            }
        }

        private void SkipButton_Click(object sender, EventArgs e)
        {
            _Update.Skip = true;

            

            var path = Program.DataPath + "\\Data\\NewVersion.json";

            Serializer<Update>.WriteToJSONFile(_Update, path);
            
            this.Close();
        }

        public static void SendMail(string recipient, string subject, string body, string attachmentFilename = "")
        {
            SmtpClient smtpClient = new SmtpClient();
            NetworkCredential basicCredential = new NetworkCredential("issuetimetrackerbugreporter@gmail.com", "Pizza12345");
            MailMessage message = new MailMessage();
            MailAddress fromAddress = new MailAddress("issuetimetrackerbugreporter@gmail.com");

            // setup up the host, increase the timeout to 5 minutes
            smtpClient.Host = "smtp.gmail.com";
            smtpClient.UseDefaultCredentials = false;
            smtpClient.Credentials = basicCredential;
            smtpClient.Timeout = (60 * 5 * 1000);
            smtpClient.EnableSsl = true;

            message.From = fromAddress;
            message.Subject = subject;
            message.IsBodyHtml = false;
            message.Body = body;
            message.To.Add(recipient);

            if (attachmentFilename != null && attachmentFilename != "" && System.IO.File.Exists(attachmentFilename))
            {
                Attachment attachment = new Attachment(attachmentFilename, MediaTypeNames.Application.Octet);
                ContentDisposition disposition = attachment.ContentDisposition;
                disposition.CreationDate = System.IO.File.GetCreationTime(attachmentFilename);
                disposition.ModificationDate = System.IO.File.GetLastWriteTime(attachmentFilename);
                disposition.ReadDate = System.IO.File.GetLastAccessTime(attachmentFilename);
                disposition.FileName = Path.GetFileName(attachmentFilename);
                disposition.Size = new FileInfo(attachmentFilename).Length;
                disposition.DispositionType = DispositionTypeNames.Attachment;
                message.Attachments.Add(attachment);
            }

            smtpClient.Send(message);
        }

        private void Updater_Load(object sender, EventArgs e)
        {

        }
    }
}
