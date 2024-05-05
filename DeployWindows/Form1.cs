using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Threading;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Header;
using CustomConfig;
using System.Net;
using System.IO.Compression;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
namespace DeployWindows
{
    public partial class Clean : Form
    {
        private Rectangle lab1, lab2, lab3, lab4, lb5, lab6, prog1, prog2, bttn1, lstbx1;
        char cLetter, tLetter;
        private Size form;
        private int totalProg, index;
        private bool isoYes = false;
        private bool selectedin = false;
        private string esdNo;
        private string esdYes;
        private string[] argss = Environment.GetCommandLineArgs();
        private string wimLoc;
        private string ApplyDir;
        private string ISOLoc, folder;
        private string ISOExtract;
        public bool isExpress, invis;
        public string topass, disks;
        // input should be like this : >DeployWindows.exe topass='--ESDMode=True --WinVer=Windows_10 --Release=10,_version_22H2_[19045.2965]_(Updated_May_2023) --Language=English' disks='0' isExpress='False'
        private void setArgs() // this method will take the args that is needed for this program to work
        {
            StringBuilder inp = new StringBuilder();
            try
            {
                foreach (string word in argss) // spaces out the command line arguments
                {
                    inp.Append(word);
                    inp.Append(" ");
                }
            } 
            
            catch
            {

            }
            string args = inp.ToString(); // save to string
            // input 
            string pattern = @"(\w+)='([^']+)'"; // this will split up the args into their corresponding variables so name='John' age='30' city='New York'


            MatchCollection matches = Regex.Matches(args, pattern); // make the matches

            foreach (Match match in matches)
            {
                string attribute = match.Groups[1].Value;
                string value = match.Groups[2].Value;

                if (attribute == "topass")
                {
                    topass = value; // sets the toPass value which will be given to MSWISO
                }
                else if (attribute == "disks")
                {
                    disks = value; // this is so we know what drive to target
                }
                else if (attribute == "isExpress")
                {
                    if (value == "False" | value == "false")
                    {
                        isExpress = false; // we need to know if its an Express install, so that we don't add junk
                    }
                    else if (value == "True" | value == "true") {

                        isExpress = true;

                    }
                }
            }
            if (matches.Count == 0)
            {
                MessageBox.Show("Error, restart program.");
            }
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        public bool checkIfExist(string path) // simple method to check if a file exists
        {
            if (File.Exists(path))
            {
                return true;
            } else
            {
                return false;
            }
        }
        private async void Clean_Load(object sender, EventArgs e)
        {
           DriveLetters letters = new DriveLetters(Environment.SystemDirectory + "\\driveLetters.txt"); // sets the drive letters from the previous setup when the user was asked to input new letters when there was a clash
            cLetter = letters.CLetter; // set the new C letter
            tLetter = letters.TLetter; // set the new T letter
            esdNo = @tLetter.ToString() +":\\contin\\test.esd"; // from here
esdYes = @tLetter.ToString() +":\\contin\\sources\\install.esd";
            esdYes = @tLetter.ToString() +":\\contin\\sources\\install.esd";
            wimLoc = @tLetter.ToString() +":\\contin\\sources\\install.wim";
            ApplyDir = @cLetter.ToString() +":\\";
            ISOLoc = @tLetter.ToString() +":\\contin\test.iso";
            ISOExtract = @tLetter.ToString() + ":\\contin\\";
            folder = @tLetter.ToString() + ":\\contin\\"; // to here are the fixed file locations
            setArgs(); // get the info that was passed on
            label3.Text = "Current Task: Downloading | Current Progress: estimating";
            label4.Text = "Do not panic if frozen! WinPE is not optimized for high loads!";
           await Task.Run(() => Downld()); // Download the ESD image using MSWISO
            label3.Text = "Current Task: Installing | Current Progress: estimating";
            progressBar2.Value = 0;
            label4.Text = "Estimating progress..."; // from here
            label5.Visible = false;
            label6.Visible = false;
            listBox1.Visible = false;
            button1.Visible = false; // to here we hide most of the UI like the text and progress bar so that only the dialog for the Windows version is shown
            await Task.Run(() => DISMDiag()); // start the process of showing the Windows versions avaliable
           
        }

        private async Task Downld() // user algo
        {
            
            ProcessStartInfo startInfo = new ProcessStartInfo // from here
            {
                FileName = folder + "\\MSWISO\\MSWISO.exe",
                Arguments = topass,
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true
            }; // to here starts MSWISO with the toPass arguments, without creating a window and redirecting the output

            using (Process process = new Process { StartInfo = startInfo })
            {
                process.Start();

                while (!process.StandardOutput.EndOfStream)
                {
                    string WhtOut = process.StandardOutput.ReadLine(); // start reading the output
                    int giveProgress;
                    string newOut = WhtOut.TrimEnd('%'); // remove the % as we only need an int
                    if (int.TryParse(newOut, out giveProgress))
                    {

                        label4.BeginInvoke(new Action(() => label4.Text = "Current progress: " + WhtOut));
                        progressBar2.BeginInvoke(new Action(() => progressBar2.Value = giveProgress));
                        totalProg = Math.Min(giveProgress / 2, 50); // As tis is task 1 of 2, we do divide by 2
                        label3.BeginInvoke(new Action(() => label3.Text = "Current Task: Downloading | Total Progress: " + totalProg + "%"));
                        progressBar1.BeginInvoke(new Action(() => progressBar1.Value = totalProg));
                    }
                    else if (WhtOut.StartsWith("No")) // sometimes there is a connection time out or the version no longer exist, so we show the user the error message
                    {
                        MessageBox.Show("It looks like Microsoft no longer hosts this version... Try to not use legacy mode if possible or a different version. Press OK to restart.");
                        Process.Start("cmd.exe", "/C wpeutil reboot");
                    }
                }

                process.WaitForExit();
                process.Close();
            }
        }
        private async Task DISMDiag() // user algo
        {
            invis = false;
            MakeInvis(invis); // make the things we don't want to show invisible
            label5.Invoke(new Action(() => label5.Visible = true)); // from here
            listBox1.Invoke(new Action(() => listBox1.Visible = true));
            button1.Invoke(new Action(() => button1.Visible = true)); // to here we show the dialog
            string arguments = "n/a";
            string exePath = Environment.GetEnvironmentVariable("SystemRoot") + @"\System32" + @"\dism.exe"; // we need to thing the path of DISM.exe
            if (isoYes) // if we have an iso then use the location for the ISO
            {
                try
                {
                    arguments = @"/get-wiminfo /wimfile:" + wimLoc;
                }
                catch (FileNotFoundException)
                {
                    arguments = @"/get-wiminfo /wimfile:" + esdYes;
                }
            }
            else // if we have an ESD then use the location for the ESD
            {
                if (File.Exists(esdNo))
                {
                    arguments = @"/get-wiminfo /wimfile:" + esdNo;
                }
               
            }
            using (Process process = new Process())
            {
                process.StartInfo.FileName = exePath; // dism
                process.StartInfo.Arguments = arguments; // gets the wim info with the location of ESD or WIM
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.CreateNoWindow = true;
                process.StartInfo.RedirectStandardOutput = true;

                process.OutputDataReceived += (sender, e) =>
                {
                    if (e.Data != null)
                    {
                        string newOut = e.Data.ToString();
                        string pattern = @"^Name :\s+(.+)$"; // we only want the names of the avaliable versions and not junk like dism version

                        Match match = Regex.Match(newOut, pattern);
                        if (match.Success)
                        {
                            string name = match.Groups[1].Value;

                            listBox1.Invoke(new Action(() =>
                            {
                                listBox1.Items.Add(name); // adds the matches for the windows versions
                            }));
                        }
                    }
                };



                await Task.Run(() =>
                {
                    do
                    {
                        Console.WriteLine("User still in progress"); // this is an intentional loop because we want to install Windows once we have all the data ready. Some people, like me, have very fast internet and we have not finishes setting up our apps or drivers yet. Installing windows without this data leads to corruption.
                        System.Threading.Thread.Sleep(3000); 
                    }
                    while (!File.Exists(Environment.SystemDirectory + "\\done.txt")); // the setup will make a txt file as a flag to show that the install was done and you can continue.
                    timeToDownload(); // starts countdown
                    process.Start();
                    process.BeginOutputReadLine();
                    process.WaitForExit();
                    
                });
            }


        }
        private async void timeToDownload() // once the setup has determined that the previous setup is done. We can start the AFK countdown. //simple user algo
        {
            await Task.Run(() =>
            {
                int time = 0;
                int timeLeft = 30; // we give the user 30 seconds to respond

                label6.BeginInvoke(new Action(() => label6.Visible = true));
                label6.BeginInvoke(new Action(() => label6.Text = tLetter.ToString() +"EST"));

                Thread countdownThread = new Thread(() =>
                {
                    while (timeLeft > 0) // keep doing this while the time is still running
                    {
                        if (!selectedin) // if the user hasn't selected a version yet
                        {
                            if (time < 30)
                            {
                                timeLeft--; // reduce the amount of time left
                                time++;
                                label6.BeginInvoke(new Action(() =>
                                {
                                    label6.Text = "You have " + timeLeft.ToString() + " seconds to choose or index " + (listBox1.Items.Count).ToString() + " is chosen."; // show the user how much time is left
                                }));

                                Thread.Sleep(1000); // wait one second
                            }
                        } else if(selectedin) // the user has selected a version which means they are present, stop the countdown
                        {
                            label6.BeginInvoke(new Action(() => label6.Visible = false));
                            break;
                        }
                    }
                    if (timeLeft == 0) // time is up, user has been AFK for too long
                    {
                        label6.BeginInvoke(new Action(() => label6.Visible = false)); // hide the warning
                        index = listBox1.Items.Count; // we select the last version in the list as its guaranteed not to be an installer 
                        button1_Click(null , null); // we simulate user input to click the continue button

                    }
                });
                countdownThread.Start();
            });
        }



        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            index = listBox1.SelectedIndex + 1; // selects the version of windows the user wants

        }

        private async void button1_Click(object sender, EventArgs e)
        {
            invis = true;
            selectedin = true; // stop the countdown
            MakeInvis(invis); // restore the UI
            label5.BeginInvoke(new Action(() =>label5.Visible = false)); // from here
            listBox1.BeginInvoke(new Action(() => listBox1.Visible = false));
            button1.BeginInvoke(new Action(() => button1.Visible = false)); // to here hide the DISMdiag
            await Task.Run(() => Install()); // start the install
        }
        

        private async Task Install()//simple user algo
        {
            bool ignoreFirst = true;
            progressBar2.BeginInvoke(new Action(() => progressBar2.Value = 0)); // clear any previous work
            string arguments;
            string exePath = Environment.GetEnvironmentVariable("SystemRoot") + @"\System32" + @"\dism.exe"; // get location of DISM
            if (isoYes) // the location if we have an ISO
            {
                MessageBox.Show(isoYes.ToString());
                try
                {
                    arguments = @"/Apply-Image /Imagefile:" + wimLoc + " /Index:" + index + @" /ApplyDir:" + ApplyDir;
                }
                catch (FileNotFoundException ex)
                {
                    arguments = @"/Apply-Image /Imagefile:" + esdYes + " /Index:" + index + @" /ApplyDir:" + ApplyDir;

                }
            }
            else // the location if we have a ESD file
            {
                arguments = @"/Apply-Image /Imagefile:" + esdNo + " /Index:" + index + @" /ApplyDir:" + ApplyDir;
            }

            using (Process process = new Process())
            {
                process.StartInfo.FileName = exePath; //DISM
                process.StartInfo.Arguments = arguments; // Code to deploy windows using the windows version the user requested using the ESD/WIM
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.CreateNoWindow = true;
                process.StartInfo.RedirectStandardOutput = true;

                process.OutputDataReceived += async (sender, e) =>
                {
                    if (e.Data != null)
                    {
                        string pattern = @"\b\d+\b"; // only want numeric items

                        Match match = Regex.Match(e.Data, pattern);

                        if (match.Success)
                        {
                            foreach (Match m in match.Groups)
                            {
                                if (!string.IsNullOrEmpty(m.Value) && int.Parse(m.Value) == 10 && ignoreFirst) // When we run DISM for the first time it shows the version and the software gets confused
                                {
                                    ignoreFirst = false;
                                }
                                else if (!string.IsNullOrEmpty(m.Value) && !ignoreFirst) // only want pure data
                                {
                                    int prog;
                                    if (int.TryParse(m.Value, out prog))
                                    {
                                        progressBar2.BeginInvoke(new Action(() => progressBar2.Value = prog));
                                        label4.BeginInvoke(new Action(() => label4.Text = "Current progress:" + prog + "%"));
                                        progressBar1.BeginInvoke(new Action(() => progressBar1.Value = 50 + ((progressBar1.Value + prog / 2) - progressBar1.Value))); // add onto the 50% that was done by downloading
                                        label3.BeginInvoke(new Action(() => label3.Text = "Current Task: Installing | Total Progress: " + progressBar1.Value + "%"));

                                    }
                                    else
                                    {
                                        Console.WriteLine("Wrong input");
                                    }
                                }
                            }
                        }
                        else
                        {

                        }
                    }
                };

                process.Start();
                process.BeginOutputReadLine();
                await Task.Run(() =>
                {
                    process.WaitForExit();
                });
            } // from here 
            label2.Text = "Done, rebooting now!";
            string deltemploc = cLetter.ToString() +":\\tempdelete.bat";
            string deltempcontents = @"@echo off" + Environment.NewLine +
             @"echo select disk " + disks + " > del.txt" + Environment.NewLine + 
             @"echo sel part 2 >> del.txt" + Environment.NewLine +
             @"echo delete partition override >> del.txt" + Environment.NewLine +
             @"echo sel part 1 >> del.txt " + Environment.NewLine +
             @"echo extend >> del.txt" + Environment.NewLine +
             @"echo exit >> del.txt" + Environment.NewLine +
             @"diskpart /s del.txt" + Environment.NewLine +
              @"del %0";
            File.WriteAllText(deltemploc, deltempcontents); // to here, this is the code that is generated to delete the temp drive for post setup // writing and reading a file
            if (!Directory.Exists(cLetter.ToString() +":\\Windows\\Panther"))
            {
                Directory.CreateDirectory(cLetter.ToString() +":\\Windows\\Panther"); // create the Panther folder so we can place the unattend script

            }
            string xmlLoc = tLetter.ToString() +":\\contin\\unattend.xml";// from here
            string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>" + Environment.NewLine +
            @"<unattend xmlns=""urn:schemas-microsoft-com:unattend"">" + Environment.NewLine +
            @"    <settings pass=""oobeSystem"">" + Environment.NewLine +
            @"        <component name=""Microsoft-Windows-Shell-Setup"" processorArchitecture=""amd64"" publicKeyToken=""31bf3856ad364e35"" language=""neutral"" versionScope=""nonSxS"" xmlns:wcm=""http://schemas.microsoft.com/WMIConfig/2002/State"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">" + Environment.NewLine +
            @"            <FirstLogonCommands>" + Environment.NewLine +
            @"                <SynchronousCommand wcm:action=""add"">" + Environment.NewLine +
            @"                    <CommandLine>C:\tempdelete.bat</CommandLine>" + Environment.NewLine +
            @"                    <Description>Deletes the T: or D: drive</Description>" + Environment.NewLine +
            @"                    <Order>1</Order>" + Environment.NewLine +
            @"                    <RequiresUserInput>true</RequiresUserInput>" + Environment.NewLine +
            @"                </SynchronousCommand>" + Environment.NewLine +
            @"            </FirstLogonCommands>" + Environment.NewLine +
            @"        </component>" + Environment.NewLine +
            @"    </settings>" + Environment.NewLine +
            @"</unattend>";
            if (!isExpress)
            {
                File.WriteAllText(xmlLoc, xmlContent); // writing and reading a file
            } // to here, we have to make an automation script to make the temp drive delete run at OOBE. When we select clean, the user doesn't want any modifications done to their computer but we still need to do clean up.
            else
            { // from here
                if (!Directory.Exists(cLetter.ToString() + ":\\Windows\\Setup\\Scripts\\"))
                {
                    Directory.CreateDirectory(cLetter.ToString() + ":\\Windows\\Setup\\Scripts\\");
                }
                try
                {
                    File.Move(tLetter.ToString() + ":\\contin\\Extras\\setup.exe", cLetter.ToString() + ":\\Windows\\Setup\\Scripts\\Ninite.exe");
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
                try
                {
                    File.Move(tLetter.ToString() + ":\\contin\\installer.bat", cLetter.ToString() + ":\\Windows\\Setup\\Scripts\\SetupComplete.cmd");
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
                try
                {
                    File.Move(tLetter.ToString() + ":\\contin\\autorun.exe", cLetter.ToString() + ":\\Windows\\Setup\\Scripts\\autorun.exe");
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
                try
                {
                    File.Move(tLetter.ToString() + ":\\contin\\test.au3", cLetter.ToString() + ":\\Windows\\Setup\\Scripts\\autorun.au3");
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
                try
                {
                    File.Move("X:\\Windows\\System32\\windowskey.txt", cLetter.ToString() + ":\\Windows\\Setup\\Scripts\\windowskey.txt");
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
                try
                {
                    File.Move("X:\\Windows\\System32\\SSID.txt", cLetter.ToString() + ":\\Windows\\Setup\\Scripts\\SSID.txt");
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
                try
                {
                    File.Move("X:\\Windows\\System32\\Password.txt", cLetter.ToString() + ":\\Windows\\Setup\\Scripts\\Password.txt");
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
                try
                {
                    File.Move("X:\\Windows\\System32\\profile.xml", cLetter.ToString() + ":\\Windows\\Setup\\Scripts\\profile.xml");
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
                try
                {
                    using (var client = new WebClient())
                    {
                        client.DownloadFile("https://github.com/eliasailenei/PortableISO/releases/download/POST/Installer.zip", tLetter.ToString() + ":\\contin\\installer.zip");
                    }
                    ZipFile.ExtractToDirectory(tLetter.ToString() + ":\\contin\\installer.zip", cLetter.ToString() + ":\\Windows\\Setup\\Scripts\\");
                    File.Delete(tLetter.ToString() + ":\\contin\\installer.zip");
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                };
            } // to here what we do is create the Scripts folder where we place our post setup things like our ninite or our Wi-Fi profile. We also download the post installer setup to move PortableDriver and get other things done during POST setup.
            
            string outFile = @"X:\bootable.bat"; // from here
            string content = "@echo off" + Environment.NewLine +
            @"ECHO Adding MBR..." + Environment.NewLine +
            @"copy " + tLetter.ToString() + ":\\contin\\unattend.xml " + cLetter.ToString() +":\\Windows\\Panther" + Environment.NewLine +
            @cLetter.ToString() +":\\Windows\\System32\\bcdboot "  + cLetter.ToString() + ":\\Windows" + Environment.NewLine +
            @"wpeutil reboot" + Environment.NewLine;
            File.WriteAllText(outFile, content);
            ProcessStartInfo diskp = new ProcessStartInfo();
            diskp.FileName = "cmd.exe";
            diskp.Arguments = "/C X:\\bootable.bat";
            diskp.CreateNoWindow = true;
            diskp.UseShellExecute = false;
            diskp.RedirectStandardOutput = true;
            diskp.RedirectStandardInput = true;
            Process diskpart = Process.Start(diskp);
            diskpart.WaitForExit();
            string debug = diskpart.StandardOutput.ReadToEnd();  // to here we make our drive bootable and restart the computer
        }
        private void MakeInvis(bool invis) // method to make the UI invisible //simple user algo
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => MakeInvis(invis)));
                return;
            }

            label1.Visible = invis;
            label2.Visible = invis;
            label3.Visible = invis;
            label4.Visible = invis;
            progressBar1.Visible = invis;
            progressBar2.Visible = invis;
        }

        public Clean()
        {
            InitializeComponent(); // load all of our UI items
            FormBorderStyle = FormBorderStyle.None; // do not add a border
            WindowState = FormWindowState.Maximized; // make it full screen
            this.Resize += Clean_rsize; // resize the UI according the screen resolution
            form = this.Size;
            lab1 = new Rectangle(label1.Location, label1.Size); // from here
            lab2 = new Rectangle(label2.Location, label2.Size);
            lab3 = new Rectangle(label3.Location, label3.Size);
            lab4 = new Rectangle(label4.Location, label4.Size);
            lb5 = new Rectangle(label5.Location, label5.Size);
            lab6 = new Rectangle(label6.Location, label6.Size);
            prog1 = new Rectangle(progressBar1.Location, progressBar1.Size);
            prog2 = new Rectangle(progressBar2.Location, progressBar2.Size);
            bttn1 = new Rectangle(button1.Location, button1.Size);
            lstbx1 = new Rectangle(listBox1.Location, listBox1.Size); // to here we make rectangles which represent the UI 
            label5.Visible = false;
            listBox1.Visible = false;
            button1.Visible = false;

        }

        private void Clean_rsize(object sender, EventArgs e)// recursive algorithm
        { // change the size of the UI elements one by one
            resizeControl(lab1, label1);
            resizeControl(lab2, label2);
            resizeControl(lab3, label3);
            resizeControl(lab4, label4);
            resizeControl(lb5, label5);
            resizeControl(lab6, label6);
            resizeControl(prog1, progressBar1);
            resizeControl(prog2, progressBar2);
            resizeControl(bttn1, button1);
            resizeControl(lstbx1, listBox1);

        }
        private void resizeControl(Rectangle r, Control c)
        {
            float xRatio = (float)this.Width / form.Width;
            float yRatio = (float)this.Height / form.Height; // get current x and y locations

            int newX = (int)(r.X * xRatio);
            int newY = (int)(r.Y * yRatio); // find the new positions

            int newWidth = (int)(r.Width * xRatio);
            int newHeight = (int)(r.Height * yRatio); // find the new height and width 

            c.Location = new Point(newX, newY); // apply the new data onto the UI elements
            c.Size = new Size(newWidth, newHeight);
        }

        private void timer1_Tick(object sender, EventArgs e)
        { // fade in effect
            if (Opacity == 1)
            {
                timer1.Stop();
            }
            Opacity += .4;
        }

        private void timer2_Tick(object sender, EventArgs e)
        { // fade out effect
            if (Opacity == 0)
            {
                this.Hide();
            }
            Opacity -= .4;
        }
    }
}
