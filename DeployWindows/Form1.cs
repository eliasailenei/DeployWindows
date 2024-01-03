﻿using System;
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



namespace DeployWindows
{
    public partial class Clean : Form
    {
        private bool invis;
        private Rectangle lab1, lab2, lab3, lab4, lab5, prog1, prog2, bttn1, lstbx1;
        private int index;
        private Size form;
        private int totalProg;
        private int prog;
        private bool isoYes = false;
        private string esdNo = @"T:\contin\test.esd";
        private string esdYes = @"T:\contin\\sources\install.esd";
        private string[] argss = Environment.GetCommandLineArgs();
        private string wimLoc = @"T:\contin\\sources\install.wim";
        private string ApplyDir = @"C:\";
        private string ISOLoc = @"T:\contin\test.iso";
        private string ISOExtract = @"T:\contin\";
        public bool isExpress = false;
        public string topass;
        public string disks;
        // input should be like this : >DeployWindows.exe topass='--ESDMode=True --WinVer=Windows_10 --Release=10,_version_22H2_[19045.2965]_(Updated_May_2023) --Language=English' disks='0' isExpress='False'

        // We need to skip the language setting in some way
        private void setArgs()
        {
            StringBuilder inp = new StringBuilder();
            try
            {
                foreach (string word in argss)
                {
                    inp.Append(word);
                    inp.Append(" ");
                }
            } 
            
            catch
            {

            }
            string args = inp.ToString();
            // input 
            string pattern = @"(\w+)='([^']+)'";

            MatchCollection matches = Regex.Matches(args, pattern);

            foreach (Match match in matches)
            {
                string attribute = match.Groups[1].Value;
                string value = match.Groups[2].Value;

                if (attribute == "topass")
                {
                    topass = value;
                }
                else if (attribute == "disks")
                {
                    disks = value;
                }
                else if (attribute == "isExpress")
                {
                    if (value == "False" | value == "false")
                    {
                        isExpress = false;
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

        private async void Clean_Load(object sender, EventArgs e)
        {
            setArgs();
            label3.Text = "Current Task: Downloading | Current Progress: estimating";
            label4.Text = "Do not panic if frozen! WinPE is not optimized for high loads!";

            await Task.Run(() => Downld());

            label3.Text = "Current Task: ISO check | Current Progress: estimating";
            progressBar2.Value = 0;
            label4.Text = "Do not panic if frozen! WinPE is not optimized for high loads!";
            await Task.Run(() => ISOcheck());
            label3.Text = "Current Task: Installing | Current Progress: estimating";
            progressBar2.Value = 0;
            label4.Text = "Estimating progress...";
            label5.Visible = false;
            listBox1.Visible = false;
            button1.Visible = false;
            await Task.Run(() => DISMDiag());
           
        }

        private async Task Downld()
        {
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = @"MSWISO\MSWISO.exe",
                Arguments = topass,
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using (Process process = new Process { StartInfo = startInfo })
            {
                process.Start();

                while (!process.StandardOutput.EndOfStream)
                {
                    string WhtOut = process.StandardOutput.ReadLine();
                    int giveProgress;
                    string newOut = WhtOut.TrimEnd('%');
                    if (int.TryParse(newOut, out giveProgress))
                    {

                        label4.BeginInvoke(new Action(() => label4.Text = "Current progress: " + WhtOut));
                        progressBar2.BeginInvoke(new Action(() => progressBar2.Value = giveProgress));
                        totalProg = Math.Min(giveProgress / 2, 50);
                        label3.BeginInvoke(new Action(() => label3.Text = "Current Task: Downloading | Total Progress: " + totalProg + "%"));
                        progressBar1.BeginInvoke(new Action(() => progressBar1.Value = totalProg));
                    }
                    else if (WhtOut.StartsWith("No"))
                    {
                        MessageBox.Show("It looks like Microsoft no longer hosts this version... Try to not use legacy mode if possible or a different version. Press OK to restart.");
                        Process.Start("cmd.exe", "/C wpeutil reboot");
                    }
                }

                process.WaitForExit();
                process.Close();
            }
        }
        private async Task DISMDiag()
        {
            invis = false;
            MakeInvis(invis);
            label5.Invoke(new Action(() => label5.Visible = true));
            listBox1.Invoke(new Action(() => listBox1.Visible = true));
            button1.Invoke(new Action(() => button1.Visible = true));
            string arguments;
            string exePath = Environment.GetEnvironmentVariable("SystemRoot") + @"\System32" + @"\dism.exe";
            if (isoYes)
            {
                MessageBox.Show(isoYes.ToString());
                try
                {
                    arguments = @"/get-wiminfo /wimfile:" + wimLoc;
                }
                catch (FileNotFoundException)
                {
                    arguments = @"/get-wiminfo /wimfile:" + esdYes;
                }
            }
            else
            {
                arguments = @"/get-wiminfo /wimfile:" + esdNo;
            }
            using (Process process = new Process())
            {
                process.StartInfo.FileName = exePath;
                process.StartInfo.Arguments = arguments;
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.CreateNoWindow = true;
                process.StartInfo.RedirectStandardOutput = true;

                process.OutputDataReceived += (sender, e) =>
                {
                    if (e.Data != null)
                    {
                        string newOut = e.Data.ToString();
                        string pattern = @"^Name :\s+(.+)$";

                        Match match = Regex.Match(newOut, pattern);
                        if (match.Success)
                        {
                            string name = match.Groups[1].Value;

                            listBox1.Invoke(new Action(() =>
                            {
                                listBox1.Items.Add(name);
                            }));
                        }
                    }
                };



                await Task.Run(() =>
                {
                    process.Start();
                    process.BeginOutputReadLine();
                    process.WaitForExit();
                });
            }


        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            index = listBox1.SelectedIndex + 1;

        }

        private async void button1_Click(object sender, EventArgs e)
        {
            invis = true;
            MakeInvis(invis);
            label5.Visible = false;
            listBox1.Visible = false;
            button1.Visible = false;
            await Task.Run(() => Install());
        }

        private async Task Install()
        {
            progressBar2.Value = 0;
            string arguments;
            string exePath = Environment.GetEnvironmentVariable("SystemRoot") + @"\System32" + @"\dism.exe";
            if (isoYes)
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
            else
            {
                arguments = @"/Apply-Image /Imagefile:" + esdNo + " /Index:" + index + @" /ApplyDir:" + ApplyDir;
            }

            using (Process process = new Process())
            {
                process.StartInfo.FileName = exePath;
                process.StartInfo.Arguments = arguments;
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.CreateNoWindow = true;
                process.StartInfo.RedirectStandardOutput = true;

                process.OutputDataReceived += async (sender, e) =>
                {
                    if (e.Data != null)
                    {
                        string pattern = @"\b\d+\b";

                        Match match = Regex.Match(e.Data, pattern);

                        if (match.Success)
                        {
                            foreach (Match m in match.Groups)
                            {
                                progressBar1.BeginInvoke(new Action(() => progressBar2.Value = int.Parse(m.Value)));
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
            string deltemploc = "C:\\tempdelete.bat";
            string deltempcontents = @"@echo off" + Environment.NewLine +
             @"echo select disk " + disks[1] + " > del.txt" + Environment.NewLine +
             @"echo sel part 2 >> del.txt" + Environment.NewLine +
             @"echo delete partition override >> del.txt" + Environment.NewLine +
             @"echo sel part 1 >> del.txt " + Environment.NewLine +
             @"echo extend >> del.txt" + Environment.NewLine +
             @"echo exit >> del.txt" + Environment.NewLine +
             @"diskpart /s del.txt" + Environment.NewLine +
             @"diskpart /s del.txt" + Environment.NewLine +
              @"del %0";
            File.WriteAllText(deltemploc, deltempcontents);
            Directory.CreateDirectory("C:\\Windows\\Panther");
            string xmlLoc = "T:\\contin\\unattend.xml";
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
                File.WriteAllText(xmlLoc, xmlContent);
            }
            string outFile = @"X:\bootable.bat";
            string content = @"
            @echo off" + Environment.NewLine +
            @"ECHO Adding MBR..." + Environment.NewLine +
            @"copy T:\contin\unattend.xml C:\Windows\Panther" + Environment.NewLine +
            @"C:\Windows\System32\bcdboot C:\Windows" + Environment.NewLine +
            @"nowpeutil.exe reboot" + Environment.NewLine;
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
            string debug = diskpart.StandardOutput.ReadToEnd(); // to here, make sub. Make function to see if express.exe is done, if done continue
        }
        private void MakeInvis(bool invis)
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

        private async Task ISOcheck()
        {
            await Task.Run(() =>
            {
                totalProg = 50;
                int progress = 1;
                progressBar2.Invoke(new Action(() => progressBar2.Value = progress));
                label4.Invoke(new Action(() => label4.Text = "Current progress: " + progress));
                totalProg = totalProg + progress;
                label3.BeginInvoke(new Action(() => label3.Text = "Current Task: ISO check | Total Progress: " + totalProg + "%"));
                string[] fileFound = Directory.GetFiles(AppDomain.CurrentDomain.BaseDirectory, "*.iso");

                if (fileFound != null)
                {

                    using (Process process = new Process())
                    {
                        process.StartInfo.FileName = "7z.exe";
                        process.StartInfo.Arguments = @"x -o" + ISOExtract + " " + ISOLoc + " -bd -bsp1 -bso1";
                        process.StartInfo.UseShellExecute = false;
                        process.StartInfo.CreateNoWindow = true;
                        process.StartInfo.RedirectStandardOutput = true;
                        process.StartInfo.RedirectStandardError = true;

                        process.OutputDataReceived += (sender, e) =>
                        {
                            if (e.Data != null)
                            {

                                Match match = Regex.Match(e.Data, @"\s(\d+)%");
                                if (match.Success)
                                {
                                    progress = int.Parse(match.Groups[1].Value);

                                    if (progress == 99)
                                    {
                                        progress++;
                                    }
                                    totalProg = 50;
                                    if (progress >= 4)
                                    {
                                        int scaledValue = (int)(15 * Math.Min(1, (progress - 4) / 96.0));
                                        totalProg = totalProg + scaledValue;
                                    }
                                    if (progress == 5)
                                    {
                                        isoYes = true;
                                    }
                                    totalProg = Math.Min(totalProg, 65);
                                    label4.BeginInvoke(new Action(() => label4.Text = "Current progress: " + progress + "%"));
                                    progressBar2.BeginInvoke(new Action(() => progressBar2.Value = progress));
                                    label3.BeginInvoke(new Action(() => label3.Text = "Current Task: ISO extract | Total Progress: " + totalProg + "%"));
                                    progressBar1.BeginInvoke(new Action(() => progressBar1.Value = totalProg));
                                }
                            }
                        };

                        process.Start();
                        process.BeginOutputReadLine();
                        process.WaitForExit();

                    }
                }
                else
                {
                    label3.BeginInvoke(new Action(() => label3.Text = "Current Task: Skipping ISO check | Total Progress: " + totalProg));
                }

            });
        }
        public Clean()
        {
            InitializeComponent();
            FormBorderStyle = FormBorderStyle.None;
            WindowState = FormWindowState.Maximized;
            this.Resize += Clean_rsize;
            form = this.Size;
            lab1 = new Rectangle(label1.Location, label1.Size);
            lab2 = new Rectangle(label2.Location, label2.Size);
            lab3 = new Rectangle(label3.Location, label3.Size);
            lab4 = new Rectangle(label4.Location, label4.Size);
            lab5 = new Rectangle(label5.Location, label5.Size);
            prog1 = new Rectangle(progressBar1.Location, progressBar1.Size);
            prog2 = new Rectangle(progressBar2.Location, progressBar2.Size);
            bttn1 = new Rectangle(button1.Location, button1.Size);
            lstbx1 = new Rectangle(listBox1.Location, listBox1.Size);
            label5.Visible = false;
            listBox1.Visible = false;
            button1.Visible = false;

        }

        private void Clean_rsize(object sender, EventArgs e)
        {

            resizeControl(lab1, label1);
            resizeControl(lab2, label2);
            resizeControl(lab3, label3);
            resizeControl(lab4, label4);
            resizeControl(lab5, label5);
            resizeControl(prog1, progressBar1);
            resizeControl(prog2, progressBar2);
            resizeControl(bttn1, button1);
            resizeControl(lstbx1, listBox1);

        }
        private void resizeControl(Rectangle r, Control c)
        {
            float xRatio = (float)this.Width / form.Width;
            float yRatio = (float)this.Height / form.Height;

            int newX = (int)(r.X * xRatio);
            int newY = (int)(r.Y * yRatio);

            int newWidth = (int)(r.Width * xRatio);
            int newHeight = (int)(r.Height * yRatio);

            c.Location = new Point(newX, newY);
            c.Size = new Size(newWidth, newHeight);
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (Opacity == 1)
            {
                timer1.Stop();
            }
            Opacity += .4;
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            if (Opacity == 0)
            {
                this.Hide();
            }
            Opacity -= .4;
        }
    }
}
