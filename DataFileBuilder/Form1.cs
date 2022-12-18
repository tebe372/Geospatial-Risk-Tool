using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO.Ports;
using System.IO;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;

namespace GeospatialRiskAnalysisTool
{
    public partial class Form1 : Form
    {
        string _dataFile = string.Empty;
        int _windDirectionIndex = 0;
        int _windSpeedIndex = 0;
        int _sampleDateTimeIndex = 0;
        string[] _attributes = null;
        const double SUNRISE = 6.0;
        const double SUNSET = 18.0;
        List<SiteObject> _sites = new List<SiteObject>();

        public Form1()
        {
            InitializeComponent();
        }

        private AboutForm aboutForm = null;
        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (aboutForm == null)
                aboutForm = new AboutForm();
            aboutForm.ShowDialog();
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Text file(*.csv)|*.csv";
            ofd.DefaultExt = ".csv";
            ofd.Multiselect = false;
            ofd.InitialDirectory = Path.Combine(Application.StartupPath, string.Format("Data"));
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                _dataFile = ofd.FileName;
                lbCalibration.Text = Path.GetFileName(_dataFile);
                this.comboBox1.Items.Clear();
                this.comboBox2.Items.Clear();
                this.comboBox3.Items.Clear();
                using ( FileStream fs = new FileStream( _dataFile, FileMode.Open))
                {
                    using(StreamReader sr = new StreamReader(fs))
                    {
                        string line = sr.ReadLine();
                        _attributes = line.Split(',');
                        this.comboBox1.Items.AddRange(_attributes);
                        this.comboBox2.Items.AddRange(_attributes);
                        this.comboBox3.Items.AddRange(_attributes);
                    }
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.listView1.Columns[0].Width = (int)Math.Ceiling(this.listView1.ClientRectangle.Width * 0.06);
            this.listView1.Columns[1].Width = (int)Math.Ceiling(this.listView1.ClientRectangle.Width * 0.08);
            this.listView1.Columns[2].Width = (int)Math.Ceiling(this.listView1.ClientRectangle.Width * 0.06);
            this.listView1.Columns[3].Width = (int)Math.Ceiling(this.listView1.ClientRectangle.Width * 0.06);
            this.listView1.Columns[4].Width = (int)Math.Ceiling(this.listView1.ClientRectangle.Width * 0.08);
            this.listView1.Columns[5].Width = (int)Math.Ceiling(this.listView1.ClientRectangle.Width * 0.16);
            this.listView1.Columns[6].Width = (int)Math.Ceiling(this.listView1.ClientRectangle.Width * 0.16);
            this.listView1.Columns[7].Width = (int)Math.Ceiling(this.listView1.ClientRectangle.Width * 0.16);
            this.listView1.Columns[8].Width = (int)Math.Ceiling(this.listView1.ClientRectangle.Width * 0.16);
            string filenPath = Path.Combine(Application.StartupPath, "Resources");
            filenPath = Path.Combine(filenPath, "metStationsHanford.csv");
            using (FileStream fs = new FileStream(filenPath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                using(StreamReader sr = new StreamReader(fs))
                {
                    string line = sr.ReadLine();
                    line = sr.ReadLine();
                    while ( line != null)
                    {
                        string[] tokens = line.Split(',');
                        SiteObject so = new SiteObject();
                        so.SiteId = tokens[0];
                        so.SiteNumber = tokens[1];
                        so.SiteName = tokens[2];
                        so.Latitude = tokens[3];
                        so.Longitude = tokens[4];
                        _sites.Add(so);
                        line = sr.ReadLine();
                    }
                }
            }
            this.comboBox4.Items.AddRange(_sites.Select(a => a.SiteName).ToArray());
            this.comboBox4.SelectedIndex = 0;
            this.comboBox7.Items.AddRange(_sites.Select(a => a.SiteName).ToArray());
            this.comboBox7.SelectedIndex = 0;
            this.richTextBox1.Text = Path.Combine(Application.StartupPath, @"Results\PTR-MS\PTR_MS_Data.csv");
            using (FileStream fs = new FileStream(this.richTextBox1.Text, FileMode.Open))
            {
                using (StreamReader sr = new StreamReader(fs))
                {
                    string line = sr.ReadLine();
                    string[] myattributes = line.Split(',');
                    this.comboBox6.Items.AddRange(myattributes);
                    this.comboBox6.Text = "acetone";
                }
            }
        }

        private void listView1_DrawColumnHeader(object sender, DrawListViewColumnHeaderEventArgs e)
        {
            e.DrawDefault = true;
        }

        private void listView1_DrawItem(object sender, DrawListViewItemEventArgs e)
        {
            ListView listView = (ListView)sender;

            // Check if e.Item is selected and the ListView has a focus.
            if (!listView.Focused && e.Item.Selected)
            {
                Rectangle rowBounds = e.Bounds;
                int leftMargin = e.Item.GetBounds(ItemBoundsPortion.Label).Left;
                Rectangle bounds = new Rectangle(leftMargin, rowBounds.Top, rowBounds.Width - leftMargin, rowBounds.Height);
                e.Graphics.FillRectangle(SystemBrushes.Highlight, bounds);
                e.Graphics.DrawString(e.Item.Text, e.Item.Font, Brushes.White, e.Item.Position);
            }
            else
                e.DrawDefault = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (_dataFile == null || _dataFile.Length == 0)
                return;

            if (this.comboBox1.Text == null || this.comboBox1.Text.Length == 0)
                return;
            if (this.comboBox2.Text == null || this.comboBox2.Text.Length == 0)
                return;
            if (this.comboBox3.Text == null || this.comboBox3.Text.Length == 0)
                return;
            var temp = this.Cursor;
            this.Cursor = Cursors.WaitCursor;
            Dictionary<string, ListViewItem> items = new Dictionary<string, ListViewItem>();
            var lines = File.ReadAllLines(_dataFile);
            var count = lines.Length;
            if (count > 1)
            {
                this.progressBar1.Maximum = count;
                this.progressBar1.Value = 0;
                for ( int i=1;i<count;i++)
                {
                    string line = lines[i];
                    this.progressBar1.Value += 1;
                    int nCount = 0;
                    if(line != null)
                    {
                        string[] attributes = line.Split(',');
                        string dateTimeString = attributes[_sampleDateTimeIndex];
                        if(!items.ContainsKey(dateTimeString))
                        {
                            try
                            {
                                dateTimeString = dateTimeString.Replace("\"", "");
                                DateTime dt = DateTime.Parse(dateTimeString);
                                ListViewItem item = new ListViewItem((dt.Year-2000).ToString());
                                double result = 0d;
                                double result2 = 0d;
                                if (!double.TryParse(attributes[_windDirectionIndex].Replace("\"", ""), out result) ||
                                    !double.TryParse(attributes[_windSpeedIndex].Replace("\"", ""), out result2))
                                    continue;
                                item.SubItems.Add(dt.Month.ToString("00"));
                                item.SubItems.Add(dt.Day.ToString("00"));
                                item.SubItems.Add(dt.Hour.ToString("00"));
                                item.SubItems.Add(dt.Minute.ToString("00"));
                                item.SubItems.Add((result*180.0/Math.PI).ToString());
                                item.SubItems.Add(attributes[_windSpeedIndex].Replace("\"", ""));
                                item.SubItems.Add(textBox3.Text.Trim());
                                string category = "4";// "D";
                                //double speed = Convert.ToDouble(attributes[_windSpeedIndex].Replace("\"", ""));
                                //if (dt.Hour >= SUNRISE && dt.Hour < SUNSET) // Day time
                                //{
                                //    if (speed < 3)
                                //        category = "C";
                                //    else if (speed < 2)
                                //        category = "B";
                                //}
                                //else
                                //{
                                //    if (speed < 3)
                                //        category = "E";
                                //}

                                item.SubItems.Add(category);
                                item.BackColor = Color.White;
                                if (nCount % 2 == 0)
                                {
                                    item.BackColor = Color.LightCyan;
                                }
                                items.Add(dateTimeString, item);
                            }
                            catch (Exception ex)
                            {

                            }
                            nCount++;
                        }
                    }
                    Application.DoEvents();
                }
                this.listView1.BeginUpdate();
                this.listView1.Items.Clear();
                this.listView1.Items.AddRange(items.Values.ToArray());
                this.listView1.EndUpdate();
            }
            this.Cursor = temp;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var temp = this.Cursor;
            DateTime dt = DateTime.Now;
            this.Cursor = Cursors.WaitCursor;
            string fileprefix = Path.GetFileNameWithoutExtension(_dataFile);
            string file = string.Format("{0}_stndata.txt", fileprefix);
            if (this.listView1.Items.Count > 0)
            {
                FolderBrowserDialog sfd = new FolderBrowserDialog();
                if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    string path = sfd.SelectedPath;
                    file = Path.Combine(path, file);
                    this.progressBar1.Maximum = this.listView1.Items.Count;
                    this.progressBar1.Value = 0;
                    using (FileStream fs = new FileStream(file, FileMode.Create))
                    {
                        using (StreamWriter sw = new StreamWriter(fs))
                        {
                            foreach (ListViewItem item in this.listView1.Items)
                            {
                                var c = 4;
                                var category = item.SubItems[8].Text;
                                if (category == "C")
                                    c = 3;
                                else if (category == "B")
                                    c = 2;
                                else if (category == "E")
                                    c = 5;

                                string line = string.Format("{0} {1:00} {2:00} {3:00} {4:00} {5} {6} {7} {8} ",
                                item.SubItems[0].Text, item.SubItems[1].Text, item.SubItems[2].Text, item.SubItems[3].Text,
                                item.SubItems[4].Text, item.SubItems[5].Text, item.SubItems[6].Text, item.SubItems[7].Text,
                                c);
                                sw.WriteLine(line);
                                this.progressBar1.Value++;
                            }
                        }
                    }
                }
                MessageBox.Show("Successfully exported data into:\n" + file, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            this.Cursor = temp;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            this._windDirectionIndex = (sender as ComboBox).SelectedIndex;
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            this._windSpeedIndex = (sender as ComboBox).SelectedIndex;
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            this._sampleDateTimeIndex = (sender as ComboBox).SelectedIndex;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.listView1.Items.Clear();
            _windDirectionIndex = 0;
            _windSpeedIndex = 0;
            _sampleDateTimeIndex = 0;
            _attributes = null;
            this._dataFile = string.Empty;
            this.lbCalibration.Text = "";
            this.comboBox1.Items.Clear();
            this.comboBox2.Items.Clear();
            this.comboBox3.Items.Clear();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if(!File.Exists(@"C:\hysplit\exec\stn2arl.exe"))
            {
                MessageBox.Show(@"C:\hysplit\exec\stn2arl.exe doesn't exist." + "\nPlease download Hysplit from https://www.ready.noaa.gov/HYSPLIT.php and install it into " + @"C:\hysplit.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            var temp = this.Cursor;
            this.Cursor = Cursors.WaitCursor;
            string fileprefix = Path.GetFileNameWithoutExtension(_dataFile);
            string file = string.Format("{0}_stndata.txt", fileprefix);
            string infile = string.Format(@"C:\HYSPLIT\working\{0}_H{1}.txt", fileprefix,comboBox4.SelectedIndex + 1);
            string outfile = string.Format(@"C:\HYSPLIT\working\{0}_H{1}.bin", fileprefix, comboBox4.SelectedIndex + 1);
            if (this.listView1.Items.Count > 0)
            {
                using (FileStream fs = new FileStream(infile, FileMode.Create))
                {
                    using (StreamWriter sw = new StreamWriter(fs))
                    {
                        foreach (ListViewItem item in this.listView1.Items)
                        {
                            var c = 4;
                            var category = item.SubItems[8].Text;
                            if (category == "C")
                                c = 3;
                            else if (category == "B")
                                c = 2;
                            else if (category == "E")
                                c = 5;

                            string line = string.Format("{0} {1:00} {2:00} {3:00} {4:00} {5} {6} {7} {8} ",
                            item.SubItems[0].Text, item.SubItems[1].Text, item.SubItems[2].Text, item.SubItems[3].Text,
                            item.SubItems[4].Text, item.SubItems[5].Text, item.SubItems[6].Text, item.SubItems[7].Text,
                            c);
                            sw.WriteLine(line);
                        }
                    }
                }
                var pp = new ProcessStartInfo(@"C:\hysplit\exec\stn2arl.exe", string.Format("{0} {1} {2} {3}", infile, outfile, this.textBox1.Text, this.textBox2.Text))
                {
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    WorkingDirectory = Application.StartupPath,
                };
                var process = Process.Start(pp);
                process.WaitForExit();
                process.Close();
                MessageBox.Show("Successfully converted to ARL data file:\n" + outfile, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.tbBinFile.Text = outfile;
                this.textBox7.Text = outfile;
            }
            this.Cursor = temp;
        }

        private void DeleteAllFile( string folder, string prefix)
        {
            string[] filePaths = Directory.GetFiles(folder, prefix + "*.txt", SearchOption.AllDirectories);
            foreach (string filePath in filePaths)
            {
                File.Delete(filePath);
            }
        }

        private void MoveAllFile(string srcfolder, string dstfolder, string prefix)
        {
            string[] filePaths = Directory.GetFiles(srcfolder, prefix + "*.txt", SearchOption.AllDirectories);
            foreach (string filePath in filePaths)
            {
                string filename = Path.GetFileName(filePath);
                FileInfo fi = new FileInfo(filePath);
                long filesize = fi.Length;
                if( filesize < 1024)
                    File.Delete(filePath);
                else
                    File.Move(filePath,Path.Combine(dstfolder, filename));
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            var tmpCursor = this.Cursor;
            Cursor = Cursors.WaitCursor;
            this.DeleteAllFile(@"C:\HYSPLIT\working", this.textBox6.Text.Trim());
            DateTime dt = new DateTime(dateTimePicker1.Value.Year, dateTimePicker1.Value.Month, dateTimePicker1.Value.Day, Convert.ToInt32(textBox4.Text), Convert.ToInt32(textBox8.Text), 0);
            if (textBox5.Text != null && textBox5.Text.Trim().Length > 0
                && textBox7.Text != null && textBox7.Text.Trim().Length > 0
                )
            {
                int nLoop = Convert.ToInt32(Math.Ceiling(Double.Parse(textBox5.Text.Trim())));
                progressBar1.Maximum = nLoop;
                progressBar1.Value = 0;
                string fileName = Path.GetFileNameWithoutExtension(textBox7.Text);
                int lastindex = fileName.LastIndexOf("_");
                string siteName = fileName.Substring(lastindex + 1, fileName.Length - lastindex - 1);
                for (int i = 0; i < nLoop; i++)
                {
                    //string outfilename = string.Format("{0}{1}.{2}{3:00}{4:00}{5:00}.txt", textBox6.Text, siteName, dt.Year, dt.Month, dt.Day, dt.Hour);
                    string outfilename = string.Format("{4}{3}{0:00}{1:00}{2:00}.txt", dt.Month, dt.Day, dt.Hour,dt.Year, textBox6.Text.Trim());
                    if (File.Exists(@"C:\HYSPLIT\exec\CONTROL"))
                    {
                        File.Delete(@"C:\HYSPLIT\exec\CONTROL");
                    }
                    if (File.Exists(@"C:\HYSPLIT\working\CONTROL"))
                    {
                        File.Delete(@"C:\HYSPLIT\working\CONTROL");
                    }
                    StringBuilder sb = new StringBuilder();
                    string starttime = string.Format("{0:00} {1:00} {2:00} {3:00} {4:00}", dt.Year - 2000, dt.Month, dt.Day, dt.Hour, dt.Minute);
                    sb.AppendLine(starttime);
                    sb.AppendLine("1");
                    sb.AppendLine(string.Format("{0} {1} {2}", textBox1.Text, textBox2.Text, textBox3.Text));
                    if( this.radioButton1.Checked)
                        sb.AppendLine("72");
                    else
                        sb.AppendLine("-72");
                    sb.AppendLine("0");
                    sb.AppendLine("10000.0");
                    sb.AppendLine("1");
                    string tmp = Path.GetDirectoryName(textBox7.Text).Replace("\\","/");
                    sb.AppendLine(tmp + "/");
                    sb.AppendLine(Path.GetFileName(textBox7.Text));
                    sb.AppendLine(@"./");
                    //sb.AppendLine("mytdump"+i);
                    sb.AppendLine(outfilename);
                    File.WriteAllText(@"C:\HYSPLIT\exec\CONTROL", sb.ToString());
                    File.WriteAllText(@"C:\HYSPLIT\working\CONTROL", sb.ToString());
                    sb.Clear();
                    var pp = new ProcessStartInfo(@"C:\hysplit\exec\hyts_std.exe")
                    {
                        CreateNoWindow = true,
                        UseShellExecute = false,
                        WorkingDirectory = @"C:\HYSPLIT\working",
                    };
                    var process = Process.Start(pp);
                    process.WaitForExit();
                    process.Close();

                    //var ppp = new ProcessStartInfo(@"C:\hysplit\exec\trajplot.exe", string.Format("-imytdump{0} -jC:\\HYSPLIT\\graphics\\arlmap -o{1}{2}", i, textBox6.Text,i))
                    //var ppp = new ProcessStartInfo(@"C:\hysplit\exec\trajplot.exe", string.Format("-i{0} -jC:\\HYSPLIT\\graphics\\arlmap -o{0}", outfilename))
                    //{
                    //    CreateNoWindow = true,
                    //    UseShellExecute = false,
                    //    WorkingDirectory = @"C:\HYSPLIT\working",
                    //};
                    //var pprocess = Process.Start(ppp);
                    //pprocess.WaitForExit();
                    //pprocess.Close();
                    progressBar1.Value++;
                    if( this.radioButton1.Checked)
                        dt = dt.AddHours((int)(Convert.ToDouble(this.comboBox5.Text) / 60 + 0.2));
                    else
                        dt = dt.AddHours((int)(Convert.ToDouble(this.comboBox5.Text) / 60 + 0.2)*-1);
                    Application.DoEvents();
                }
                string dataFolder = Path.Combine(Application.StartupPath, @"Results\MetData"); 
                this.DeleteAllFile(dataFolder, "");
                this.MoveAllFile(@"C:\HYSPLIT\working", dataFolder, this.textBox6.Text.Trim());
                MessageBox.Show("Results are saved under C:\\HYSPLIT\\working as " + textBox6.Text + "##.txt", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            Cursor = tmpCursor;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Text file(*.bin)|*.bin";
            ofd.DefaultExt = ".bin";
            ofd.Multiselect = false;
            ofd.InitialDirectory = @"C:\HYSPLIT\working";
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                textBox7.Text = ofd.FileName;
            }
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            var site = _sites.FirstOrDefault(a => a.SiteName == comboBox4.Text);
            if( site != null)
            {
                this.textBox1.Text = site.Latitude;
                this.textBox2.Text = site.Longitude;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            /*
             *  C:\Users\houh226\AppData\Local\Programs\Python\Python310\Scripts\pip3.exe install numpy
                C:\Users\houh226\AppData\Local\Programs\Python\Python310\Scripts\pip3.exe install openpyxl
                C:\Users\houh226\AppData\Local\Programs\Python\Python310\Scripts\pip3.exe install scipy
                C:\Users\houh226\AppData\Local\Programs\Python\Python310\Scripts\pip3.exe install matplotlib
                C:\Users\houh226\AppData\Local\Programs\Python\Python310\Scripts\pip3.exe install mpld3
                C:\Users\houh226\AppData\Local\Programs\Python\Python310\Scripts\pip3.exe install basemap
                C:\Users\houh226\AppData\Local\Programs\Python\Python310\Scripts\pip3.exe install plotly
                C:\Users\houh226\AppData\Local\Programs\Python\Python310\Scripts\pip3.exe install pandas
                C:\Users\houh226\AppData\Local\Programs\Python\Python310\Scripts\pip3.exe install basemap-data-hires
             */
            if (this.richTextBox2.Text.Trim().Length == 0 || !File.Exists(this.richTextBox2.Text.Trim()))
            {
                MessageBox.Show(this.richTextBox2.Text.Trim() + " doesn't exist.\nPlease install it.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            var cursor = this.Cursor;
            this.Cursor = Cursors.WaitCursor;

            string dataFolder = Path.Combine(Application.StartupPath, @"Results\MetData");
            int count = Directory.GetFiles(dataFolder, "*.txt", SearchOption.TopDirectoryOnly).Select(f => Path.GetFileName(f)).Count();
            this.textBox9.Text = String.Format("Data processing may take {0} minutes.", (int)Math.Ceiling(count / 60.0));
            Process p = new Process();
            p.StartInfo = new ProcessStartInfo();
            p.StartInfo.FileName = this.richTextBox2.Text;
            p.StartInfo.Arguments = String.Format(" cpsdf.py {0} {1} {2} {3}", (int)Math.Floor( double.Parse(this.textBox11.Text)),
                (int)Math.Floor(double.Parse(this.textBox10.Text)),this.comboBox6.Text.Trim(), this.comboBox7.Text.Trim());
            p.StartInfo.WorkingDirectory = Path.Combine(Application.StartupPath, "Results");
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.CreateNoWindow = false;
            p.Start();
            p.WaitForExit();
            p.Close();
            p.Dispose();
            this.textBox9.Text = "";
            GC.Collect();
            this.Cursor = cursor;
        }

        private void helpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start(Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "Resources/GeospatialRiskAnalysisToolManual.pdf"));
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog fb = new OpenFileDialog();
            fb.Filter = "Text file(*.csv)|*.csv";
            fb.DefaultExt = ".csv";
            fb.Multiselect = false;
            fb.InitialDirectory = Path.Combine(Application.StartupPath, @"Results\PTR-MS");
            fb.FileName = "PTR-MS.csv";
            fb.Multiselect = false;
            if (fb.ShowDialog() != DialogResult.Cancel)
            {
                this.comboBox6.Items.Clear();
                using (FileStream fs = new FileStream(fb.FileName, FileMode.Open))
                {
                    using (StreamReader sr = new StreamReader(fs))
                    {
                        string line = sr.ReadLine();
                        string[] myattributes = line.Split(',');
                        this.comboBox6.Items.AddRange(myattributes);
                        this.comboBox6.Text = "acetone";
                    }
                }
            }
            this.richTextBox1.Text = fb.FileName;
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog fb = new OpenFileDialog();
            fb.InitialDirectory = @"C:\Users";
            fb.FileName = "python.exe"; 
            fb.Multiselect = false;
            if (fb.ShowDialog() != DialogResult.Cancel)
            {
                this.richTextBox2.Text = fb.FileName;
            }
        }

         
        private void button8_Click(object sender, EventArgs e)
        {
            var cursor = this.Cursor;
            this.Cursor = Cursors.WaitCursor;
            string dataFolder = Path.Combine(Application.StartupPath, @"Results\MetData");
            List<string> filePaths = Directory.GetFiles(dataFolder, "*.txt", SearchOption.TopDirectoryOnly).Select(f => Path.GetFileName(f)).ToList();
            Dictionary<string, string> ptrmsdic = new Dictionary<string, string>();
            int index = 0;
            using (FileStream fs = new FileStream(this.richTextBox1.Text, FileMode.Open))
            {
                using (StreamReader sr = new StreamReader(fs))
                {
                    string line = sr.ReadLine();
                    string[] myattributes = line.Split(',');
                    for( int i = 0; i < myattributes.Length; i++)
                    {
                        if( myattributes[i] == this.comboBox6.Text)
                        {
                            index = i;
                            break;
                        }
                    }
                    while((line = sr.ReadLine())!= null)
                    {
                        string[] tokens = line.Split(',');
                        DateTime dt = DateTime.Parse(tokens[0]);
                        string key = string.Format("{0:0000}{1:00}{2:00}{3:00}.txt", dt.Year, dt.Month, dt.Day, dt.Hour);
                        ptrmsdic.Add(key, tokens[index]);
                    }                    
                }
            }
            Dictionary<string, string> myptrmsdic = new Dictionary<string, string>();
            foreach ( var filename in filePaths)
            {
                string key = filename.Substring(filename.Length - 14);
                if(ptrmsdic.ContainsKey(key) && ptrmsdic[key] != "NULL")
                    myptrmsdic.Add(filename, ptrmsdic[key]);
            }
            int nCount = 0;
            foreach (var data in myptrmsdic)
            {
                ListViewItem item = new ListViewItem(data.Key);
                item.SubItems.Add(data.Value);
                if( nCount %2 != 0)
                {
                    item.BackColor = Color.LightCyan;
                }
                listView2.Items.Add(item);
                Application.DoEvents();
                nCount++;
            }
            string listFile = Path.Combine(dataFolder, "list.xlsx");
            File.Delete(listFile);
            Excel.Application xlApp = new Excel.Application();
            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!");
                return;
            }

            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            nCount = 1;
            foreach (var data in myptrmsdic)
            {
                xlWorkSheet.Cells[nCount, 1] = data.Key;
                xlWorkSheet.Cells[nCount, 2] = data.Value;
                nCount++;
            }
            xlWorkBook.SaveAs(listFile, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, misValue,
            misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
            this.Cursor = cursor;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
            var site = _sites.FirstOrDefault(a => a.SiteName == comboBox7.Text);
            if (site != null)
            {
                this.textBox11.Text = site.Latitude;
                this.textBox10.Text = site.Longitude;
            }
        }
    }
}
