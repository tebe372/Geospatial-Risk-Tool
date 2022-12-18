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

namespace GeospatialRiskAnalysisTool
{
    public partial class MLForm : Form
    {
        string _dataFile = string.Empty;
        int _windDirectionIndex = 0;
        int _windSpeedIndex = 0;
        int _sampleDateTimeIndex = 0;
        string[] _attributes = null;
        const double SUNRISE = 6.0;
        const double SUNSET = 18.0;

        public MLForm()
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
                //TODO
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
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
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var temp = this.Cursor;
            DateTime dt = DateTime.Now;
            this.Cursor = Cursors.WaitCursor;
            
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
        }

        private void button4_Click(object sender, EventArgs e)
        {
            var temp = this.Cursor;
            this.Cursor = Cursors.WaitCursor;
            
            this.Cursor = temp;
        }
    }
}
