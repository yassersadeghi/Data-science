using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace statics
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            chart1.Series[0].Name = "module / atm";

            foreach (Form1.personel p in Form1.P)
            {
                chart1.Series["module / atm"].Points.AddXY(p.name, ((double)p.module ) / p.atms);

            }

            chart1.Series["module / atm"]["PixelPointWidth"] = "23";
            chart1.ChartAreas[0].AxisX.LabelStyle.Angle = -45;
            Font f = new Font("Times New Roman", 12);
            chart1.ChartAreas[0].AxisX.LabelStyle.Font = f;
            chart1.ChartAreas[0].AxisX.Interval = 1;
           // chart1.Series[0].IsValueShownAsLabel = true;
            chart1.ChartAreas[0].Area3DStyle.Enable3D = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SaveFileDialog d = new SaveFileDialog();

            if (d.ShowDialog() == DialogResult.OK)
            {
                chart1.SaveImage(d.FileName + ".jpg", ChartImageFormat.Jpeg);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
