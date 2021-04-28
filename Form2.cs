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
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            
           
            chart1.Series[0].Name = "log / atm";
  
            foreach (Form1.personel p in Form1.P )
            {
                chart1.Series["log / atm"].Points.AddXY(p.name, ((double)p.log) / p.atms);
                
            }
         
            chart1.Series["log / atm"]["PixelPointWidth"] = "23";
            chart1 .ChartAreas[0].AxisX.LabelStyle.Angle = -45;
            Font f = new Font("Times New Roman", 12);
            chart1.ChartAreas[0].AxisX.LabelStyle.Font = f;
            chart1.ChartAreas[0].AxisX.Interval = 1;

           // chart1.Series[0].IsValueShownAsLabel = true;
            chart1.ChartAreas[0].Area3DStyle.Enable3D = true;

            StripLine stripline1 = new StripLine();
            stripline1.Interval = 0;
            stripline1.IntervalOffset = 0.3;
            stripline1.StripWidth = 0.005;
            stripline1.BackColor = Color.Yellow;
            chart1.ChartAreas[0].AxisY.StripLines.Add(stripline1);

            StripLine stripline2 = new StripLine();
            stripline2.Interval = 0;
            stripline2.IntervalOffset = 0.5;
            stripline2.StripWidth = 0.005;
            stripline2.BackColor = Color.Red;
            chart1.ChartAreas[0].AxisY.StripLines.Add(stripline2);


        }

        private void button1_Click(object sender, EventArgs e)
        {
            SaveFileDialog d = new SaveFileDialog();
            
            if(d.ShowDialog() == DialogResult.OK)
            {
                chart1.SaveImage(d.FileName +".jpg", ChartImageFormat.Jpeg);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
