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
    public partial class Form4 : Form
    {
        public Form4()
        {
            InitializeComponent();
        }

        private void Form4_Load(object sender, EventArgs e)
        {
            int de, st, clamp, shutter, clutch, cassette, zak, d_board, crdr, jprn, rprn, cutter, pc, epp;
            int monitor = 0;
            de = st = clamp = shutter = clutch = cassette = zak = d_board = crdr = jprn = rprn = cutter = pc = epp = 0;
         
            
                
            de =( de + Form1.chartmodules[0]);
            de =( de + Form1.chartmodules[1]);
            st = st + Form1.chartmodules[2];
            st = st + Form1.chartmodules[3];
            clamp =  Form1.chartmodules[4];
            shutter =  Form1.chartmodules[5];
            clutch= Form1.chartmodules[6];
            cassette= Form1.chartmodules[7]+ Form1.chartmodules[8];
            zak= Form1.chartmodules[9];
            d_board= Form1.chartmodules[10];
            crdr= Form1.chartmodules[11]+ Form1.chartmodules[12]+ Form1.chartmodules[13];
            jprn= Form1.chartmodules[14]+ Form1.chartmodules[15];
            rprn= Form1.chartmodules[16]+ Form1.chartmodules[17]+ Form1.chartmodules[18]+ Form1.chartmodules[19];
            cutter= Form1.chartmodules[20];
            pc= Form1.chartmodules[21]+ Form1.chartmodules[22]+ Form1.chartmodules[23]+ Form1.chartmodules[24];
            epp= Form1.chartmodules[25]+ Form1.chartmodules[26]+ Form1.chartmodules[27];
            monitor= Form1.chartmodules[28]+ Form1.chartmodules[29]+ Form1.chartmodules[30]+ Form1.chartmodules[31];

            chart1.Series[0].Name = "modules";

           // chart1.Series[0].SmartLabelStyle.Enabled = true;
            chart1.Series[0].IsValueShownAsLabel = true;
            chart1.ChartAreas[0].Area3DStyle.Enable3D = true;
            

           // chart1.ChartAreas[0].Area3DStyle.PointDepth = 4;
            chart1.Series["modules"].Points.AddXY("Double Extractor", de);
            chart1.Series["modules"].Points.AddXY("Stacker", st);
            chart1.Series["modules"].Points.AddXY("Clamp", clamp);
            chart1.Series["modules"].Points.AddXY("Shutter", shutter);
            chart1.Series["modules"].Points.AddXY("Clutch", clutch);
            chart1.Series["modules"].Points.AddXY("Cassette", cassette);
            chart1.Series["modules"].Points.AddXY("ZAC", zak);
            chart1.Series["modules"].Points.AddXY("Distributer Board", d_board);
            chart1.Series["modules"].Points.AddXY("Card Reader", crdr);
            chart1.Series["modules"].Points.AddXY("Journal Printer", jprn);
            chart1.Series["modules"].Points.AddXY("Receipt Printer", rprn);
            chart1.Series["modules"].Points.AddXY("Cutter", cutter);
            chart1.Series["modules"].Points.AddXY("PC", pc);
            chart1.Series["modules"].Points.AddXY("EPP", epp);
            chart1.Series["modules"].Points.AddXY("Monitor", monitor);


            chart1.Series["modules"]["PixelPointWidth"] = "23";
            chart1.ChartAreas[0].AxisX.LabelStyle.Angle = 30;
            Font f = new Font("Times New Roman", 8);
            
            chart1.ChartAreas[0].AxisX.LabelStyle.Font = f;
            chart1.ChartAreas[0].AxisX.Interval = 1;
            chart1.ChartAreas[0].AxisY.Interval = 5;
            chart1.ChartAreas[0].AxisX.LabelStyle.ForeColor = Color.Blue;

        
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
