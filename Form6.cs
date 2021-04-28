using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Data.SqlTypes;

namespace statics
{
    public partial class Form6 : Form
    {
        public  struct fp
        {
            public string name;
            public string job;
            public string type;
            public string unit;
            public string date;
        }
        private string office;
        private string[,] logs;
        private int I;
        public Form6(string o , string[,]  f ,int i)
        {
            InitializeComponent();
            logs  = f;
            office = o;
            I = i;

        }

        private void Form6_Load(object sender, EventArgs e)
        {
            this.Text = office +" : لیست اعلام خرابیهای بعد از بازدید در بازه بیست روزه";
            dataGridView1.Columns.Add("name", "کارشناس مراجعه کننده");
            dataGridView1.Columns.Add("job", "شماره کار");
            dataGridView1.Columns.Add("unit", "تجهیزات");
            dataGridView1.Columns.Add("date", "تاریخ و ساعت");
            dataGridView1.Columns[0].Width = 180;
            dataGridView1.Columns[2].Width = 370;
            dataGridView1.Columns[1].Width = 140;
            dataGridView1.Columns[3].Width = 140;
            if (logs.Length >0)
            for(int j=0; j < I  ; j++)
            {
                DataGridViewRow R = (DataGridViewRow)dataGridView1.Rows[0].Clone();
                R.Cells[0].Value = logs[j,0];
                R.Cells[1].Value = logs[j, 1];
                R.Cells[2].Value = logs[j, 2];
                R.Cells[3].Value = logs[j, 3];
                dataGridView1.Rows.Add(R);
            }
            
        }
    }
}
