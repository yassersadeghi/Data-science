using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace statics
{
    public partial class Form7 : Form
    {
       

        public Form7(string name , string [,] wtimeunstandards , int I)
        {
            InitializeComponent();
            this.Text += "   " + name;

            dataGridView1.Columns.Add("job", "شماره کار");
            dataGridView1.Columns.Add("time", "کارکرد");

            if (wtimeunstandards.Length > 0)
                for (int j = 0; j < I ; j++)
                {
                    DataGridViewRow R = (DataGridViewRow)dataGridView1.Rows[0].Clone();
                    R.Cells[0].Value = wtimeunstandards[j,0];
                    R.Cells[1].Value = wtimeunstandards[j,1];
                    
                    dataGridView1.Rows.Add(R);
                }

        }
    }
}
