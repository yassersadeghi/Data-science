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

namespace statics
{
    public partial class Form5 : Form
    {
        public Form5()
        {
            InitializeComponent();
        }
        public bool combo1, combo2;
        public static  int year,year0, month0, month1;
        public string syear ,syear0, smonth0 , smonth1;
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            combo1 = true;
            if (combo1 && combo2)
                button1.Enabled = true;
            
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            combo2 = true;
            if (combo1 && combo2)
                button1.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            year = System.Convert.ToInt32(comboBox2.Text.Trim().ToString());
            syear = comboBox2.Text.Trim().ToString();
            year0 = year;
            

            month1 = comboBox1.SelectedIndex + 1;
            if (month1 == 1)
            {
                month0 = 12;
                year0 = year - 1;
            }
            else
                month0 = month1 - 1;

            if (month0 < 10)
                smonth0 = "0" + month0.ToString().Trim();
            else
                smonth0= month0.ToString().Trim();

            if (month1 < 10)
                smonth1 = "0" + month1.ToString().Trim();
            else
                smonth1 = month1.ToString().Trim();

            syear0 = year0.ToString().Trim();

            // MessageBox.Show(syear + "/" + smonth0 + " to " + smonth1);


            string Constr = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\db.mdb";

            

            OleDbConnection conn = new OleDbConnection(Constr);
         
           OleDbCommand comm = new OleDbCommand("DELETE * FROM GR", conn);
            conn.Open();
            comm.ExecuteNonQuery();
            conn.Close();
            string c = @"  [وضعيت] <> 'Canceled' AND ";
            OleDbCommand comm2 = new OleDbCommand("INSERT INTO GR SELECT * FROM GeneralReport WHERE "+c + statemenet("[زمان مراجعه]", syear0.Trim() , syear.Trim() , smonth0.Trim(), smonth1.Trim()), conn);
            conn.Open();
            comm2.ExecuteNonQuery();
            conn.Close();


            
           
            this.Close();
        }

        private void Form5_Load(object sender, EventArgs e)
        {
            
            button1.Enabled = false;
            combo1 = combo2 = false;
        }

        private string statemenet(string f , string y0 ,  string y , string m0 , string m1)
        {
            string s="";
            s =@" ( "+ f+ " NOT LIKE '%" + y0 + "/" + m0 + "/" + "20%' ) AND (";
            s = s + @" ( " + f + " LIKE '%" + y0 + "/" + m0 + "/" + "2%' ) OR ";
            s = s + @" ( " + f + " LIKE '%" + y0 + "/" + m0 + "/" + "3%' ) OR ";
            s = s + @" ( " + f + " LIKE '%" + y + "/" + m1 + "/" + "0%' ) OR ";
            s = s + @" ( " + f + " LIKE '%" + y + "/" + m1 + "/" + "1%' ) OR ";
            s = s + @" ( " + f + " LIKE '%" + y + "/" + m1 + "/" + "20%' ) )";

            return s;
        }
    }
}
