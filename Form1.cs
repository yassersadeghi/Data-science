using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace statics
{


    public partial class Form1 : Form
    {
        public static List<personel> P = new List<personel>();
        public static List<string> modules = new List<string>();
        public static int[] chartmodules = new int[35];

        public Form1()
        {
            InitializeComponent();

        }
        public class personel
        {
            public string name;
            public string office;
            public string zone;
            public int atms;
            public int eq;
            public int log;
            public int relog1;
            public int relog;
            public int outrelog;
            public int ftt;
            public int service;
            public int sla;
            public int priod;
            public int module;
            public int logafterpriod;
            public int logafterpriod2;
            public TimeSpan Wtime;
            public TimeSpan Wtimetotal;
            public int Wtimeunstandard;
            public struct m
            {
                public string name;
                public string search;
                public int total;


            }

            public struct w
            {
                public string jobno;
                public TimeSpan wtime;


            }

            public List<m> allmodules = new List<m>();
            public List<fp> logpriods = new List<fp>();
            public List<w> wtimeunstandards = new List<w>();

            public personel(string n, string o, string z, int at)
            {
                name = n;
                office = o;
                zone = z;
                atms = at;
                module = sla = log = 0;
                logpriods = new List<fp>();
                Wtime = new TimeSpan(0, 0, 0);
                Wtimetotal = new TimeSpan(0, 0, 0);
                Wtimeunstandard = 0;

                allmodules.Add(new m { name = "Double Extractor 1&2", search = "Double Extractor, CMD-V4, 1&2 ]", total = 0 });
                allmodules.Add(new m { name = "Double Extractor 3&4", search = "Double Extractor, CMD-V4, 3&4 ]", total = 0 });
                allmodules.Add(new m { name = "Stacker", search = "Stacker, CMD-V4, Xe/USB ]", total = 0 });
                allmodules.Add(new m { name = "Stacker Without Single Reject", search = "Stacker Without Single Reject, CMD-V4, Xe/USB ]", total = 0 });
                allmodules.Add(new m { name = "Clamp", search = "CMD-V4 Clamping Transport Mechanism, Xe/USB", total = 0 });
                allmodules.Add(new m { name = "Shutter RL 2X50", search = "Shutter, CMD-V4 Dispenser Horizontal RL, 2X50", total = 0 });
                allmodules.Add(new m { name = "Clutch Assy", search = "Clutch Assy, Double Extractor Xe/USB", total = 0 });
                allmodules.Add(new m { name = "Cassette Cash Out With LCS", search = "Cassette, Cash Out, With LCS/Without Lock ]", total = 0 });
                allmodules.Add(new m { name = "Cassette Reject Retract", search = "Cassette, Reject-Retract, Without Lock Xe/USB", total = 0 });
                allmodules.Add(new m { name = "CMD-V4 Controller With USB", search = "CMD-V4 Controller, With USB, Xe/USB ]", total = 0 });
                allmodules.Add(new m { name = "Distributor Board 4X", search = "Distributor Board, 4X", total = 0 });
                allmodules.Add(new m { name = "Card Reader CHD-V2X", search = "Card Reader, CHD-V2X ]", total = 0 });
                allmodules.Add(new m { name = "Card Reader CHD-V2XU", search = "Card Reader, CHD-V2XU ]", total = 0 });
                allmodules.Add(new m { name = "Card Reader CHD-V2CU", search = "Card Reader, CHD-V2CU ]", total = 0 });
                allmodules.Add(new m { name = "Journal Printer TP06", search = "Journal Printer, TP06, Xe/USB ]", total = 0 });
                allmodules.Add(new m { name = "Journal Printer NP06", search = "Journal Printer, NP06, Xe/USB ]", total = 0 });
                allmodules.Add(new m { name = "Receipt Printer NP07", search = "Receipt Printer, NP07 ]", total = 0 });
                allmodules.Add(new m { name = "Receipt Printer TP07", search = "Receipt Printer, TP07 ]", total = 0 });
                allmodules.Add(new m { name = "Receipt Printer TP28", search = "Receipt Printer, TP28 ]", total = 0 });
                allmodules.Add(new m { name = "Receipt Printer TP13", search = "Receipt Printer, TP13 ]", total = 0 });
                allmodules.Add(new m { name = "Cutter Assd NP07/TP07", search = "Cutter Assd, NP07/TP07", total = 0 });

                allmodules.Add(new m { name = "PC P4", search = "PC, P4, 2x50 Xe ]", total = 0 });
                allmodules.Add(new m { name = "PC P4 3.4", search = "PC, P4 3.4, 2x50 USB ]", total = 0 });
                allmodules.Add(new m { name = "PC P4 GA-P41T", search = "PC, P4 GA-P41T, 2xx0 Xe/USB ]", total = 0 });
                allmodules.Add(new m { name = "PC 5G", search = "SWAP-PC, 5G Cel-G1820 ProCash TPMen, PC280/PC285 ]", total = 0 });
                allmodules.Add(new m { name = "EPP V5.0", search = "EPP, V5.0", total = 0 });
                allmodules.Add(new m { name = "EPP J6 300", search = "EPP, J6 300", total = 0 });
                allmodules.Add(new m { name = "EPP V7.0", search = "EPP, V7.0", total = 0 });
                allmodules.Add(new m { name = "Monitor 12.1 Sunlight", search = " Monitor, 12.1\" Sunlight", total = 0 });
                allmodules.Add(new m { name = "Monitor 12.1 SVGA", search = "Monitor, 12.1\", SVGA", total = 0 });
                allmodules.Add(new m { name = "Monitor 15 Sinocan", search = " Monitor, 15\" Sinocan T06-15OF With SE-CP, Faradis105", total = 0 });
                allmodules.Add(new m { name = "Monitor 12.1 Sunlight GDS", search = "Monitor, 12.1\" Sunlight GDS", total = 0 });



            }

            public OleDbDataReader read(string command)
            {
                string Constr = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\db.mdb";
                OleDbConnection conn = new OleDbConnection(Constr);
                OleDbCommand comm = new OleDbCommand(command, conn);
                conn.Open();
                OleDbDataReader r = comm.ExecuteReader();
                conn.Close();

                return r;

            }

            public m[] getallmodules()
            {
                string a, command;
                string m = @" 'عادي' ";
                a = " and ";
                command = "select [قطعات مصرفی] from GR where [کارشناس] = '" + name + "' " + a + "[محصول] = 'ATM'" + a + "[شرايط انجام کار] = " + m;
                string Constr = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\db.mdb";
                OleDbConnection conn = new OleDbConnection(Constr);
                conn.Open();
                OleDbDataAdapter da = new OleDbDataAdapter(command, conn);
                DataTable dt = new DataTable();
                da.Fill(dt);
                m[] arr = allmodules.ToArray();
                for (int i = 0; i < dt.Rows.Count; i++)
                    for (int y = 0; y < arr.Length; y++)
                        if (dt.Rows[i][0].ToString().Contains(allmodules[y].search))
                            arr[y].total = arr[y].total + 1;

                conn.Close();
                dt.Dispose();
                da.Dispose();
                conn.Dispose();
                return arr;
            }

            public m[] getallmodulesbyzone()
            {
                string a, command;
                string m = @" 'عادي' ";
                a = " and ";
                command = "select [قطعات مصرفی] from GR where [متولی اقدام] = '" + office + "' and ( [حوزه] = '" + zone + "' or [حوزه] is null )" + a + "[محصول] = 'ATM'";
                string Constr = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\db.mdb";
                OleDbConnection conn = new OleDbConnection(Constr);
                conn.Open();
                OleDbDataAdapter da = new OleDbDataAdapter(command, conn);
                DataTable dt = new DataTable();
                da.Fill(dt);
                m[] arr = allmodules.ToArray();
                for (int i = 0; i < dt.Rows.Count; i++)
                    for (int y = 0; y < arr.Length; y++)
                        if (dt.Rows[i][0].ToString().Contains(allmodules[y].search))
                            arr[y].total = arr[y].total + 1;

                conn.Close();
                dt.Dispose();
                da.Dispose();
                conn.Dispose();
                return arr;
            }
            public void getlogs()
            {
                string a, n, m, b, command;

                n = @"'رفع اشكال'";
                a = " and ";
                m = @" 'عادي' ";
                b = @"'مشتري' , 'گروه بازاريابي و فروش' , 'مديريت روابط و خدمات مشتريان' , 'سامانه مديريت ارتباط با مشتري'";
                command = "select * from GR where [متولی اقدام] = '" + office + "' and ( [حوزه] = '" + zone + "' or [حوزه] is null )  and [نوع کار] = " + n + a + "[محصول] = 'ATM'" + a + "[شرايط انجام کار] = " + m + a + "[درخواست كننده] IN ( " + b + " )";
                string Constr = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\db.mdb";
                OleDbConnection conn = new OleDbConnection(Constr);
                conn.Open();
                OleDbDataAdapter da = new OleDbDataAdapter(command, conn);
                DataTable dt = new DataTable();
                da.Fill(dt);
                log = dt.Rows.Count;
                conn.Close();
                dt.Dispose();
                da.Dispose();
                conn.Dispose();


            }

            public void getwtime()
            {
                string command;
                string jn = @" , [شماره کار]";
                command = "select [مدت اجراء]" + jn + " from GR where [کارشناس] = '" + name + "' ";
                string Constr = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\db.mdb";
                OleDbConnection conn = new OleDbConnection(Constr);
                conn.Open();
                OleDbDataAdapter da = new OleDbDataAdapter(command, conn);
                DataTable dt = new DataTable();
                da.Fill(dt);

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string wt = dt.Rows[i][0].ToString().Trim();
                    string j = dt.Rows[i][1].ToString().Trim();
                    if ((wt != "") && (wt != null) && (wt != "0:0"))
                    {
                        if (wt.Length > 7) wt = "10000:00";

                        Int32  h = System.Convert.ToInt16(wt.Substring(0, wt.IndexOf(':')));
                        Int32 m = System.Convert.ToInt16(wt.Substring(wt.IndexOf(':') + 1, wt.Length - wt.IndexOf(':') - 1));
                        TimeSpan temp = new TimeSpan(h, m, 0);
                        if (h < 12)
                        {

                            Wtimetotal = Wtimetotal.Add(temp);
                            Wtime = Wtime.Add(temp);
                        }
                        else
                        {
                            Wtimetotal = Wtimetotal.Add(temp);
                            Wtimeunstandard++;
                            wtimeunstandards.Add(new w { jobno = j, wtime = temp });
                        }
                    }
                }

                conn.Close();
                dt.Dispose();
                da.Dispose();
                conn.Dispose();


            }

            public void getftt()
            {
                string a, n, m, b, c, command;
                c = @" AND [وضعيت] <> 'Canceled' ";
                n = @"'رفع اشكال'";
                a = " and ";
                m = @" 'عادي' ";
                //  b = @"'مشتري' , 'گروه بازاريابي و فروش' , 'مديريت روابط و خدمات مشتريان' , 'سامانه مديريت ارتباط با مشتري'";
                command = "select * from GR where [کارشناس] = '" + name + "' and [شرايط انجام کار] = " + m + " and [نوع کار] = " + n + a + "[مرجع اقدام] = 'OMS' and [شماره گزارش کار] not like '#%' AND [شماره کار] not like [شماره گزارش کار] " + c;
                string Constr = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\db.mdb";
                OleDbConnection conn = new OleDbConnection(Constr);
                conn.Open();
                OleDbDataAdapter da = new OleDbDataAdapter(command, conn);
                DataTable dt = new DataTable();
                da.Fill(dt);
                ftt = dt.Rows.Count;
                conn.Close();
                dt.Dispose();
                da.Dispose();
                conn.Dispose();

            }

            public void getservice()
            {
                string n, command;
                n = @"'رفع اشكال'";

                command = "select * from GR where [کارشناس] = '" + name + "' and [نوع کار] <> " + n;
                string Constr = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\db.mdb";
                OleDbConnection conn = new OleDbConnection(Constr);
                conn.Open();
                OleDbDataAdapter da = new OleDbDataAdapter(command, conn);
                DataTable dt = new DataTable();
                da.Fill(dt);
                service = dt.Rows.Count;
                conn.Close();
                dt.Dispose();
                da.Dispose();
                conn.Dispose();

            }

            public void getsla()
            {
                string a, n, m, b, command;
                n = @"'رفع اشكال'";
                a = " and ";
                m = @" 'عادي' ";

                b = @"'مشتري' , 'گروه بازاريابي و فروش' , 'مديريت روابط و خدمات مشتريان' , 'سامانه مديريت ارتباط با مشتري'";
                command = "select * from GR where [کارشناس] = '" + name + "' and [نوع کار] = " + n + a + "[محصول] = 'ATM' and " + "[شرايط انجام کار] like " + m + a + "[درخواست كننده] IN ( " + b + " ) and ( [دلیل بروز تاخیر] <> NULL )";
                string Constr = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\db.mdb";
                OleDbConnection conn = new OleDbConnection(Constr);
                conn.Open();
                OleDbDataAdapter da = new OleDbDataAdapter(command, conn);
                DataTable dt = new DataTable();
                da.Fill(dt);
                sla = dt.Rows.Count;
                conn.Close();
                dt.Dispose();
                da.Dispose();
                conn.Dispose();

            }

            public void getpriod()
            {
                string a, n, command;
                n = @"'بازديد دوره‌اي'";
                a = " and ";

                command = "select * from GR where [متولی اقدام] = '" + office + "' and ( [حوزه] = '" + zone + "' or [حوزه] is null )" + " and [نوع کار] = " + n + a + "[محصول] = 'ATM'";
                string Constr = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\db.mdb";
                OleDbConnection conn = new OleDbConnection(Constr);
                conn.Open();
                OleDbDataAdapter da = new OleDbDataAdapter(command, conn);
                DataTable dt = new DataTable();
                da.Fill(dt);
                priod = dt.Rows.Count;
                conn.Close();
                dt.Dispose();
                da.Dispose();
                conn.Dispose();

            }

            public void getrelog()
            {
                string a, n, m, b, command;

                n = @"'رفع اشكال'";
                a = " and ";
                m = @" 'عادي' ";
                b = @"'مشتري' , 'گروه بازاريابي و فروش' , 'مديريت روابط و خدمات مشتريان' , 'سامانه مديريت ارتباط با مشتري'";
                command = "select [تجهیزات شعبه] from GR where [متولی اقدام] = '" + office + "' and ( [حوزه] like '" + zone + "' or [حوزه] is null )  and [نوع کار] = " + n + a + "[محصول] = 'ATM'" + a + "[شرايط انجام کار] like " + m + a + "[درخواست كننده] IN ( " + b + " )  group by [تجهیزات شعبه] having count(*) > 3";
                string Constr = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\db.mdb";
                OleDbConnection conn = new OleDbConnection(Constr);
                conn.Open();
                OleDbDataAdapter da = new OleDbDataAdapter(command, conn);
                DataTable dt = new DataTable();
                da.Fill(dt);
                relog = dt.Rows.Count;
                conn.Close();
                dt.Dispose();
                da.Dispose();
                conn.Dispose();

            }
            public void getrelog1()
            {
                string a, n, m, b, command;

                n = @"'رفع اشكال'";
                a = " and ";
                m = @" 'عادي' ";
                b = @"'مشتري' , 'گروه بازاريابي و فروش' , 'مديريت روابط و خدمات مشتريان' , 'سامانه مديريت ارتباط با مشتري'";
                command = "select [تجهیزات شعبه] from GR where [متولی اقدام] = '" + office + "' and ( [حوزه] like '" + zone + "' or [حوزه] is null )  and [نوع کار] = " + n + a + "[محصول] = 'ATM'" + a + "[شرايط انجام کار] like " + m + a + "[درخواست كننده] IN ( " + b + " )  group by [تجهیزات شعبه] having count(*) > 1";
                string Constr = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\db.mdb";
                OleDbConnection conn = new OleDbConnection(Constr);
                conn.Open();
                OleDbDataAdapter da = new OleDbDataAdapter(command, conn);
                DataTable dt = new DataTable();
                da.Fill(dt);
                relog1 = dt.Rows.Count;
                conn.Close();
                dt.Dispose();
                da.Dispose();
                conn.Dispose();

            }

            public void getmodules()
            {
                string a, n, m, b, command;
                n = @"'رفع اشكال'";
                a = " and ";
                m = @" 'عادي' ";
                b = @"'مشتري' , 'گروه بازاريابي و فروش' , 'مديريت روابط و خدمات مشتريان' , 'سامانه مديريت ارتباط با مشتري'";
                command = "select [قطعات مصرفی] from GR where [متولی اقدام] = '" + office + "' and (( [حوزه] = '" + zone + "') or ( [حوزه] is null) )  and [نوع کار] = " + n + a + "[محصول] = 'ATM'" + a + "[شرايط انجام کار] = " + m + a + "([درخواست كننده]  IN ( " + b + " )) and [قطعات مصرفی] <> NULL ";
                string Constr = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\db.mdb";
                OleDbConnection conn = new OleDbConnection(Constr);
                conn.Open();
                OleDbDataAdapter da = new OleDbDataAdapter(command, conn);
                DataTable dt = new DataTable();
                da.Fill(dt);

                for (int i = 0; i < dt.Rows.Count; i++)
                    foreach (string md in modules)
                        if ((dt.Rows[i][0].ToString() != null) && (dt.Rows[i][0].ToString().Trim().Contains(md)))
                            module = module + 1;


                conn.Close();
                dt.Dispose();
                da.Dispose();
                conn.Dispose();


            }

            public void getrelogout()
            {
                string a, n, m, b, command;
                n = @"'رفع اشكال'";
                a = " and ";
                m = @" 'عادي' ";
                b = @"'مشتري' , 'گروه بازاريابي و فروش' , 'مديريت روابط و خدمات مشتريان' , 'سامانه مديريت ارتباط با مشتري'";
                command = "select [تجهیزات شعبه] from GR where [تجهیزات شعبه] <> null and [متولی اقدام] = '" + office + "' and ( [حوزه] like '" + zone + "' or [حوزه] is null )  and [نوع کار] = " + n + a + "[محصول] = 'ATM'" + a + "[شرايط انجام کار] like " + m + a + "[درخواست كننده] IN ( " + b + " ) group by [تجهیزات شعبه] having count(*) > 1";


                string Constr = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\db.mdb";
                OleDbConnection conn = new OleDbConnection(Constr);
                conn.Open();
                OleDbDataAdapter da = new OleDbDataAdapter(command, conn);
                DataTable dt = new DataTable();
                da.Fill(dt);
                string s = "";
                for (int i = 0; i < dt.Rows.Count; i++)
                    s = s + " '" + dt.Rows[i][0].ToString() + "'   ,";
                if (s != "")
                    s = s.Remove(s.Length - 2, 2);
                if (s == "")
                    s = "'no eq'";

                dt.Dispose();
                da.Dispose();

                conn.Close();
                conn.Open();
                command = "select * from GR where ( [متولی اقدام] = '" + office + "') and ( [حوزه] like '" + zone + "' or [حوزه] is null )  and [نوع کار] = " + n + a + "[محصول] = 'ATM'" + a + "[شرايط انجام کار] like " + m + a + "[درخواست كننده] IN ( " + b + " ) and ( [تجهیزات شعبه] IN (" + s + "))";
                OleDbDataAdapter DA = new OleDbDataAdapter(command, conn);
                DataTable DT = new DataTable();
                DA.Fill(DT);
                conn.Close();

                List<fp> lp = new List<fp>();

                for (int i = 0; i < DT.Rows.Count; i++)
                    if (DT.Rows[i]["کارشناس"].ToString() != name)
                    {
                        fp p = new fp();
                        p.name = DT.Rows[i]["کارشناس"].ToString();
                        p.job = DT.Rows[i]["شماره کار"].ToString();
                        p.type = DT.Rows[i]["شرح خرابی"].ToString();
                        p.unit = DT.Rows[i]["تجهیزات شعبه"].ToString();
                        p.date = DT.Rows[i]["زمان پذيرش توسط شرکت"].ToString();
                        lp.Add(p);
                    }


                for (int i = 0; i < DT.Rows.Count; i++)
                    foreach (fp p in lp)
                        if ((p.job != DT.Rows[i]["شماره کار"].ToString()) && (p.type == DT.Rows[i]["شرح خرابی"].ToString()) && (p.unit == DT.Rows[i]["تجهیزات شعبه"].ToString()))
                        {
                            PersianDateTime pd1 = new PersianDateTime(System.Convert.ToInt32(p.date.Substring(0, 4)), System.Convert.ToInt32(p.date.Substring(5, 2)), System.Convert.ToInt32(p.date.Substring(8, 2)));
                            PersianDateTime pd2 = new PersianDateTime(System.Convert.ToInt32(DT.Rows[i]["زمان پذيرش توسط شرکت"].ToString().Substring(0, 4)), System.Convert.ToInt32(DT.Rows[i]["زمان پذيرش توسط شرکت"].ToString().Substring(5, 2)), System.Convert.ToInt32(DT.Rows[i]["زمان پذيرش توسط شرکت"].ToString().Substring(8, 2)));
                            PersianDateTime pd3 = pd1.AddDays(3);

                            //  if ((-2 < DateTime.Compare(pd1.ToDateTime(), pd2.ToDateTime())) && (DateTime.Compare(pd1.ToDateTime(), pd2.ToDateTime()) < 1))
                            if ((pd1.ToDateTime() < pd2.ToDateTime()) && (pd2.ToDateTime() <= pd3.ToDateTime()))
                                foreach (personel pr in P)
                                    if (pr.name == p.name)
                                        pr.outrelog = pr.outrelog + 1;

                        }



            }

            public void getlogafterpriod()
            {
                string a, n, m, b, l, command;

                n = @"'بازديد دوره‌اي'";
                l = @"'رفع اشكال'";
                a = " and ";
                m = @" 'عادي' ";
                b = @"'مشتري' , 'گروه بازاريابي و فروش' , 'مديريت روابط و خدمات مشتريان' , 'سامانه مديريت ارتباط با مشتري'";
                command = "select [تجهیزات شعبه] from GeneralReport where [تجهیزات شعبه] <> null and [متولی اقدام] = '" + office + "' and ( [حوزه] like '" + zone + "' or [حوزه] is null )  and [نوع کار] = " + n + a + "[محصول] = 'ATM'" + a + "[شرايط انجام کار] like " + m + a + "[درخواست كننده] IN ( " + b + " ) ";


                string Constr = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\db.mdb";
                OleDbConnection conn = new OleDbConnection(Constr);
                conn.Open();
                OleDbDataAdapter da = new OleDbDataAdapter(command, conn);
                DataTable dt = new DataTable();
                da.Fill(dt);
                string s = "";
                for (int i = 0; i < dt.Rows.Count; i++)
                    s = s + " '" + dt.Rows[i][0].ToString() + "'   ,";
                if (s != "")
                    s = s.Remove(s.Length - 2, 2);
                if (s == "")
                    s = "'no eq'";

                dt.Dispose();
                da.Dispose();
                conn.Close();

                conn.Open();
                command = "select * from GR where ( [متولی اقدام] = '" + office + "') and ( [حوزه] like '" + zone + "' or [حوزه] is null )  and [نوع کار] = " + l + a + "[محصول] = 'ATM'" + a + "[شرايط انجام کار] like " + m + a + "[درخواست كننده] IN ( " + b + " ) and ( [تجهیزات شعبه] IN (" + s + "))";
                OleDbDataAdapter DA = new OleDbDataAdapter(command, conn);
                DataTable DT = new DataTable();
                DA.Fill(DT);
                conn.Close();

                List<fp> lp = new List<fp>();

                for (int i = 0; i < DT.Rows.Count; i++)
                // if (DT.Rows[i]["کارشناس"].ToString() != name)
                {
                    fp p = new fp();
                    p.name = DT.Rows[i]["کارشناس"].ToString();
                    p.job = DT.Rows[i]["شماره کار"].ToString();
                    p.type = DT.Rows[i]["شرح خرابی"].ToString();
                    p.unit = DT.Rows[i]["تجهیزات شعبه"].ToString();
                    p.date = DT.Rows[i]["زمان پذيرش توسط شرکت"].ToString();
                    lp.Add(p);
                }

                command = "select * from GeneralReport where [تجهیزات شعبه] <> null and [متولی اقدام] = '" + office + "' and ( [حوزه] like '" + zone + "' or [حوزه] is null )  and [نوع کار] = " + n + a + "[محصول] = 'ATM'" + a + "[شرايط انجام کار] like " + m + a + "[درخواست كننده] IN ( " + b + " ) ";

                conn.Open();
                OleDbDataAdapter DA1 = new OleDbDataAdapter(command, conn);
                DataTable DS = new DataTable();
                DA1.Fill(DS);



                for (int i = 0; i < DS.Rows.Count; i++)
                    foreach (fp p in lp)
                        if ((p.job != DS.Rows[i]["شماره کار"].ToString()) && (p.unit == DS.Rows[i]["تجهیزات شعبه"].ToString()) && (DS.Rows[i]["زمان مراجعه"].ToString().Length > 8))
                        {
                            PersianDateTime pd1 = new PersianDateTime(System.Convert.ToInt32(p.date.Substring(0, 4)), System.Convert.ToInt32(p.date.Substring(5, 2)), System.Convert.ToInt32(p.date.Substring(8, 2)));
                            PersianDateTime pd2 = new PersianDateTime(System.Convert.ToInt32(DS.Rows[i]["زمان مراجعه"].ToString().Substring(0, 4)), System.Convert.ToInt32(DS.Rows[i]["زمان مراجعه"].ToString().Substring(5, 2)), System.Convert.ToInt32(DS.Rows[i]["زمان مراجعه"].ToString().Substring(8, 2)));
                            PersianDateTime pd3 = pd1.AddDays(-20);

                            //  if ((-2 < DateTime.Compare(pd1.ToDateTime(), pd2.ToDateTime())) && (DateTime.Compare(pd1.ToDateTime(), pd2.ToDateTime()) < 1))
                            if ((pd1.ToDateTime() > pd2.ToDateTime()) && (pd2.ToDateTime() > pd3.ToDateTime()))
                            {
                                logafterpriod = logafterpriod + 1;
                                logpriods.Add(p);
                                PersianDateTime pdEdge = new PersianDateTime(Form5.year0, Form5.month0, 21);
                                if ((pd2.ToDateTime() < pdEdge.ToDateTime()) && (DS.Rows[i]["ملاحظات"].ToString().Trim() != "counted"))
                                {
                                    priod = priod + 1;
                                    DS.Rows[i]["ملاحظات"] = "counted";

                                }
                            }

                        }
                conn.Close();
            }

        }

        public string Constr;
        private void bindingSource1_CurrentChanged(object sender, EventArgs e)
        {


        }


        private void Form1_Load(object sender, EventArgs e)
        {

            menuStrip1.Items[2].Enabled = false;
            toolStripMenuItem2.Enabled = false;
            this.WindowState = FormWindowState.Maximized;

            modules.Add("Double Extractor, CMD-V4, 1&2 ]");
            modules.Add("Double Extractor, CMD-V4, 3&4 ]");

            modules.Add("Stacker, CMD-V4, Xe/USB ]");
            modules.Add("Stacker Without Single Reject, CMD-V4, Xe/USB ]");

            modules.Add("Cassette, Cash Out, With LCS/Without Lock ]");
            modules.Add("Cassette, Cash Out, Without LCS/Without Lock ]");

            modules.Add("Card Reader, CHD-V2X ]");
            modules.Add("Card Reader, CHD-V2XU ]");
            modules.Add("Card Reader, CHD-V2CU ]");

            modules.Add("Journal Printer, TP06, Xe/USB ]");
            modules.Add("Journal Printer, NP06, Xe/USB ]");

            modules.Add("Receipt Printer, NP07 ]");
            modules.Add("Receipt Printer, TP07 ]");
            modules.Add("Receipt Printer, TP28 ]");
            modules.Add("Receipt Printer, TP13 ]");

            modules.Add("CMD-V4 Controller, With USB, Xe/USB ]");

            modules.Add("PC, P4, 2x50 Xe ]");
            modules.Add("PC, P4 3.4, 2x50 USB ]");
            modules.Add("PC, P4 GA-P41T, 2xx0 Xe/USB ]");
            modules.Add("SWAP-PC, 5G Cel-G1820 ProCash TPMen, PC280/PC285 ]");




            Constr = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\db.mdb";



            OleDbConnection conn = new OleDbConnection(Constr);
            string d = @"[متولی اقدام]";
            OleDbCommand comm = new OleDbCommand("UPDATE GeneralReport SET [حوزه] = 'NONE' WHERE ((( [حوزه]  IS NULL ) OR ( [حوزه] = '' )  ) AND (" + d + " = 'دفتر استاني فارس')) ", conn);
            conn.Open();
            comm.ExecuteNonQuery();
            conn.Close();
            string bazdid = "بازديد دوره‌اي";
            comm.CommandText = "DELETE FROM GeneralReport WHERE [نوع کار] = '" + bazdid + "' AND ( [شماره گزارش کار] like '#%' OR [شماره کار]  like [شماره گزارش کار] )";
            conn.Open();
            comm.ExecuteNonQuery();
            conn.Close();

            OleDbDataAdapter da = new OleDbDataAdapter("select * from GeneralReport", conn);
            conn.Open();

            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.AutoGenerateColumns = true;
            dataGridView1.DataSource = dt;

            conn.Close();




        }

        private void آمارعملکردToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Constr = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\db.mdb";

            OleDbConnection conn3 = new OleDbConnection(Constr);
            OleDbCommand comm3 = new OleDbCommand("UPDATE GeneralReport SET [ملاحظات] = '' ", conn3);
            conn3.Open();
            comm3.ExecuteNonQuery();
            conn3.Close();

            OleDbConnection conn = new OleDbConnection(Constr);
            OleDbDataAdapter da = new OleDbDataAdapter("select * from personnel", conn);
            conn.Open();
            DataTable dt = new DataTable();
            da.Fill(dt);

            P = new List<personel>();

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                P.Add(new personel(dt.Rows[i][1].ToString(), dt.Rows[i][2].ToString(), dt.Rows[i][3].ToString(), System.Convert.ToInt32(dt.Rows[i][4].ToString().Trim())));
            }

            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();
            dataGridView1.Columns.Add("name", " نام کارشناس");
            dataGridView1.Columns.Add("zone", "  منطقه حوزه پوشش");
            dataGridView1.Columns.Add("log", "تعداد اعلام خرابی");
            dataGridView1.Columns.Add("relog", "تعداد اعلام خرابی بالاتر از 3 مورد");
            dataGridView1.Columns.Add("sla", "تعداد رفع اشکالهای خارج از SLA");
            dataGridView1.Columns.Add("service", " تعداد کارهای غیر از رفع اشکال خودپرداز");
            dataGridView1.Columns.Add("priod", "بازدید ادواری");
            dataGridView1.Columns.Add("outrelog", "تکرار خرابی خارج از منطقه تحت پوشش");
            dataGridView1.Columns.Add("ftt", " تعداد لاگهای بسته شده در OMS");
            dataGridView1.Columns.Add("module", "تعداد قطعات مصرفی");
            dataGridView1.Columns.Add("logpriod", "تعداد اعلام خرابی بعد از بازدید");
            dataGridView1.Columns.Add("wtimetotal", " مجموع کارکرد");
            dataGridView1.Columns.Add("wtime", " مجموع کارکرد استاندارد");
            dataGridView1.Columns.Add("wtimeunst", " تعداد موارد کارکرد غیراستاندارد");
            dataGridView1.Columns.Add("relog1", "تعداد اعلام خرابی بالاتر از یک مورد");
            dataGridView1.Columns[0].Width = 200;
            dataGridView1.Columns[1].Width = 200;
            dataGridView1.Columns[3].Width = 150;
            dataGridView1.Columns[4].Width = 150;
            dataGridView1.Columns[5].Width = 200;
            dataGridView1.Columns[6].Width = 120;
            dataGridView1.Columns[7].Width = 150;
            dataGridView1.Columns[8].Width = 150;
            dataGridView1.Columns[10].Width = 150;
            dataGridView1.Columns[14].Width = 150;
            dataGridView1.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[7].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[8].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[9].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[10].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[11].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[12].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[13].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[14].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[3].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[4].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[5].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[6].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[7].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[8].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[9].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[10].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[11].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[12].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[13].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[14].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            for (int i = 2; i < dataGridView1.ColumnCount; i++)
            {
                dataGridView1.Columns[i].DefaultCellStyle.Font = new Font("Arial", 18F, GraphicsUnit.Pixel);
            }


            foreach (personel p in P)
            {
                p.getftt();
                p.getlogs();
                p.getpriod();
                p.getrelog();
                p.getrelog1();
                p.getservice();
                p.getsla();
                p.getrelogout();
                p.getmodules();
                p.getlogafterpriod();
                p.getwtime();
                DataGridViewRow R = (DataGridViewRow)dataGridView1.Rows[0].Clone();
                R.Cells[0].Value = p.name;
                R.Cells[1].Value = (p.zone + " " + p.office).Trim();
                if ((p.office == "دفتر اقماري زند 1") || p.office == "دفتر اقماري شيراز شمال")
                    R.Cells[1].Value = (p.office).Trim();
                R.Cells[2].Value = p.log;
                R.Cells[3].Value = p.relog;
                R.Cells[4].Value = p.sla;
                R.Cells[5].Value = p.service;
                R.Cells[6].Value = p.priod;
                R.Cells[7].Value = p.outrelog;
                R.Cells[8].Value = p.ftt;
                R.Cells[9].Value = p.module;
                R.Cells[10].Value = p.logafterpriod;
                string zero = "";
                if (p.Wtimetotal.Minutes < 10) zero = "0";
                R.Cells[11].Value = ((p.Wtimetotal.Days * 24) + p.Wtimetotal.Hours).ToString() + ":" + zero + p.Wtimetotal.Minutes.ToString();


                zero = "";
                if (p.Wtime.Minutes < 10) zero = "0";
                R.Cells[12].Value = ((p.Wtime.Days * 24) + p.Wtime.Hours).ToString() + ":" + zero + p.Wtime.Minutes.ToString();
                R.Cells[13].Value = p.Wtimeunstandard;
                R.Cells[14].Value = p.relog1;
                dataGridView1.Rows.Add(R);
            }

            conn.Close();
            menuStrip1.Items[2].Enabled = true;
        }

        private void آمارToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void آمارقطعاتمصرفیToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Constr = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\db.mdb";
            OleDbConnection conn = new OleDbConnection(Constr);
            OleDbDataAdapter da = new OleDbDataAdapter("select * from personnel", conn);
            conn.Open();
            DataTable dt = new DataTable();
            da.Fill(dt);

            List<personel> pr = new List<personel>();

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                pr.Add(new personel(dt.Rows[i][1].ToString(), dt.Rows[i][2].ToString(), dt.Rows[i][3].ToString(), System.Convert.ToInt32(dt.Rows[i][4].ToString().Trim())));
            }

            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();
            dataGridView1.Columns.Add("name", " نام کارشناس");
            dataGridView1.Columns.Add("zone", "  منطقه تحت پوشش");

            foreach (personel.m mod in pr[0].allmodules)
                dataGridView1.Columns.Add(mod.name, mod.name);

            dataGridView1.Columns.Add("TOTAL", "TOTAL");

            /*  
              dataGridView1.Columns.Add("log", "تعداد اعلام خرابی");
              dataGridView1.Columns.Add("relog", "تعداد اعلام خرابی بالاتر از 4 مورد");
              dataGridView1.Columns.Add("sla", "تعداد رفع اشکالهای خارج از SLA");
              dataGridView1.Columns.Add("service", " تعداد خدمات موردی");
              dataGridView1.Columns.Add("priod", "بازدید ادواری");
              dataGridView1.Columns.Add("outrelog", "تکرار خرابی خارج از منطقه تحت پوشش");
              dataGridView1.Columns.Add("ftt", " تعداد لاگهای بسته شده در OMS");
              dataGridView1.Columns.Add("module", "تعداد قطعات مصرفی");
              */

            //  for (int j = 0; j < chartmodules.Length; j++)
            //    chartmodules[j] = 0;


            foreach (personel p in pr)
            {
                int t = 0;
                personel.m[] pt = p.getallmodules();
                DataGridViewRow R = (DataGridViewRow)dataGridView1.Rows[0].Clone();
                R.Cells[0].Value = p.name;
                R.Cells[1].Value = (p.zone + " " + p.office).Trim();
                if ((p.office == "دفتر اقماري زند 1") || p.office == "دفتر اقماري شيراز شمال")
                    R.Cells[1].Value = (p.office).Trim();
                for (int i = 0; i < pt.Length; i++)
                {


                    R.Cells[i + 2].Value = pt[i].total;
                    t = t + pt[i].total;

                }

                R.Cells[R.Cells.Count - 1].Value = t;

                dataGridView1.Rows.Add(R);
            }




            dataGridView1.Columns[0].Width = 200;
            dataGridView1.Columns[1].Width = 200;
            dataGridView1.Columns[5].Width = 110;
            for (int i = 2; i < dataGridView1.Columns.Count; i++)
            {
                dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView1.Columns[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView1.Columns[i].DefaultCellStyle.Font = new Font("Arial", 18F, GraphicsUnit.Pixel);

            }

        }

        private void خروجToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ذخیرهسازیToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count < 30)
            {
                Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
                Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
                app.Visible = true;
                worksheet = workbook.Sheets["Sheet1"];
                worksheet = workbook.ActiveSheet;
                worksheet.Name = "Exported from gridview";
                for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                {
                    worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                }

                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                    }
                }

                app.Quit();
            }
        }

        private void dataGridView1_ColumnSortModeChanged(object sender, DataGridViewColumnEventArgs e)
        {

        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Constr = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\db.mdb";
            OleDbConnection conn = new OleDbConnection(Constr);
            OleDbDataAdapter da = new OleDbDataAdapter("select * from personnel", conn);
            conn.Open();
            DataTable dt = new DataTable();
            da.Fill(dt);

            List<personel> pr = new List<personel>();

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                pr.Add(new personel(dt.Rows[i][1].ToString(), dt.Rows[i][2].ToString(), dt.Rows[i][3].ToString(), System.Convert.ToInt32(dt.Rows[i][4].ToString().Trim())));
            }

            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();
            dataGridView1.Columns.Add("name", " نام کارشناس");
            dataGridView1.Columns.Add("zone", "  منطقه تحت پوشش");

            foreach (personel.m mod in pr[0].allmodules)
                dataGridView1.Columns.Add(mod.name, mod.name);

            dataGridView1.Columns.Add("TOTAL", "TOTAL");



            for (int j = 0; j < chartmodules.Length; j++)
                chartmodules[j] = 0;

            foreach (personel p in pr)
            {
                int t = 0;
                personel.m[] pt = p.getallmodulesbyzone();
                DataGridViewRow R = (DataGridViewRow)dataGridView1.Rows[0].Clone();
                R.Cells[0].Value = p.name;
                R.Cells[1].Value = (p.zone + " " + p.office).Trim();
                if ((p.office == "دفتر اقماري زند 1") || p.office == "دفتر اقماري شيراز شمال")
                    R.Cells[1].Value = (p.office).Trim();
                for (int i = 0; i < pt.Length; i++)
                {
                    R.Cells[i + 2].Value = pt[i].total;
                    t = t + pt[i].total;
                    chartmodules[i] = chartmodules[i] + pt[i].total;
                }

                R.Cells[R.Cells.Count - 1].Value = t;

                dataGridView1.Rows.Add(R);
            }




            dataGridView1.Columns[0].Width = 200;
            dataGridView1.Columns[1].Width = 200;
            dataGridView1.Columns[5].Width = 110;
            for (int i = 2; i < dataGridView1.Columns.Count; i++)
            {
                dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView1.Columns[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView1.Columns[i].DefaultCellStyle.Font = new Font("Arial", 18F, GraphicsUnit.Pixel);

            }


            toolStripMenuItem2.Enabled = true;
        }

        private void نمودارنسبتخرابیToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form2 f = new statics.Form2();
            f.Show();
        }

        private void نمودارToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void نمودارمصرفقطعهToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form3 f = new statics.Form3();
            f.Show();

        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            Form4 f = new statics.Form4();
            f.Show();
        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            Form5 f = new statics.Form5();
            f.StartPosition = FormStartPosition.CenterScreen;
            f.Show();
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if ((dataGridView1.ColumnCount < 20) && dataGridView1.CurrentCell.ColumnIndex == 10)
            {
                string person = dataGridView1.CurrentRow.Cells[0].Value.ToString().Trim();
                var query = from c in P
                            where c.name == person
                            select c;



                foreach (personel p in query)
                {
                    string[,] s = new string[p.logpriods.Count + 1, 4];

                    for (int j = 0; j < p.logpriods.Count; j++)
                    {
                        s[j, 0] = p.logpriods[j].name;
                        s[j, 1] = p.logpriods[j].job;
                        s[j, 2] = p.logpriods[j].unit;
                        s[j, 3] = p.logpriods[j].date;
                    }


                    Form6 f6 = new Form6(p.office, s, p.logpriods.Count);
                    f6.Show();

                }


            }

            if ((dataGridView1.ColumnCount < 20) && dataGridView1.CurrentCell.ColumnIndex == 13)
            {
                string person = dataGridView1.CurrentRow.Cells[0].Value.ToString().Trim();
                var query = from c in P
                            where c.name == person
                            select c;



                foreach (personel p in query)
                {
                    string[,] s = new string[p.wtimeunstandards.Count + 1, 2];

                    for (int j = 0; j < p.wtimeunstandards.Count; j++)
                    {
                        s[j, 1] = p.wtimeunstandards[j].wtime.ToString();
                        s[j, 0] = p.wtimeunstandards[j].jobno;

                    }

                    Form7 f7 = new Form7(p.name, s , p.wtimeunstandards.Count );
                    f7.Show();

                }
            }
        }

        public struct fp
        {
            public string name;
            public string job;
            public string type;
            public string unit;
            public string date;
        }


    }
}
