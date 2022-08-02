using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;




namespace 熱處理歷史資料讀取作業
{
    public partial class Form1 : Form
    {

        public static DateTime minTime = DateTime.Now ;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            radioButton1.Checked = true;
            timer1.Enabled = true;
            dateTimePicker1.Value = DateTime.Today.AddDays(-1);

            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "yyyy/MM/dd";

            


        }
        
        private void button1_Click(object sender, EventArgs e)
        {

            DataTable DataTable = new DataTable();
            DataTable.Rows.Clear();
            dataGridView1.DataSource = DataTable;

            
            minTime = dateTimePicker1.Value;

            
       
          
                try
                {
                    if (radioButton3.Checked == true)
                    {
                        string[] dirs = new string[0];
                        if (comboBox1.Text == "熱處理1")
                        {
                            dirs = Directory.GetFiles(@"\\192.168.18.1\s7034be\MDB\" + minTime.ToString("yyyy") + @"\" + minTime.ToString("MM") + @"\", "HIS" + minTime.ToString("dd") + ".mdb");
                        }
                        else if (comboBox1.Text == "熱處理2")
                        {
                            dirs = Directory.GetFiles(@"\\192.168.18.2\s7033be\MDB\" + minTime.ToString("yyyy") + @"\" + minTime.ToString("MM") + @"\", "HIS" + minTime.ToString("dd") + ".mdb");

                        }
                        string file = @" ";

                        foreach (var HIS in dirs)
                        {
                            file = HIS.ToString();
                            label3.Text = HIS.ToString();
                        }

                        //Console.WriteLine(file);

                        string l_provider = "Microsoft.Jet.OLEDB.4.0";

                        if (radioButton1.Checked == true)
                            l_provider = "Microsoft.Jet.OLEDB.4.0";
                        else if (radioButton2.Checked == true)
                            l_provider = "Microsoft.ACE.OLEDB.12.0";

                        OleDbConnection connection = new OleDbConnection("Provider=" + l_provider + "; Data Source=" + file + ";jet oledb:database password=;");

                        connection.Open();
                        //MessageBox.Show("開啟動作後，連結狀態：" + connection.State.ToString(), "連結狀態");

                        String queryString;
                        string[] a = new string[0];

                        queryString = "Select * from HISTORY where M_NO is not null and [L-T] like '%00' order by [L-T];";



                        OleDbCommand command = new OleDbCommand(queryString, connection);
                        OleDbDataAdapter dataAdpter = new OleDbDataAdapter(queryString, connection);
                        DataSet DataSetValue = new DataSet();
                        DataSetValue.Clear();
                        dataAdpter.Fill(DataSetValue);

                        DataTable = new DataTable();

                        //Console.WriteLine(DataSetValue.Tables[0].Columns[0].Caption);


                        for (int i = 0; i < DataSetValue.Tables[0].Columns.Count; i++)
                        {
                            DataTable.Columns.Add(DataSetValue.Tables[0].Columns[i].Caption);

                        }

                        //Console.WriteLine(DataSetValue.Tables[0].Columns.Count);
                        Array.Resize(ref a, DataSetValue.Tables[0].Columns.Count);

                        for (int i = 0; i < DataSetValue.Tables[0].Rows.Count; i++)
                        {
                            for (int j = 0; j < DataSetValue.Tables[0].Columns.Count; j++)
                            {
                                //Console.WriteLine(DataSetValue.Tables[0].Rows[i][DataSetValue.Tables[0].Columns[j].Caption].ToString());
                                a[j] = DataSetValue.Tables[0].Rows[i][DataSetValue.Tables[0].Columns[j].Caption].ToString();
                                if (j == 0)
                                {
                                    DateTime DateObject = DateTime.Parse(a[0]);
                                    a[0] = DateObject.ToString("yyyy/MM/dd HH:mm:ss");
                                    
                                }

                            }
                            DataTable.Rows.Add(a);
                        }

                        

                        dataGridView1.DataSource = DataTable;
                        dataGridView1.AllowUserToAddRows = false;

                        //自動調整寬度
                        dataGridView1.AutoResizeColumns();

                        //queryString = "INSERT INTO DATA ([Value1], [Value2]) VALUES ('" + Value1.Text + "','" + Value2.Text + "')";
                        //command = new OleDbCommand(queryString, connection);
                        //Console.WriteLine(dataGridView1.RowCount);

                        connection.Close();
                    }
                    else if (radioButton4.Checked == true)
                    {

                        string[] dirs = new string[0];
                        if (comboBox1.Text == "熱處理1")
                        {
                            dirs = Directory.GetFiles(@"\\192.168.18.1\s7034be\MDB\" + minTime.ToString("yyyy") + @"\" + minTime.ToString("MM") + @"\", "HISDMP" + minTime.ToString("dd") + ".mdb");
                        }
                        else if (comboBox1.Text == "熱處理2")
                        {
                            dirs = Directory.GetFiles(@"\\192.168.18.2\s7033be\MDB\" + minTime.ToString("yyyy") + @"\" + minTime.ToString("MM") + @"\", "HISDMP" + minTime.ToString("dd") + ".mdb");

                        }
                        string file = @" ";

                        foreach (var HIS in dirs)
                        {
                            file = HIS.ToString();
                            label3.Text = HIS.ToString();
                        }

                        string l_provider = "Microsoft.Jet.OLEDB.4.0";

                        if (radioButton1.Checked == true)
                            l_provider = "Microsoft.Jet.OLEDB.4.0";
                        else if (radioButton2.Checked == true)
                            l_provider = "Microsoft.ACE.OLEDB.12.0";

                        OleDbConnection connection = new OleDbConnection("Provider=" + l_provider + "; Data Source=" + file + ";jet oledb:database password=;");

                        connection.Open();
                        //MessageBox.Show("開啟動作後，連結狀態：" + connection.State.ToString(), "連結狀態");

                        String queryString;
                        string[] a = new string[0];

                        queryString = "Select * from DPM where H_V <> 0 and [L-T] in (select min([L-T]) from DPM group by H_HOUR,H_MIN) order by [L-T];";


                        OleDbCommand command = new OleDbCommand(queryString, connection);
                        OleDbDataAdapter dataAdpter = new OleDbDataAdapter(queryString, connection);
                        DataSet DataSetValue = new DataSet();
                        DataSetValue.Clear();
                        dataAdpter.Fill(DataSetValue);

                        DataTable = new DataTable();

                        //Console.WriteLine(DataSetValue.Tables[0].Columns[0].Caption);


                        for (int i = 0; i < DataSetValue.Tables[0].Columns.Count; i++)
                        {
                            DataTable.Columns.Add(DataSetValue.Tables[0].Columns[i].Caption);

                        }

                        //Console.WriteLine(DataSetValue.Tables[0].Columns.Count);
                        Array.Resize(ref a, DataSetValue.Tables[0].Columns.Count);

                        for (int i = 0; i < DataSetValue.Tables[0].Rows.Count; i++)
                        {
                            for (int j = 0; j < DataSetValue.Tables[0].Columns.Count; j++)
                            {
                                //Console.WriteLine(DataSetValue.Tables[0].Rows[i][DataSetValue.Tables[0].Columns[j].Caption].ToString());
                                a[j] = DataSetValue.Tables[0].Rows[i][DataSetValue.Tables[0].Columns[j].Caption].ToString();
                                if (j == 0)
                                {
                                    DateTime DateObject = DateTime.Parse(a[0]);
                                    a[0] = DateObject.ToString("yyyy/MM/dd HH:mm:ss");

                                }
                                
                            }
                             DataTable.Rows.Add(a); 
                            
                        }

                        
                        dataGridView1.DataSource = DataTable;
                        dataGridView1.AllowUserToAddRows = false;

                        //自動調整寬度
                        dataGridView1.AutoResizeColumns();

                        //queryString = "INSERT INTO DATA ([Value1], [Value2]) VALUES ('" + Value1.Text + "','" + Value2.Text + "')";
                        //command = new OleDbCommand(queryString, connection);
                        //Console.WriteLine(dataGridView1.RowCount);

                        connection.Close();
                    }

                }
                catch (System.IO.DirectoryNotFoundException)
                {
                    MessageBox.Show("無當天紀錄");
                }
               
        }
        private void btnExcel_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFD = new SaveFileDialog();
            saveFD.InitialDirectory = "C:";
            saveFD.Title = "Save as Excel File";
            saveFD.FileName = "";
            saveFD.Filter = "Excel File(2003)|*.xls|Excel File(2007)|*.xlsx";
            if (saveFD.ShowDialog() != DialogResult.Cancel)
            {
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Application.Workbooks.Add(Type.Missing);

                excel.get_Range("A:AZ", Type.Missing).NumberFormatLocal = "@";  //設定A-AE欄儲存格格式為文字
                for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                {
                    excel.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                }
                for (int i = 0; i < dataGridView1.Rows.Count; i++) 
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        if (dataGridView1[j, i].Value == null)
                        {
                            this.dataGridView1[j, i].Value = "0";
                        }
                        string value = dataGridView1[j, i].Value.ToString();
                        excel.Cells[i + 2, j + 1] = value;
                    }
                }
                excel.ActiveWorkbook.SaveCopyAs(saveFD.FileName.ToString());
                excel.ActiveWorkbook.Saved = true;
                excel.Quit();
            }
            MessageBox.Show("EXCEL產製完成");
        }


        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {

        }
        private void timer1_Tick_1(object sender, EventArgs e)
        {
            DateTime dt = DateTime.Now;
            label1.Text = dt.ToString("yyyy/MM/dd tt hh:mm:ss");
        }

       
        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            if (dateTimePicker1.Value.Date >= DateTime.Today)
            {
                MessageBox.Show("當日資料無法讀取");
                dateTimePicker1.Value = DateTime.Today.AddDays(-1);
                return;
            }
            
            
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }
    }
}
