using Npgsql;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Runtime.InteropServices;   //Required for Excel
using Excel = Microsoft.Office.Interop.Excel;   //Required for Excel
//Right Click Your Project=> Add Reference=> COM =>
//Tick 1 item ( Microsoft Excel 16.0 Oject Library ). And Then Click Ok.

namespace Postgresql_Fetch_Data
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }
        private void Select()
        {
            string host_ip, host_port, username, password;
            string db_name, db_table;

            host_ip = textBox3.Text; host_port = textBox4.Text;
            username = textBox1.Text; password = textBox2.Text;
            //db_name = textBox6.Text;
            db_name = comboBox1.Text;
            //textBox5.Text = comboBox1.Text;
            db_table = textBox5.Text;

            string Connection_String = "Server=" + host_ip +
                ";Port=" + host_port +
                ";User Id=" + username +
                ";Password=" + password +
                ";Database=" + db_name;

            string sql; NpgsqlCommand cmd; DataTable dt;

            try
            {
                NpgsqlConnection npgsqlConnection = new NpgsqlConnection(Connection_String);
                npgsqlConnection.Open();

                sql = "SELECT * FROM " + db_table + ";" ;
                cmd = new NpgsqlCommand(sql, npgsqlConnection);

                dt = new DataTable();
                dt.Load(cmd.ExecuteReader());

                npgsqlConnection.Close();

                dataGridView1.DataSource = null;
                dataGridView1.DataSource = dt;
            }

            catch (Exception ex)
            {
                MessageBox.Show("Error: "+ex.Message);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Select();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string host_ip, host_port, username, password;
            string db_name, db_table;

            host_ip = textBox3.Text; host_port = textBox4.Text;
            username = textBox1.Text; password = textBox2.Text;

            string Connection_String = "Server=" + host_ip +
                ";Port=" + host_port +
                ";User Id=" + username +
                ";Password=" + password;

            string sql; NpgsqlCommand cmd; DataTable dt;

            try
            {
                NpgsqlConnection npgsqlConnection = new NpgsqlConnection(Connection_String);
                npgsqlConnection.Open();

                sql = "SELECT datname FROM pg_database;";
                cmd = new NpgsqlCommand(sql, npgsqlConnection);

                dt = new DataTable();
                dt.Load(cmd.ExecuteReader());

                npgsqlConnection.Close();

                int dt_rows = 0;
                dt_rows = dt.Rows.Count;

                comboBox1.Items.Clear();
                foreach(DataRow row in dt.Rows)
                {
                    richTextBox1.AppendText("\n" + row[0].ToString());
                    comboBox1.Items.Add(row[0].ToString());
                }
                int a = comboBox1.Items.Count;
                if (a > 0) { comboBox1.SelectedIndex = a-1; }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 0)
            {
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "CSV (*.csv)|*.csv";
                sfd.FileName = "Voter Response.csv";
                bool fileError = false;
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    if (File.Exists(sfd.FileName))
                    {
                        try
                        {
                            File.Delete(sfd.FileName);
                        }
                        catch (IOException ex)
                        {
                            fileError = true;
                            MessageBox.Show("It wasn't possible to write the data to the disk." + ex.Message);
                        }
                    }
                    if (!fileError)
                    {
                        try
                        {
                            int columnCount = dataGridView1.Columns.Count;
                            string columnNames = "";
                            string[] outputCsv = new string[dataGridView1.Rows.Count + 1];
                            for (int i = 0; i < columnCount; i++)
                            {
                                columnNames += dataGridView1.Columns[i].HeaderText.ToString() + ",";
                            }
                            outputCsv[0] += columnNames;

                            for (int i = 1; i < dataGridView1.Rows.Count; i++)
                            {
                                for (int j = 0; j < columnCount; j++)
                                {
                                    outputCsv[i] += dataGridView1.Rows[i - 1].Cells[j].Value.ToString() + ",";
                                }
                            }

                            File.WriteAllLines(sfd.FileName, outputCsv, Encoding.UTF8);
                            DialogResult a = MessageBox.Show("Data Exported Successfully !!!\n\n" +
                                "Do you want to open now", "Info", MessageBoxButtons.YesNo, MessageBoxIcon.Information);

                            if (a == DialogResult.Yes)
                            {
                                try
                                {
                                    System.Diagnostics.Process.Start(sfd.FileName);//launch a PDF with the default associated application
                                }
                                catch (Exception Ex)
                                {
                                    MessageBox.Show(Ex.ToString());
                                }
                            }

                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error :" + ex.Message);
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("No Record To Export !!!", "Info");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            SaveFileDialog savedlg = new SaveFileDialog();
            savedlg.Filter = "ASCII Text file (.txt)|.txt|" +
            "Excel file (.xls)|.xls| All files (.)|.";
            savedlg.FilterIndex = 2;
            if (savedlg.ShowDialog() == DialogResult.OK)
            {

            }
            else return;

            //Start Excel and get Application object.
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            //xlApp.Visible = true;  //To show the excell in live

            //Get a new workbook.
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet1;
            Excel.Worksheet xlWorkSheet2;

            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet1 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlWorkSheet1.Name = "Result";

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed or version not supported!!", "Save Aborted");
                return;
            }

            int tRows = dataGridView1.RowCount - 1;
            int tCols = dataGridView1.ColumnCount - 1;
            //Data Transfer from dataGridView1 Header to Excel Cells
            for (int j = 0; j <= tCols; j++)
                xlWorkSheet1.Cells[1, j + 1] = dataGridView1.Columns[j].HeaderText;
            //Data Transfer from dataGridView1 to Excel Cells
            for (int i = 0; i < tRows; i++)
                for (int j = 0; j <= tCols; j++)
                    xlWorkSheet1.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
            //Auto Colum Fit                
            for (int j = 0; j <= tCols; j++)
                xlWorkSheet1.Columns[j + 1].AutoFit();
            //xlWorkSheet1.Columns[j].Visible = true;

            Excel.Range Rng = xlWorkSheet1.Range["A1" ,"B10"];

            Rng.Style.Font.Size = 12 ;
            Rng.Style.Font.Bold = false; 
            Rng.Style.Font.Italic = true;
            Rng.Style.Font.Color = Color.Blue ;
            
            xlWorkBook.SaveAs(@savedlg.FileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet1);
            //Marshal.ReleaseComObject(xlWorkSheet2);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            DialogResult a1;
            a1 = MessageBox.Show("Excel file created , you can find the file at " + @savedlg.FileName + "\n\n Do you want to Open Now", "Status and Option", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (a1 == DialogResult.Yes) System.Diagnostics.Process.Start(@savedlg.FileName);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string host_ip, host_port, username, password;
            string db_name, db_table;

            host_ip = textBox3.Text; host_port = textBox4.Text;
            username = textBox1.Text; password = textBox2.Text;
            //db_name = textBox6.Text;
            db_name = comboBox1.Text;
            //textBox5.Text = comboBox1.Text;
            db_table = textBox5.Text;

            string Connection_String = "Server=" + host_ip +
                ";Port=" + host_port +
                ";User Id=" + username +
                ";Password=" + password +
                ";Database=" + db_name;

            string sql; NpgsqlCommand cmd; DataTable dt;
            try
            {
                NpgsqlConnection npgsqlConnection = new NpgsqlConnection(Connection_String);
                npgsqlConnection.Open();

                sql = "SELECT * FROM pg_catalog.pg_tables;";
                cmd = new NpgsqlCommand(sql, npgsqlConnection);

                dt = new DataTable();
                dt.Load(cmd.ExecuteReader());

                npgsqlConnection.Close();

                int dt_rows = 0;
                dt_rows = dt.Rows.Count;

                comboBox2.Items.Clear();
                foreach (DataRow row in dt.Rows)
                {
                    richTextBox1.AppendText("\n" + row[2].ToString());
                    if (row[0].ToString() =="public")
                    comboBox2.Items.Add(row[1].ToString());
                }
                int a = comboBox2.Items.Count;
                if (a > 0) { comboBox2.SelectedIndex = 0; }
            }

            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox6.Text = comboBox1.Text;
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox5.Text = comboBox2.Text;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
