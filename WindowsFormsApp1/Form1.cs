using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
using System.Runtime.InteropServices;
using System.IO;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            dataGridView1.CellClick += dataGridView1_CellContentClick;
            button9.Click += button9_Click;
            btnExport.Click += btnExport_Click;


        }

        SqlConnection con = new SqlConnection("Data Source=LAPTOP-VCETEBAE\\SQLEXPRESS;Initial Catalog=CRUD_SP_DB;Integrated Security=True");
        

        private void button1_Click_1(object sender, EventArgs e)
        {
            con.Open();
            String status = "";
            if (radioButton1.Checked == true)
            {
                status = radioButton1.Text;
            }
            else
            {
                status = radioButton2.Text;
            }

            SqlCommand cmd = new SqlCommand("dbo.SP_Product_Insert", con);
            cmd.CommandType = CommandType.StoredProcedure;

            cmd.Parameters.AddWithValue("@ProductId", int.Parse(textBox1.Text));
            cmd.Parameters.AddWithValue("@ItemName", textBox2.Text);
            cmd.Parameters.AddWithValue("@Color", comboBox1.Text);
            cmd.Parameters.AddWithValue("@Status", status);
            cmd.Parameters.AddWithValue("@ExpiryDate", DateTime.Parse(dateTimePicker1.Text));

            cmd.ExecuteNonQuery();
            con.Close();

            MessageBox.Show("Inserted Successfully");

            // Clear the input fields
            textBox1.Text = "";
            textBox2.Text = "";
            comboBox1.SelectedIndex = -1;
            radioButton1.Checked = false;
            radioButton2.Checked = false;
            dateTimePicker1.Value = DateTime.Now;

            LoadAllRecords();
        }

        void LoadAllRecords()
        {
            SqlCommand cmd = new SqlCommand("exec dbo.SP_Product_View", con);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //LoadAllRecords();
         
        }

        private void button2_Click(object sender, EventArgs e)
        {
            con.Open();
            String status = "";
            if (radioButton1.Checked == true)
            {
                status = radioButton1.Text;

            }
            else
            {
                status = radioButton2.Text;
            }

            SqlCommand cmd = new SqlCommand("dbo.SP_Product_Update", con);
            cmd.CommandType = CommandType.StoredProcedure;

            cmd.Parameters.AddWithValue("@ProductId", int.Parse(textBox1.Text));
            cmd.Parameters.AddWithValue("@ItemName", textBox2.Text);
            cmd.Parameters.AddWithValue("@Color", comboBox1.Text);
            cmd.Parameters.AddWithValue("@Status", status);
            cmd.Parameters.AddWithValue("@ExpiryDate", DateTime.Parse(dateTimePicker1.Text));

            cmd.ExecuteNonQuery();
            con.Close();

            MessageBox.Show("Updated Successfully");

            LoadAllRecords();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are You Sure To Delete?", "Delete", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                con.Open();

                SqlCommand cmd = new SqlCommand("dbo.SP_Product_Delete", con);
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.AddWithValue("@ProductId", int.Parse(textBox1.Text));

                cmd.ExecuteNonQuery();
                con.Close();

                MessageBox.Show("Deleted Successfully");

                // Clear the input fields
                //textBox1.Text = "";
                //textBox2.Text = "";
                //comboBox1.SelectedIndex = -1;
                //radioButton1.Checked = false;
                //radioButton2.Checked = false;
                //dateTimePicker1.Value = DateTime.Now;

                LoadAllRecords();
            }
}

        private void button4_Click(object sender, EventArgs e)
        {
            SqlCommand cmd = new SqlCommand("exec dbo.SP_Product_Search '" + int.Parse(textBox1.Text) + "'", con);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            da.Fill(dt);

            if (dt.Rows.Count > 0) // check if any rows are returned by the search
            {
                // show the data in the DataGridView
                dataGridView1.DataSource = dt;

                // show the data in the input fields
                textBox2.Text = dt.Rows[0]["ItemName"].ToString();
                comboBox1.Text = dt.Rows[0]["Color"].ToString();
                radioButton1.Checked = dt.Rows[0]["Status"].ToString() == "Ready";
                radioButton2.Checked = dt.Rows[0]["Status"].ToString() == "UnUsed";
                dateTimePicker1.Value = Convert.ToDateTime(dt.Rows[0]["ExpiryDate"]);
            }
            else
            {
                MessageBox.Show("No records found for the specified Product ID.");
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0) // make sure a valid row is clicked
            {
                DataGridViewRow row = dataGridView1.Rows[e.RowIndex];

                // show the data in the input fields
                textBox1.Text = row.Cells["ProductId"].Value.ToString();
                textBox2.Text = row.Cells["ItemName"].Value.ToString();
                comboBox1.Text = row.Cells["Color"].Value.ToString();
                radioButton1.Checked = row.Cells["Status"].Value.ToString() == "Ready";
                radioButton2.Checked = row.Cells["Status"].Value.ToString() == "UnUsed";
                dateTimePicker1.Value = Convert.ToDateTime(row.Cells["ExpiryDate"].Value);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {

            textBox1.Text = "";
            textBox2.Text = "";
            comboBox1.SelectedIndex = -1;
            radioButton1.Checked = false;
            radioButton2.Checked = false;
            dateTimePicker1.Value = DateTime.Now;

        }
       
        private void button6_Click(object sender, EventArgs e)
        {
            con.Open();
            DataTable dt = new DataTable();

            SqlCommand cmd = new SqlCommand("GetProductInfo2", con);
            cmd.CommandType = CommandType.StoredProcedure;

            // Add parameters based on user input
            if (!string.IsNullOrEmpty(textBoxItemName.Text))
            {
                cmd.Parameters.AddWithValue("@ItemName", textBoxItemName.Text);
            }
            if (!string.IsNullOrEmpty(comboBox2.Text))
            {
                cmd.Parameters.AddWithValue("@Color", comboBox2.Text);
            }
            string status = null;
            if (radioButton4.Checked)
            {
                status = "Ready";
            }
            else if (radioButton3.Checked)
            {
                status = "UnUsed";
            }
            cmd.Parameters.AddWithValue("@Status", status);

            SqlDataAdapter sda = new SqlDataAdapter(cmd);
            sda.Fill(dt);

            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("No record found.");
            }
            else
            {
                dataGridView1.DataSource = dt;
            }

            con.Close();

        }

        private void button7_Click(object sender, EventArgs e)
        {
            textBoxItemName.Text = "";
            
            comboBox2.SelectedIndex = -1;
            radioButton4.Checked = false;
            radioButton3.Checked = false;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            dataGridView1.SelectAll();
            DataObject copydata = dataGridView1.GetClipboardContent();
            if (copydata != null) Clipboard.SetDataObject(copydata);
            Microsoft.Office.Interop.Excel.Application xlapp = new Microsoft.Office.Interop.Excel.Application();
            xlapp.Visible = true;
            Microsoft.Office.Interop.Excel.Workbook xlwbook;
            Microsoft.Office.Interop.Excel.Worksheet xlsheet;
            object miseddata = System.Reflection.Missing.Value;
            xlwbook = xlapp.Workbooks.Add(miseddata);
            xlsheet = (Microsoft.Office.Interop.Excel.Worksheet)xlwbook.Worksheets.get_Item(1);

            // Set the header text
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                xlsheet.Cells[1, i + 1] = dataGridView1.Columns[i].HeaderText;
            }

            Microsoft.Office.Interop.Excel.Range xlr = (Microsoft.Office.Interop.Excel.Range)xlsheet.Cells[2, 1];
            xlr.Select();
            xlsheet.PasteSpecial(xlr, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            LoadAllRecords();
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
           
                // Create an instance of the SaveFileDialog class
                SaveFileDialog saveFileDialog = new SaveFileDialog();

                // Set the filter to limit file types to CSV files
                saveFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*";

                // If the user selects a file, export the data to the file
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // Create an instance of the StreamWriter class
                    StreamWriter streamWriter = new StreamWriter(saveFileDialog.FileName);

                    // Write the column headers to the file
                    for (int i = 0; i < dataGridView1.Columns.Count; i++)
                    {
                        streamWriter.Write(dataGridView1.Columns[i].HeaderText);
                        if (i < dataGridView1.Columns.Count - 1)
                        {
                            streamWriter.Write(",");
                        }
                    }
                    streamWriter.WriteLine();

                    // Loop through the rows of the DataGridView and write each row to the file
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataGridView1.Columns.Count; j++)
                        {
                            streamWriter.Write(dataGridView1.Rows[i].Cells[j].Value);
                            if (j < dataGridView1.Columns.Count - 1)
                            {
                                streamWriter.Write(",");
                            }
                        }
                        streamWriter.WriteLine();
                    }

                    // Close the StreamWriter
                    streamWriter.Close();

                    // Display a message indicating that the export was successful
                    MessageBox.Show("Data exported successfully.");
                }
            

        }
    }
 }
