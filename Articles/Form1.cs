using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

/*
 * Programs that needs to be installed:
 * SQLExpress
 * Microsoft Access Database Engine
*/

namespace Articles
{
    public partial class Form1 : Form
    {
        private readonly string _connectionString;
        private readonly DataTable _articlesTable;
        private string filePath;

        public Form1()
        {
            InitializeComponent();

            InitializeUI();

            _connectionString = @"Data Source=.\SQLEXPRESS;AttachDbFilename=|DataDirectory|\Database2.mdf;Integrated Security=True;User Instance=True";
            _articlesTable = new DataTable("Articles");
        }

        private async Task LoadArticlesAsync()
        {
            using (var connection = new SqlConnection(_connectionString))
            using (var adapter = new SqlDataAdapter("SELECT * FROM Articles", connection))
            {
                await connection.OpenAsync();
                adapter.Fill(_articlesTable);
            }

            dataGridView1.DataSource = _articlesTable;
            dataGridView1.Columns["Id"].Visible = false;
            dataGridView1.Columns["category"].Visible = false;
            dataGridView1.Columns["Reference"].HeaderText = "Reference";
            dataGridView1.Columns["Author"].HeaderText = "Author";
            dataGridView1.Columns["location"].HeaderText = "Location";
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Microsoft Sans Serif", 8f, FontStyle.Bold);
            dataGridView1.BackgroundColor = SystemColors.Control;
        }

        private void InitializeUI()
        {
            pictureBox1.BackColor = Color.Transparent;
            panel1.BackColor = Color.FromArgb(125, Color.Red);
            label1.BackColor = Color.FromArgb(125, Color.Red);
            label2.BackColor = Color.FromArgb(125, Color.Red);
            label3.BackColor = Color.FromArgb(125, Color.Red);
            label4.BackColor = Color.FromArgb(125, Color.Red);
            label5.BackColor = Color.FromArgb(125, Color.Red);
            label6.BackColor = Color.FromArgb(125, Color.Red);
            label7.BackColor = Color.FromArgb(125, Color.Red);
        }

        private async void Form1_Load(object sender, EventArgs e)
        {
            label1.Focus();
            InitializeUI();
            await LoadArticlesAsync();
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }

        private void textBox1_Enter(object sender, EventArgs e)
        {
            if (textBox1.Text == "Search")
            {
                textBox1.Text = "";
                textBox1.ForeColor = Color.Black;
            }
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                textBox1.Text = "Search";
                textBox1.ForeColor = Color.Silver;
            }
            _articlesTable.DefaultView.RowFilter = "";
        }
        private void btnImport_Click(object sender, EventArgs e)
        {
            // Open a file dialog to allow the user to select an Excel file to import.
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                // Set the filePath variable to the path of the selected Excel file.
                filePath = openFileDialog1.FileName;

                // Read the data from the selected Excel file and store it in a DataTable.
                using (OleDbConnection connection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + openFileDialog1.FileName + ";Extended Properties=Excel 12.0;"))
                {
                    connection.Open();
                    OleDbDataAdapter adapter = new OleDbDataAdapter("SELECT * FROM [Sheet1$]", connection);
                    DataTable importedTable = new DataTable();
                    adapter.Fill(importedTable);

                    // Delete all the data in the "Articles" table.
                    string connectionString = @"Data Source=.\SQLEXPRESS;AttachDbFilename=|DataDirectory|\Database2.mdf;Integrated Security=True;User Instance=True";
                    using (SqlConnection sqlConnection = new SqlConnection(connectionString))
                    {
                        sqlConnection.Open();
                        using (SqlCommand sqlCommand = new SqlCommand("DELETE FROM Articles", sqlConnection))
                        {
                            sqlCommand.ExecuteNonQuery();
                        }
                    }

                    // Insert the imported data into the "Articles" table.
                    using (SqlConnection sqlConnection = new SqlConnection(connectionString))
                    {
                        sqlConnection.Open();
                        using (SqlBulkCopy bulkCopy = new SqlBulkCopy(sqlConnection))
                        {
                            bulkCopy.DestinationTableName = "Articles";
                            bulkCopy.WriteToServer(importedTable);
                        }
                    }

                    // Refresh your DataGridView with the replaced data.
                    RefreshDataGridView();
                }
            }
        }

        private void RefreshDataGridView()
        {
            // Refresh your DataGridView with the replaced data.
            using (SqlConnection sqlConnection = new SqlConnection(_connectionString))
            {
                sqlConnection.Open();
                using (SqlCommand sqlCommand = new SqlCommand("SELECT * FROM Articles", sqlConnection))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter(sqlCommand);
                    DataTable dataTable = new DataTable();
                    adapter.Fill(dataTable);
                    dataGridView1.DataSource = dataTable;
                }
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            // Filter your DataGridView based on the search text.
            string filter = string.Format("Category LIKE '%{0}%' OR Article LIKE '%{0}%'", textBox1.Text);
            ((DataTable)dataGridView1.DataSource).DefaultView.RowFilter = filter;

            _articlesTable.DefaultView.RowFilter = "";
        }

    }

}