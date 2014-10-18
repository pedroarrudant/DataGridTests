using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        public void Form1_Load(object sender, EventArgs e)
        {
            OleDbConnection conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\database.accdb;Persist Security Info=False;");
            try
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand("SELECT * FROM Pessoa", conn);
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                dataGridView1.AutoResizeColumns();
                dataGridView1.Columns[0].ReadOnly = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro! " + ex);
                this.textBox1.Text = Convert.ToString(ex);
            }
        }

        /*private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox1.Text = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
        }*/
        private void DataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            string alter = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
            string campo = dataGridView1.Columns[e.ColumnIndex].Name;
            string index = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
            try
            {
                OleDbConnection conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\database.accdb;Persist Security Info=False;");
                conn.Open();
                string comando = "UPDATE Pessoa SET " + campo + " = '" + alter + "' WHERE ID = " + index + " ;";
                OleDbCommand cmd = new OleDbCommand(comando, conn);
                /*cmd.Parameters.AddWithValue("?", campo);
                cmd.Parameters.AddWithValue("?", alter);
                cmd.Parameters.AddWithValue("?", index);*/
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                cmd.ExecuteNonQuery();
                OleDbCommand cmd2 = new OleDbCommand("SELECT * FROM Pessoa", conn);
                OleDbDataAdapter da2 = new OleDbDataAdapter(cmd2);
                DataTable dt = new DataTable();
                da2.Fill(dt);
                dataGridView1.DataSource = dt;
                dataGridView1.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro! " + ex);
                this.textBox1.Text = Convert.ToString(ex);
            }

        }
    }
}
