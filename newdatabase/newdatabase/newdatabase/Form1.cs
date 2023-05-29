using System.Data.OleDb;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace newdatabase
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

        OleDbConnection conn;
        private void button1_Click(object sender, EventArgs e)
        {
            conn = new OleDbConnection();
            conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Users\\salih.morina\\Desktop\\newdatabase\\Database1.accdb";
            conn.Open();
            MessageBox.Show("Connected successfully");
        }

        private void button2_Click(object sender, EventArgs e)
        {


            using (OleDbCommand comm = new OleDbCommand())
            {
                comm.CommandText = "insert into Tabell1 (Namn,Klass,Telefonnummer,[E-post]) values(@Namn, @Klass, @Telefonnummer, @Epost)";

                // Check if the connection is null and open it if necessary
                if (conn == null)
                {
                    conn = new OleDbConnection();
                    conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Users\\salih.morina\\Desktop\\newdatabase\\Database1.accdb";
                    conn.Open();
                }

                comm.Connection = conn;
                comm.Parameters.AddWithValue("@Namn", textBox1.Text);
                comm.Parameters.AddWithValue("@Klass", textBox2.Text);
                comm.Parameters.AddWithValue("@Telefonnummer", textBox3.Text);
                comm.Parameters.AddWithValue("@E-post", textBox4.Text);

                comm.ExecuteNonQuery();

                // Close the connection if it was opened in this event handler
                if (conn.State == ConnectionState.Open)
                    conn.Close();

                MessageBox.Show("Data Inserted");
            }


        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            using (OleDbConnection conn = new OleDbConnection())
            {
                conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Users\\salih.morina\\Desktop\\newdatabase\\Database1.accdb";

                try
                {
                    conn.Open();
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = conn;
                    string query = "delete from Tabell1 where Namn=@Namn";
                    cmd.CommandText = query;
                    cmd.Parameters.AddWithValue("@Namn", textBox1.Text);
                    MessageBox.Show(query);
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Data Deleted");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message);
                }
                finally
                {
                    conn.Close();
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            using (OleDbConnection conn = new OleDbConnection())
            {
                conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Users\\salih.morina\\Desktop\\newdatabase\\Database1.accdb";

                try
                {
                    conn.Open();
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = conn;
                    string query = "SELECT * FROM Tabell1";
                    cmd.CommandText = query;
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    dataGridView1.DataSource = dt;
                    MessageBox.Show("Data Loaded");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message);
                }
                finally
                {
                    conn.Close();
                }
            }

        }
    }
}
