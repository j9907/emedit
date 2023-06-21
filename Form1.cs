using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Oracle.ManagedDataAccess.Client;


namespace emedit
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                OracleConnection con = DBConnection.DBCon();
                con.Open();

                string sql = "SELECT * FROM CUSTOM_MST";

                OracleDataAdapter adapter = new OracleDataAdapter(sql, DBConnection.DBCon());
                DataTable data_table = new DataTable();

                adapter.Fill(data_table);
                dataGridView1.DataSource = data_table;

                con.Close();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
