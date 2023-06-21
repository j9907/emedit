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
using Oracle.ManagedDataAccess.Client;

namespace emedit
{
    public partial class Form9 : Form
    {
        // 거래처 조회 화면
        OracleConnection Conn;
        OracleDataAdapter adapter;
        OracleCommand cmd;

        public Form9()
        {
            InitializeComponent();
        }

        // form9 load시 실행 함수
        private void Form9_Load(object sender, EventArgs e)
        {
            // DB 데이터 GridView에 불러오기
            try
            {
                Conn = DBConnection.DBCon();
                Conn.Open();

                string sql = "SELECT TRDCD, CMPYREGNO, CMPYNM, TRDSTATE, PRSNTNM, CMPYTEL, CMPYFAX, WRTDT, WRTID, TRDITEMS FROM CUSTOM_MST";

                adapter = new OracleDataAdapter(sql, Conn);
                DataTable data_table = new DataTable();

                adapter.Fill(data_table);
                dataGridView1.DataSource = data_table;

                Conn.Close();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //조회 버튼
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                Conn = DBConnection.DBCon();
                Conn.Open();

                //textBox1,2,3,4,5,6의 내용과 같은 TRDCD 데이터를 CUSTOM_MST 테이블에서 읽어오기
                string sql = "SELECT TRDCD, CMPYREGNO, CMPYNM, TRDSTATE, PRSNTNM, CMPYTEL, CMPYFAX, WRTDT, WRTID, TRDITEMS FROM CUSTOM_MST WHERE 1 =1 ";
                if (textBox1.Text != "") sql += " and trdcd like '%" + textBox1.Text + "'";
                if (textBox2.Text != "") sql += " and prsntnm like '%" + textBox2.Text + "'";
                if (textBox3.Text != "") sql += " and wrtid like '%" + textBox3.Text + "'";
                if (textBox4.Text != "") sql += " and cmpytel like '%" + textBox4.Text + "'";
                if (textBox5.Text != "") sql += " and cmpynm like '%" + textBox5.Text + "'";
                if (textBox6.Text != "") sql += " and wrtdt between to_char(to_date('" + textBox6.Text + "'),'yyyymmdd') and to_char(sysdate,'YYYYMMDD')";

                adapter = new OracleDataAdapter(sql, Conn);
                DataTable dataSet = new DataTable();
                adapter.Fill(dataSet);
                dataGridView1.DataSource = dataSet;

                textBox1.Text = "";
                textBox2.Text = "";
                textBox3.Text = "";
                textBox4.Text = "";
                textBox5.Text = "";

                Conn.Close();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //GridView의 셀 클릭 시 데이터 textBox에 출력
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox7.Text = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
            textBox8.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
            textBox9.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
            textBox10.Text = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
            textBox11.Text = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
            textBox12.Text = dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString();
            textBox13.Text = dataGridView1.Rows[e.RowIndex].Cells[6].Value.ToString();
            textBox14.Text = dataGridView1.Rows[e.RowIndex].Cells[9].Value.ToString();
        }

        //수정 버튼
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                Conn = DBConnection.DBCon();
                Conn.Open();

                string sql = "UPDATE CUSTOM_MST SET " +
                    "TRDCD = '" + textBox7.Text + "', " +
                    "CMPYREGNO = '" + textBox8.Text + "', " +
                    "CMPYNM = '" + textBox9.Text + "', " +
                    "TRDSTATE = '" + textBox10.Text + "'," +
                    "PRSNTNM = '" + textBox11.Text + "', " +
                    "CMPYTEL = '" + textBox12.Text + "', " +
                    "CMPYFAX = '" + textBox13.Text + "', " +
                    "TRDITEMS = '" + textBox14.Text + "' " +
                    "WHERE TRDCD = '" + textBox7.Text + "' ";

                cmd = new OracleCommand(sql, Conn);
                cmd.ExecuteNonQuery();

                textBox7.Text = "";
                textBox8.Text = "";
                textBox9.Text = "";
                textBox10.Text = "";
                textBox11.Text = "";
                textBox12.Text = "";
                textBox13.Text = "";
                textBox14.Text = "";

                GridRetrival();
                MessageBox.Show("정보가 수정되었습니다.");
                Conn.Clone();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //거래처 정보 수정 후 GridView에 정보 새로고침
        private void GridRetrival()
        {
            Conn = DBConnection.DBCon();
            Conn.Open();

            string sql = "SELECT TRDCD, CMPYREGNO, CMPYNM, TRDSTATE, PRSNTNM, CMPYTEL, CMPYFAX, WRTDT, WRTID, TRDITEMS FROM CUSTOM_MST";

            adapter = new OracleDataAdapter(sql, Conn);
            DataTable data_table = new DataTable();

            adapter.Fill(data_table);
            dataGridView1.DataSource = data_table;

            Conn.Close();
        }
    }
}
