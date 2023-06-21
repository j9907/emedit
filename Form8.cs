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
using emedit.DTO;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace emedit
{
    // 프로젝트 조회 화면
    public partial class Form8 : Form
    {
        private OracleConnection conn;
        private OracleCommand cmd;
        private OracleDataAdapter adapter;
        TaskDTO taskdto = new TaskDTO();
        string sql2;


        public Form8()
        {
            InitializeComponent();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        // form8 load시 실행되는 함수
        private void Form8_Load(object sender, EventArgs e)
        {
            //OracleDataReader myReader;
            string sql = null;

            try
            {
                conn = DBConnection.DBCon();
                conn.Open();
                sql = "select PRJTNO AS 프로젝트NO, PRJTNM AS 프로젝트명, USEFLG AS 사용여부, STARTDT AS 시작일자, ENDDT AS 종료일자, TRDCD AS 거래처코드, DISPSEQ AS 화면정렬순서,WRTUSRID AS 작성자ID, WRTDT AS 작성일자 from project_mst";
                OracleDataAdapter ad = new OracleDataAdapter();
                ad.SelectCommand = new OracleCommand(sql, conn);
                DataSet ds = new DataSet();
                ad.Fill(ds);
                dataGridView1.DataSource = ds.Tables[0].DefaultView;
                conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void GridRetrival() // 테이블을 수정하고 그리드뷰를 재출력하여 수정된 테이블을 볼수있도록 하는 함수
        {
            sql2 = "select PRJTNO AS 프로젝트NO, PRJTNM AS 프로젝트명, USEFLG AS 사용여부, STARTDT AS 시작일자, ENDDT AS 종료일자, TRDCD AS 거래처코드, DISPSEQ AS 화면정렬순서,WRTUSRID AS 작성자ID, WRTDT AS 작성일자 from project_mst";
            conn = DBConnection.DBCon();
            cmd = new OracleCommand(sql2, conn);

            adapter = new OracleDataAdapter(cmd);

            DataTable table = new DataTable();
            adapter.Fill(table);

            dataGridView1.DataSource = table.DefaultView;
        }

        // datagridview1(조회화면) 의 데이터 클릭 시 실행되는 함수
        private void dataGirdView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {

            if (e.RowIndex >= 0 && e.ColumnIndex >= 0) // 유효한 셀을 클릭했을 때
            {

                if (dataGridView1.CurrentRow.Cells[8].Value is DateTime)
                {
                    DateTime dateValue = (DateTime)dataGridView1.CurrentRow.Cells[8].Value;
                    string dateString = dateValue.ToShortDateString();

                    textBox2.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString(); //
                    textBox1.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                    textBox12.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                    textBox11.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                    textBox6.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                    textBox3.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
                    textBox9.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
                    textBox8.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
                    textBox5.Text = dateString;
                }





            }
        }

        // 수정 버튼 클릭 시 실행 함수
        private void button1_Click(object sender, EventArgs e) 
        {

            // prjtnm은 WHERE문에서 수정되기전 값을 비교하기 위해 사용
            // prjtnm2는 수정할 값을 입력하기 위해 사용
            string prjtno, prjtnm, prjtnm2, useflg, startdt, enddt, trdcd, wrtusrid, wrtdt;
            string query;

            prjtno = textBox2.Text;
            prjtnm = textBox12.Text;
            prjtnm2 = textBox1.Text;
            useflg = textBox11.Text;
            startdt = textBox6.Text;
            enddt = textBox3.Text;
            trdcd = textBox9.Text;
            wrtusrid = textBox8.Text;
            wrtdt = textBox5.Text;

            query = "update project_mst " +
                    " set prjtno = '" + prjtno + "', " +
                    " prjtnm = N'" + prjtnm2 + "', " +
                    " useflg= '" + useflg + "', " +
                    " startdt = N'" + startdt + "', " +
                    " enddt = N'" + enddt + "', " +
                    " trdcd = '" + trdcd + "', " +
                    " wrtusrid = '" + wrtusrid + "', " +
                    " wrtdt = '" + wrtdt + "' " +
                    " WHERE prjtnm = '" + prjtnm + "' ";

            try
            {

                conn = DBConnection.DBCon();
                cmd = new OracleCommand(query, conn);
                conn.Open();



                cmd.ExecuteNonQuery();   //insert, update, delete문에 사용
                GridRetrival();
                conn.Close();

                textBox2.Text = "";
                textBox1.Text = "";
                textBox12.Text = "";
                textBox11.Text = "";
                textBox6.Text = "";
                textBox3.Text = "";
                textBox9.Text = "";
                textBox8.Text = "";
                textBox5.Text = "";
                MessageBox.Show("수정되었습니다", "수정",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }
        // datagirdview(조회화면)의 데이터를 더블클릭했을때 실행되는 함수
        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewRow row = dataGridView1.SelectedRows[0];
            string data = row.Cells[0].Value.ToString();
            //MessageBox.Show(row.Cells[1].Value.ToString());
            taskdto.prjtno = data;
            this.Visible = false;
        }

        // 프로젝트번호를 반환하는 함수
        public string prjtno()
        {
            return taskdto.prjtno;
        }
        //조회 버튼
        private void button3_Click(object sender, EventArgs e)  
        {
            string sql3;

            try
            {

                conn = DBConnection.DBCon();
                conn.Open();

                sql3 = "SELECT PRJTNO AS 프로젝트NO, PRJTNM AS 프로젝트명, USEFLG AS 사용여부, STARTDT AS 시작일자, ENDDT AS 종료일자, TRDCD AS 거래처코드, DISPSEQ AS 화면정렬순서,WRTUSRID AS 작성자ID, WRTDT AS 작성일자 FROM project_mst " +
                    " where 1= 1 ";

                if (textBox4.Text != "")
                {
                    sql3 = sql3 + " and PRJTNO LIKE '%" + textBox4.Text + "%' ";
                }
                if (textBox7.Text != "")
                {
                    sql3 = sql3 + " and PRJTNM LIKE '%" + textBox7.Text + "%' ";
                }
                if (textBox10.Text != "")
                {
                    sql3 = sql3 + " and TRDCD LIKE '%" + textBox10.Text + "%' ";
                }


                adapter = new OracleDataAdapter(sql3, conn);
                DataTable dataset = new DataTable();
                adapter.Fill(dataset);
                dataGridView1.DataSource = dataset;

                textBox4.Text = ""; textBox7.Text = ""; textBox10.Text = "";

                conn.Close();


            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        // 취소 버튼 클릭 시 실행되는 함수
        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();

        }
    }
}
