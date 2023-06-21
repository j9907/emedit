using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using emedit.DAO;
using emedit.DTO;
using Oracle.ManagedDataAccess.Client;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;


namespace emedit
{
    public partial class Form1 : Form
    {
        // DBConnection
        OracleConnection Conn;
        OracleDataAdapter adapter;
        UserDTO userdto = new UserDTO();
        string TASK_ID;
        string TASKDET_ID;
        TaskDAO dao = new TaskDAO();

        public Form1()
        {
            InitializeComponent();
        }

        // Form1 Load 시 실행 함수
        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                //DataGridView에 Task_inf 테이블 조회
                Conn = DBConnection.DBCon();
                Conn.Open();

                string sql = "SELECT TASKID AS ID, to_char(to_date(REGDT,'YYYY-MM-DD'),'YYYY-MM-DD') AS 접수일, to_char(to_date(schsttm,'HH24:MI:SS'),'HH24:MI:SS') as 등록시간,HEAD AS 제목, REQNM AS 요청자, STATE AS 상태 FROM task_inf order by taskid desc";


                OracleDataAdapter adapter = new OracleDataAdapter(sql, DBConnection.DBCon());
                DataTable data_table = new DataTable();
                
                adapter.Fill(data_table);
                dataGridView1.DataSource = data_table;

                Conn.Close();
                DataGridViewColumn col = dataGridView1.Columns["제목"];
                DataGridViewColumn col2 = dataGridView1.Columns["ID"];
                DataGridViewColumn col3 = dataGridView1.Columns["등록시간"];
                DataGridViewColumn col4 = dataGridView1.Columns["접수일"];
                col2.Width = 50;
                col3.Width = 80; col4.Width = 80;
                col.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        // 테이블을 수정하고 그리드뷰를 재출력하여 수정된 테이블을 볼수있도록 하는 함수
        private void GridRetrival() 
        {
            try
            {
                Conn = DBConnection.DBCon();
                Conn.Open();

                string sql = "SELECT TASKID AS ID, to_char(to_date(REGDT,'YYYY-MM-DD'),'YYYY-MM-DD') AS 접수일, to_char(to_date(schsttm,'HH24:MI:SS'),'HH24:MI:SS') as 등록시간,HEAD AS 제목, REQNM AS 요청자, STATE AS 상태 FROM task_inf order by taskid desc";


                OracleDataAdapter adapter = new OracleDataAdapter(sql, DBConnection.DBCon());
                DataTable data_table = new DataTable();

                adapter.Fill(data_table);
                dataGridView1.DataSource = data_table;

                Conn.Close();
                DataGridViewColumn col = dataGridView1.Columns["제목"];
                DataGridViewColumn col2 = dataGridView1.Columns["ID"];
                DataGridViewColumn col3 = dataGridView1.Columns["등록시간"];
                DataGridViewColumn col4 = dataGridView1.Columns["접수일"];
                col2.Width = 50;
                col3.Width = 80; col4.Width = 80;
                col.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
       
        // MenuStrip 클릭 시 Form 이동 함수
        private void 거래처등록ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            Form4 Form4 = new Form4();
            Form4.Show();
        }

        private void 프로젝트등록ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            Form6 Form6 = new Form6(Conn);
            Form6.Show();
        }

        //DataGridView1 의 데이터 클릭 시 실행되는 함수
        // DataGridView2에 선택한 데이터의 taskid 와 같은 데이터의 taskdet_inf_renew 테이블 조회
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex > -1)
            {
                TASK_ID = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();

                try
                {
                    string sql = "select TASKDETID AS ID,to_char(to_date(STARTDT,'YYYY-MM-DD'),'YYYY-MM-DD')AS 처리일, RESULTINFO AS 작업결과, STATE AS 상태,USRID AS 담당자 from taskdet_inf_renew WHERE ( taskid = '"+TASK_ID+"' ) order by TASKDETID asc";

                    string filesql = "select filenm,svrpath from pgm_attach_file_inf WHERE ( seq LIKE '%" + TASK_ID + "%' )"; //파일
                    OracleDataAdapter adapter2 = new OracleDataAdapter(filesql, Conn); //파일

                    DataTable dt = new DataTable(); //파일
                    OracleDataAdapter adapter = new OracleDataAdapter(sql, Conn);

                    DataSet ds = new DataSet();
                    adapter.Fill(ds);
                    adapter2.Fill(dt); //파일
                    dataGridView2.DataSource = ds.Tables[0];
                    Conn.Close();
                    textBox1.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
                    textBox6.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();

                    textBox5.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
                    textBox4.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                    DataGridViewColumn col = dataGridView2.Columns["작업결과"];
                    col.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                    DataGridViewColumn col1 = dataGridView2.Columns["ID"];
                    DataGridViewColumn col2 = dataGridView2.Columns["처리일"];
                    DataGridViewColumn col3 = dataGridView2.Columns["상태"];
                    DataGridViewColumn col4 = dataGridView2.Columns["담당자"];
                    col1.Width = 30;  col2.Width = 80; col3.Width = 60; col4.Width = 80;
                    textBox7.Text = string.Empty; //파일

                    foreach (DataRow row in dt.Rows) //파일
                    {
                        string fileName = row["filenm"].ToString(); //파일
                        string filepath = row["svrpath"].ToString();  //파일

                        textBox7.AppendText($"{fileName}");  //파일

                    } //파일
                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
               
                
            }

        }

        // DataGirdView2 를 채출력 하는 함수
        private void GridRetrival2()
        {
                try
                {
                    string sql = "select TASKDETID AS ID,to_char(to_date(STARTDT,'YYYY-MM-DD'),'YYYY-MM-DD')AS 처리일, RESULTINFO AS 작업결과, STATE AS 상태,USRID AS 담당자 from taskdet_inf_renew WHERE ( taskid = '"+TASK_ID+"' ) order by TASKDETID asc";

                    OracleDataAdapter adapter = new OracleDataAdapter(sql, Conn);

                    DataSet ds = new DataSet();
                    adapter.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];
                    Conn.Close();
                    textBox1.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
                    textBox6.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();

                    textBox5.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
                    textBox4.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                    DataGridViewColumn col = dataGridView2.Columns["작업결과"];
                    col.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                    DataGridViewColumn col1 = dataGridView2.Columns["ID"];
                    DataGridViewColumn col2 = dataGridView2.Columns["처리일"];
                    DataGridViewColumn col3 = dataGridView2.Columns["상태"];
                    DataGridViewColumn col4 = dataGridView2.Columns["담당자"];
                    col1.Width = 30; col2.Width = 80; col3.Width = 60; col4.Width = 80;

                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }


            
        }

        // 회신 버튼 클릭 시 실행 함수
        private void replyButton_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            Form5 Form5 = new Form5();
            Form5.Show();
        }

        // MenuStrip 클릭 시 실행함수
        private void 태스크등록ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            Form2 Form2 = new Form2();
            Form2.Show();
        }

        private void 거래처조회ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form9 Form9 = new Form9();
            Form9.Show();
        }

        private void 종료ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        // 조회 버튼 클릭 시 실행함수
        private void button2_Click(object sender, EventArgs e)
        {
            string sql2;

            try
            {

                Conn = DBConnection.DBCon();
                Conn.Open();

                sql2 = "SELECT TASKID AS ID, to_char(to_date(REGDT,'YYYY-MM-DD'),'YYYY-MM-DD') AS 접수일, to_char(to_date(schsttm,'HH24:MI:SS'),'HH24:MI:SS') as 등록시간,HEAD AS 제목, REQNM AS 요청자, STATE AS 상태 FROM task_inf " +
                    " where 1= 1 ";

                if (textBox9.Text != "")
                {
                    sql2 = sql2 + " and REGDT between to_char(to_date('"+textBox9.Text+"','YYYYMMDD'),'YYYYMMDD') and to_char(sysdate,'YYYYMMDD')";
                }
                if (textBox3.Text != "")
                {
                    sql2 = sql2 + " and SCHSTDT between to_char(to_date('"+textBox3.Text+"','YYYYMMDD'),'YYYYMMDD') and to_char(sysdate,'YYYYMMDD')";
                }
                if (textBox10.Text != "")
                {
                    sql2 = sql2 + " and REQNM LIKE '%" + textBox10.Text + "%' ";
                }
                if (textBox8.Text != "")
                {
                    sql2 = sql2 + " and USRID LIKE '%" + textBox8.Text + "%' ";
                }
                if (textBox2.Text != "")
                {
                    sql2 = sql2 + " and HEAD LIKE '%" + textBox2.Text + "%' ";
                }
                if (comboBox1.Text != "")
                {
                    sql2 = sql2 + " and STATE LIKE '%" + comboBox1.Text + "%' ";
                }
                sql2 += "order by taskid desc";
                adapter = new OracleDataAdapter(sql2, Conn);
                DataTable dataset = new DataTable();
                adapter.Fill(dataset);
                dataGridView1.DataSource = dataset;

                textBox9.Text = ""; textBox3.Text = ""; textBox10.Text = ""; textBox8.Text = ""; textBox2.Text = "";
                DataGridViewColumn col = dataGridView1.Columns["제목"];
                DataGridViewColumn col2 = dataGridView1.Columns["ID"];
                DataGridViewColumn col3 = dataGridView1.Columns["등록시간"];
                DataGridViewColumn col4 = dataGridView1.Columns["접수일"];
                col2.Width = 50;
                col3.Width = 80; col4.Width = 80;
                col.DefaultCellStyle.WrapMode = DataGridViewTriState.True;

                Conn.Close();

            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        // DataGridView1에 우클릭 시 나타나는 ContextMenuStrip 메뉴 (수정) 실행 함수
        private void Task1_Click(object sender, EventArgs e)
        {
            this.Close();
            Form2 frm2 = new Form2(TASK_ID);
            frm2.Show();
        }

        private void 종료ToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            Application.Exit();
        }

        /*private void 태스크상세등록ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
            Form10 f10 = new Form10();
            f10.Show();
        }*/

        //  DataGridView1에 우클릭 시 나타나는 ContextMenuStrip 메뉴 (상태수정) 실행 함수
        private void MenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(sender.ToString() + "으로 변경합니다. ", "상태 수정", MessageBoxButtons.YesNo) == DialogResult.Yes){
                int result = dao.updatestate(sender.ToString(), TASK_ID);
                //MessageBox.Show(result.ToString());
                GridRetrival();
            }
            
            
        }
        //  DataGridView2에 우클릭 시 나타나는 ContextMenuStrip 메뉴 (상태수정) 실행 함수
        private void subMemuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(sender.ToString() + "으로 변경합니다. ", "상태 수정", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                int result = dao.updatesubstate(sender.ToString(), TASKDET_ID);

                //MessageBox.Show(result.ToString());
                GridRetrival2();
            }
            
        }

        // DataGridView2 선택시 선택한 taskdet_inf_renew 테이블의 taskdet_id를 저장하는 함수
        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            TASKDET_ID = dataGridView2.Rows[e.RowIndex].Cells[0].Value.ToString();
            
        }

        //회신 버튼 클릭 시 실행 함수
        private void replyButton_Click_1(object sender, EventArgs e)
        {
            this.Close();
            Form5 f5 = new Form5(TASK_ID);
            f5.Show();
        }

        // DataGridView1에 우클릭 시 나타나는 ContextMenuStrip 메뉴 (작업결과등록) 실행 함수
        private void 작업결과등록ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
            Form10 f10 = new Form10(TASK_ID);
            f10.Show();
        }

        // 열기 버튼 클릭 시 실행되는 함수
        private void button1_Click(object sender, EventArgs e) 
        {


            string filenm;
            filenm = textBox7.Text;
            string filesql = "select filenm,svrpath from pgm_attach_file_inf WHERE ( filenm LIKE '%" + filenm + "%' )";
            OracleDataAdapter adapter = new OracleDataAdapter(filesql, Conn);
            DataTable dt = new DataTable();
            adapter.Fill(dt);


            foreach (DataRow row in dt.Rows)
            {
                string fileName = row["filenm"].ToString();
                string filepath = row["svrpath"].ToString();
                OpenFile(filepath + fileName);

            }
            Conn.Close();
        }

        // 열기 버튼 실행 시 파일을 여는 함수
        private void OpenFile(string filePath) 
        {
            try
            {
                // 파일이 존재하는지 확인
                if (File.Exists(filePath))
                {

                    Process.Start(filePath);
                }
                else
                {
                    MessageBox.Show("파일이 존재하지 않습니다.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }
    }
}
