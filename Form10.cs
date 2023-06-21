using emedit.DTO;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics.Eventing.Reader;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace emedit
{
    public partial class Form10 : Form
    {
        // 태스크 상세(작업결과) 등록 페이지
        static UserDTO userdto = new UserDTO();
        string userid = "";
        string Task_id = "";
        public Form10()
        {
            InitializeComponent();
        }

        public Form10(string Taskid)
        {
            InitializeComponent();
            this.Task_id = Taskid;
            //MessageBox.Show(Task_id);
        }

        // Form10 로드 시 실행되는 함수
        private void Form10_Load(object sender, EventArgs e)
        {
            dataGridView1.ClearSelection();
            dataGridView1.CurrentCell = null;
            dataGridView2.CurrentCell = null;
            

            try
            {
                
                OracleConnection con = DBConnection.DBCon();
                con.Open();

                // 데이터그리드뷰에 테이블 데이터 표시
                string sql = "SELECT TASKID AS ID, PRJTNO AS PJNO, REGDT AS 접수일자,REQINFO AS 내용,  USRID AS 담당자, STATE AS 상태,REQNM AS 요청자 FROM TASK_INF where taskid = " + Task_id;

                OracleDataAdapter adapter = new OracleDataAdapter(sql, DBConnection.DBCon());

                DataTable data_table = new DataTable();

                adapter.Fill(data_table);
                dataGridView1.DataSource = data_table;
                gridviewds();
                // 컬럼명 변경
                /*
                dataGridView1.Columns["TASKID"].HeaderText = "태스크ID";
                dataGridView1.Columns["PRJTNO"].HeaderText = "프로젝트NO";
                //dataGridView1.Columns["PRJID"].HeaderText = "프로젝트";
                dataGridView1.Columns["REGDT"].HeaderText = "접수일자";
               // dataGridView1.Columns["SCHSTDT"].HeaderText = "등록일자";
              //  dataGridView1.Columns["SCHSTTM"].HeaderText = "등록시간";
                dataGridView1.Columns["USRID"].HeaderText = "담당자";
               // dataGridView1.Columns["SCHEDDT"].HeaderText = "종료일자";
                //dataGridView1.Columns["SCHEDTM"].HeaderText = "종료시간";
                dataGridView1.Columns["STATE"].HeaderText = "상태";
                //dataGridView1.Columns["LEVEL"].HeaderText = "공개등급";
                //dataGridView1.Columns["ORGTASKID"].HeaderText = "원본태스크ID";
                //dataGridView1.Columns["WRTUSRID"].HeaderText = "작성자";
               // dataGridView1.Columns["WRTDT"].HeaderText = "작성일자";
               // dataGridView1.Columns["HEAD"].HeaderText = "제목";
                dataGridView1.Columns["REQINFO"].HeaderText = "내용";
                dataGridView1.Columns["REQNM"].HeaderText = "요청자";
               // dataGridView1.Columns["CALLNUMBER"].HeaderText = "전화번호";
               // dataGridView1.Columns["EMAIL"].HeaderText = "이메일";*/

                // 콤보박스에 상태 옵션 추가
                state.Items.Add("등록");
                state.Items.Add("진행중");
                state.Items.Add("완료");
                state.Items.Add("취소");

                // 콤보박스에 상태 옵션 추가
                method.Items.Add("P");
                method.Items.Add("D");
                method.Items.Add("E");
                method.Items.Add("S");

                // 데이터 그리드뷰2 조회 함수
                string sql2 = "SELECT TASKID, TASKDETID, REQNM, STARTDT, STARTTM, ENDDT, ENDTM, USRID, TURNATM, METHOD, STATE, WRTDT, RESULTINFO, HEAD, WORKDAY, WRTUSRID FROM TASKDET_INF_RENEW where taskid = " + Task_id;
                OracleDataAdapter adapter2 = new OracleDataAdapter(sql2, DBConnection.DBCon());
                DataTable data_table2 = new DataTable();
                adapter2.Fill(data_table2);
                dataGridView2.DataSource = data_table2;

                dataGridView2.Columns["TASKID"].HeaderText = "태스크ID";
                dataGridView2.Columns["TASKDETID"].HeaderText = "상세태스크ID";
                dataGridView2.Columns["REQNM"].HeaderText = "요청자명";
                dataGridView2.Columns["STARTDT"].HeaderText = "작업시작일자";
                dataGridView2.Columns["STARTTM"].HeaderText = "작업시작시간";
                dataGridView2.Columns["ENDDT"].HeaderText = "작업종료일자";
                dataGridView2.Columns["ENDTM"].HeaderText = "작업종료시간";
                dataGridView2.Columns["USRID"].HeaderText = "작업자";
                dataGridView2.Columns["TURNATM"].HeaderText = "소요시간";
                dataGridView2.Columns["METHOD"].HeaderText = "작업방법";
                dataGridView2.Columns["STATE"].HeaderText = "상태";
                dataGridView2.Columns["WRTDT"].HeaderText = "작성일자";
                dataGridView2.Columns["RESULTINFO"].HeaderText = "작업결과";
                dataGridView2.Columns["HEAD"].HeaderText = "제목";
                dataGridView2.Columns["WORKDAY"].HeaderText = "작업시간";
                dataGridView2.Columns["WRTUSRID"].HeaderText = "작성자ID";
                con.Close();
                wrtusrid.Text = userdto.usrid;
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }
        // 취소 버튼 선택 시 실행되는 함수
        private void cancelbutton_Click(object sender, EventArgs e)
        {
            Close();
            Form1 showForm1 = new Form1();
            showForm1.Show();
        }

        
       /* private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                DataGridViewRow row = dataGridView1.Rows[e.RowIndex];
                string TASKID = row.Cells["TASKID"].Value.ToString();

                taskid.Text = TASKID;

            }
        }*/

        // 태스크 상세 정보 테이블에 입력한 데이터를 삽입하는 함수
        private void InsertTaskDetInfo(string TASKID, string TASKDETID, string REQNM, string STARTDT, string STARTTM, string ENDDT, string ENDTM, string USRID, string TURNATM, string METHOD, string STATE, string WRTDT, string RESULTINFO, string HEAD, string WRTUSRID, OracleConnection connection,string COMPFLG)
        {
            // SQL 쿼리 작성
            string query = "INSERT INTO TASKDET_INF_RENEW (TASKID, TASKDETID, REQNM, STARTDT, STARTTM, ENDDT, ENDTM, USRID, TURNATM, METHOD, STATE, WRTDT, RESULTINFO, HEAD,  WRTUSRID,COMPFLG) VALUES (:TASKID, :TASKDETID, :REQNM, :STARTDT, :STARTTM, :ENDDT, :ENDTM, :USRID, :TURNATM, :METHOD, :STATE, :WRTDT, :RESULTINFO, :HEAD,  :WRTUSRID,:COMPFLG)";

            // OracleCommand 객체 생성
            using (OracleCommand command = new OracleCommand(query, connection))
            {
                // 파라미터 추가
                
                command.Parameters.Add(new OracleParameter("TASKID", TASKID));
                command.Parameters.Add(new OracleParameter("TASKDETID", TASKDETID));
                command.Parameters.Add(new OracleParameter("REQNM", REQNM));
                command.Parameters.Add(new OracleParameter("STARTDT", STARTDT));
                command.Parameters.Add(new OracleParameter("STARTTM", STARTTM));
                command.Parameters.Add(new OracleParameter("ENDDT", ENDDT));
                command.Parameters.Add(new OracleParameter("ENDTM", ENDTM));
                command.Parameters.Add(new OracleParameter("USRID", USRID));
                command.Parameters.Add(new OracleParameter("TURNATM", TURNATM));
                command.Parameters.Add(new OracleParameter("METHOD", METHOD));
                command.Parameters.Add(new OracleParameter("STATE", STATE));
                command.Parameters.Add(new OracleParameter("WRTDT", WRTDT));
                command.Parameters.Add(new OracleParameter("RESULTINFO", RESULTINFO));
                command.Parameters.Add(new OracleParameter("HEAD", HEAD));
              //  command.Parameters.Add(new OracleParameter("WORKDAY", WORKDAY));
                command.Parameters.Add(new OracleParameter("WRTUSRID", WRTUSRID));
                command.Parameters.Add(new OracleParameter("COMPFLG", COMPFLG));

                try
                {
                    // 쿼리 실행
                    int rowsAffected = command.ExecuteNonQuery();

                    // 등록 결과 확인
                    if (rowsAffected > 0)
                    {
                        // 새로운 데이터를 데이터그리드뷰2에 표시하기 위해 데이터를 다시 로드
                        string sql2 = "SELECT TASKID, TASKDETID, REQNM, STARTDT, STARTTM, ENDDT, ENDTM, USRID, TURNATM, METHOD, STATE, WRTDT, RESULTINFO, HEAD, WORKDAY, WRTUSRID FROM TASKDET_INF_RENEW where taskid = "+ Task_id;
                        OracleDataAdapter adapter2 = new OracleDataAdapter(sql2, DBConnection.DBCon());
                        DataTable data_table2 = new DataTable();
                        adapter2.Fill(data_table2);
                        dataGridView2.DataSource = data_table2;

                    }

                    MessageBox.Show("값이 등록되었습니다.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("등록 중 오류가 발생했습니다: " + ex.Message);
                }
            }
        }

        // 등록 버튼 클릭 시 실행되는 클릭함수
        private void replybutton_Click(object sender, EventArgs e)
        {
            string TASKID = taskid.Text;
            string TASKDETID = taskdetid.Text;
            string REQNM = reqnm.Text;
            string STARTDT = startdt.Text;
            string STARTTM = starttm.Text;
            string ENDDT = enddt.Text;
            string ENDTM = endtm.Text;
            string USRID = usrid.Text;
            string TURNATM = turnatm.Text;
            string METHOD = method.SelectedItem.ToString();
            string STATE = state.SelectedItem.ToString();
            string WRTDT = wrtdt.Text;
            string RESULTINFO = resultinfo.Text;
            string HEAD = head.Text;
            //string WORKDAY = workday.Text;
            string WRTUSRID = wrtusrid.Text;
            string COMPFLG = "";
            if(openbtn.Checked == true) COMPFLG = "0";
            else if(closebtn.Checked == true) COMPFLG= "1";
            MessageBox.Show(COMPFLG);

            // 데이터베이스 연결
            using (OracleConnection connection = DBConnection.DBCon())
            {
                connection.Open();

                // 값을 DB에 INSERT
                InsertTaskDetInfo(TASKID, TASKDETID, REQNM, STARTDT, STARTTM, ENDDT, ENDTM, USRID, TURNATM, METHOD, STATE, WRTDT, RESULTINFO, HEAD, WRTUSRID, connection,COMPFLG);

                // 데이터베이스 연결 종료
                connection.Close();
            }
        }
        

        private void 종료ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void label18_Click(object sender, EventArgs e)
        {

        }
        // DataGridView1 (태스크정보 datagridview) 선택시 실행되는 함수
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            taskid.Text =  dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
            Delegates.usersend10();
            this.userid = userdto.usrid;
            wrtusrid.Text = userid;
            setText();
            reqnm.Text = dataGridView1.Rows[e.RowIndex].Cells[6].Value.ToString();

        }

        // 로그인한 사용자의 정보를 가져오는 함수
        public void SetUser(UserDTO dto)
        {
            userdto = dto;
        }

        // 작성할 textbox에 placeholder 기능을 구현한 함수
        public void setText()
        {
            wrtdt.Text = DateTime.Now.ToString("yyyy/MM/dd");
            startdt.Text = "YYYYMMDD";
            startdt.ForeColor = Color.Gray;
            enddt.Text = "YYYYMMDD";
            enddt.ForeColor = Color.Gray;
            turnatm.Text = "(분)";
            turnatm.ForeColor = Color.Gray;
            starttm.Text = "HH24MISS";
            starttm.ForeColor = Color.Gray;
            endtm.Text = "HH24MISS";
            endtm.ForeColor = Color.Gray;
        }

        // girdview의 크기를 조정하는 함수
        public void gridviewds()
        {
            dataGridView1.Columns["ID"].Width = 30;
            dataGridView1.Columns["PJNO"].Width = 30;
            dataGridView1.Columns["상태"].Width = 30;
            dataGridView1.Columns["담당자"].Width = 50;
            dataGridView1.Columns["요청자"].Width = 50;
            dataGridView1.Columns["접수일자"].Width = 80;
            dataGridView1.Columns["내용"].DefaultCellStyle.WrapMode = DataGridViewTriState.True;
        }

        // 작성할 textbox에 placeholder 기능을 구현한 함수
        private void placeholder(object sender, EventArgs e)
        {
            string textname = ((TextBox)sender).Name;
           
            
            switch (textname)
            {
                case "turnatm":
                    if (turnatm.Text == "(분)") turnatm.Text = ""; turnatm.ForeColor = Color.Black;
                    
                    break;
                case "startdt":
                    if (((TextBox)sender).Text == "YYYYMMDD") ((TextBox)sender).Text = ""; ((TextBox)sender).ForeColor = Color.Black;
                    
                    break;
                case "enddt":
                    if (((TextBox)sender).Text == "YYYYMMDD") ((TextBox)sender).Text = ""; ((TextBox)sender).ForeColor = Color.Black;
                    break;
                case "starttm":
                    if (starttm.Text == "HH24MISS") starttm.Text = ""; starttm.ForeColor = Color.Black;
                    break;
                case "endtm":
                    if (endtm.Text == "HH24MISS") endtm.Text = ""; endtm.ForeColor = Color.Black;
                    break;
            }
        }

        // 작성할 textbox에 placeholder 기능을 구현한 함수
        private void leaveplaceholder(object sender, EventArgs e)
        {
            string textname = ((TextBox)sender).Name;
            switch (textname)
            {
                case "turnatm":
                    
                    if (turnatm.Text == "") turnatm.Text = "(분)"; turnatm.ForeColor = Color.Gray;
                    break;
                case "startdt":
                   
                    if (((TextBox)sender).Text == "") ((TextBox)sender).Text = "YYYYMMDD"; ((TextBox)sender).ForeColor = Color.Gray;
                    break;
                case "enddt":
                    
                    if (((TextBox)sender).Text == "") ((TextBox)sender).Text = "YYYYMMDD"; ((TextBox)sender).ForeColor = Color.Gray;
                    break;
                case "starttm":
                    if (starttm.Text == "") starttm.Text = "HH24MISS"; starttm.ForeColor = Color.Gray;
                    break;
                case "endtm":
                    if (endtm.Text == "") endtm.Text = "HH24MISS"; endtm.ForeColor = Color.Gray;
                    break;
            }
        }
    }
}
