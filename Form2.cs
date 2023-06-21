using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using emedit.DAO;
using emedit.DTO;

namespace emedit
{
    public partial class Form2 : Form
    {
        // 태스크(요청사항)등록 화면
        readonly private Form8 f8 = new Form8();
        OracleConnection conn;
        private string filePath;
        private string filename = null;

        
        static UserDTO userdto = new UserDTO();
        TaskDAO taskdao = new TaskDAO();
        TaskDTO taskdto = new TaskDTO();
        string Taskid = "0";

        public Form2()
        {
            InitializeComponent();
        }

        public Form2(string taskid)
        {
            InitializeComponent();
            this.Taskid = taskid;
            //if(taskid != "0") MessageBox.Show(taskid);
            

        }

        // combobox1.item.Add로 combobox에 추가

        // Form2 가 Load시 실행되는 함수
        private void Form2_Load(object sender, EventArgs e)
        {
            // combobox에 데이터를 데이터베이스에서 조회하여 가져오는 함수
            taskdao.combobox_item("STATE", statecombo);
            taskdao.combobox_item("LEVEL", levelcombo);
            comboinit();
            // 등록
            if (Taskid == "0")
            {
                wrtdttext.Text = DateTime.Now.ToString("yyyy/MM/dd");
                
                start_dt.Text = "YYYYMMDD";
                start_dt.ForeColor = Color.Gray;
                Delegates.usersend2();
                wrtuser_id.Text = userdto.usrname;
            }
            // 수정
            else
            {
                TaskDTO dto = new TaskDTO();
                dto = taskdao.selectTask(Taskid);
                //MessageBox.Show(dto.taskid);
                int state = selectindex(statecombo, dto.state);
                statecombo.SelectedIndex = state;
                int level = selectindex(levelcombo, dto.level);
                levelcombo.SelectedIndex = level;
                inputdata(dto);
            }
        }

        // 요청구분 combobox에 데이터를 입력하는 함수

        public void comboinit()
        {
            combocall.Items.Add("공급과정");
            combocall.Items.Add("무상보증기간");
            combocall.Items.Add("영업");
            combocall.Items.Add("유상계약");
            combocall.SelectedIndex = 0;
        }

        // 수정 시 선택한 combobox 데이터의 index 값을 반환하는 함수
        public int selectindex(ComboBox combobox,string text)
        {
            int index = combobox.FindString(text);
            return index;
        }

        // 업로드 버튼 클릭 시 실행되는 함수
        private void button3_Click(object sender, EventArgs e)
        {
            uploadtext.Text = showFileOpenDialog();
        }

        // 업로드 버튼 클릭 시 선택한 파일의 경로와 파일명을 변수에 저장하는 함수
        public string showFileOpenDialog()
        {
            OpenFileDialog ofd = openFileDialog1;
            DialogResult dr = ofd.ShowDialog();

            if (dr == DialogResult.OK)
            {
                filename = ofd.SafeFileName;
                filePath = ofd.FileName.Replace(filename, "");

            }

            return filename;
        }

        // 프로젝트NO 조회 버튼 클릭 시 실행되는 함수
        private void button4_Click(object sender, EventArgs e)
        {
            f8.ShowDialog();
            prjt_no.Text = f8.prjtno();
        }

        // TaskDTO 클래스에 입력한 값을 저장하는 함수
        private TaskDTO textinsert()
        {
            TaskDTO dto = new TaskDTO();
            {
                
                dto.prjtno = prjt_no.Text;
                dto.regdt = start_dt.Text;
                dto.userid = user_Id.Text;
                
                dto.level = levelcombo.SelectedValue.ToString();
                dto.wrtuserid = wrtuser_id.Text;
                dto.wrtdt = wrtdttext.Text;
                dto.state = statecombo.SelectedValue.ToString();
                dto.receiptflg = combocall.Text.ToString();
                dto.reqnm = textBox10.Text;
                dto.head = textBox7.Text;
                dto.reqinfo = textBox5.Text;
                dto.email = emailbox.Text;
                if(Taskid != "0") dto.taskid = Taskid;
            }

            return dto;
        }

        // 데이터 수정 시 데이터베이스에 저장한 값을 textbox에 저장하는 함수
        private void inputdata(TaskDTO dto)
        {
             
             prjt_no.Text = dto.prjtno;
            start_dt.Text = dto.regdt;
            user_Id.Text = dto.userid;
            
            wrtuser_id.Text = dto.wrtuserid;
            wrtdttext.Text = dto.wrtdt;
            textBox10.Text = dto.reqnm;
            textBox7.Text = dto.head;   
            textBox5.Text = dto.reqinfo;
            emailbox.Text = dto.email;
        }
        
        // 등록 버튼 선택 시 실행되는 함수
        private void button1_Click(object sender, EventArgs e)
        {
            int fileresult = 0;
            // 입력한 데이터를 dto에 전달
            taskdto = textinsert();
            int result = 0;
            int[] results = new int[2];
            string seq = "";
            // 태스크 정보를 등록할 시 실행하는 함수
            if (Taskid == "0") {
                // task_inf 데이터베이스에 삽입하는 함수
                results = taskdao.inserttask(taskdto);
                
                result = results[0];
                seq = results[1].ToString();
                if (!string.IsNullOrEmpty(uploadtext.Text))
                {
                    //MessageBox.Show(seq);
                    // 파일 설명을 데이터베이스에 저장하는 함수
                    fileresult = taskdao.uploadfile(filename, filePath, seq);
                }
            }
            // 태스크 정보를 수정할 시 실행하는 함수
            else result = taskdao.updatetask(taskdto);
            if (result == 1 || fileresult == 1 && result == 1)
            {
                MessageBox.Show("등록이 완료되었습니다!");
                if(Taskid == "0")
                {
                    // 모든 등록이 완료된 후 요청자에게 이메일 전송
                    taskdto.usremail = taskdao.reademail(user_Id.Text);
                    sendmail sm = new sendmail();
                    sm.sendtomail(taskdto);
                }
                this.Close();
                Form1 f1 = new Form1();
                f1.ShowDialog();
            }
            else
            {
                MessageBox.Show("등록이 실패되었습니다. 다시 시도해 주십시오");
            }
        }

        // 취소 버튼 클릭 시 실행되는 함수
        private void button2_Click(object sender, EventArgs e)
        {
            Close();
            Form1 Form1 = new Form1();
            Form1.Show();
           
        }

        // placeholder 구현 함수
        private void start_dt_Enter(object sender, EventArgs e)
        {
            if (start_dt.Text == "YYYYMMDD") start_dt.Text = ""; start_dt.ForeColor = Color.Black;
        }

        private void start_dt_Leave(object sender, EventArgs e)
        {
            if (start_dt.Text == "") start_dt.Text = "YYYYMMDD"; start_dt.ForeColor = Color.Gray;
        }

        // 로그인한 사용자의 데이터를 저장하는 함수
        public void SetUser(UserDTO dto)
        {
            userdto = dto;
        
        }

        private void label13_Click(object sender, EventArgs e)
        {

        }
    }
}
