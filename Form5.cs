using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using emedit.DAO;
using Oracle.ManagedDataAccess.Client;

namespace emedit
{
    public partial class Form5 : Form
    {
        // 태스크 상세 조회 및 회신 화면 
        private string searchKeyword = string.Empty;
        string taskid = "";
        string filepathname = string.Empty;
        public Form5()
        {
            InitializeComponent();
        }

        public Form5(string taskid)
        {
            InitializeComponent();
            this.taskid = taskid;
        }

        // form5 load 시 실행되는 화면
        private void Form5_Load(object sender, EventArgs e)
        {
            try
            {
                OracleConnection con = DBConnection.DBCon();
                con.Open();

                string sql = "SELECT TASKID AS ID,TASKDETID AS DETID,REQNM AS 요청자,HEAD AS 제목, RESULTINFO AS 작업결과,STARTDT AS 시작일자," +
                    "ENDDT AS 종료일자,USRID AS 작업자,TURNATM AS 소요시간,METHOD AS 작업방법,  " +
                    "STATE AS 상태,WRTDT AS 작성일자, COMPFLG AS 구분 FROM TASKDET_INF_RENEW where taskid = " + taskid;

                OracleDataAdapter adapter = new OracleDataAdapter(sql, DBConnection.DBCon());
                DataTable data_table = new DataTable();

                adapter.Fill(data_table);
                dataGridView1.DataSource = data_table;

                con.Close();
                display();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        // DataGireView1을 의 데이터를 선택 시 실행되는 함수
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
          

            if (e.RowIndex >= 0) // 유효한 행을 클릭했을 때만 실행
            {
                DataGridViewRow selectedRow = dataGridView1.Rows[e.RowIndex];

                //필요한 셀 값을 가져와 텍스트 박스에 설정
                string cellValue1 = selectedRow.Cells["요청자"].Value.ToString(); //요청자명
                string cellValue2 = selectedRow.Cells["제목"].Value.ToString(); //제목
                string cellValue3 = selectedRow.Cells["작업자"].Value.ToString(); //작업자
                string cellValue4 = selectedRow.Cells["작업방법"].Value.ToString(); //작업방법
                string cellValue5 = selectedRow.Cells["작업결과"].Value.ToString(); //작업결과
                //string cellValue6 = selectedRow.Cells["WORKDAY"].Value.ToString(); //작업시간
                string cellValue7 = selectedRow.Cells["시작일자"].Value.ToString(); //작업시작일자
                string cellValue8 = selectedRow.Cells["종료일자"].Value.ToString(); //작업종료일자
                //string cellValue9 = selectedRow.Cells["WRTUSRID"].Value.ToString(); //작성자ID
                string cellValue10 = selectedRow.Cells["작성일자"].Value.ToString(); //작성일자
                string cellValue11 = selectedRow.Cells["구분"].Value.ToString(); // 공개구분

                //
               
                //MessageBox.Show(cellValue11);
                // 필요한 셀 값들을 추가로 가져오거나 처리
                string formattedText =
                $"요청자명: {cellValue1}\r\n제목: {cellValue2}\r\n" +
                $"작업자: {cellValue3}\r\n작업방법: {cellValue4}\r\n" +
                $"작업결과: {cellValue5}\r\n" +
                $"작업시작일자: {cellValue7}\r\n작업종료일자: {cellValue8}\r\n" +
                $"작성일자: {cellValue10}\r\n";
                if(cellValue11 == "1") textBox1.Text = formattedText;
                if(cellValue11 == "0" || cellValue11 == "") textBox2.Text = formattedText;
            }
        }

        //키워드 검색
        private void dtsTextbox_TextChanged(object sender, EventArgs e)
        {
            // 검색어를 가져온다.
            searchKeyword = dtsTextbox.Text;

            // 검색어가 비어 있을 경우 모든 행의 색상을 기본 색상으로 설정한다.
            if (string.IsNullOrEmpty(searchKeyword))
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    row.DefaultCellStyle.BackColor = Color.White; // 기본 색상으로 설정
                }
                return;
            }

            // 데이터 그리드 뷰의 모든 행을 순회하면서 검색어와 일치하는 행을 특정한 색상으로 표시한다.
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                bool isMatch = false;

                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Value != null && cell.Value.ToString().Contains(searchKeyword))
                    {
                        isMatch = true;
                        break;
                    }
                }

                if (isMatch)
                {
                    row.DefaultCellStyle.BackColor = Color.Yellow; // 특정한 색상으로 설정
                }
                else
                {
                    row.DefaultCellStyle.BackColor = Color.White; // 기본 색상으로 설정
                }
            }
        }


        // 첨부파일 Listbox에 delete 키 선택시 실행되는 이벤트 함수
        private void listBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                // Delete 키를 눌렀을 때 선택된 파일 삭제
                if (listBox1.SelectedIndex != -1)
                {
                    string selectedFile = listBox1.SelectedItem.ToString();
                    File.Delete(selectedFile);
                    listBox1.Items.RemoveAt(listBox1.SelectedIndex);
                }
            }
        }

        // Listbox에 있는 파일 명을 두번 클릭 시 실행되는 이벤트 함수
        private void listBox1_DoubleClick(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex != -1)
            {
                string selectedFile = listBox1.SelectedItem.ToString();
                File.Delete(selectedFile);
                listBox1.Items.RemoveAt(listBox1.SelectedIndex);
            }
        }

        // 파일 선택 버튼 클릭 시 실행되는 함수
        private void fileselectbutton_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true; // 여러 파일 선택을 허용합니다.

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string[] fileNames = openFileDialog.FileNames;
                foreach (string fileName in fileNames)
                {
                    filepathname = fileName;
                    // 파일 경로에서 파일 이름을 추출하여 ListBox에 추가합니다.
                    string displayFileName = Path.GetFileName(fileName);
                    listBox1.Items.Add(displayFileName);
                }
            }
        }

        // 회신 버튼 클릭 시 실행 함수
        private void replybutton_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("회신하시겠습니까?", "알림", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            if (dr == DialogResult.OK)
            {
                TaskDAO dao = new TaskDAO();
                string email = dao.getemail(taskid);
                sendmail sendmail = new sendmail();
                if (filepathname == "" || filepathname == string.Empty) filepathname = "x";
                sendmail.sendtorev(email, textBox2.Text, filepathname);
                MessageBox.Show("회신 완료되었습니다.");
            }
            else
            {
                MessageBox.Show("회신 취소되었습니다.");
            }
        }

        // 취소 버튼 클릭 시 실행되는 함수

        private void cancelbutton_Click(object sender, EventArgs e)
        {
            Close();
            Form1 showForm1 = new Form1();
            showForm1.Show();
        }

        // DatagirdView 디자인
        private void display()
        {
            dataGridView1.Columns["ID"].Width = 30;
            dataGridView1.Columns["DETID"].Width = 40;
            dataGridView1.Columns["요청자"].Width = 60;
            //dataGridView1.Columns["제목"].Width = 50;
            dataGridView1.Columns["작업자"].Width = 60;
            dataGridView1.Columns["작업방법"].Width = 80;
            dataGridView1.Columns["작업결과"].Width = 80;
            dataGridView1.Columns["시작일자"].Width = 80;
            dataGridView1.Columns["종료일자"].Width = 80;
            dataGridView1.Columns["소요시간"].Width = 60;
            dataGridView1.Columns["작성일자"].Width = 80;
            dataGridView1.Columns["상태"].Width = 50;
            dataGridView1.Columns["구분"].Width = 40;
            dataGridView1.Columns["작업결과"].DefaultCellStyle.WrapMode = DataGridViewTriState.True;
        }
    }
}
