using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using emedit.DTO;
using Oracle.ManagedDataAccess.Client;

namespace emedit
{
    public partial class Form4 : Form
    {
        // 거래처 조회 화면
        OracleConnection Conn;
        OracleCommand cmd;
        static UserDTO userdto = new UserDTO();

        public Form4()
        {
            InitializeComponent();
        }

        // form4 load시 실행되는 함수
        private void Form4_Load(object sender, EventArgs e)
        {
            combobox_item("TRDSTATE", comboBox1);
            Delegates.usersend4();
            textBox4.Text = userdto.usrid;
        }

        //콤보 박스에 DB값 읽어 오기
        public void combobox_item(string value, ComboBox co)
        {
            string sql = null;
            OracleDataReader reader;

            try
            {
                Conn = DBConnection.DBCon();
                Conn.Open();

                sql = "select basecd,basenm from base_cod where basecatg = '" + value + "' order by basecd asc";
                reader = DBConnection.DataSelect(sql, Conn);
                Dictionary<string, string> item1 = new Dictionary<string, string>();

                while (reader.Read())
                {
                    item1.Add(reader.GetString(0), reader.GetString(1));
                }

                co.DataSource = new BindingSource(item1, null);
                co.DisplayMember = "name";
                co.ValueMember = "value";
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        // 등록 버튼 클릭 시 거래처 정보 DB에 저장
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                Conn = DBConnection.DBCon();
                Conn.Open();

                //DB CUSTOM_MST에 거래처 코드, 사업자 등록번호, 대표자명, 거래처 구분, 등록자 ID, 거래처명, 전화번호, FAX번호, 주거래 내용 등록
                string sql = "INSERT INTO CUSTOM_MST (TRDCD, CMPYREGNO, CMPYNM, TRDSTATE, PRSNTNM, CMPYTEL, CMPYFAX, WRTDT, WRTID, TRDITEMS)" +
                    "values ('" + textBox1.Text + "', '" + textBox2.Text + "', '" + textBox5.Text + "', '" + comboBox1.SelectedIndex.ToString()+ "', '" + textBox3.Text + "', '" + textBox7.Text + "', '" + textBox8.Text + "',to_char(sysdate,'YYYYMMDD'), '" + textBox4.Text + "', '" + textBox6.Text + "')";

                cmd = new OracleCommand(sql, Conn);
                cmd.ExecuteNonQuery();

                MessageBox.Show("등록되었습니다.");
                Conn.Close();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            //Form1로 이동
            this.Visible = false;
            Form1 Form1 = new Form1();
            Form1.Show();
        }

        //취소 버튼 클릭 시 실행 함수
        private void button2_Click(object sender, EventArgs e)
        {
            Close();

            Form1 Form1 = new Form1();
            Form1.Show();
        }

       /* //거래처 조회로 이동
        private void 거래처조회ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form9 Form9 = new Form9();
            Form9.Show();
        }*/
        // 로그인한 사용자의 정보를 저장하는 함수
        public void SetUser(UserDTO dto)
        {
            userdto = dto;
           
        }
    }
}
