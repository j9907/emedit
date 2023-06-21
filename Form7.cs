using emedit.DTO;
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



namespace emedit
{
    // 로그인한 사용자 정보를 전달하는 이벤트핸들러
    public delegate void UserEventHandler(UserDTO dto);
    public partial class Form7 : LoginForm.LoginForm
    {
        // 로그인 화면
        private OracleConnection LocalConn;
        public UserEventHandler UserEventHandler;
        UserDTO dto = new UserDTO();

        public Form7()
        {
            InitializeComponent();
        }

        private void Form7_Load(object sender, EventArgs e)
        {

        }
        // 로그인 버튼 클릭 시 실행되는 함수
        private void button1_Click(object sender, EventArgs e)
        {

            string id = base.textid.Text;
            string pw = base.textpwd.Text;
            string shapw = SHA256Hash(pw);
            OracleDataReader myReader;
            string sql = null;

            try
            {
                LocalConn = DBConnection.DBCon();
                LocalConn.Open();
                sql = "select PASSWD,USRID,USRKORNM from USR_MST where USRID = '" + id + "'";
                myReader = DBConnection.DataSelect(sql, LocalConn);
                if (myReader.Read())
                {
                    string row = myReader.GetString(0);
                    dto.usrid = myReader.GetString(1);
                    dto.usrname = myReader.GetString(2);
                    bool result = isSame(pw, row);
                    if (result)
                    {
                        this.Visible = false;
                        Form1 Form1 = new Form1();
                        Delegates dele = new Delegates();
                        this.UserEventHandler += new UserEventHandler(dele.SetUser);
                        UserEventHandler(dto);
                        Form1.Show();
                    }
                    else
                    {
                        MessageBox.Show("로그인 실패");
                    }
                }
                else
                {
                    MessageBox.Show("비밀번호가 틀렸거나 잘못된 아이디 입니다.");
                }

            }
            catch (Exception ex) { }
        }

        // sha256으로 비밀번호를 암호화하는 함수
        public static string SHA256Hash(string data)
        {
            System.Security.Cryptography.SHA256 sha = new System.Security.Cryptography.SHA256Managed();
            byte[] hash = sha.ComputeHash(Encoding.ASCII.GetBytes(data));
            StringBuilder stringBuilder = new StringBuilder();
            foreach (byte b in hash)
            {
                stringBuilder.AppendFormat("{0:x2}", b);
            }
            return stringBuilder.ToString();
        }

        // 암호화된 비밀번호가 데이터베이스의 암호화된 비밀번호와 일치하는지 확인하는 함수
        public static bool isSame(string pw1, string pw2)
        {
            string shapw = SHA256Hash(pw1);
            if (pw2 == shapw)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

    }
}
