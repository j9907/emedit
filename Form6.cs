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
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace emedit
{
    public partial class Form6 : Form
    {
        // 프로젝트 등록 화면
        public Form6()
        {
            InitializeComponent();
        }
        OracleConnection Conn;
        OracleCommand cmd;
        public static UserDTO userdto = new UserDTO();


        public Form6(OracleConnection Conn)
        {
            InitializeComponent();
            this.Conn = Conn;
        }

        private void Form6_Load(object sender, EventArgs e)
        {
            Delegates.usersend6();
            textBox8.Text = userdto.usrid;
        }

        //등록 버튼
        private void button1_Click(object sender, EventArgs e)
        {
            string prjtno, prjtnm, startdt, enddt, trdcd, wrtusrid, wrtdt;
            int dcount = 0;
            prjtno = textBox2.Text;
            prjtnm = textBox1.Text;
            startdt = textBox6.Text;
            enddt = textBox3.Text;
            trdcd = textBox9.Text;
            wrtusrid = textBox8.Text;
            wrtdt = textBox5.Text;

            string query = "INSERT INTO project_mst (PRJTNO,PRJTNM,USEFLG,STARTDT,ENDDT,TRDCD,DISPSEQ,WRTUSRID,WRTDT) " +
                "VALUES('" + prjtno + "', '" + prjtnm + "','" + 1 + "' , N'" + startdt + "' , N'" + enddt + "' , N'" + trdcd + "' , '" + dcount + "' , '" + wrtusrid + "' , N'" + wrtdt + "' ) ";

            cmd = new OracleCommand(query, Conn);

            try
            {
                Conn.Open();
                cmd.ExecuteNonQuery();
                dcount++;
                MessageBox.Show("등록되었습니다.");
                textBox2.Text = ""; textBox1.Text = ""; textBox6.Text = ""; textBox3.Text = ""; textBox9.Text = ""; textBox8.Text = ""; textBox5.Text = "";
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            finally
            {
                Conn.Close();
            }
        }

        
       /* private void button2_Click(object sender, EventArgs e)
        {
            Close();

            Form1 Form1 = new Form1();
            Form1.Show();
        }*/
        // 로그인한 사용자의 정보를 저장
        public void SetUser(UserDTO dto)
        {
            userdto = dto;

        }

        //취소 버튼 클릭 시 실행 함수

        private void button2_Click_1(object sender, EventArgs e)
        {
            this.Close();
            Form1 f1 = new Form1();
            f1.Show();
        }
    }
}

