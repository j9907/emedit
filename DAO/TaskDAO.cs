using emedit.DTO;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Data.SqlTypes;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace emedit.DAO
{
    public class TaskDAO
    {
        private OracleConnection conn;

        // task와 관련한 전반적인 함수를 작성한 class 파일

        // form2(태스크 등록)화면에 combobox에 추가할 데이터를 데이터베이스에서 조회하여 반환
        public void combobox_item(string value, ComboBox co)
        {
            string sql = null;
            OracleDataReader reader;
            try
            {
                conn = DBConnection.DBCon();
                conn.Open();
                sql = "select basecd,basenm from base_cod where basecatg = '" + value + "' order by basecd asc";
                reader = DBConnection.DataSelect(sql, conn);
                // MessageBox.Show("reader");
                Dictionary<string, string> item1 = new Dictionary<string, string>();

                while (reader.Read())
                {

                    item1.Add(reader.GetString(0), reader.GetString(1));


                }

                co.DataSource = new BindingSource(item1, null);
                co.DisplayMember = "name";
                co.ValueMember = "value";

                conn.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    
        // task_inf 테이블에 데이터를 삽입하고, 삽입한 taskid를 조회하여 int형 배열에 저장해 반환
        public int[] inserttask(TaskDTO dto)
        {
            conn = DBConnection.DBCon();
            conn.Open();
            string sql = "insert into task_inf(taskid,prjtno,regdt,schstdt,schsttm,usrid,state,\"LEVEL\",wrtusrid,wrtdt,head,reqinfo,RECEIPTFLG,reqnm,email) " +
                "values (task_inf_seq.NEXTVAL,'"+dto.prjtno+"','"+dto.regdt+ "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'HH24MISS'),'" +dto.userid +"'" +
                ",'" +dto.state+ "', '" +dto.level+ "', '" +dto.wrtuserid+"','"+dto.wrtdt+"','" +dto.head+"','" +dto.reqinfo+"','"+dto.receiptflg+"','"+dto.reqnm+"','"+dto.email+"')";
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = conn;
            cmd.CommandText = sql;
            int result = cmd.ExecuteNonQuery();
            string seq = selectseq();
            int[] rul = new int[2] { result, Int32.Parse(seq) };
            conn.Close();
            return rul;
        }

        // form2(태스크등록)화면에 첨부파일이 존재할 경우 파일의 정보를 데이터베이스에 저장하는 함수
        public int uploadfile(string filename, string filepath,string taskid)
        {
            conn = DBConnection.DBCon();
            conn.Open();
            string sql = "insert into PGM_ATTACH_FILE_INF values (file_inf_seq.NEXTVAL,'"+taskid+"','0','"+filepath+"','"+filename+"',to_char(sysdate,'yyyymmdd'))";
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = conn;
            cmd.CommandText = sql;
            int result = cmd.ExecuteNonQuery();
            
            conn.Close();
            return result;
        }
    
        // task_inf 테이블의 수정 시 실행되는 함수
        public int updatetask(TaskDTO dto)
        {
            conn = DBConnection.DBCon();
            conn.Open();
            string sql = "update task_inf  set prjtno = :value1 , state = :value2 , \"LEVEL\" = :value3 , head = :value5, " +
                "reqinfo = :value6 , RECEIPTFLG = :value7 , reqnm = :value8 where taskid = " + dto.taskid;
            OracleCommand cmd = new OracleCommand(sql,conn);
            cmd.Parameters.Add("value1", dto.prjtno);
            cmd.Parameters.Add("value2", dto.state);
            cmd.Parameters.Add("value3", dto.level);
            cmd.Parameters.Add("value5", dto.head);
            cmd.Parameters.Add("value6", dto.reqinfo);
            cmd.Parameters.Add("value7", dto.receiptflg);
            cmd.Parameters.Add("value8", dto.reqnm);

            int result = cmd.ExecuteNonQuery();
            
            return result;
        }

        // task_inf 테이블을 조회하는 함수
        public TaskDTO selectTask(string TaskId)
        {
            conn = DBConnection.DBCon();
            conn.Open();
            
            string sql = "select taskid,prjtno,regdt,usrid,state,\"LEVEL\",wrtusrid,wrtdt,head,reqinfo,RECEIPTFLG,reqnm,email from Task_inf where taskid = " + TaskId;
            
            OracleCommand myCommand = new OracleCommand(sql, conn);
            OracleDataReader reader = myCommand.ExecuteReader();
            // MessageBox.Show("reader");
            TaskDTO dto = new TaskDTO();
            while (reader.Read())
            {
                
                dto.taskid = reader.GetString(0);
                dto.prjtno = reader.GetString(1);
                dto.regdt = reader.GetString(2);
                dto.userid = reader.GetString(3);
                dto.state = reader.GetString(4);
                dto.level = reader.GetString(5);
                dto.wrtuserid = reader.GetString(6);
                dto.wrtdt = reader.GetString(7);
                dto.head = reader.GetString(8);
                dto.reqinfo = reader.GetString(9);
                dto.receiptflg = reader.GetString(10);
                dto.reqnm = reader.GetString(11);
                dto.email = reader.GetString(12);
            }
            return dto;
        }

        // task_inf 테이블의 상태 컬럼을 수정하는 함수
        public int updatestate(string state,string task_id)
        {
            OracleConnection conn = DBConnection.DBCon();
            conn.Open();

            string sql = "update task_inf set state = '" + state +"' where taskid = '" + task_id + "'";

            OracleCommand myCommand = new OracleCommand(sql, conn);
            int result = myCommand.ExecuteNonQuery();
            return result;
        }

        //taskdet_inf_renew 테이블의 상테 컬럼을 수정하는 함수
        public int updatesubstate(string state,string task_id)
        {
           
            OracleConnection conn = DBConnection.DBCon();
            conn.Open();

            string sql = "update  taskdet_inf_renew set state = '" + state + "' where taskdetid = '" + task_id + "'";

            OracleCommand myCommand = new OracleCommand(sql, conn);
            int result = myCommand.ExecuteNonQuery();
            return result;
        }

        // 로그인한 사용자의 email 정보를 조회하는 함수

        public string reademail(string usrid)
        {
            OracleConnection conn = DBConnection.DBCon();
            conn.Open();

            string sql = "select emailid from usr_mst where usrkornm = '" + usrid + "'";
            string str = "";
            OracleCommand myCommand = new OracleCommand( sql, conn);
            OracleDataReader reader = myCommand.ExecuteReader();
            while(reader.Read())
            {
                 str = reader.GetString(0);  
            }

            return str;
        }

        // 입력된 task_inf 의 taskid 값을 조회하는 함수
        
        public string selectseq()
        {       
            string sql = "select task_inf_seq.CURRVAL from dual";
            string str = "";
            OracleCommand myCommand = new OracleCommand(sql, conn);
            OracleDataReader reader = myCommand.ExecuteReader(); 
            if(reader.Read()) str = reader.GetString(0);
           
            return str;

        }

        // task_inf 테이블의 데이터 중 email을 조회하는 함수
        public string getemail(string taskid)
        {
            OracleConnection conn = DBConnection.DBCon();
            conn.Open();
            string sql = "select email from task_inf where taskid = " + taskid;
            string str = "";
            OracleCommand oracleCommand = new OracleCommand(sql, conn);
            OracleDataReader reader = oracleCommand.ExecuteReader();
            if (reader.Read()) str = reader.GetString(0);
            conn.Close();
            return str;
        }
    }
}
