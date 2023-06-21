using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Oracle.ManagedDataAccess.Client;

namespace emedit
{

    public class DBConnection
    {
       // OracleConnection Conn;

        // 데이터베이스에 연결하는 connection class
        public static OracleConnection DBCon()
        {
            OracleConnection Conn;

            string ConStr = "data source=(DESCRIPTION =(ADDRESS = (PROTOCOL = TCP)(HOST = 121.125.32.217)(PORT = 1521))" +
                "(CONNECT_DATA =(SERVER = DEDICATED)(SID = umcit))); User Id=ksedu; Password=ksedu";
            Conn = new OracleConnection(ConStr);
            return Conn;
        }
        public static OracleDataReader DataSelect(string sql, OracleConnection Conn)
        {
            try
            {
                OracleCommand myCommand = new OracleCommand(sql, Conn);
                return myCommand.ExecuteReader();
            }
            catch (Exception ex)
            {
                MessageBox.Show(sql + "\n" + ex.Message);
                return null;
            }
        }

        
    }
}
