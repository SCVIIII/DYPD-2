using Autodesk.AutoCAD.DatabaseServices;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static DYPD_2.FClass;

namespace DYPD_2
{
    public static class USRSQLTools
    {
        /// <summary>
        /// 检测是否自动登录成功
        /// </summary>
        /// <param name="db"></param>
        /// <returns></returns>
        public static string auto_login(Database db)
        {
            //230804修改:取消ini文件,改为由程序预设内容自动登录
            //获取服务器地址
            string strServer = "sh-cdb-o3ki0rsa.sql.tencentcdb.com";
            //获取登录用户
            string strUserID = "common_01";
            //获取登录密码
            string strPwd = "dq123456";
            //获取端口
            string strPort = "63558";
            //获取库名
            //string strDataBase = OperatorFile.GetIniFileString("MySQL", "DataBase", "", Path.GetDirectoryName(db.OriginalFileName) + "\\ERP.ini");
            //数据库连接信息
            string connStr = "server = " + strServer + " ;user = " + strUserID + " ;port=" + strPort + ";password = " + strPwd;

            //创建数据库连接
            MySqlConnection connection;
            //连接MYSQL数据库
            try
            {
                connection = new MySqlConnection(connStr);
                connection.Open();
                connection.Close();
                return connStr;
                //ini文件存在,且登录成功
            }
            catch (MySqlException ex)
            {
                //ini文件存在,但登录失败
                connection = null;
                return "";
            }
            
        } //end of auto_login


        /// <summary>
        /// 读取MYSQL表，查询电缆规格，并转换为CABLE_CLASS类
        /// 电缆表：
        /// </summary>
        /// <param name="args"></param>
        public static CABLE_CLASS query_Cable(MySqlConnection conn, int Izd)
        {
            //数据库连接
            if (conn.State == System.Data.ConnectionState.Closed)
            {
                conn.Open();
            }

            //返回值
            CABLE_CLASS result = new CABLE_CLASS();
            //数据库查询语句
            string query = "SELECT * FROM base.dypd_ercipei_cableswitch WHERE Izd= " + Izd;

            //数据库查询
            using (var command = new MySqlCommand(query, conn))
            {
                using (var reader = command.ExecuteReader())
                {

                    while (reader.Read())
                    {
                        result.Izd = Izd;
                        result.Cable = reader["Cable"].ToString();
                        result.CablePE = reader["CablePE"].ToString();
                        result.SC = reader["SC"].ToString();
                        result.MCB = reader["MCB"].ToString();
                        result.INS = reader["INS"].ToString();
                        result.FU = reader["FU"].ToString();
                        result.CT = reader["CT"].ToString();

                    }
                } //end of using (var reader = command.ExecuteReader())
            } //end of using (var command = new MySqlCommand(query, conn))

            return result;
        } //end of public static void sql_CBClass(MySqlConnection conn)




        /// <summary>
        /// 断路器表（施耐德）：dypd_ercipei_mcb_schneider
        /// </summary>
        /// <param name="conn"></param>
        /// <param name="Izd"></param>
        /// <returns></returns>
        public static MCB_CLASS query_MCB(MySqlConnection conn, int Izd)
        {

            
            //数据库连接
            if (conn.State == System.Data.ConnectionState.Closed)
            {
                conn.Open();
            }

            //新建返回值
            MCB_CLASS result = new MCB_CLASS();
            //数据库查询语句
            string query = "SELECT * FROM base.dypd_ercipei_mcb_schneider WHERE Izd= " + Izd;

            //数据库查询
            using (var command = new MySqlCommand(query, conn))
            {
                using (var reader = command.ExecuteReader())
                {

                    while (reader.Read())
                    {
                        result.Izd = Izd;
                        result.Info1 = reader["Info1"].ToString();
                        result.Info2 = reader["Info2"].ToString();
                        result.Info3 = reader["Info3"].ToString();
                        result.Info4 = reader["Info4"].ToString();
                    }
                } //end of using (var reader = command.ExecuteReader())
            } //end of using (var command = new MySqlCommand(query, conn))

            //返回结果
            return result;
        } //end of public static MCB_CLASS query_MCB(MySqlConnection conn, int Izd)






    } // end of public static class USRSQLTools
} //end of namespace
