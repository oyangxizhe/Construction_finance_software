using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Globalization;
using System.Data.SqlClient;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using XizheC;

namespace XizheC
{
    public class CDETAIL_CONSTRUCTION_LOG
    {

        private string _getsql;
        public string getsql
        {
            set { _getsql = value; }
            get { return _getsql; ; }

        }
        private string _CLID;
        public string CLID
        {
            set { _CLID = value; }
            get { return _CLID; }

        }
        private string _PROJECT_NAME;
        public string PROJECT_NAME
        {
            set { _PROJECT_NAME = value; }
            get { return _PROJECT_NAME; }

        }
        private string _PROJECT_DATE;
        public string PROJECT_DATE
        {
            set { _PROJECT_DATE = value; }
            get { return _PROJECT_DATE; }

        }
        private string _PROJECT_PART;
        public string PROJECT_PART
        {
            set { _PROJECT_PART = value; }
            get { return _PROJECT_PART; }

        }
        private string _TEMPERATURE;
        public string TEMPERATURE
        {
            set { _TEMPERATURE = value; }
            get { return _TEMPERATURE; }

        }
        private string _WIND_SCALE;
        public string WIND_SCALE
        {
            set { _WIND_SCALE = value; }
            get { return _WIND_SCALE; }

        }
        private string _CONTENT_O;
        public string CONTENT_O
        {
            set { _CONTENT_O = value; }
            get { return _CONTENT_O; }

        }
        private string _CONTENT_T;
        public string CONTENT_T
        {
            set { _CONTENT_T = value; }
            get { return _CONTENT_T; }

        }
        private string _CONTENT_TH;
        public string CONTENT_TH
        {
            set { _CONTENT_TH = value; }
            get { return _CONTENT_TH; }

        }
        private string _MAKERID;
        public string MAKERID
        {
            set { _MAKERID = value; }
            get { return _MAKERID; }

        }
        private string _HANDLER;
        public string HANDLER
        {
            set { _HANDLER = value; }
            get { return _HANDLER; }

        }
        private string _PROJECT_MANAGE;
        public string PROJECT_MANAGE
        {
            set { _PROJECT_MANAGE = value; }
            get { return _PROJECT_MANAGE; }

        }
        private bool _IFExecutionSUCCESS;
        public bool IFExecution_SUCCESS
        {
            set { _IFExecutionSUCCESS = value; }
            get { return _IFExecutionSUCCESS; }

        }
        private string _ErrowInfo;
        public string ErrowInfo
        {

            set { _ErrowInfo = value; }
            get { return _ErrowInfo; }

        }
        private string _CONSTRUCTION_UNIT_DATE;
        public string CONSTRUCTION_UNIT_DATE
        {
            set { _CONSTRUCTION_UNIT_DATE = value; }
            get { return _CONSTRUCTION_UNIT_DATE; }

        }
        string sql = @"
SELECT
CLID AS 日志单号,
PROJECT_NAME AS 工程名称,
PROJECT_DATE AS 施工日期,
PROJECT_PART AS 工程部位,
TEMPERATURE AS 温度,
WIND_SCALE AS 风级,
CONTENT_O AS 施工质量,
CONTENT_T AS 检验情况,
CONTENT_TH AS 安全问题,
PROJECT_MANAGE AS 项目负责人,
HANDLER AS 经手人,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.MAKERID) AS 制单人,
DATE AS 制单日期
FROM 
DETAIL_CONSTRUCTION_LOG A
";


        string sql1 = @"
INSERT INTO
DETAIL_CONSTRUCTION_LOG(
CLID,
PROJECT_NAME,
PROJECT_DATE,
PROJECT_PART,
TEMPERATURE,
WIND_SCALE,
CONTENT_O,
CONTENT_T,
CONTENT_TH,
PROJECT_MANAGE,
HANDLER,
MAKERID,
DATE,
YEAR,
MONTH
) VALUES 

(
@CLID,
@PROJECT_NAME,
@PROJECT_DATE,
@PROJECT_PART,
@TEMPERATURE,
@WIND_SCALE,
@CONTENT_O,
@CONTENT_T,
@CONTENT_TH,
@PROJECT_MANAGE,
@HANDLER,
@MAKERID,
@DATE,
@YEAR,
@MONTH
)

";
        string sql2 = @"
UPDATE DETAIL_CONSTRUCTION_LOG SET 
CLID=@CLID,
PROJECT_NAME=@PROJECT_NAME,
PROJECT_DATE=@PROJECT_DATE,
TEMPERATURE=@TEMPERATURE,
WIND_SCALE=@WIND_SCALE,
CONTENT_O=@CONTENT_O,
CONTENT_T=@CONTENT_T,
CONTENT_TH=@CONTENT_TH,
PROJECT_MANAGE=@PROJECT_MANAGE,
HANDLER=@HANDLER,
MAKERID=@MAKERID,
DATE=@DATE,
YEAR=@YEAR,
MONTH=@MONTH


";
        basec bc = new basec();
        DataTable dt = new DataTable();
        DataTable dto = new DataTable();
        public CDETAIL_CONSTRUCTION_LOG()
        {
            IFExecution_SUCCESS = true;
            getsql = sql;
          

        }
        public string GETID()
        {
            string v1 = bc.numYM(10, 4, "0001", "SELECT * FROM DETAIL_CONSTRUCTION_LOG", "CLID", "CL");
            string GETID = "";
            if (v1 != "Exceed Limited")
            {
                GETID = v1;
            }
            return GETID;
        }
        #region save IDVALUE
        public void save(string TABLENAME, string COLUMNID, string COLUMNNAME, string IDVALUE, string NAMEVALUE, string INFOID, string INFONAME)
        {
            
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            string v1 = bc.getOnlyString("SELECT " + COLUMNID + " FROM " + TABLENAME + " WHERE " + COLUMNID + "='" + IDVALUE + "'");
            string v2 = bc.getOnlyString("SELECT " + COLUMNID + " FROM " + TABLENAME + " WHERE " + COLUMNNAME + "='" + NAMEVALUE + "'");
            //string varMakerID;
            if (!bc.exists("SELECT " + COLUMNID + " FROM " + TABLENAME + " WHERE " + COLUMNID + "='" + IDVALUE + "'"))
            {
              SQlcommandE(sql1, IDVALUE, NAMEVALUE);
            }
            else
            {
                SQlcommandE(sql2 + " WHERE " + COLUMNID + "='" + IDVALUE + "'", IDVALUE, NAMEVALUE);
            }
            IFExecution_SUCCESS = true;

          
        }
        #endregion

        #region SQlcommandE
        protected void SQlcommandE(string sql, string IDVALUE, string NAMEVALUE)
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            SqlConnection sqlcon = bc.getcon();
            SqlCommand sqlcom = new SqlCommand(sql, sqlcon);
            sqlcom.Parameters.Add("@CLID", SqlDbType.VarChar, 20).Value = CLID;
            sqlcom.Parameters.Add("@PROJECT_NAME", SqlDbType.VarChar, 20).Value = PROJECT_NAME;
            sqlcom.Parameters.Add("@PROJECT_DATE", SqlDbType.VarChar, 20).Value = PROJECT_DATE;
            sqlcom.Parameters.Add("@PROJECT_PART", SqlDbType.VarChar, 20).Value = PROJECT_PART;
            sqlcom.Parameters.Add("@TEMPERATURE", SqlDbType.VarChar, 20).Value = TEMPERATURE;
            sqlcom.Parameters.Add("@WIND_SCALE", SqlDbType.VarChar, 20).Value = WIND_SCALE;
            sqlcom.Parameters.Add("@CONTENT_O", SqlDbType.VarChar, 20).Value = CONTENT_O;
            sqlcom.Parameters.Add("@CONTENT_T", SqlDbType.VarChar, 20).Value = CONTENT_T;
            sqlcom.Parameters.Add("@CONTENT_TH", SqlDbType.VarChar, 20).Value = CONTENT_TH;
            sqlcom.Parameters.Add("@PROJECT_MANAGE", SqlDbType.VarChar, 20).Value = PROJECT_MANAGE;
            sqlcom.Parameters.Add("@HANDLER", SqlDbType.VarChar, 20).Value = HANDLER;
            sqlcom.Parameters.Add("@MAKERID", SqlDbType.VarChar).Value = MAKERID;
            sqlcom.Parameters.Add("@DATE", SqlDbType.VarChar).Value = varDate;
            sqlcom.Parameters.Add("@YEAR", SqlDbType.VarChar).Value = year;
            sqlcom.Parameters.Add("@MONTH", SqlDbType.VarChar).Value = month;
            sqlcom.Parameters.Add("@DAY", SqlDbType.VarChar).Value = year;
            sqlcon.Open();
            sqlcom.ExecuteNonQuery();
            sqlcon.Close();
        }
        #endregion
    }
}
