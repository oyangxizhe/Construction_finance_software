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
    public class CPROJECT_VISA
    {

        private string _getsql;
        public string getsql
        {
            set { _getsql = value; }
            get { return _getsql; ; }

        }
        private string _PWD;
        public string PWD
        {
            set { _PWD = value; }
            get { return _PWD; }

        }
        private string _PVID;
        public string PVID
        {
            set { _PVID = value; }
            get { return _PVID; }

        }
        private string _PROJECT_NAME;
        public string PROJECT_NAME
        {
            set { _PROJECT_NAME = value; }
            get { return _PROJECT_NAME; }

        }
        private string _PROJECT_DART;
        public string PROJECT_DART
        {
            set { _PROJECT_DART = value; }
            get { return _PROJECT_DART; }

        }
        private string _VISA_BECAUSE_AND_CONTENT;
        public string VISA_BECAUSE_AND_CONTENT
        {
            set { _VISA_BECAUSE_AND_CONTENT = value; }
            get { return _VISA_BECAUSE_AND_CONTENT; }

        }
        private string _DETAIL_CONSTRUCTION_UNIT;
        public string DETAIL_CONSTRUCTION_UNIT
        {
            set { _DETAIL_CONSTRUCTION_UNIT = value; }
            get { return _DETAIL_CONSTRUCTION_UNIT; }

        }
        private string _DETAIL_CONSTRUCTION_UNIT_PROJECT_MANAGE;
        public string DETAIL_CONSTRUCTION_UNIT_PROJECT_MANAGE
        {
            set { _DETAIL_CONSTRUCTION_UNIT_PROJECT_MANAGE = value; }
            get { return _DETAIL_CONSTRUCTION_UNIT_PROJECT_MANAGE; }

        }
        private string _HANDLER;
        public string HANDLER
        {
            set { _HANDLER = value; }
            get { return _HANDLER; }

        }
        private string _SUPVISTION_DEPART_ADVANCE;
        public string SUPVISTION_DEPART_ADVANCE
        {
            set { _SUPVISTION_DEPART_ADVANCE = value; }
            get { return _SUPVISTION_DEPART_ADVANCE; }

        }
        private string _SUPVISOR;
        public string SUPVISOR
        {
            set { _SUPVISOR = value; }
            get { return _SUPVISOR; }

        }
        private string _GENERAL_SUPVISOR;
        public string GENERAL_SUPVISOR
        {
            set { _GENERAL_SUPVISOR = value; }
            get { return _GENERAL_SUPVISOR; }

        }
        private string _HANDL_DATE;
        public string HANDL_DATE
        {
            set { _HANDL_DATE = value; }
            get { return _HANDL_DATE; }

        }
        private string _SUPVISTION_DATE;
        public string SUPVISTION_DATE
        {
            set { _SUPVISTION_DATE = value; }
            get { return _SUPVISTION_DATE; }

        }
        private string _CONSTRUCTION_UNIT;
        public string CONSTRUCTION_UNIT
        {
            set { _CONSTRUCTION_UNIT = value; }
            get { return _CONSTRUCTION_UNIT; }

        }
        private string _REP;
        public string REP
        {
            set { _REP = value; }
            get { return _REP; }

        }
        private string _CONSTRUCTION_UNIT_PROJECT_MANAGE;
        public string CONSTRUCTION_UNIT_PROJECT_MANAGE
        {
            set { _CONSTRUCTION_UNIT_PROJECT_MANAGE = value; }
            get { return _CONSTRUCTION_UNIT_PROJECT_MANAGE; }

        }
        private string _MAKERID;
        public string MAKERID
        {
            set { _MAKERID = value; }
            get { return _MAKERID; }

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
PVID AS 签证单号,
PROJECT_NAME AS 工程名称,
PROJECT_DART AS 工程部位,
VISA_BECAUSE_AND_CONTENT AS 签证原因和内容,
DETAIL_CONSTRUCTION_UNIT AS 施工单位,
DETAIL_CONSTRUCTION_UNIT_PROJECT_MANAGE AS 项目负责人,
HANDLER AS 经手人,
HANDL_DATE AS 施工方签字日期,
SUPVISTION_DEPART_ADVANCE AS 项目监理机构意见,
SUPVISOR AS 监理工程师,
GENERAL_SUPVISOR AS 总监理工程师,
SUPVISTION_DATE AS 监理方签字日期,
CONSTRUCTION_UNIT AS 建设单位,
REP AS 现场代表,
CONSTRUCTION_UNIT_PROJECT_MANAGE AS 负责人,
CONSTRUCTION_UNIT_DATE AS 建设方签字日期,
(SELECT ENAME FROM EMPLOYEEINFO WHERE EMID=A.MAKERID) AS 制单人,
DATE AS 制单日期
FROM 
PROJECT_VISA A
";


        string sql1 = @"
INSERT INTO
PROJECT_VISA(
PVID,
PROJECT_NAME,
PROJECT_DART,
VISA_BECAUSE_AND_CONTENT,
DETAIL_CONSTRUCTION_UNIT,
DETAIL_CONSTRUCTION_UNIT_PROJECT_MANAGE,
HANDLER,
HANDL_DATE,
SUPVISTION_DEPART_ADVANCE,
SUPVISOR,
GENERAL_SUPVISOR,
SUPVISTION_DATE,
CONSTRUCTION_UNIT,
REP,
CONSTRUCTION_UNIT_PROJECT_MANAGE,
CONSTRUCTION_UNIT_DATE,
MAKERID,
DATE,
YEAR,
MONTH,
DAY
) VALUES 

(
@PVID,
@PROJECT_NAME,
@PROJECT_DART,
@VISA_BECAUSE_AND_CONTENT,
@DETAIL_CONSTRUCTION_UNIT,
@DETAIL_CONSTRUCTION_UNIT_PROJECT_MANAGE,
@HANDLER,
@HANDL_DATE,
@SUPVISTION_DEPART_ADVANCE,
@SUPVISOR,
@GENERAL_SUPVISOR,
@SUPVISTION_DATE,
@CONSTRUCTION_UNIT,
@REP,
@CONSTRUCTION_UNIT_PROJECT_MANAGE,
@CONSTRUCTION_UNIT_DATE,
@MAKERID,
@DATE,
@YEAR,
@MONTH,
@DAY
)

";
        string sql2 = @"
UPDATE PROJECT_VISA SET 
PVID=@PVID,
PROJECT_NAME=@PROJECT_NAME,
PROJECT_DART=@PROJECT_DART,
VISA_BECAUSE_AND_CONTENT=@VISA_BECAUSE_AND_CONTENT,
DETAIL_CONSTRUCTION_UNIT=@DETAIL_CONSTRUCTION_UNIT,
DETAIL_CONSTRUCTION_UNIT_PROJECT_MANAGE=@DETAIL_CONSTRUCTION_UNIT_PROJECT_MANAGE,
HANDLER=@HANDLER,
HANDL_DATE=@HANDL_DATE,
SUPVISTION_DEPART_ADVANCE=@SUPVISTION_DEPART_ADVANCE,
SUPVISOR=@SUPVISOR,
GENERAL_SUPVISOR=@GENERAL_SUPVISOR,
SUPVISTION_DATE=@SUPVISTION_DATE,
CONSTRUCTION_UNIT=@CONSTRUCTION_UNIT,
REP=@REP,
CONSTRUCTION_UNIT_PROJECT_MANAGE=@CONSTRUCTION_UNIT_PROJECT_MANAGE,
CONSTRUCTION_UNIT_DATE=@CONSTRUCTION_UNIT_DATE,
MAKERID=@MAKERID,
DATE=@DATE,
YEAR=@YEAR,
MONTH=@MONTH,
DAY=@DAY

";
    basec bc = new basec();
        DataTable dt = new DataTable();
        DataTable dto = new DataTable();
        public CPROJECT_VISA()
        {
            IFExecution_SUCCESS = true;
            getsql = sql;
          

        }
        public string GETID()
        {
            string v1 = bc.numYM(10, 4, "0001", "SELECT * FROM PROJECT_VISA", "PVID", "PV");
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
            sqlcom.Parameters.Add("@PVID", SqlDbType.VarChar).Value = PVID;
            sqlcom.Parameters.Add("@PROJECT_NAME", SqlDbType.VarChar).Value = PROJECT_NAME;
            sqlcom.Parameters.Add("@PROJECT_DART", SqlDbType.VarChar).Value = PROJECT_DART;
            sqlcom.Parameters.Add("@VISA_BECAUSE_AND_CONTENT", SqlDbType.VarChar).Value = VISA_BECAUSE_AND_CONTENT;
            sqlcom.Parameters.Add("@DETAIL_CONSTRUCTION_UNIT", SqlDbType.VarChar).Value = DETAIL_CONSTRUCTION_UNIT;
            sqlcom.Parameters.Add("@DETAIL_CONSTRUCTION_UNIT_PROJECT_MANAGE", SqlDbType.VarChar).Value = DETAIL_CONSTRUCTION_UNIT_PROJECT_MANAGE;
            sqlcom.Parameters.Add("@HANDLER", SqlDbType.VarChar).Value = HANDLER;
            sqlcom.Parameters.Add("@HANDL_DATE", SqlDbType.VarChar).Value = HANDL_DATE;
            sqlcom.Parameters.Add("@SUPVISTION_DEPART_ADVANCE", SqlDbType.VarChar).Value = SUPVISTION_DEPART_ADVANCE;
            sqlcom.Parameters.Add("@SUPVISOR", SqlDbType.VarChar).Value = SUPVISOR;
            sqlcom.Parameters.Add("@GENERAL_SUPVISOR", SqlDbType.VarChar).Value = GENERAL_SUPVISOR;
            sqlcom.Parameters.Add("@SUPVISTION_DATE", SqlDbType.VarChar).Value = SUPVISTION_DATE;
            sqlcom.Parameters.Add("@CONSTRUCTION_UNIT", SqlDbType.VarChar).Value = CONSTRUCTION_UNIT;
            sqlcom.Parameters.Add("@REP", SqlDbType.VarChar).Value = REP;
            sqlcom.Parameters.Add("@CONSTRUCTION_UNIT_PROJECT_MANAGE", SqlDbType.VarChar).Value = CONSTRUCTION_UNIT_PROJECT_MANAGE;
            sqlcom.Parameters.Add("@CONSTRUCTION_UNIT_DATE", SqlDbType.VarChar).Value = CONSTRUCTION_UNIT_DATE;
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
