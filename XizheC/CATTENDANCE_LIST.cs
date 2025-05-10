using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;
using System.Data.SqlClient;
using XizheC;
using System.Windows.Forms;

namespace XizheC
{
    public class CATTENDANCE_LIST
    {
        basec bc = new basec();
        #region nature
        private string _EMID;
        public string EMID
        {
            set { _EMID = value; }
            get { return _EMID; }

        }
     
        private string _ALID;
        public string ALID
        {
            set { _ALID = value; }
            get { return _ALID; }

        }
      
        private string _sql;
        public string sql
        {
            set { _sql = value; }
            get { return _sql; }

        }
        private string _sqlo;
        public string sqlo
        {
            set { _sqlo = value; }
            get { return _sqlo; }

        }
        private string _sqlt;
        public string sqlt
        {
            set { _sqlt = value; }
            get { return _sqlt; }

        }
        private string _sqlth;
        public string sqlth
        {
            set { _sqlth = value; }
            get { return _sqlth; }

        }
        private string _sqlf;
        public string sqlf
        {
            set { _sqlf = value; }
            get { return _sqlf; }

        }
        private string _sqlfi;
        public string sqlfi
        {
            set { _sqlfi = value; }
            get { return _sqlfi; }

        }

        private string _sqlsi;
        public string sqlsi
        {
            set { _sqlsi = value; }
            get { return _sqlsi; }

        }
        private string _MAKERID;
        public string MAKERID
        {
            set { _MAKERID = value; }
            get { return _MAKERID; }

        }
        private string _ALKEY;
        public string ALKEY
        {
            set { _ALKEY = value; }
            get { return _ALKEY; }

        }
        private  bool _IFExecutionSUCCESS;
        public  bool IFExecution_SUCCESS
        {
            set { _IFExecutionSUCCESS = value; }
            get { return _IFExecutionSUCCESS; }

        }

        private string _SN;
        public string SN
        {
            set { _SN = value; }
            get { return _SN; }

        }
        private string _ErrowInfo;
        public string ErrowInfo
        {

            set { _ErrowInfo = value; }
            get { return _ErrowInfo; }

        }
        private string _REMARK;
        public string REMARK
        {
            set { _REMARK = value; }
            get { return _REMARK; }

        }
        private string _POSITION;
        public string POSITION
        {
            set { _POSITION = value; }
            get { return _POSITION; }

        }
        private string _TIME;
        public string TIME
        {

            set { _TIME = value; }
            get { return _TIME; }

        }
        private string _ATTENDANCE_DATE;
        public string ATTENDANCE_DATE
        {
            set { _ATTENDANCE_DATE = value; }
            get { return _ATTENDANCE_DATE; }

        }
        #endregion
        DataTable dt = new DataTable();
        #region sql
        string setsql = @"
SELECT 
A.ALID AS 出勤编号,
A.SN AS 项次,
D.ATTENDANCE_DATE AS 出勤日期,
D.TIME AS 时段,
A.EMID AS 员工工号,
B.EName AS 员工姓名,
A.POSITION AS 职务,
C.ENAME AS  制单人,
D.DATE AS 制单日期
FROM
ATTENDANCE_LIST_DET A 
LEFT JOIN EmployeeInfo B ON A.EMID=B.EMID
LEFT JOIN ATTENDANCE_LIST_MST D ON A.ALID=D.ALID
LEFT JOIN EMPLOYEEINFO C ON D.MAKERID=C.EMID

";


        string setsqlo = @"
INSERT INTO ATTENDANCE_LIST_DET
(
ALKEY,
ALID,
SN,
EMID,
POSITION,
YEAR,
MONTH,
DAY
)
VALUES
(
@ALKEY,
@ALID,
@SN,
@EMID,
@POSITION,
@YEAR,
@MONTH,
@DAY

)


";

        string setsqlt = @"

INSERT INTO ATTENDANCE_LIST_MST
(
ALID,
TIME,
ATTENDANCE_DATE,
DATE,
MAKERID,
YEAR,
MONTH

)
VALUES
(
@ALID,
@TIME,
@ATTENDANCE_DATE,
@DATE,
@MAKERID,
@YEAR,
@MONTH

)
";
        string setsqlth = @"
UPDATE ATTENDANCE_LIST_MST SET 
TIME=@TIME,
ATTENDANCE_DATE=@ATTENDANCE_DATE,
DATE=@DATE,
YEAR=@YEAR,
MONTH=@MONTH

";

        string setsqlf = @"
SELECT 
cast(0   as   bit)   as   复选框,
A.EMID AS 员工工号,
A.ENAME AS 员工姓名,
A.POSITION AS 职务
FROM EmployeeInfo A
";
        string setsqlfi = @"

";
        string setsqlsi = @"


";
        #endregion
        public CATTENDANCE_LIST()
        {
            string year, month, day;
            year = DateTime.Now.ToString("yy");
            month = DateTime.Now.ToString("MM");
            day = DateTime.Now.ToString("dd");
            //GETID =bc.numYM(10, 4, "0001", "SELECT * FROM WORKORDER_PICKING_MST", "WPID", "WP");

            sql = setsql;
            sqlo = setsqlo;
            sqlt = setsqlt;
            sqlth = setsqlth;
            sqlf = setsqlf;
            sqlfi = setsqlfi;
            sqlsi = setsqlsi;
        }
        #region GetTableInfo
        public DataTable GetTableInfo(DataTable dtx1)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("序号", typeof(string));
            dt.Columns.Add("出勤编号", typeof(string));
            dt.Columns.Add("项次", typeof(string));
            dt.Columns.Add("出勤日期", typeof(string));
            dt.Columns.Add("时段", typeof(string));
            dt.Columns.Add("员工工号", typeof(string));
            dt.Columns.Add("员工姓名", typeof(string));
            dt.Columns.Add("职务", typeof(string));
            dt.Columns.Add("出勤人数", typeof(string));
            dt.Columns.Add("制单人", typeof(string));
            dt.Columns.Add("制单日期", typeof(string));
            if(dtx1.Rows .Count >0)
            {
                int i=1;
                foreach (DataRow dr in dtx1.Rows )
                {
                    
                    DataRow dr1 = dt.NewRow();
                    dr1["序号"] = i.ToString();
                    dr1["出勤编号"] = dr["出勤编号"].ToString();
                    dr1["项次"] = dr["项次"].ToString();
                    dr1["出勤日期"] = dr["出勤日期"].ToString();
                    dr1["时段"] = dr["时段"].ToString();
                    dr1["员工工号"] = dr["员工工号"].ToString();
                    dr1["员工姓名"] = dr["员工姓名"].ToString();
                    dr1["职务"] = dr["职务"].ToString();
                    dr1["制单人"] = dr["制单人"].ToString();
                    dr1["制单日期"] = dr["制单日期"].ToString();
                    dt.Rows.Add(dr1);
                    i = i + 1;
                }
            }

            foreach (DataRow dr in dt.Rows)
            {

                DataTable dtx2 = bc.GET_DT_TO_DV_TO_DT(dtx1, "", string.Format("出勤日期='{0}'", dr["出勤日期"].ToString()));
                if (dtx2.Rows.Count > 0)
                {
                    dr["出勤人数"] = dtx2.Rows.Count.ToString();
                }
            }
            return dt;
        }
 
        #endregion
        public string GETID()
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
            string v1 = bc.numYM(10, 4, "0001", "select * from ATTENDANCE_LIST_MST", "ALID", "AL");
            string GETID = "";
            if (v1 != "Exceed Limited")
            {
                GETID = v1;
              
            }
            return GETID;
        }
        #region SQlcommandE_DET
        public  void SQlcommandE_DET(string sql)
        {
            ALKEY = bc.numYMD(20, 12, "000000000001", "SELECT * FROM ATTENDANCE_LIST_DET", "ALKEY", "AL");
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss").Replace ("-","/");
            SqlConnection sqlcon = bc.getcon();
            sqlcon.Open();
            SqlCommand sqlcom = new SqlCommand(sql, sqlcon);
            sqlcom.Parameters.Add("@ALKEY", SqlDbType.VarChar, 20).Value = ALKEY;
            sqlcom.Parameters.Add("@ALID", SqlDbType.VarChar, 20).Value = ALID;
            sqlcom.Parameters.Add("@SN", SqlDbType.VarChar, 20).Value = SN;
            sqlcom.Parameters.Add("@EMID", SqlDbType.VarChar, 20).Value = EMID;
            sqlcom.Parameters.Add("@POSITION", SqlDbType.VarChar, 20).Value = POSITION;
            sqlcom.Parameters.Add("@YEAR", SqlDbType.VarChar, 20).Value = year;
            sqlcom.Parameters.Add("@MONTH", SqlDbType.VarChar, 20).Value = month;
            sqlcom.Parameters.Add("@DAY", SqlDbType.VarChar, 20).Value = day;
            sqlcom.ExecuteNonQuery();
            sqlcon.Close();
        }
        #endregion
        #region SQlcommandE_MST
        public  void SQlcommandE_MST(string sql)
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss").Replace("-", "/");
            SqlConnection sqlcon = bc.getcon();
            SqlCommand sqlcom = new SqlCommand(sql, sqlcon);
            sqlcon.Open();
            sqlcom.Parameters.Add("@ALID", SqlDbType.VarChar, 20).Value = ALID;
            sqlcom.Parameters.Add("@TIME", SqlDbType.VarChar, 20).Value = TIME;
            sqlcom.Parameters.Add("@ATTENDANCE_DATE", SqlDbType.VarChar, 20).Value = ATTENDANCE_DATE;
            sqlcom.Parameters.Add("@DATE", SqlDbType.VarChar, 20).Value = varDate;
            sqlcom.Parameters.Add("@MAKERID", SqlDbType.VarChar, 20).Value = MAKERID;
            sqlcom.Parameters.Add("@YEAR", SqlDbType.VarChar, 20).Value = year;
            sqlcom.Parameters.Add("@MONTH", SqlDbType.VarChar, 20).Value = month;
            sqlcom.ExecuteNonQuery();
            sqlcon.Close();
        }
        #endregion
    
    
    }
}
