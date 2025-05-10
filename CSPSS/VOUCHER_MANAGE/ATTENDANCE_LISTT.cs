using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using XizheC;

namespace CSPSS.VOUCHER_MANAGE
{
    public partial class ATTENDANCE_LISTT : Form
    {
        DataTable dt = new DataTable();
        basec bc=new basec ();
        private string _IDO;
        public string IDO
        {
            set { _IDO = value; }
            get { return _IDO; }

        }
     
        private string _ADD_OR_UPDATE;
        public string ADD_OR_UPDATE
        {
            set { _ADD_OR_UPDATE = value; }
            get { return _ADD_OR_UPDATE; }
        }
        private bool _IFExecutionSUCCESS;
        public bool IFExecution_SUCCESS
        {
            set { _IFExecutionSUCCESS = value; }
            get { return _IFExecutionSUCCESS; }

        }
        private static bool _IF_DOUBLE_CLICK;
        public static bool IF_DOUBLE_CLICK
        {
            set { _IF_DOUBLE_CLICK = value; }
            get { return _IF_DOUBLE_CLICK; }

        }
        private static string _WAREID;
        public static string WAREID
        {
            set { _WAREID = value; }
            get { return _WAREID; }

        }
        private static string _CO_WAREID;
        public static string CO_WAREID
        {
            set { _CO_WAREID = value; }
            get { return _CO_WAREID; }

        }
        private static string _WNAME;
        public static string WNAME
        {
            set { _WNAME = value; }
            get { return _WNAME; }

        }
        private static string _STID;
        public static string STID
        {
            set { _STID = value; }
            get { return _STID; }

        }
        private static string _STEP_ID;
        public static string STEP_ID
        {
            set { _STEP_ID = value; }
            get { return _STEP_ID; }

        }
        private static string _STEP;
        public static string STEP
        {
            set { _STEP = value; }
            get { return _STEP; }

        }
        private  delegate bool dele(string a1,string a2);
        private delegate void delex();
        ATTENDANCE_LIST F1 = new ATTENDANCE_LIST();
        protected int M_int_judge, i;
        protected int select;
        CATTENDANCE_LIST cATTENDANCE_LIST = new CATTENDANCE_LIST();
       
        public ATTENDANCE_LISTT()
        {
            InitializeComponent();
        }
        public ATTENDANCE_LISTT(ATTENDANCE_LIST FRM)
        {
            InitializeComponent();
            F1 = FRM;

        }
        private void ATTENDANCE_LISTT_Load(object sender, EventArgs e)
        {
           
            textBox1.Text = IDO;
            bind();
        }
        private void dgvClientInfo_DoubleClick(object sender, EventArgs e)
        {
   
        }
        private void dgvClientInfo_CellClick(object sender, DataGridViewCellEventArgs e)
        {
         
     
          
        }

        public void ClearText()
        {
            comboBox1.Text = "";
            dateTimePicker1.Text = DateTime.Now.ToString("yyyy/MM/dd").Replace("-", "/");
        }
        private void btnSearch_Click(object sender, EventArgs e)
        {
            bind();
            try
            {
           
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        #region bind
        private void bind()
        {
            dataGridView1.EditMode = DataGridViewEditMode.EditOnEnter;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.ContextMenuStrip = contextMenuStrip1;
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "yyyy/MM/dd";
            hint.Location = new Point(400, 100);
            hint.ForeColor = Color.Red;
        
            if (bc.GET_IFExecutionSUCCESS_HINT_INFO(IFExecution_SUCCESS) != "")
            {

                hint.Text = bc.GET_IFExecutionSUCCESS_HINT_INFO(IFExecution_SUCCESS);
            }
            else
            {
                hint.Text = "";
            }
            this.Text = "出勤信息";
            GetTableInfo();
        }
        #endregion
        #region GetTableInfo
        public void GetTableInfo()
        {
            dt = new DataTable();
            dt.Columns.Add("序号", typeof(string ));
            dt.Columns.Add("复选框", typeof(bool));
            dt.Columns.Add("项次", typeof(string));
            dt.Columns.Add("员工工号", typeof(string));
            dt.Columns.Add("员工姓名", typeof(string));
            dt.Columns.Add("职务", typeof(string));
            DataTable dtx = bc.getdt(cATTENDANCE_LIST.sqlf);
            if (dtx.Rows.Count > 0)
            {
                int i = 1;
                foreach (DataRow dr in dtx.Rows)
                {
                    //MessageBox.Show(dr["员工姓名"].ToString());
                    DataRow dr1 = dt.NewRow();
                    dr1["序号"] = i.ToString();
                    dr1["员工工号"] = dr["员工工号"].ToString();
                    dr1["员工姓名"] = dr["员工姓名"].ToString();
                    dr1["职务"] = dr["职务"].ToString();
                    dt.Rows.Add(dr1);
                    i = i + 1;
                }
            }
          
            dtx= bc.getdt(cATTENDANCE_LIST.sql  + " where A.ALID='" +textBox1 .Text + "' ORDER BY  A.ALID ASC ");
            if (dt.Rows.Count > 0)
            {
                if (dtx.Rows.Count > 0)
                {
                    comboBox1.Text = dtx.Rows[0]["时段"].ToString();
                    foreach (DataRow dr in dt.Rows)
                    {

                        foreach (DataRow dr1 in dtx.Rows)
                        {
                            if (dr["员工工号"].ToString() == dr1["员工工号"].ToString())
                            {
                                dr["复选框"] = true;
                                dr["职务"] = dr1["职务"].ToString();
                                dr["项次"] = dr1["项次"].ToString();
                                break;

                            }
                        }
                    }
                }
            }
        
            dataGridView1.DataSource = dt;
            if (dt.Rows.Count > 0)
            {
                dgvStateControl();
            }
        }

        #endregion;
        private void btnAdd_Click(object sender, EventArgs e)
        {
            add();
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            btnSave.Enabled = true;
            M_int_judge = 1;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            btnSave.Focus();
   
            if (juage())
            {
                IFExecution_SUCCESS = false;
            }
            else if (juage_if_exists_select())
            {
                if (bc.exists("SELECT * FROM ATTENDANCE_LIST_DET WHERE ALID='" + textBox1.Text + "' "))
                {
                    basec.getcoms("DELETE ATTENDANCE_LIST_DET WHERE ALID='" + textBox1.Text + "'");
                }
                string year = DateTime.Now.ToString("yy");
                string month = DateTime.Now.ToString("MM");
                string day = DateTime.Now.ToString("dd");
                string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss").Replace("-", "/");
                cATTENDANCE_LIST.ALID = textBox1.Text;
                cATTENDANCE_LIST.TIME = comboBox1.Text;
                cATTENDANCE_LIST.ATTENDANCE_DATE = dateTimePicker1.Text;
                cATTENDANCE_LIST.MAKERID = LOGIN.EMID;
                if (!bc.exists("SELECT ALID FROM ATTENDANCE_LIST_MST WHERE ALID='" +textBox1 .Text + "'"))
                {
                    cATTENDANCE_LIST.SQlcommandE_MST(cATTENDANCE_LIST.sqlt);
                    IFExecution_SUCCESS = true;
                }
                else
                {
                    cATTENDANCE_LIST.SQlcommandE_MST(cATTENDANCE_LIST.sqlth + " WHERE ALID='" +textBox1 .Text  + "'");
                    IFExecution_SUCCESS = true;
                }
                save();
                if (IFExecution_SUCCESS == true && ADD_OR_UPDATE == "ADD")
                {
                    add();
                }

                F1.load();
            }
            else
            {
                hint.Text = "至少要选中一个员工才能保存";



            }
            try
            {
          

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);


            }
        }
        private void add()
        {
            ClearText();
            textBox1.Text = cATTENDANCE_LIST.GETID();
            bind();
            ADD_OR_UPDATE = "ADD";
        }
        private void save()
        {
            btnSave.Focus();
            if (dt.Rows.Count > 0)
            {
                int i = 1;
                foreach (DataRow dr in dt.Rows)
                {
                    if (dr["复选框"].ToString() == "True")
                    {

                        cATTENDANCE_LIST.EMID = dr["员工工号"].ToString();
                        cATTENDANCE_LIST.SN = i.ToString();
                        cATTENDANCE_LIST.POSITION = dr["职务"].ToString();
                        cATTENDANCE_LIST.SQlcommandE_DET(cATTENDANCE_LIST.sqlo);
                        hint.Text = cATTENDANCE_LIST.ErrowInfo;
                        if (IFExecution_SUCCESS)
                        {
                            bind();
                        }
                        i = i + 1;
                    }
                }
            }
           
            try
            {
       
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
        }
        private bool juage()
        {
            
            bool b = false;
            if (textBox1.Text == "")
            {
                hint.Text = "出勤编号不能为空！";
                b = true;
            }

            return b;
       
        }
        private bool juage_if_exists_select()
        {
            bool b = false;
            foreach (DataRow dr in dt.Rows)
            {
                if (dr["复选框"].ToString() == "True")
                {
                    b = true;
                    break;
                }
            }
            return b;

        }
        private void btnDel_Click(object sender, EventArgs e)
        {
            string id = Convert.ToString(dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value).Trim();
         
                basec.getcoms("DELETE ATTENDANCE_LIST_MST WHERE ALID='" + textBox1.Text + "'");
                basec.getcoms("DELETE ATTENDANCE_LIST_DET WHERE ALID='" + textBox1.Text + "'");
                bind();
                ClearText();
                textBox1.Text = "";
                F1.load();
            
            try
            {
             
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        #region override enter
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Enter && ((!(ActiveControl is System.Windows.Forms.TextBox) ||
                !((System.Windows.Forms.TextBox)ActiveControl).AcceptsReturn)))
            {
                SendKeys.SendWait("{Tab}");
                return true;
            }
            if (keyData == (Keys.Enter | Keys.Shift))
            {
                SendKeys.SendWait("+{Tab}");

                return true;
            }
            if (keyData == (Keys.F7))
            {

                //double_info();

                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }
        #endregion
        #region dgvStateControl
        private void dgvStateControl()
        {
            int i;
            dataGridView1.RowHeadersDefaultCellStyle.BackColor = Color.Lavender;
       
            int numCols1 = dataGridView1.Columns.Count;
            //dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;/*自动调整DATAGRIDVIEW的列宽*/
         
            for (i = 0; i < numCols1; i++)
            {

                dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                //this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
                dataGridView1.EnableHeadersVisualStyles = false;
                dataGridView1.Columns[i].HeaderCell.Style.BackColor = Color.Lavender;
                if (dataGridView1.Columns[i].DataPropertyName.ToString() =="复选框")
                {
                    dataGridView1.Columns[i].ReadOnly = false;
                }
                else
                {
                    dataGridView1.Columns[i].ReadOnly = true;

                }
                dataGridView1.Columns[i].DefaultCellStyle .Alignment =DataGridViewContentAlignment .MiddleCenter;
            }
   
            for (i = 0; i < dataGridView1.Columns.Count; i++)
            {
                dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.OldLace;
                i = i + 1;
            }
            dataGridView1.Columns["序号"].Width = 40;
            dataGridView1.Columns["复选框"].Width = 50;
            dataGridView1.Columns["项次"].Width = 40;
            dataGridView1.Columns["员工工号"].Width = 70;
            dataGridView1.Columns["员工姓名"].Width = 70;
            dataGridView1.Columns["职务"].Width = 70;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
        }
        #endregion


        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

   

        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
   
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void contextMenuStrip1_Click(object sender, EventArgs e)
        {
       
        }

        private void 删除此项ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string v1 = dt.Rows[dataGridView1.CurrentCell.RowIndex][0].ToString();
            string sql2 = "DELETE FROM ATTENDANCE_LIST_DET WHERE ALID='" + textBox1.Text + "' AND SN='" + v1 + "'";
            if (dt.Rows.Count > 0)
            {

                if (MessageBox.Show("确定要删除该条信息吗？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                    if (!bc.exists("SELECT * FROM ATTENDANCE_LIST_DET WHERE ALID='" + textBox1.Text + "' AND SN='"+v1+"'"))
                    {
                        hint.Text = "此条记录还未写入数据库";
                    }
                    else  if (bc.juageOne("SELECT * FROM ATTENDANCE_LIST_DET WHERE ALID='" + textBox1.Text + "'"))
                    {

                        basec.getcoms(sql2);
                        string sql3 = "DELETE ATTENDANCE_LIST_MST WHERE ALID='" + textBox1.Text + "'";
                        basec.getcoms(sql3);
                        basec.getcoms("DELETE REMARK WHERE ALID='" + textBox1.Text + "'");
                        IFExecution_SUCCESS = false;
                        bind();
                    }
                    else
                    {

                        basec.getcoms(sql2);
                      
                        IFExecution_SUCCESS = false;
                        bind();
                    }
                }
             
             
            }
            try
            {


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
        }
    }
}
