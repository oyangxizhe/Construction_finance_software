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
    public partial class ATTENDANCE_LIST : Form
    {
        DataTable dt = new DataTable();
        basec bc=new basec ();
        private string _IDO;
        public string IDO
        {
            set { _IDO = value; }
            get { return _IDO; }

        }
        private static string _EMID;
        public static string EMID
        {
            set { _EMID = value; }
            get { return _EMID; }

        }
        private static string _ENAME;
        public static string ENAME
        {
            set { _ENAME = value; }
            get { return _ENAME; }

        }
        private int _GET_DATA_INT;
        public int GET_DATA_INT
        {
            set { _GET_DATA_INT = value; }
            get { return _GET_DATA_INT; }

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
        protected int M_int_judge, i;
        protected int select;
        CATTENDANCE_LIST cATTENDANCE_LIST = new CATTENDANCE_LIST();
        public ATTENDANCE_LIST()
        {
            InitializeComponent();
        }
        private void ATTENDANCE_LIST_Load(object sender, EventArgs e)
        {
            DataTable dtx = bc.getdt("SELECT * FROM EMPLOYEEINFO ");
            AutoCompleteStringCollection inputInfoSource = new AutoCompleteStringCollection();
            comboBox1.Items.Clear();
            foreach (DataRow dr in dtx.Rows)
            {
                comboBox1.Items.Add(dr["ENAME"].ToString());
                inputInfoSource.Add(dr["ENAME"].ToString());
            }
            dtx = bc.getdt("SELECT * FROM EMPLOYEEINFO");
            AutoCompleteStringCollection inputInfoSource1 = new AutoCompleteStringCollection();
            comboBox2.Items.Clear();
            foreach (DataRow dr in dtx.Rows)
            {
                comboBox2.Items.Add(dr["POSITION"].ToString());
                inputInfoSource1.Add(dr["POSITION"].ToString());
            }
            this.comboBox2.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.comboBox2.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource;
            this.comboBox2.AutoCompleteCustomSource = inputInfoSource1;
           //bind();
        }

        #region bind
        private void bind()
        {
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "yyyy/MM/dd";
            dateTimePicker2.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.CustomFormat = "yyyy/MM/dd";
            dateTimePicker3.Format = DateTimePickerFormat.Custom;
            dateTimePicker3.CustomFormat = "yyyy/MM/dd";
            if (bc.getOnlyString("SELECT UNAME FROM USERINFO WHERE USID='" + LOGIN.USID + "'") == "admin")
            {
                btnToExcel.Visible = true;
            }
            else
            {
                btnToExcel.Visible = false;
            }
            StringBuilder stb = new StringBuilder();
            stb.Append(cATTENDANCE_LIST.sql);
            stb.Append(" WHERE A.ALID LIKE '%" + textBox1.Text + "%' ");
            stb.AppendFormat(" AND B.ENAME LIKE '%{0}%'", comboBox1.Text);
            stb.AppendFormat(" AND A.POSITION LIKE '%{0}%'", comboBox2.Text);
            stb.Append(" AND D.TIME LIKE '%" + comboBox3.Text + "%'");
            dataGridView1.AllowUserToAddRows = false;
            //dataGridView1.ContextMenuStrip = contextMenuStrip1;
            if (checkBox1 .Checked)
            {

                stb.AppendFormat(" AND D.DATE BETWEEN '{0}' AND '{1}'",dateTimePicker1.Text+" 0:00:00" ,dateTimePicker2 .Text+" 23:59:59" );
            }
            if (checkBox2.Checked)
            {

                stb.AppendFormat(" AND D.ATTENDANCE_DATE= '{0}'", dateTimePicker3.Text);
            }
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
            search_o(stb.ToString());
  
            this.Text = "出勤信息";
        }
        #endregion
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
        #region search_o()
        public void search_o(string sql)
        {
            string sqlo = " ORDER BY A.ALID ASC";
            //string v7 = bc.getOnlyString("SELECT SCOPE FROM SCOPE_OF_AUTHORIZATION WHERE USID='" + LOGIN.USID + "'");
            string v7 = "Y";
            if (v7 == "Y")
            {

                dt = bc.getdt(sql + sqlo);

            }
            else if (v7 == "GROUP")
            {

                dt = bc.getdt(sql + @" AND D.MAKERID IN (SELECT EMID FROM USERINFO A WHERE USER_GROUP IN 
 (SELECT USER_GROUP FROM USERINFO WHERE USID='" + LOGIN.USID + "'))" + sqlo);
            }
            else
            {
                dt = bc.getdt(sql + " AND D.MAKERID='" + LOGIN.EMID + "'" + sqlo);

            }
            if (v7 == "Y")
            {
               // btnToExcel.Visible = true;
            
            }
            else
            {
                //btnToExcel.Visible = false;
              
            }
            dt = cATTENDANCE_LIST.GetTableInfo(dt);
            if (dt.Rows.Count > 0)
            {
                dataGridView1.DataSource = dt;
                dgvStateControl();
            }
            else
            {
                hint.Text = "找不到所要搜索项！";
                dataGridView1.DataSource = null;

            }
        }
        #endregion
        private void btnAdd_Click(object sender, EventArgs e)
        {
          
            ATTENDANCE_LISTT FRM = new ATTENDANCE_LISTT(this);
            FRM.IDO = cATTENDANCE_LIST.GETID();
            FRM.ADD_OR_UPDATE = "ADD";
            FRM.Show();
          
        }
        public void load()
        {
            bind();
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

            dataGridView1.Columns["序号"].Width = 40;
            dataGridView1.Columns["出勤编号"].Width = 70;
            dataGridView1.Columns["制单人"].Width = 70;
            dataGridView1.Columns["制单日期"].Width = 120;
            dataGridView1.Columns["项次"].Width = 40;
            dataGridView1.Columns["出勤日期"].Width = 70;
            dataGridView1.Columns["时段"].Width = 70;
            dataGridView1.Columns["员工工号"].Width = 60;
            dataGridView1.Columns["员工姓名"].Width = 70;
            dataGridView1.Columns["职务"].Width = 60;
            dataGridView1.Columns["出勤人数"].Width = 60;
            for (i = 0; i < numCols1; i++)
            {

                dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                //this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
                dataGridView1.EnableHeadersVisualStyles = false;
                dataGridView1.Columns[i].HeaderCell.Style.BackColor = Color.Lavender;
                dataGridView1.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
   
            for (i = 0; i < dataGridView1.Columns.Count; i++)
            {
                dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.OldLace;
                dataGridView1.Columns[i].ReadOnly = true;
                i = i + 1;
               
            }
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
        }
        #endregion

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
   

            if (select != 0)
            {
               
                    int intCurrentRowNumber = this.dataGridView1.CurrentCell.RowIndex;
                    string s1 = this.dataGridView1.Rows[intCurrentRowNumber].Cells[0].Value.ToString().Trim();
                    string s2 = this.dataGridView1.Rows[intCurrentRowNumber].Cells[1].Value.ToString().Trim();
                    string s3 = this.dataGridView1.Rows[intCurrentRowNumber].Cells[2].Value.ToString().Trim();
                    string s4 = this.dataGridView1.Rows[intCurrentRowNumber].Cells[3].Value.ToString().Trim();
  
                    /*CSPSS.VOUCHER_MANAGE.FrmSellTableT.data4[0] = "doubleclick";
                    CSPSS.VOUCHER_MANAGE.FrmSellTableT.data1[0] = s1;
                    CSPSS.VOUCHER_MANAGE.FrmSellTableT.data1[1] = s2;
                    CSPSS.VOUCHER_MANAGE.FrmSellTableT.data1[2] = s3;
                    CSPSS.VOUCHER_MANAGE.FrmSellTableT.data1[3] = s4;
                    CSPSS.VOUCHER_MANAGE.FrmSellTableT.data1[4] = s5;*/
                    if (select == 1)
                    {
                        
                    }
                    this.Close();
                

            }
            else
            {
                ATTENDANCE_LISTT FRM = new ATTENDANCE_LISTT(this);
                FRM.IDO = dt.Rows[dataGridView1.CurrentCell.RowIndex]["出勤编号"].ToString();
                FRM.ADD_OR_UPDATE = "UPDATE";
                FRM.Show();
            }
        }

        private void btnToExcel_Click(object sender, EventArgs e)
        {
            if (dt.Rows.Count > 0)
            {

                bc.dgvtoExcel(dataGridView1, "出勤信息");

            }
            else
            {
                MessageBox.Show("没有数据可导出！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        public void WORKORDER_USE()
        {
            select = 1;

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

   
    }
}
