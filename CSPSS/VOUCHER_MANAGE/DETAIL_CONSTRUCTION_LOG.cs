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
    public partial class DETAIL_CONSTRUCTION_LOG : Form
    {
        DataTable dt = new DataTable();
        DataTable dt1 = new DataTable();
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
        basec bc = new basec();
        CDETAIL_CONSTRUCTION_LOG cDETAIL_CONSTRUCTION_LOG = new CDETAIL_CONSTRUCTION_LOG();
        protected int M_int_judge, i;
        protected int select;
        public DETAIL_CONSTRUCTION_LOG()
        {
            InitializeComponent();
        }
        
        private void DETAIL_CONSTRUCTION_LOG_Load(object sender, EventArgs e)
        {

            Bind();

        }
        private void Bind()
        {
            dataGridView1.AllowUserToAddRows = false;
            textBox1.Text = IDO;
            dt = basec.getdts(cDETAIL_CONSTRUCTION_LOG .getsql );
            dataGridView1.DataSource = dt;
            textBox2.Focus();
            dgvStateControl();
            hint.Location = new Point(400,100);
            hint.ForeColor = Color.Red;
            if (bc.GET_IFExecutionSUCCESS_HINT_INFO(IFExecution_SUCCESS) != "")
            {
                hint.Text = bc.GET_IFExecutionSUCCESS_HINT_INFO(IFExecution_SUCCESS);
            }
            else
            {
                hint.Text = "";
            }
            comboBox1.DataSource = bc.RETURN_ADD_EMPTY_COLUMN("EMPLOYEEINFO", "ENAME");
            comboBox1.DisplayMember = "ENAME";
            comboBox2.DataSource = bc.RETURN_ADD_EMPTY_COLUMN("EMPLOYEEINFO", "ENAME");
            comboBox2.DisplayMember = "ENAME";

            this.WindowState= FormWindowState.Maximized;
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
          
            dateTimePicker1.CustomFormat = "yyyy/MM/dd";

            groupBox3.Text = "施工质量、安全、进度、机械、劳动力情况：";
            groupBox4.Text = "材料进场、使用、检验；试块留置、送检：";
            groupBox5.Text = "存在质量、安全问题及处理意见：";
        }
        #region dgvStateControl
        private void dgvStateControl()
        {
            int i;
            dataGridView1.RowHeadersDefaultCellStyle.BackColor = Color.Lavender;
            int numCols1 = dataGridView1.Columns.Count;
            for (i = 0; i < numCols1; i++)
            {

                dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                if (i == 1)
                {
                    dataGridView1.Columns[i].Width = 70;

                }
                else if (i == 6)
                {
                    dataGridView1.Columns[i].Width = 120;

                }
                else if (i == 4)
                {
                    dataGridView1.Columns[i].Width = 90;

                }
                else
                {
                    dataGridView1.Columns[i].Width = 60;

                }
            
                dataGridView1.EnableHeadersVisualStyles = false;
                dataGridView1.Columns[i].HeaderCell.Style.BackColor = Color.Lavender;

            }
            for (i = 0; i < dataGridView1.Columns.Count; i++)
            {
                dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.OldLace;
                i = i + 1;
            }
            for (i = 0; i < dataGridView1.Columns.Count; i++)
            {
                dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridView1.Columns[i].ReadOnly = true;

            }
            dataGridView1.Columns["日志单号"].Width = 80;
            dataGridView1.Columns["工程名称"].Width = 80;
            dataGridView1.Columns["施工日期"].Width = 80;
            dataGridView1.Columns["工程部位"].Width = 80;
            dataGridView1.Columns["温度"].Width = 80;
            dataGridView1.Columns["风级"].Width = 80;
            dataGridView1.Columns["施工质量"].Width = 150;
            dataGridView1.Columns["检验情况"].Width = 150;
            dataGridView1.Columns["安全问题"].Width = 150;
            dataGridView1.Columns["项目负责人"].Width = 80;
            dataGridView1.Columns["经手人"].Width = 80;
            dataGridView1.Columns["制单人"].Width = 70;
            dataGridView1.Columns["制单日期"].Width = 120;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
        }
        #endregion
    
        #region save
        private void save()
        {
            string year = DateTime.Now.ToString("yy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string varDate = DateTime.Now.ToString("yyy/MM/dd HH:mm:ss");
            cDETAIL_CONSTRUCTION_LOG.CLID = textBox1.Text;
            cDETAIL_CONSTRUCTION_LOG.PROJECT_NAME = textBox2.Text;
            cDETAIL_CONSTRUCTION_LOG.PROJECT_DATE = dateTimePicker1.Text;
            cDETAIL_CONSTRUCTION_LOG.PROJECT_PART = textBox3.Text;
            cDETAIL_CONSTRUCTION_LOG.TEMPERATURE = textBox4.Text;
            cDETAIL_CONSTRUCTION_LOG.WIND_SCALE = textBox5.Text;
            cDETAIL_CONSTRUCTION_LOG.CONTENT_O = textBox6.Text;
            cDETAIL_CONSTRUCTION_LOG.CONTENT_T = textBox7.Text;
            cDETAIL_CONSTRUCTION_LOG.CONTENT_TH = textBox8.Text;
            cDETAIL_CONSTRUCTION_LOG.PROJECT_MANAGE = comboBox1.Text;
            cDETAIL_CONSTRUCTION_LOG.HANDLER = comboBox2.Text;
            cDETAIL_CONSTRUCTION_LOG.MAKERID = LOGIN.EMID;
            cDETAIL_CONSTRUCTION_LOG.save("DETAIL_CONSTRUCTION_LOG", "CLID", "PROJECT_NAME", textBox1.Text, textBox2.Text, "单号", "");
            if (cDETAIL_CONSTRUCTION_LOG.IFExecution_SUCCESS)
            {
                IFExecution_SUCCESS = cDETAIL_CONSTRUCTION_LOG.IFExecution_SUCCESS;
                Bind();
            }
            else
            {
                hint.Text = cDETAIL_CONSTRUCTION_LOG.ErrowInfo;
            }
        }
        #endregion
        #region juage()
        private bool juage()
        {
            bool b = false;
            if (textBox1.Text == "")
            {
                b = true;

                hint.Text = "单号不能为空！";
             
            }
            return b;
        }
        #endregion
        public void ClearText()
        {
            textBox2.Text = "";
            textBox3.Text = "";
            textBox6.Text = "";
     
            textBox7.Text = "";
            textBox8.Text = "";
            comboBox1.Text = "";
            comboBox2.Text = "";
    
            string v1 = DateTime.Now.ToString("yyyy/MM/dd").Replace("-", "/");
            dateTimePicker1.Value = Convert.ToDateTime(v1);
     
        
        }


        private void btnAdd_Click(object sender, EventArgs e)
        {

            add();
        }
        private void add()
        {

            textBox1.Text = cDETAIL_CONSTRUCTION_LOG.GETID();
            ClearText();
            textBox2.Focus();

        }
      

        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                if (juage())
                {

                }
                else
                {
                    save();
                    if (IFExecution_SUCCESS)
                    {
                        add();
                    }

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            }
         
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            try
            {


                dt = bc.getdt(cDETAIL_CONSTRUCTION_LOG .getsql +" WHERE A.CLID LIKE '%"+textBox20.Text +"%' AND A.PROJECT_NAME LIKE '%"+textBox21 .Text +"%'");
                if (dt.Rows.Count > 0)
                {
                    dataGridView1.DataSource = dt;
                    dgvStateControl();
                }
                else
                {

                    hint.Text = "没有找到相关信息！";
                    dataGridView1.DataSource = null;
                }
            
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            string id = Convert.ToString(dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value).Trim();
                IFExecution_SUCCESS = false;
                string strSql = "DELETE FROM DETAIL_CONSTRUCTION_LOG WHERE CLID='" + id + "'";
                basec.getcoms(strSql);
                Bind();
                ClearText();
            
            try
            {
            
            }
            catch (Exception)
            {


            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        #region override enter
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Enter && ((!(ActiveControl is System.Windows.Forms.TextBox) ||
                !((System.Windows.Forms.TextBox)ActiveControl).AcceptsReturn)))
            {

                if (dataGridView1.CurrentCell.ColumnIndex == 7 &&
                    dataGridView1["借方原币金额", dataGridView1.CurrentCell.RowIndex].Value.ToString() != null)
                {

                    SendKeys.SendWait("{Tab}");
                    SendKeys.SendWait("{Tab}");
                }
                else if (dataGridView1.CurrentCell.ColumnIndex == 9)
                {
                    SendKeys.SendWait("{Tab}");
                    SendKeys.SendWait("{Tab}");
                    SendKeys.SendWait("{Tab}");
                }
                else
                {

                    SendKeys.SendWait("{Tab}");
                }
                return true;
            }
            if (keyData == (Keys.Enter | Keys.Shift))
            {
                SendKeys.SendWait("+{Tab}");

                return true;
            }
            if (keyData == (Keys.F7))
            {

                dataGridView1.Focus();

                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }
        #endregion

        private void label17_Click(object sender, EventArgs e)
        {

        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView1_Click(object sender, EventArgs e)
        {
            string v1 = Convert.ToString(dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value).Trim();
            int i = dataGridView1.CurrentCell.RowIndex;
            textBox1.Text = dt.Rows[i]["日志单号"].ToString();
            textBox2.Text = dt.Rows[i]["工程名称"].ToString();
            dateTimePicker1 .Text   = dt.Rows[i]["施工日期"].ToString();
            textBox3.Text = dt.Rows[i]["工程部位"].ToString();
            textBox4.Text = dt.Rows[i]["温度"].ToString();
            textBox5.Text = dt.Rows[i]["风级"].ToString();
            textBox6.Text = dt.Rows[i]["施工质量"].ToString();
            textBox7.Text = dt.Rows[i]["检验情况"].ToString();
            textBox8.Text = dt.Rows[i]["安全问题"].ToString();
            comboBox1.Text  = dt.Rows[i]["项目负责人"].ToString();
            comboBox2.Text  = dt.Rows[i]["经手人"].ToString();
        }

        private void label19_Click(object sender, EventArgs e)
        {

        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox6_TextChanged_1(object sender, EventArgs e)
        {

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
    }
}
