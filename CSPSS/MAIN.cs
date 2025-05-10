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
using System.IO;
using System.Threading;



namespace CSPSS
{
    public partial class MAIN : Form
    {
         DataTable dt = new DataTable();
         DataTable dt2 = new DataTable();
         basec bc = new basec();
         CUSER cuser = new CUSER();
         CEMPLOYEE_INFO cemplyee_info = new CEMPLOYEE_INFO();
         Color c2 = System.Drawing.ColorTranslator.FromHtml("#4a7bb8");
         Color c3 = System.Drawing.ColorTranslator.FromHtml("#24ade5");
         CDEPART cdepart = new CDEPART();
         CPOSITION cposition = new CPOSITION();
         CVOUCHER cvoucher = new CVOUCHER();
         CUSER_GROUP cuser_group = new CUSER_GROUP();
         CABSTRACT cabstract = new CABSTRACT();
         CPROJECT_VISA cproject_visa = new CPROJECT_VISA();
         CDETAIL_CONSTRUCTION_LOG cdetail_construction_log = new CDETAIL_CONSTRUCTION_LOG();
         CATTENDANCE_LIST cATTENDANCE_LIST = new CATTENDANCE_LIST();
         StringBuilder sqb = new StringBuilder();
         
         private string _SAMPLE_ID;
         public string SAMPLE_ID
         {
             set { _SAMPLE_ID = value; }
             get { return _SAMPLE_ID; }

         }
         private string _ID;
         public string ID
         {
             set { _ID = value; }
             get { return _ID; }
         }
         private string _PNID;
         public string PNID
         {
             set { _PNID = value; }
             get { return _PNID; }
         }
         private string _PROJECT_NAME;
         public string PROJECT_NAME
         {
             set { _PROJECT_NAME = value; }
             get { return _PROJECT_NAME; }

         }
         CFileInfo cfileinfo = new CFileInfo();

        public MAIN()
        {
            InitializeComponent();
        }
        #region bind1
        private void bind1()
        {
            this.Icon = CSPSS.Resource1.xz_200X200;
            timer1.Enabled = true;
            timer1.Interval = 1000;
            pictureBox1.BackColor = c2;
            notifyIcon1.Icon = CSPSS.Resource1.xz_200X200;
            notifyIcon1.Text = "财务管理系统";
            pictureBox1.Image = CSPSS.Resource1.company;
            sqb = new StringBuilder();
            sqb.AppendFormat("财务管理系统");
            sqb.AppendFormat(" Version 1.0.0 ");
            //string v5 = AppDomain.CurrentDomain.BaseDirectory.ToString() + "Version.xml";
            sqb.AppendFormat("当前版本更新日期：{0}", "20161216_1700");
            this.Text = sqb.ToString();
            dt = bc.getdt("SELECT * from RightList where USID = '" + LOGIN.USID + "'");
            SHOW_TREEVIEW(dt);
            menuStrip1.Font = new Font("宋体", 9);
            this.WindowState = FormWindowState.Maximized;
            toolStripStatusLabel1.Text = "||当前用户：" + LOGIN.UNAME;
            toolStripStatusLabel2.Text = "||所属部门：" + LOGIN.DEPART;
            toolStripStatusLabel3.Text = "||登录时间：" + DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString() + " || 技术支持：苏州好用软件有限公司";
            listView1.BackColor = c2;
            listView1.ForeColor = Color.White;
            listView1.Font = new Font("新宋体", 11);
            listView2.BorderStyle = BorderStyle.None;
            
            imageList1.Images.Add(CSPSS.Resource1._1);
            imageList1.Images.Add(CSPSS.Resource1._2);
            imageList1.Images.Add(CSPSS.Resource1._3);
            imageList1.Images.Add(CSPSS.Resource1._4);
            imageList1.Images.Add(CSPSS.Resource1._5);
            imageList1.Images.Add(CSPSS.Resource1._6);
            imageList1.Images.Add(CSPSS.Resource1._7);
            imageList1.Images.Add(CSPSS.Resource1._8);
            imageList1.Images.Add(CSPSS.Resource1._9);
            imageList1.Images.Add(CSPSS.Resource1._10);
            imageList1.Images.Add(CSPSS.Resource1._11);
            imageList1.Images.Add(CSPSS.Resource1._12);
            imageList1.Images.Add(CSPSS.Resource1._13);
            imageList1.Images.Add(CSPSS.Resource1._14);
            imageList1.Images.Add(CSPSS.Resource1._15);
            imageList1.Images.Add(CSPSS.Resource1._16);
            imageList1.Images.Add(CSPSS.Resource1._17);
            imageList1.Images.Add(CSPSS.Resource1._18);
            imageList1.Images.Add(CSPSS.Resource1._19);
            imageList1.Images.Add(CSPSS.Resource1._20);
            imageList1.Images.Add(CSPSS.Resource1._21);
            imageList1.Images.Add(CSPSS.Resource1._22);
            imageList1.Images.Add(CSPSS.Resource1._23);
            imageList1.Images.Add(CSPSS.Resource1._24);
            imageList1.Images.Add(CSPSS.Resource1._25);
            imageList1.Images.Add(CSPSS.Resource1._26);
            imageList1.Images.Add(CSPSS.Resource1._27);
            imageList1.Images.Add(CSPSS.Resource1._28);
            imageList1.Images.Add(CSPSS.Resource1._29);
            imageList1.Images.Add(CSPSS.Resource1._30);
            imageList1.Images.Add(CSPSS.Resource1._31);
            imageList1.Images.Add(CSPSS.Resource1._32);
            imageList1.Images.Add(CSPSS.Resource1._33);
            imageList1.Images.Add(CSPSS.Resource1._34);
            imageList1.Images.Add(CSPSS.Resource1._35);
            imageList1.Images.Add(CSPSS.Resource1._36);
            imageList1.Images.Add(CSPSS.Resource1._37);
            imageList1.Images.Add(CSPSS.Resource1._38);
            imageList1.Images.Add(CSPSS.Resource1._39);
            imageList1.Images.Add(CSPSS.Resource1._40);
            imageList1.Images.Add(CSPSS.Resource1._41);
            imageList1.Images.Add(CSPSS.Resource1._42);

            imageList1.ColorDepth = ColorDepth.Depth32Bit;/*防止图片失真*/
            listView1.View = View.SmallIcon;
            listView2.View = View.LargeIcon;
            imageList1.ImageSize = new Size(48, 48);/*set imglist size*/
            listView1.SmallImageList = imageList1;
            listView2.LargeImageList = imageList1;
        }
        #endregion

        #region load
        private void MAIN_Load(object sender, EventArgs e)
        {
            try
            {
                bind1();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion
        #region show_treeview
        private void SHOW_TREEVIEW(DataTable dt)
        {
            dt = bc.GET_DT_TO_DV_TO_DT(dt, "NODEID ASC", "PARENT_NODEID=0");
            if (dt.Rows.Count > 0)
            {
                for(int i=0;i<dt.Rows.Count ;i++)
                {
                    ListViewItem lvi = listView1.Items.Add(dt.Rows[i]["NODE_NAME"].ToString());
                    lvi.ImageIndex = Convert.ToInt32(dt.Rows[i]["NODEID"].ToString()) - 1;/*NEED THIS SO CAN SHOW*/
                }
                DataTable  dtx = bc.GET_DT_TO_DV_TO_DT(dt, "", "NODE_NAME='账务管理'");
                if (dtx.Rows.Count > 0)
                {
                    click(dtx.Rows[0]["NODE_NAME"].ToString());
                    if(listView1.Items.Count ==1)
                    {
                        listView1.Items[0].BackColor = c3;
                    }
                    else
                    {
                        
                        listView1.Items[1].BackColor = c3;
                    }
                }
                else
                {
                    click(dt.Rows[0]["NODE_NAME"].ToString());
                    listView1.Items[0].BackColor = c3;
                }
            }
        }
        #endregion

        #region show_treeview_O
        private void SHOW_TREEVIEW_O(string NODEID)
        {
          
            dt2 = bc.getdt("SELECT * FROM RIGHTLIST WHERE PARENT_NODEID='" + NODEID  + "'AND  USID = '" + LOGIN.USID + "' ORDER BY NODEID ASC");
            if (dt2.Rows.Count > 0)
            {
                for (int i = 0; i < dt2.Rows.Count; i++)
                {
                    ListViewItem lvi = listView2.Items.Add(dt2.Rows[i]["NODE_NAME"].ToString());
                    lvi.ImageIndex = Convert.ToInt32(dt2.Rows[i]["NODEID"].ToString()) - 1;/*NEED THIS SO CAN SHOW*/
                }
            }
            else
            {
                MessageBox.Show(NODEID +","+LOGIN.USID );

            }
        }
        #endregion

         private void 退出系统ToolStripMenuItem1_Click(object sender, EventArgs e)
         {
             if (MessageBox.Show("确定要退出本系统吗？", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.OK)
             {
                 EXIT();
             }
             else
             {
                 
             }
         }
         private void listView1_Click(object sender, EventArgs e)
         {
             try
             {
                 string v1 = listView1.SelectedItems[0].SubItems[0].Text.ToString();/*get selectitem value*/
                 click(v1);
             }
             catch (Exception)
             {


             }
            
         }
         private void click(string NODE_NAME)
         {
            
             listView2.Items.Clear();
             string id = bc.getOnlyString("SELECT NODEID FROM RIGHTLIST WHERE NODE_NAME='" + NODE_NAME + "'");
             SHOW_TREEVIEW_O(id);

             foreach (ListViewItem lvi in listView1.Items)
             {
                 if (lvi.Selected)
                 {
                     lvi.BackColor = c3;
                     pictureBox1.Focus();/*SELECTED AFTER MOVE FOCUS*/
                 }
                 else
                 {
                     lvi.BackColor = c2;
                 }

             }

         }
         #region listview2
         private void listView2_Click(object sender, EventArgs e)
         {
             string v1 = listView2.SelectedItems[0].SubItems[0].Text.ToString();/*get selectitem value*/
             #region v1
             if (v1 == "科目信息维护")
             {
                 CSPSS.BASE_INFO.ACCOUNTANT_COURSE FRM = new CSPSS.BASE_INFO.ACCOUNTANT_COURSE();
                 FRM.Show();


             }
             else if (v1 == "员工出勤表")
             {
                 CSPSS.VOUCHER_MANAGE.ATTENDANCE_LIST FRM = new CSPSS.VOUCHER_MANAGE.ATTENDANCE_LIST();
                 FRM.IDO = cATTENDANCE_LIST.GETID();
                 FRM.ShowDialog();

             }
             else if (v1 == "施工日志")
             {
                 CSPSS.VOUCHER_MANAGE.DETAIL_CONSTRUCTION_LOG FRM = new CSPSS.VOUCHER_MANAGE.DETAIL_CONSTRUCTION_LOG();
                 FRM.IDO = cdetail_construction_log.GETID();
                 FRM.ShowDialog();

             }
             else if (v1 == "工程签证单")
             {
                 CSPSS.VOUCHER_MANAGE.PROJECT_VISA FRM = new CSPSS.VOUCHER_MANAGE.PROJECT_VISA();
                 FRM.IDO = cproject_visa.GETID();
                 FRM.ShowDialog();

             }
             else if (v1 == "员工信息维护")
             {
                 CSPSS.BASE_INFO.EMPLOYEE_INFO FRM = new CSPSS.BASE_INFO.EMPLOYEE_INFO();
                 FRM.IDO = cemplyee_info.GETID();
                 FRM.ShowDialog();

             }
             else if (v1 == "摘要信息维护")
             {
                 CSPSS.BASE_INFO.ABSTRACT FRM = new CSPSS.BASE_INFO.ABSTRACT();
                 FRM.IDO = cabstract.GETID();
                 FRM.ShowDialog();

             }
             else if (v1 == "部门信息维护")
             {
                 CSPSS.BASE_INFO.DEPART FRM = new CSPSS.BASE_INFO.DEPART();
                 FRM.IDO = cdepart.GETID();
                 FRM.ShowDialog();

             }
             else if (v1 == "职务信息维护")
             {
                 CSPSS.BASE_INFO.POSITION FRM = new CSPSS.BASE_INFO.POSITION();
                 FRM.IDO = cposition.GETID();
                 FRM.ShowDialog();

             }
             else if (v1 == "服务器IP")
             {
                 CSPSS.BASE_INFO.UPLOADFILE_DOMAIN FRM = new CSPSS.BASE_INFO.UPLOADFILE_DOMAIN();

                 FRM.ShowDialog();

             }
             else if (v1 == "录入凭证作业")
             {
                 CSPSS.VOUCHER_MANAGE.VOUCHERT FRM = new CSPSS.VOUCHER_MANAGE.VOUCHERT();
                 FRM.IDO = cvoucher.GETID();
                 FRM.ShowDialog();

             }
             else if (v1 == "凭证查询作业")
             {
                 CSPSS.VOUCHER_MANAGE.VOUCHER FRM = new CSPSS.VOUCHER_MANAGE.VOUCHER();
                 FRM.ShowDialog();

             }
             else if (v1 == "用户帐户")
             {
                 CSPSS.USER_MANAGE.USER_INFO FRM = new CSPSS.USER_MANAGE.USER_INFO();
                 FRM.IDO = cuser.GETID();
                 FRM.ADD_OR_UPDATE = "ADD";
                 FRM.ShowDialog();

             }
             else if (v1 == "更改密码")
             {
                 CSPSS.USER_MANAGE.EDIT_PWD FRM = new CSPSS.USER_MANAGE.EDIT_PWD();
                 FRM.ShowDialog();

             }
             else if (v1 == "权限管理")
             {
                 CSPSS.USER_MANAGE.EDIT_RIGHT FRM = new CSPSS.USER_MANAGE.EDIT_RIGHT();
                 FRM.ShowDialog();

             }
             else if (v1 == "用户组信息")
             {
                 CSPSS.USER_MANAGE.USER_GROUP FRM = new CSPSS.USER_MANAGE.USER_GROUP();
                 FRM.IDO = cuser_group.GETID();
                 FRM.ShowDialog();

             }
             #endregion
         }
         #endregion
         private void notifyIcon1_Click(object sender, EventArgs e)
         {
             click();//托盘单击事件
  
         }
         private void notifyIcon1_BalloonTipClicked(object sender, EventArgs e)
         {
             click();//气泡单击事件
             showform();
         }
         private void notifyIcon1_BalloonTipClosed(object sender, EventArgs e)
         {
             click();//气泡关闭单击事件
             //MessageBox.Show("ok");
         }
         private void showform()
         {
           
         

         }
         private void click()
         {
            
             //basec.getcoms("UPDATE REMIND SET RECEIVE_STATUS='Y' WHERE RIID='" + ID + "' AND NOTICE_MAKERID='" + LOGIN.EMID + "'");
             timer2.Enabled = false;
      
             this.WindowState = FormWindowState.Maximized;
             ContextMenu c = new ContextMenu();
             MenuItem s = new MenuItem("退出");
             c.MenuItems.Add(s);
             notifyIcon1.ContextMenu = c;
             notifyIcon1.Icon = new Icon(System.IO.Path.GetFullPath("Image/xz 200X200.ico"));
             s.Click += new EventHandler(notify_Click);
             this.Show();

         }
         private void notify_Click(object sender, EventArgs e)
         {
             EXIT();
         }
         private void EXIT()
         {
             this.Dispose();
             notifyIcon1.Dispose();
             Application.Exit();
             string varDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss").Replace("-", "/");
             //bc.getcom(@"UPDATE AUTHORIZATION_USER SET STATUS='N' ,LEAVE_DATE='" + varDate + "'WHERE AUID='" + LOGIN.AUID + "'");
         }
         
        
     
         private void MAIN_FormClosing(object sender, FormClosingEventArgs e)
         {
             e.Cancel = true;
             this.Hide();
         }

         private void MAIN_FormClosed(object sender, FormClosedEventArgs e)
         {
          

         }
   
         private void groupBox1_Paint(object sender, PaintEventArgs e)
         {
             e.Graphics.Clear(this.c2);
         }
    }
}
