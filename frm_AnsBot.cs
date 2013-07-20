using System;
using System.IO;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;

namespace WindowsApplication5
{
	/// <summary>
	/// Form1 的摘要描述。
	/// </summary>
	public class frm_AnsBot : System.Windows.Forms.Form
	{
		//--自設變數開始
		private DataSet dSet;				//資料子集
		private string strFileName ;		//使用者report檔案位置
		private string strAns;				//是非題解答
		private OleDbConnection oleconn;	//連結控制用
		private OleDbDataAdapter oleAdap;	//資料暫存用
		private OleDbCommand oleCmd ;
		private OleDbDataReader oleReader ;
		private string strSrcDb ;			//資料庫位置
		private int iStarPlace;				//起始位置設定(內定,未提供使用者更改)
		private int iLeng;					//語言切換
		private int iLen;					//關鍵字長度
		private bool bStyle;					//題型true:選擇 false:是非

		//--自設變數結尾
		private System.Windows.Forms.Label lbeName;
		private System.Windows.Forms.TextBox txtName;
		private System.Windows.Forms.Button btnTest;
		private System.Windows.Forms.Button btnAns;
		private System.Windows.Forms.TextBox txtNo1;
		private System.Windows.Forms.Label lbeNo1;
		private System.Windows.Forms.TextBox txtNo2;
		private System.Windows.Forms.TextBox txtNo3;
		private System.Windows.Forms.TextBox txtNo4;
		private System.Windows.Forms.TextBox txtNo5;
		private System.Windows.Forms.TextBox txtNo6;
		private System.Windows.Forms.TextBox txtNo7;
		private System.Windows.Forms.Label lbeNo2;
		private System.Windows.Forms.Label lbeNo3;
		private System.Windows.Forms.Label lbeNo4;
		private System.Windows.Forms.Label lbeNo7;
		private System.Windows.Forms.Label lbeNo6;
		private System.Windows.Forms.Label lbeNo5;
		private System.Windows.Forms.TabControl tabControl1;
		private System.Windows.Forms.TabPage tabPage2;
		private System.Windows.Forms.TabPage tabPage3;
		private System.Windows.Forms.TextBox txtLike;
		private System.Windows.Forms.Button btnLike;
		private System.Windows.Forms.Label lbeException;
		private System.Windows.Forms.ComboBox cmbLen;
		private System.Windows.Forms.ComboBox cmbLeg;
		private System.Windows.Forms.TabPage tabPage1;
		private System.Windows.Forms.Button btnStrClear;
		private System.Windows.Forms.Button btnAnsClear;
		private System.Windows.Forms.ComboBox cmbPlace;
		private System.Windows.Forms.Button btnUpdate;
		private System.Windows.Forms.TextBox txtUdStr;
		private System.Windows.Forms.Label lbePlace;
		private System.Windows.Forms.Label lbeQ;
		private System.Windows.Forms.TextBox txtQNum;
		private System.Windows.Forms.ComboBox cmbStyle;
		private System.Windows.Forms.ComboBox cmbStyle2;
		public System.Windows.Forms.DataGrid dataGrid2;
		private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn1;
		private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn2;
		private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn3;
		private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn4;
		private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn5;
		private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn6;
		private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn7;
		public System.Windows.Forms.DataGrid dataGrid1;
		private System.Windows.Forms.DataGridTableStyle dataGridTableStyle1;
		private System.Windows.Forms.DataGridTableStyle dataGridTableStyle2;
		private System.Windows.Forms.DataGridTableStyle dataGridTableStyle3;
		private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn8;
		private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn9;
		private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn10;
		private System.Windows.Forms.Timer tmeExAns;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.TextBox txtNo8;
		private System.Windows.Forms.TextBox txtNo10;
		private System.Windows.Forms.Label lbeNo8;
		private System.Windows.Forms.Label llbeNo9;
		private System.Windows.Forms.Label lbeNo10;
		private System.Windows.Forms.Label lbeNo11;
		private System.Windows.Forms.TextBox txtNo11;
		private System.Windows.Forms.TextBox txtNo9;
		private System.Windows.Forms.TextBox txtNo12;
		private System.Windows.Forms.Label lbeNo12;
		private System.Windows.Forms.Button button1;
		private System.ComponentModel.IContainer components;

		public frm_AnsBot()
		{
			//
			// Windows Form 設計工具支援的必要項
			//
			InitializeComponent();

			//
			// TODO: 在 InitializeComponent 呼叫之後加入任何建構函式程式碼
			//
		}

		/// <summary>
		/// 清除任何使用中的資源。
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form 設計工具產生的程式碼
		/// <summary>
		/// 此為設計工具支援所必須的方法 - 請勿使用程式碼編輯器修改
		/// 這個方法的內容。
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frm_AnsBot));
			this.lbeName = new System.Windows.Forms.Label();
			this.txtName = new System.Windows.Forms.TextBox();
			this.btnTest = new System.Windows.Forms.Button();
			this.btnAns = new System.Windows.Forms.Button();
			this.txtNo1 = new System.Windows.Forms.TextBox();
			this.lbeNo1 = new System.Windows.Forms.Label();
			this.txtNo2 = new System.Windows.Forms.TextBox();
			this.txtNo3 = new System.Windows.Forms.TextBox();
			this.txtNo4 = new System.Windows.Forms.TextBox();
			this.txtNo5 = new System.Windows.Forms.TextBox();
			this.txtNo6 = new System.Windows.Forms.TextBox();
			this.txtNo7 = new System.Windows.Forms.TextBox();
			this.lbeNo2 = new System.Windows.Forms.Label();
			this.lbeNo3 = new System.Windows.Forms.Label();
			this.lbeNo4 = new System.Windows.Forms.Label();
			this.lbeNo7 = new System.Windows.Forms.Label();
			this.lbeNo6 = new System.Windows.Forms.Label();
			this.lbeNo5 = new System.Windows.Forms.Label();
			this.tabControl1 = new System.Windows.Forms.TabControl();
			this.tabPage2 = new System.Windows.Forms.TabPage();
			this.dataGrid1 = new System.Windows.Forms.DataGrid();
			this.dataGridTableStyle1 = new System.Windows.Forms.DataGridTableStyle();
			this.dataGridTextBoxColumn1 = new System.Windows.Forms.DataGridTextBoxColumn();
			this.dataGridTextBoxColumn2 = new System.Windows.Forms.DataGridTextBoxColumn();
			this.dataGridTableStyle2 = new System.Windows.Forms.DataGridTableStyle();
			this.dataGridTextBoxColumn3 = new System.Windows.Forms.DataGridTextBoxColumn();
			this.dataGridTextBoxColumn4 = new System.Windows.Forms.DataGridTextBoxColumn();
			this.dataGridTextBoxColumn5 = new System.Windows.Forms.DataGridTextBoxColumn();
			this.dataGridTextBoxColumn6 = new System.Windows.Forms.DataGridTextBoxColumn();
			this.dataGridTextBoxColumn7 = new System.Windows.Forms.DataGridTextBoxColumn();
			this.tabPage1 = new System.Windows.Forms.TabPage();
			this.cmbStyle2 = new System.Windows.Forms.ComboBox();
			this.dataGrid2 = new System.Windows.Forms.DataGrid();
			this.dataGridTableStyle3 = new System.Windows.Forms.DataGridTableStyle();
			this.dataGridTextBoxColumn8 = new System.Windows.Forms.DataGridTextBoxColumn();
			this.dataGridTextBoxColumn9 = new System.Windows.Forms.DataGridTextBoxColumn();
			this.dataGridTextBoxColumn10 = new System.Windows.Forms.DataGridTextBoxColumn();
			this.cmbPlace = new System.Windows.Forms.ComboBox();
			this.btnUpdate = new System.Windows.Forms.Button();
			this.txtUdStr = new System.Windows.Forms.TextBox();
			this.lbePlace = new System.Windows.Forms.Label();
			this.txtQNum = new System.Windows.Forms.TextBox();
			this.lbeQ = new System.Windows.Forms.Label();
			this.tabPage3 = new System.Windows.Forms.TabPage();
			this.lbeException = new System.Windows.Forms.Label();
			this.txtLike = new System.Windows.Forms.TextBox();
			this.btnLike = new System.Windows.Forms.Button();
			this.cmbLen = new System.Windows.Forms.ComboBox();
			this.cmbLeg = new System.Windows.Forms.ComboBox();
			this.cmbStyle = new System.Windows.Forms.ComboBox();
			this.tmeExAns = new System.Windows.Forms.Timer(this.components);
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.txtNo8 = new System.Windows.Forms.TextBox();
			this.txtNo10 = new System.Windows.Forms.TextBox();
			this.lbeNo8 = new System.Windows.Forms.Label();
			this.llbeNo9 = new System.Windows.Forms.Label();
			this.lbeNo10 = new System.Windows.Forms.Label();
			this.lbeNo11 = new System.Windows.Forms.Label();
			this.txtNo11 = new System.Windows.Forms.TextBox();
			this.txtNo9 = new System.Windows.Forms.TextBox();
			this.txtNo12 = new System.Windows.Forms.TextBox();
			this.lbeNo12 = new System.Windows.Forms.Label();
			this.button1 = new System.Windows.Forms.Button();
			this.tabControl1.SuspendLayout();
			this.tabPage2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.dataGrid1)).BeginInit();
			this.tabPage1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.dataGrid2)).BeginInit();
			this.tabPage3.SuspendLayout();
			this.SuspendLayout();
			// 
			// lbeName
			// 
			this.lbeName.Location = new System.Drawing.Point(8, 16);
			this.lbeName.Name = "lbeName";
			this.lbeName.Size = new System.Drawing.Size(64, 23);
			this.lbeName.TabIndex = 0;
			this.lbeName.Text = "玩家名稱:";
			// 
			// txtName
			// 
			this.txtName.Location = new System.Drawing.Point(64, 8);
			this.txtName.Name = "txtName";
			this.txtName.Size = new System.Drawing.Size(112, 22);
			this.txtName.TabIndex = 1;
			this.txtName.Text = "chat";
			// 
			// btnTest
			// 
			this.btnTest.Location = new System.Drawing.Point(184, 8);
			this.btnTest.Name = "btnTest";
			this.btnTest.Size = new System.Drawing.Size(40, 23);
			this.btnTest.TabIndex = 2;
			this.btnTest.Text = "測試";
			this.btnTest.Click += new System.EventHandler(this.btnTest_Click);
			// 
			// btnAns
			// 
			this.btnAns.Enabled = false;
			this.btnAns.Location = new System.Drawing.Point(304, 8);
			this.btnAns.Name = "btnAns";
			this.btnAns.Size = new System.Drawing.Size(112, 23);
			this.btnAns.TabIndex = 2;
			this.btnAns.Text = "開始解題";
			this.btnAns.Click += new System.EventHandler(this.btnAns_Click);
			// 
			// txtNo1
			// 
			this.txtNo1.Location = new System.Drawing.Point(32, 40);
			this.txtNo1.Name = "txtNo1";
			this.txtNo1.Size = new System.Drawing.Size(24, 22);
			this.txtNo1.TabIndex = 5;
			this.txtNo1.Text = "";
			// 
			// lbeNo1
			// 
			this.lbeNo1.Location = new System.Drawing.Point(16, 48);
			this.lbeNo1.Name = "lbeNo1";
			this.lbeNo1.Size = new System.Drawing.Size(24, 23);
			this.lbeNo1.TabIndex = 0;
			this.lbeNo1.Text = "1.";
			// 
			// txtNo2
			// 
			this.txtNo2.ForeColor = System.Drawing.SystemColors.WindowText;
			this.txtNo2.Location = new System.Drawing.Point(72, 40);
			this.txtNo2.Name = "txtNo2";
			this.txtNo2.Size = new System.Drawing.Size(24, 22);
			this.txtNo2.TabIndex = 5;
			this.txtNo2.Text = "";
			// 
			// txtNo3
			// 
			this.txtNo3.Location = new System.Drawing.Point(112, 40);
			this.txtNo3.Name = "txtNo3";
			this.txtNo3.Size = new System.Drawing.Size(24, 22);
			this.txtNo3.TabIndex = 5;
			this.txtNo3.Text = "";
			// 
			// txtNo4
			// 
			this.txtNo4.Location = new System.Drawing.Point(152, 40);
			this.txtNo4.Name = "txtNo4";
			this.txtNo4.Size = new System.Drawing.Size(24, 22);
			this.txtNo4.TabIndex = 5;
			this.txtNo4.Text = "";
			// 
			// txtNo5
			// 
			this.txtNo5.Location = new System.Drawing.Point(192, 40);
			this.txtNo5.Name = "txtNo5";
			this.txtNo5.Size = new System.Drawing.Size(24, 22);
			this.txtNo5.TabIndex = 5;
			this.txtNo5.Text = "";
			// 
			// txtNo6
			// 
			this.txtNo6.Location = new System.Drawing.Point(232, 40);
			this.txtNo6.Name = "txtNo6";
			this.txtNo6.Size = new System.Drawing.Size(24, 22);
			this.txtNo6.TabIndex = 5;
			this.txtNo6.Text = "";
			// 
			// txtNo7
			// 
			this.txtNo7.Location = new System.Drawing.Point(272, 40);
			this.txtNo7.Name = "txtNo7";
			this.txtNo7.Size = new System.Drawing.Size(24, 22);
			this.txtNo7.TabIndex = 5;
			this.txtNo7.Text = "";
			// 
			// lbeNo2
			// 
			this.lbeNo2.Location = new System.Drawing.Point(56, 48);
			this.lbeNo2.Name = "lbeNo2";
			this.lbeNo2.Size = new System.Drawing.Size(24, 23);
			this.lbeNo2.TabIndex = 0;
			this.lbeNo2.Text = "2.";
			// 
			// lbeNo3
			// 
			this.lbeNo3.Location = new System.Drawing.Point(96, 48);
			this.lbeNo3.Name = "lbeNo3";
			this.lbeNo3.Size = new System.Drawing.Size(24, 23);
			this.lbeNo3.TabIndex = 0;
			this.lbeNo3.Text = "3.";
			// 
			// lbeNo4
			// 
			this.lbeNo4.Location = new System.Drawing.Point(136, 48);
			this.lbeNo4.Name = "lbeNo4";
			this.lbeNo4.Size = new System.Drawing.Size(24, 23);
			this.lbeNo4.TabIndex = 0;
			this.lbeNo4.Text = "4.";
			// 
			// lbeNo7
			// 
			this.lbeNo7.Location = new System.Drawing.Point(256, 48);
			this.lbeNo7.Name = "lbeNo7";
			this.lbeNo7.Size = new System.Drawing.Size(24, 23);
			this.lbeNo7.TabIndex = 0;
			this.lbeNo7.Text = "7.";
			// 
			// lbeNo6
			// 
			this.lbeNo6.Location = new System.Drawing.Point(216, 48);
			this.lbeNo6.Name = "lbeNo6";
			this.lbeNo6.Size = new System.Drawing.Size(24, 23);
			this.lbeNo6.TabIndex = 0;
			this.lbeNo6.Text = "6.";
			// 
			// lbeNo5
			// 
			this.lbeNo5.Location = new System.Drawing.Point(176, 48);
			this.lbeNo5.Name = "lbeNo5";
			this.lbeNo5.Size = new System.Drawing.Size(24, 23);
			this.lbeNo5.TabIndex = 0;
			this.lbeNo5.Text = "5.";
			// 
			// tabControl1
			// 
			this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.tabControl1.Controls.Add(this.tabPage2);
			this.tabControl1.Controls.Add(this.tabPage1);
			this.tabControl1.Controls.Add(this.tabPage3);
			this.tabControl1.Location = new System.Drawing.Point(8, 144);
			this.tabControl1.Name = "tabControl1";
			this.tabControl1.SelectedIndex = 0;
			this.tabControl1.Size = new System.Drawing.Size(408, 160);
			this.tabControl1.TabIndex = 7;
			// 
			// tabPage2
			// 
			this.tabPage2.Controls.Add(this.dataGrid1);
			this.tabPage2.Location = new System.Drawing.Point(4, 21);
			this.tabPage2.Name = "tabPage2";
			this.tabPage2.Size = new System.Drawing.Size(400, 135);
			this.tabPage2.TabIndex = 1;
			this.tabPage2.Text = "符合題組";
			// 
			// dataGrid1
			// 
			this.dataGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.dataGrid1.DataMember = "";
			this.dataGrid1.HeaderForeColor = System.Drawing.SystemColors.ControlText;
			this.dataGrid1.Location = new System.Drawing.Point(0, 8);
			this.dataGrid1.Name = "dataGrid1";
			this.dataGrid1.Size = new System.Drawing.Size(400, 128);
			this.dataGrid1.TabIndex = 6;
			this.dataGrid1.TableStyles.AddRange(new System.Windows.Forms.DataGridTableStyle[] {
																								  this.dataGridTableStyle1,
																								  this.dataGridTableStyle2});
			// 
			// dataGridTableStyle1
			// 
			this.dataGridTableStyle1.DataGrid = this.dataGrid1;
			this.dataGridTableStyle1.GridColumnStyles.AddRange(new System.Windows.Forms.DataGridColumnStyle[] {
																												  this.dataGridTextBoxColumn1,
																												  this.dataGridTextBoxColumn2});
			this.dataGridTableStyle1.HeaderForeColor = System.Drawing.SystemColors.ControlText;
			this.dataGridTableStyle1.MappingName = "Result";
			// 
			// dataGridTextBoxColumn1
			// 
			this.dataGridTextBoxColumn1.Format = "";
			this.dataGridTextBoxColumn1.FormatInfo = null;
			this.dataGridTextBoxColumn1.HeaderText = "解答";
			this.dataGridTextBoxColumn1.MappingName = "Ans";
			this.dataGridTextBoxColumn1.Width = 35;
			// 
			// dataGridTextBoxColumn2
			// 
			this.dataGridTextBoxColumn2.Format = "";
			this.dataGridTextBoxColumn2.FormatInfo = null;
			this.dataGridTextBoxColumn2.HeaderText = "問題描述";
			this.dataGridTextBoxColumn2.MappingName = "PrbCH";
			this.dataGridTextBoxColumn2.Width = 800;
			// 
			// dataGridTableStyle2
			// 
			this.dataGridTableStyle2.DataGrid = this.dataGrid1;
			this.dataGridTableStyle2.GridColumnStyles.AddRange(new System.Windows.Forms.DataGridColumnStyle[] {
																												  this.dataGridTextBoxColumn3,
																												  this.dataGridTextBoxColumn4,
																												  this.dataGridTextBoxColumn5,
																												  this.dataGridTextBoxColumn6,
																												  this.dataGridTextBoxColumn7});
			this.dataGridTableStyle2.HeaderForeColor = System.Drawing.SystemColors.ControlText;
			this.dataGridTableStyle2.MappingName = "Result2";
			// 
			// dataGridTextBoxColumn3
			// 
			this.dataGridTextBoxColumn3.Format = "";
			this.dataGridTextBoxColumn3.FormatInfo = null;
			this.dataGridTextBoxColumn3.HeaderText = "問題描述";
			this.dataGridTextBoxColumn3.MappingName = "PrbCH";
			this.dataGridTextBoxColumn3.Width = 225;
			// 
			// dataGridTextBoxColumn4
			// 
			this.dataGridTextBoxColumn4.Format = "";
			this.dataGridTextBoxColumn4.FormatInfo = null;
			this.dataGridTextBoxColumn4.HeaderText = "解答1";
			this.dataGridTextBoxColumn4.MappingName = "Ans1";
			this.dataGridTextBoxColumn4.Width = 75;
			// 
			// dataGridTextBoxColumn5
			// 
			this.dataGridTextBoxColumn5.Format = "";
			this.dataGridTextBoxColumn5.FormatInfo = null;
			this.dataGridTextBoxColumn5.HeaderText = "解答2";
			this.dataGridTextBoxColumn5.MappingName = "Ans2";
			this.dataGridTextBoxColumn5.Width = 75;
			// 
			// dataGridTextBoxColumn6
			// 
			this.dataGridTextBoxColumn6.Format = "";
			this.dataGridTextBoxColumn6.FormatInfo = null;
			this.dataGridTextBoxColumn6.HeaderText = "解答3";
			this.dataGridTextBoxColumn6.MappingName = "Ans3";
			this.dataGridTextBoxColumn6.Width = 75;
			// 
			// dataGridTextBoxColumn7
			// 
			this.dataGridTextBoxColumn7.Format = "";
			this.dataGridTextBoxColumn7.FormatInfo = null;
			this.dataGridTextBoxColumn7.HeaderText = "解答4";
			this.dataGridTextBoxColumn7.MappingName = "Ans4";
			this.dataGridTextBoxColumn7.Width = 75;
			// 
			// tabPage1
			// 
			this.tabPage1.Controls.Add(this.cmbStyle2);
			this.tabPage1.Controls.Add(this.dataGrid2);
			this.tabPage1.Controls.Add(this.cmbPlace);
			this.tabPage1.Controls.Add(this.btnUpdate);
			this.tabPage1.Controls.Add(this.txtUdStr);
			this.tabPage1.Controls.Add(this.lbePlace);
			this.tabPage1.Controls.Add(this.txtQNum);
			this.tabPage1.Controls.Add(this.lbeQ);
			this.tabPage1.Location = new System.Drawing.Point(4, 21);
			this.tabPage1.Name = "tabPage1";
			this.tabPage1.Size = new System.Drawing.Size(400, 135);
			this.tabPage1.TabIndex = 3;
			this.tabPage1.Text = "手動更新題庫";
			// 
			// cmbStyle2
			// 
			this.cmbStyle2.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbStyle2.ItemHeight = 12;
			this.cmbStyle2.Items.AddRange(new object[] {
														   "選擇(Qx)",
														   "是非(Qz)"});
			this.cmbStyle2.Location = new System.Drawing.Point(8, 32);
			this.cmbStyle2.Name = "cmbStyle2";
			this.cmbStyle2.Size = new System.Drawing.Size(56, 20);
			this.cmbStyle2.TabIndex = 0;
			this.cmbStyle2.Text = "題組";
			// 
			// dataGrid2
			// 
			this.dataGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.dataGrid2.DataMember = "";
			this.dataGrid2.HeaderForeColor = System.Drawing.SystemColors.ControlText;
			this.dataGrid2.Location = new System.Drawing.Point(0, 64);
			this.dataGrid2.Name = "dataGrid2";
			this.dataGrid2.Size = new System.Drawing.Size(400, 72);
			this.dataGrid2.TabIndex = 5;
			this.dataGrid2.TableStyles.AddRange(new System.Windows.Forms.DataGridTableStyle[] {
																								  this.dataGridTableStyle3});
			// 
			// dataGridTableStyle3
			// 
			this.dataGridTableStyle3.DataGrid = this.dataGrid2;
			this.dataGridTableStyle3.GridColumnStyles.AddRange(new System.Windows.Forms.DataGridColumnStyle[] {
																												  this.dataGridTextBoxColumn8,
																												  this.dataGridTextBoxColumn9,
																												  this.dataGridTextBoxColumn10});
			this.dataGridTableStyle3.HeaderForeColor = System.Drawing.SystemColors.ControlText;
			this.dataGridTableStyle3.MappingName = "Result3";
			// 
			// dataGridTextBoxColumn8
			// 
			this.dataGridTextBoxColumn8.Format = "";
			this.dataGridTextBoxColumn8.FormatInfo = null;
			this.dataGridTextBoxColumn8.HeaderText = "題號";
			this.dataGridTextBoxColumn8.MappingName = "PrbNo";
			this.dataGridTextBoxColumn8.Width = 75;
			// 
			// dataGridTextBoxColumn9
			// 
			this.dataGridTextBoxColumn9.Format = "";
			this.dataGridTextBoxColumn9.FormatInfo = null;
			this.dataGridTextBoxColumn9.HeaderText = "顯示字詞";
			this.dataGridTextBoxColumn9.MappingName = "PrbCH";
			this.dataGridTextBoxColumn9.Width = 75;
			// 
			// dataGridTextBoxColumn10
			// 
			this.dataGridTextBoxColumn10.Format = "";
			this.dataGridTextBoxColumn10.FormatInfo = null;
			this.dataGridTextBoxColumn10.HeaderText = "判定字詞";
			this.dataGridTextBoxColumn10.MappingName = "PrbGB";
			this.dataGridTextBoxColumn10.Width = 75;
			// 
			// cmbPlace
			// 
			this.cmbPlace.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPlace.ItemHeight = 12;
			this.cmbPlace.Items.AddRange(new object[] {
														  "顯示字",
														  "判斷字"});
			this.cmbPlace.Location = new System.Drawing.Point(72, 8);
			this.cmbPlace.Name = "cmbPlace";
			this.cmbPlace.Size = new System.Drawing.Size(56, 20);
			this.cmbPlace.TabIndex = 0;
			this.cmbPlace.Text = "位置";
			// 
			// btnUpdate
			// 
			this.btnUpdate.Location = new System.Drawing.Point(216, 32);
			this.btnUpdate.Name = "btnUpdate";
			this.btnUpdate.Size = new System.Drawing.Size(80, 23);
			this.btnUpdate.TabIndex = 9;
			this.btnUpdate.Text = "更新";
			this.btnUpdate.Click += new System.EventHandler(this.btnUpdate_Click);
			// 
			// txtUdStr
			// 
			this.txtUdStr.Location = new System.Drawing.Point(72, 32);
			this.txtUdStr.Name = "txtUdStr";
			this.txtUdStr.Size = new System.Drawing.Size(136, 22);
			this.txtUdStr.TabIndex = 8;
			this.txtUdStr.Text = "";
			// 
			// lbePlace
			// 
			this.lbePlace.Location = new System.Drawing.Point(8, 16);
			this.lbePlace.Name = "lbePlace";
			this.lbePlace.TabIndex = 0;
			this.lbePlace.Text = "更新位置:";
			// 
			// txtQNum
			// 
			this.txtQNum.Location = new System.Drawing.Point(160, 8);
			this.txtQNum.Name = "txtQNum";
			this.txtQNum.Size = new System.Drawing.Size(48, 22);
			this.txtQNum.TabIndex = 1;
			this.txtQNum.Text = "";
			// 
			// lbeQ
			// 
			this.lbeQ.Location = new System.Drawing.Point(128, 16);
			this.lbeQ.Name = "lbeQ";
			this.lbeQ.TabIndex = 0;
			this.lbeQ.Text = "題號:";
			// 
			// tabPage3
			// 
			this.tabPage3.Controls.Add(this.lbeException);
			this.tabPage3.Location = new System.Drawing.Point(4, 21);
			this.tabPage3.Name = "tabPage3";
			this.tabPage3.Size = new System.Drawing.Size(400, 135);
			this.tabPage3.TabIndex = 2;
			this.tabPage3.Text = "其他資訊";
			// 
			// lbeException
			// 
			this.lbeException.Location = new System.Drawing.Point(8, 8);
			this.lbeException.Name = "lbeException";
			this.lbeException.Size = new System.Drawing.Size(392, 120);
			this.lbeException.TabIndex = 0;
			this.lbeException.Text = "mail:wkunhui@gmail.com\r\nV2.1:修改對話中若使用@會造成程式當機問題。\r\nV3." +
				"0:紀錄當前已公布答案但未存於資料庫的值\r\nV3.1:修改選擇第8題後無法出現提試問題\r\nV3.2:修改ox題目無選項時提供相似題庫";
			// 
			// txtLike
			// 
			this.txtLike.Location = new System.Drawing.Point(136, 112);
			this.txtLike.Name = "txtLike";
			this.txtLike.Size = new System.Drawing.Size(200, 22);
			this.txtLike.TabIndex = 8;
			this.txtLike.Text = "";
			// 
			// btnLike
			// 
			this.btnLike.Location = new System.Drawing.Point(336, 112);
			this.btnLike.Name = "btnLike";
			this.btnLike.Size = new System.Drawing.Size(80, 23);
			this.btnLike.TabIndex = 9;
			this.btnLike.Text = "關鍵字搜索";
			this.btnLike.Click += new System.EventHandler(this.btnLike_Click);
			// 
			// cmbLen
			// 
			this.cmbLen.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbLen.Items.AddRange(new object[] {
														"1個字",
														"2個字",
														"3個字",
														"4個字",
														"5個字",
														"6個字",
														"7個字",
														"8個字",
														"9個字"});
			this.cmbLen.Location = new System.Drawing.Point(232, 8);
			this.cmbLen.Name = "cmbLen";
			this.cmbLen.Size = new System.Drawing.Size(64, 20);
			this.cmbLen.TabIndex = 0;
			this.cmbLen.Text = "相似度";
			// 
			// cmbLeg
			// 
			this.cmbLeg.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbLeg.ItemHeight = 12;
			this.cmbLeg.Items.AddRange(new object[] {
														"繁體",
														"簡体"});
			this.cmbLeg.Location = new System.Drawing.Point(16, 112);
			this.cmbLeg.Name = "cmbLeg";
			this.cmbLeg.Size = new System.Drawing.Size(56, 20);
			this.cmbLeg.TabIndex = 0;
			this.cmbLeg.Text = "語言";
			// 
			// cmbStyle
			// 
			this.cmbStyle.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbStyle.ItemHeight = 12;
			this.cmbStyle.Items.AddRange(new object[] {
														  "選擇",
														  "是非"});
			this.cmbStyle.Location = new System.Drawing.Point(80, 112);
			this.cmbStyle.Name = "cmbStyle";
			this.cmbStyle.Size = new System.Drawing.Size(56, 20);
			this.cmbStyle.TabIndex = 0;
			this.cmbStyle.Text = "題組";
			// 
			// tmeExAns
			// 
			this.tmeExAns.Tick += new System.EventHandler(this.tmeExAns_Tick);
			// 
			// label1
			// 
			this.label1.Font = new System.Drawing.Font("新細明體", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(136)));
			this.label1.ForeColor = System.Drawing.Color.Red;
			this.label1.Location = new System.Drawing.Point(16, 96);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(296, 16);
			this.label1.TabIndex = 10;
			this.label1.Text = "report檔案使用中,無法刪除!!";
			this.label1.Visible = false;
			// 
			// label2
			// 
			this.label2.Font = new System.Drawing.Font("新細明體", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(136)));
			this.label2.ForeColor = System.Drawing.Color.Red;
			this.label2.Location = new System.Drawing.Point(16, 96);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(296, 16);
			this.label2.TabIndex = 10;
			this.label2.Text = "無符合資料!!";
			this.label2.Visible = false;
			// 
			// label3
			// 
			this.label3.Font = new System.Drawing.Font("新細明體", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(136)));
			this.label3.ForeColor = System.Drawing.Color.Red;
			this.label3.Location = new System.Drawing.Point(8, 96);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(504, 16);
			this.label3.TabIndex = 11;
			this.label3.Visible = false;
			// 
			// txtNo8
			// 
			this.txtNo8.Location = new System.Drawing.Point(32, 72);
			this.txtNo8.Name = "txtNo8";
			this.txtNo8.Size = new System.Drawing.Size(64, 22);
			this.txtNo8.TabIndex = 5;
			this.txtNo8.Text = "";
			// 
			// txtNo10
			// 
			this.txtNo10.Location = new System.Drawing.Point(192, 72);
			this.txtNo10.Name = "txtNo10";
			this.txtNo10.Size = new System.Drawing.Size(64, 22);
			this.txtNo10.TabIndex = 5;
			this.txtNo10.Text = "";
			// 
			// lbeNo8
			// 
			this.lbeNo8.Location = new System.Drawing.Point(16, 80);
			this.lbeNo8.Name = "lbeNo8";
			this.lbeNo8.Size = new System.Drawing.Size(24, 23);
			this.lbeNo8.TabIndex = 0;
			this.lbeNo8.Text = "8.";
			// 
			// llbeNo9
			// 
			this.llbeNo9.Location = new System.Drawing.Point(96, 80);
			this.llbeNo9.Name = "llbeNo9";
			this.llbeNo9.Size = new System.Drawing.Size(24, 23);
			this.llbeNo9.TabIndex = 0;
			this.llbeNo9.Text = "9.";
			// 
			// lbeNo10
			// 
			this.lbeNo10.Location = new System.Drawing.Point(176, 80);
			this.lbeNo10.Name = "lbeNo10";
			this.lbeNo10.Size = new System.Drawing.Size(24, 23);
			this.lbeNo10.TabIndex = 0;
			this.lbeNo10.Text = "10";
			// 
			// lbeNo11
			// 
			this.lbeNo11.Location = new System.Drawing.Point(256, 80);
			this.lbeNo11.Name = "lbeNo11";
			this.lbeNo11.Size = new System.Drawing.Size(24, 23);
			this.lbeNo11.TabIndex = 0;
			this.lbeNo11.Text = "11";
			// 
			// txtNo11
			// 
			this.txtNo11.Location = new System.Drawing.Point(272, 72);
			this.txtNo11.Name = "txtNo11";
			this.txtNo11.Size = new System.Drawing.Size(64, 22);
			this.txtNo11.TabIndex = 5;
			this.txtNo11.Text = "";
			// 
			// txtNo9
			// 
			this.txtNo9.Location = new System.Drawing.Point(112, 72);
			this.txtNo9.Name = "txtNo9";
			this.txtNo9.Size = new System.Drawing.Size(64, 22);
			this.txtNo9.TabIndex = 5;
			this.txtNo9.Text = "";
			// 
			// txtNo12
			// 
			this.txtNo12.Location = new System.Drawing.Point(352, 72);
			this.txtNo12.Name = "txtNo12";
			this.txtNo12.Size = new System.Drawing.Size(64, 22);
			this.txtNo12.TabIndex = 5;
			this.txtNo12.Text = "";
			// 
			// lbeNo12
			// 
			this.lbeNo12.Location = new System.Drawing.Point(336, 80);
			this.lbeNo12.Name = "lbeNo12";
			this.lbeNo12.Size = new System.Drawing.Size(24, 23);
			this.lbeNo12.TabIndex = 0;
			this.lbeNo12.Text = "12";
			// 
			// button1
			// 
			this.button1.Location = new System.Drawing.Point(304, 40);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(112, 23);
			this.button1.TabIndex = 2;
			this.button1.Text = "清除暫存資料庫";
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// frm_AnsBot
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 15);
			this.ClientSize = new System.Drawing.Size(424, 310);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtNo1);
			this.Controls.Add(this.txtName);
			this.Controls.Add(this.txtNo2);
			this.Controls.Add(this.txtNo3);
			this.Controls.Add(this.txtNo4);
			this.Controls.Add(this.txtNo5);
			this.Controls.Add(this.txtNo6);
			this.Controls.Add(this.txtNo7);
			this.Controls.Add(this.txtLike);
			this.Controls.Add(this.txtNo8);
			this.Controls.Add(this.txtNo10);
			this.Controls.Add(this.txtNo11);
			this.Controls.Add(this.txtNo9);
			this.Controls.Add(this.txtNo12);
			this.Controls.Add(this.tabControl1);
			this.Controls.Add(this.btnTest);
			this.Controls.Add(this.lbeName);
			this.Controls.Add(this.btnAns);
			this.Controls.Add(this.lbeNo1);
			this.Controls.Add(this.lbeNo2);
			this.Controls.Add(this.lbeNo3);
			this.Controls.Add(this.lbeNo4);
			this.Controls.Add(this.lbeNo7);
			this.Controls.Add(this.lbeNo6);
			this.Controls.Add(this.lbeNo5);
			this.Controls.Add(this.cmbLen);
			this.Controls.Add(this.cmbLeg);
			this.Controls.Add(this.btnLike);
			this.Controls.Add(this.cmbStyle);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.lbeNo8);
			this.Controls.Add(this.lbeNo10);
			this.Controls.Add(this.llbeNo9);
			this.Controls.Add(this.lbeNo11);
			this.Controls.Add(this.lbeNo12);
			this.Controls.Add(this.button1);
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Name = "frm_AnsBot";
			this.Text = "解題達人[繁]v3.2";
			this.Load += new System.EventHandler(this.frm_AnsBot_Load);
			this.tabControl1.ResumeLayout(false);
			this.tabPage2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.dataGrid1)).EndInit();
			this.tabPage1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.dataGrid2)).EndInit();
			this.tabPage3.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// 應用程式的主進入點。
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new frm_AnsBot());
		}

		private void frm_AnsBot_Load(object sender, System.EventArgs e)
		{
			this.cmbStyle2.SelectedIndex = 1;
			this.cmbPlace.SelectedIndex = 0;
			this.cmbLeg.SelectedIndex = 0;
			this.cmbStyle.SelectedIndex = 1;
			this.cmbLen.SelectedIndex = 4;
			this.iLen = 0;
			this.iLeng = 1;
			this.iStarPlace = 2;
			this.bStyle = false;
			this.strSrcDb = Application.StartupPath +@"\Ansdb3.0.mdb";
			this.oleconn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + this.strSrcDb );

		
		}

		private void btnLike_Click(object sender, System.EventArgs e)
		{
			this.bStyle = this.cmbStyle.SelectedIndex == 1 ? false:true;		
			this.iLeng = this.cmbLeg.SelectedIndex;			
			this.iStarPlace = 0;
			this.iLen = this.txtLike.Text.Length ;
			fnSearch(this.txtLike.Text);
		}
		public void fnReAns(string strTmp,int iNum)
		{
			switch(iNum)
			{
				case 0:
					this.txtNo1.Text = strTmp;
					break;
				case 1:
					this.txtNo2.Text = strTmp;
					break;
				case 2:
					this.txtNo3.Text = strTmp;
					break;
				case 3:
					this.txtNo4.Text = strTmp;
					break;
				case 4:
					this.txtNo5.Text = strTmp;
					break;
				case 5:
					this.txtNo6.Text = strTmp;
					break;
				case 6:
					this.txtNo7.Text = strTmp;
					break;
				case 7:
					this.txtNo8.Text = strTmp;
					break;
				case 8:
					this.txtNo9.Text = strTmp;
					break;
				case 9:
					this.txtNo10.Text = strTmp;
					break;
				case 10:
					this.txtNo11.Text = strTmp;
					break;
				case 11:
					this.txtNo12.Text = strTmp;
					break;
			}

		}

		private void fnSearch(string strTmp)
		{
			string strData = null;
			string strComm = null;
			string strPrbLeg = null;
			string strFnTmp = null;
			strPrbLeg = this.iLeng == 0 ? "PrbCH":"PrbCH";
			
			try 
			{
				this.oleconn.Open();

				do
				{
					if((this.iStarPlace+this.iLen) > strTmp.Length) 
					{
						this.label2.Visible = true;
						break;
					}
					else
					{
						this.label2.Visible = false;
					}
					strFnTmp = strTmp.Substring(this.iStarPlace,this.iLen);
					this.oleCmd = !this.bStyle ? new OleDbCommand( @"SELECT COUNT(*) From Data WHERE "+strPrbLeg+" Like '%"+strFnTmp+"%'", this.oleconn):new OleDbCommand( @"SELECT COUNT(*) From sData WHERE "+strPrbLeg+" Like '%"+strFnTmp+"%'", this.oleconn);
					this.iStarPlace++;

				}while(!(Convert.ToInt32(this.oleCmd.ExecuteScalar()) > 0));
				strComm = iLen == 0 ? "SELECT PrbNo,Ans,PrbCH,PrbCH From Data WHERE "+strPrbLeg+" Like '%"+strTmp+"%'": "SELECT PrbNo,Ans,PrbCH,PrbCH From Data WHERE "+strPrbLeg+" Like '%"+strFnTmp+"%'";
				strComm = !this.bStyle ? strComm : "SELECT PrbNo,Ans1,Ans2,Ans3,Ans4,PrbCH,PrbCH From sData WHERE "+strPrbLeg+" Like '%"+strFnTmp+"%'";

				this.oleAdap = new OleDbDataAdapter(strComm, this.oleconn );
				this.dSet=new DataSet(); 
				if( this.tabControl1.SelectedIndex == 0)
				{
					if(this.bStyle )
					{//true 選擇
						this.oleAdap.Fill(this.dSet,"Result2");
						this.dataGrid1.DataSource=dSet.Tables["Result2"];
					}
					else
					{//false 是非
						this.oleAdap.Fill(this.dSet,"Result");
						this.dataGrid1.DataSource = dSet.Tables["Result"];
					}
				}
				else
				{
					if(this.tabControl1.SelectedIndex == 1 )
					{
						this.oleAdap.Fill(this.dSet,"Result3");
						this.dataGrid2.DataSource=dSet.Tables["Result3"];
					}
				}
			}
			catch( Exception ex ) 
			{
//				MessageBox.Show( ex.Message+strComm );
			}
			this.oleconn.Close();
			
		}
		private void fnUpdate()
		{
			string strComm = null;
			OleDbCommand oleCmd = null;
			OleDbDataReader oleReader = null;
			bool bResult = false;
			string strPrbLeg = this.cmbPlace.SelectedIndex == 0 ? "PrbCH":"PrbCH";
			string strData = this.cmbStyle2.SelectedIndex == 1 ? "Data":"SData";

			strComm = @"UPDATE "+strData+" Set "+strPrbLeg+" = '"+this.txtUdStr.Text+"' WHERE PrbNo ='"+this.txtQNum.Text+"'";
			try 
			{
				this.oleconn.Open();
				oleCmd = new OleDbCommand(strComm, this.oleconn);
				oleCmd.ExecuteNonQuery();
				this.oleAdap = new OleDbDataAdapter("SELECT * From "+strData+" WHERE PrbNo ='"+this.txtQNum.Text+"'", this.oleconn );
				this.dSet=new DataSet(); 
				this.oleAdap.Fill(this.dSet,"Result3");
				this.dataGrid2.DataSource=dSet.Tables["Result3"];


			}
			catch( Exception ex ) 
			{
//				MessageBox.Show( ex.Message+strComm );
			}
			this.oleconn.Close();
			
		}
		private bool bfnSearch(string strTmp)
		{
			string strComm = null;
			bool bResult = false;
			strComm = @"SELECT COUNT(*) From Data WHERE PrbCH Like '%"+strTmp+"%'";
			try 
			{
				this.oleconn.Open();
				this.oleCmd = new OleDbCommand(strComm, this.oleconn);
				if(Convert.ToInt32(this.oleCmd.ExecuteScalar()) > 0)
				{
					strComm = "SELECT Ans From Data WHERE PrbCH Like '%"+strTmp+"%'";
					this.oleCmd = new OleDbCommand(strComm, this.oleconn);
					this.strAns = Convert.ToString(this.oleCmd.ExecuteScalar());
					bResult = true;
				}
				else
				{
					bResult = false;
				}			

			}
			catch( Exception ex ) 
			{
//				MessageBox.Show( ex.Message+strComm );
			}
			this.oleconn.Close();
			return bResult;
			
		}
		private bool bfnSearchX(string strTmp)
		{
			string strComm = null;
			bool bResult = false;
			strComm = @"SELECT COUNT(*) From XData WHERE Quention Like '%"+strTmp+"%'";
			try 
			{
				this.oleconn.Open();
				this.oleCmd = new OleDbCommand(strComm, this.oleconn);
				if(Convert.ToInt32(this.oleCmd.ExecuteScalar()) > 0)
				{
					strComm = "SELECT Ans From XData WHERE Quention Like '%"+strTmp+"%'";
					this.oleCmd = new OleDbCommand(strComm, this.oleconn);
					this.strAns = Convert.ToString(this.oleCmd.ExecuteScalar());
					bResult = true;
				}
				else
				{
					bResult = false;
				}			

			}
			catch( Exception ex ) 
			{
								MessageBox.Show( ex.Message+strComm );
			}
			this.oleconn.Close();
			return bResult;
			
		}
		private void fnAutoUpdate(string Ans,string AnsContent,string Question )
		{
//			strTmp = strTmp.Replace(i+"題：","@").Split('@')[1].Replace("題：","@").Split('@')[0];	
//			MessageBox.Show(strTmp);
			string strComm = null;
			OleDbCommand oleCmd = null;
			strComm = @"INSERT INTO XData (Ans,AnsContent,Quention) VALUES('"+Ans+"','"+AnsContent+"','"+Question+"')" ;
			try 
			{
				this.oleconn.Open();
				oleCmd = new OleDbCommand(strComm, this.oleconn);
				oleCmd.ExecuteNonQuery();

			}
			catch( Exception ex ) 
			{
				//				MessageBox.Show( ex.Message+strComm );
			}
			this.oleconn.Close();
			
		}
		private bool bfnSearchSel(string strTmp,string strQTmp,int iAns)
		{
			string strComm = null;
			bool bResult = false;
			string strPrbLeg = null;
			string strFnTmp = null;
			strFnTmp = strQTmp.Substring(this.iStarPlace,this.iLen);
			strPrbLeg = "Ans1";
			strPrbLeg = iAns == 2 ? "Ans2":strPrbLeg;
			strPrbLeg = iAns == 3 ? "Ans3":strPrbLeg;
			strPrbLeg = iAns == 4 ? "Ans4":strPrbLeg;
			strComm = @"SELECT COUNT(*) From sData WHERE "+strPrbLeg+" Like '%"+strTmp+"%' AND PrbCH Like '%"+strFnTmp+"%'";
			try 
			{
				this.oleconn.Open();
				this.oleCmd = new OleDbCommand(strComm, this.oleconn);
				if(Convert.ToInt32(this.oleCmd.ExecuteScalar()) > 0)
				{
					bResult = true;
				}
				else
				{
					bResult = false;
				}			

			}
			catch( Exception ex ) 
			{
				//				MessageBox.Show( ex.Message+strComm );
			}
			this.oleconn.Close();
			return bResult;
			
		}
		private bool bfnSearchSelX(string strTmp,string Quention,string sAns)
		{
			string strComm = null;
			bool bResult = false;

			strComm = @"SELECT COUNT(*) From XData WHERE Ans Like '%"+sAns+"%'AND AnsContent Like '%"+strTmp+"%'AND Quention Like '%"+Quention+"%'";
			try 
			{
				this.oleconn.Open();
				this.oleCmd = new OleDbCommand(strComm, this.oleconn);
				if(Convert.ToInt32(this.oleCmd.ExecuteScalar()) > 0)
				{
					bResult = true;
				}
				else
				{
					bResult = false;
				}			

			}
			catch( Exception ex ) 
			{
				//				MessageBox.Show( ex.Message+strComm );
			}
			this.oleconn.Close();
			return bResult;
			
		}
		public  bool fnCheckAns(string strName)
		{
			string strTmp = null;		
			this.label3.Visible = false;
			string strTmpSel = null,strQTmpSel = null,strTmpAns = null;
			string[] strStatues = {"無(99%)","1(99%)","2(99%)","12(50%)","3(99%)","13(50%)","23(50%)","123(33%)","4(99%)","14(50%)","24(50%)","124(33%)","34(50%)","134(33%)","234(33%)","無(99%)"};
			int iTmpCont = 0,iRespond = 0;
			try
			{
				FileInfo fileDir = new FileInfo(Application.StartupPath +@"\"+strName+".txt");
//				FileInfo fileDir = new FileInfo(@"D:\GAME\Gravity\sRO\Report\"+strName+".txt");
				if(fileDir.Exists)
				{
					FileStream fs = fileDir.OpenRead();
					StreamReader srRead = new StreamReader(fs,System.Text.Encoding.Default);
					string strTmpAll = srRead.ReadToEnd();
					strQTmpSel = strTmpAll;
					fs = fileDir.OpenRead();
					srRead = new StreamReader (fs,System.Text.Encoding.Default);

					if(fs.Length>0)
					{
						while(srRead.Peek() >= 0)
						{
							strTmp = srRead.ReadLine();
							if (strTmp==null)
								continue;
//							string strTmpEx = strTmp; //暫存strTmp
								for(int i = 1;i < 13;i++)
								{
//									strTmp = strTmpEx;
									if( i < 8)
									{
										this.bStyle = false;
										if(strTmp.IndexOf("：第"+i+"題：") > 0)
										{
											strTmp = strTmp.Replace('@','-').Replace("題：","@").Split('@')[1];	
											if(bfnSearch(strTmp))
											{
												fnReAns(this.strAns,i-1);
												this.iStarPlace = 0;
												this.iLen = strTmp.Length ;
												fnSearch(strTmp);

											}
											else if(bfnSearchX(strTmp))
											{
												fnReAns(this.strAns,i-1);
												this.iStarPlace = 2;
												this.iLen = this.cmbLen.SelectedIndex+1;
												fnSearch(strTmp);
											}
											else
											{
												if(strTmpAll.IndexOf("第"+i+"題：") > 0)
												{
													
													iTmpCont = strTmpAll.Replace('@','-').Replace("第"+i+"題：","@").Split('@').Length;
													strTmp = strTmpAll.Replace('@','-').Replace("第"+i+"題：","@").Split('@')[iTmpCont-1].Replace("機智問答進行至第","@").Split('@')[0];//.Replace('@','-').Replace('\r','@').Split('@')[0];//.Replace("思考時間","@").Split('@')[0];//.Replace("選項","@")
													if(strTmp.IndexOf("機智問答主持人：X") > 0)
													{
														strTmp = strTmp.Replace('\r','@').Split('@')[0];
														fnAutoUpdate("錯",null,strTmp);
													}
													else if(strTmp.IndexOf("機智問答主持人：O") > 0)
													{
														strTmp = strTmp.Replace('\r','@').Split('@')[0];
														fnAutoUpdate("對",null,strTmp);			
													}
													
												}
												fnReAns("無",i-1);
												this.iStarPlace = 2;
												this.iLen = this.cmbLen.SelectedIndex+1;
												fnSearch(strTmp);
											}
	
										}
									}
									else
									{
										this.bStyle = true;
										this.iStarPlace = 2;
										this.iLen = this.cmbLen.SelectedIndex+1;
										if(strTmp.IndexOf(i+"題：") > 0)
										{
											strTmp = strTmp.Replace('@','-').Replace("題：","@").Split('@')[1];	
											fnSearch(strTmp);
										}
										if(strQTmpSel.IndexOf(i+"題：") > 0)
										{
											iRespond = 0;
											iTmpCont = strQTmpSel.Replace('@','-').Replace(i+"題：","@").Split('@').Length;
											strQTmpSel = strQTmpSel.Replace('@','-').Replace(i+"題：","@").Split('@')[iTmpCont-1];//.Replace("思考時間","@").Split('@')[0];//.Replace("選項","@")
											strTmp = strQTmpSel.Replace('@','-').Replace('\r','@').Split('@')[0];//.Replace("思考時間","@").Split('@')[0];//.Replace("選項","@")
											if(strQTmpSel.IndexOf("選項一：") > 0)
											{
												strTmpSel = strQTmpSel.Replace('@','-').Replace("選項一：","@").Split('@')[1].Replace('\r','@').Split('@')[0];
												iRespond += bfnSearchSel(strTmpSel,strTmp,1) ? 1:0;
											}
											if(strQTmpSel.IndexOf("選項二：") > 0)
											{
												strTmpSel = strQTmpSel.Replace('@','-').Replace("選項二：","@").Split('@')[1].Replace('\r','@').Split('@')[0];
												iRespond += bfnSearchSel(strTmpSel,strTmp,2) ? 2:0;
											}
											if(strQTmpSel.IndexOf("選項三：") > 0)
											{
												strTmpSel = strQTmpSel.Replace('@','-').Replace("選項三：","@").Split('@')[1].Replace('\r','@').Split('@')[0];
												iRespond += bfnSearchSel(strTmpSel,strTmp,3) ? 4:0;
											}
											if(strQTmpSel.IndexOf("選項四：") > 0)
											{
												strTmpSel = strQTmpSel.Replace('@','-').Replace("選項四：","@").Split('@')[1].Replace('\r','@').Split('@')[0];
												iRespond += bfnSearchSel(strTmpSel,strTmp,4) ? 8:0;
											}			
											if(iRespond == 0 || iRespond == 15)
											{
												strTmpSel = strQTmpSel.Replace('@','-').Replace("選項一：","@").Split('@')[1].Replace('\r','@').Split('@')[0];
												iRespond = bfnSearchSelX(strTmpSel,strTmp,"1") ? 1:iRespond;
												strTmpSel = strQTmpSel.Replace('@','-').Replace("選項二：","@").Split('@')[1].Replace('\r','@').Split('@')[0];
												iRespond = bfnSearchSelX(strTmpSel,strTmp,"2") ? 2:iRespond;
												strTmpSel = strQTmpSel.Replace('@','-').Replace("選項三：","@").Split('@')[1].Replace('\r','@').Split('@')[0];
												iRespond = bfnSearchSelX(strTmpSel,strTmp,"3") ? 4:iRespond;
												strTmpSel = strQTmpSel.Replace('@','-').Replace("選項四：","@").Split('@')[1].Replace('\r','@').Split('@')[0];
												iRespond = bfnSearchSelX(strTmpSel,strTmp,"4") ? 8:iRespond;
												
											}
											if(iRespond == 0 || iRespond == 15)
											{
												strTmpAns = strQTmpSel.Replace("機智問答進行至第","@").Split('@')[0];
												if(strTmpAns.IndexOf("選項1：") > 0)
												{
													strTmpSel = strTmpAns.Replace("選項1：","@").Split('@')[1].Replace('\r','@').Split('@')[0];
													if(!bfnSearchSelX(strTmpSel,strTmp,"1"))
													fnAutoUpdate("1",strTmpSel,strTmp);
												}
												if(strTmpAns.IndexOf("選項2：") > 0)
												{
													strTmpSel = strTmpAns.Replace("選項2：","@").Split('@')[1].Replace('\r','@').Split('@')[0];
													if(!bfnSearchSelX(strTmpSel,strTmp,"2"))
													fnAutoUpdate("2",strTmpSel,strTmp);
												}
												if(strTmpAns.IndexOf("選項3：") > 0)
												{
													strTmpSel = strTmpAns.Replace("選項3：","@").Split('@')[1].Replace('\r','@').Split('@')[0];
													if(!bfnSearchSelX(strTmpSel,strTmp,"3"))
													fnAutoUpdate("3",strTmpSel,strTmp);
												}
												if(strTmpAns.IndexOf("選項4：") > 0)
												{
													strTmpSel = strTmpAns.Replace("選項4：","@").Split('@')[1].Replace('\r','@').Split('@')[0];
													if(!bfnSearchSelX(strTmpSel,strTmp,"4"))
													fnAutoUpdate("4",strTmpSel,strTmp);
												}
											}
													
											fnReAns(strStatues[iRespond],i-1);
										}

									}
								}


						}
					}
					fs.Close();		
					if(fileDir.Exists)
					{
						try
						{
							File.Delete(Application.StartupPath +@"\"+strName+".txt");
							this.label1.Visible = false;
						}
						catch(Exception e)
						{
							this.label1.Visible = true;

						}
					}
					else
					{
						MessageBox.Show(@"檔案刪除錯誤");
					}
					

				}
				else
				{
//					MessageBox.Show("ERROR,找不到"+Application.StartupPath +@"\"+strName+@".txt檔案,請於遊戲中重新輸入/report");
				}
				return true;
			}
			catch(Exception e)
			{
				this.label3.Text = e.Message;
				this.label3.Visible = true;
//				MessageBox.Show(e.Message,"Error",MessageBoxButtons.OK,MessageBoxIcon.Error);
				return false;
			}
		}

		private void btnTest_Click(object sender, System.EventArgs e)
		{
			FileInfo fileDir = new FileInfo(Application.StartupPath +@"\"+this.txtName.Text+".txt");
			if(fileDir.Exists)
			{
				this.txtName.Enabled = false;
				this.btnAns.Enabled = true;
				MessageBox.Show(@"檔案正確,解題開始");
			}
			else
			{
                MessageBox.Show(@"檔案錯誤");
				this.btnAns.Enabled = false;
			}
		
		}

		private void btnAns_Click(object sender, System.EventArgs e)
		{
			this.tmeExAns.Enabled = !this.tmeExAns.Enabled;
			this.txtName.Enabled = this.tmeExAns.Enabled ? this.txtName.Enabled:!this.txtName.Enabled;
			this.btnAns.Text = !this.tmeExAns.Enabled ? "開始解題":"結束解題";
			this.btnTest.Enabled = !this.tmeExAns.Enabled;

		}

//		private void dataGrid1_CurrentCellChanged(object sender, System.EventArgs e)
//		{
//			int iRownr=this.dataGrid1.CurrentCell.RowNumber;
//			object cellvalue1=this.dataGrid1[iRownr, 0];
//			fnReAns(cellvalue1.ToString(),this.cmbNum.SelectedIndex);
//
//		}

		private void btnStrClear_Click(object sender, System.EventArgs e)
		{
			this.txtLike.Text ="";
		}

		private void btnAnsClear_Click(object sender, System.EventArgs e)
		{
			this.txtNo1.Text = "";
			this.txtNo2.Text = "";
			this.txtNo3.Text = "";
			this.txtNo4.Text = "";
			this.txtNo5.Text = "";
			this.txtNo6.Text = "";
			this.txtNo7.Text = "";			
		}

		private void btnUpdate_Click(object sender, System.EventArgs e)
		{
			fnUpdate();		
		}

		private void tmeExAns_Tick(object sender, System.EventArgs e)
		{
			this.iLeng = 1;
			fnCheckAns(this.txtName.Text);		

		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			string strComm = null;

			strComm = @"DELETE FROM XData" ;
			try 
			{
				this.oleconn.Open();
				oleCmd = new OleDbCommand(strComm, this.oleconn);
				oleCmd.ExecuteNonQuery();
				MessageBox.Show( "清除成功!!");

			}
			catch( Exception ex ) 
			{
				MessageBox.Show( ex.Message+strComm );
			}
			this.oleconn.Close();
		}


	}
}
