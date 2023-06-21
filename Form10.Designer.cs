namespace emedit
{
    partial class Form10
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.enddt = new System.Windows.Forms.TextBox();
            this.startdt = new System.Windows.Forms.TextBox();
            this.reqnm = new System.Windows.Forms.TextBox();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.taskdetid = new System.Windows.Forms.TextBox();
            this.taskid = new System.Windows.Forms.TextBox();
            this.method = new System.Windows.Forms.ComboBox();
            this.state = new System.Windows.Forms.ComboBox();
            this.turnatm = new System.Windows.Forms.TextBox();
            this.endtm = new System.Windows.Forms.TextBox();
            this.starttm = new System.Windows.Forms.TextBox();
            this.usrid = new System.Windows.Forms.TextBox();
            this.resultinfo = new System.Windows.Forms.TextBox();
            this.head = new System.Windows.Forms.TextBox();
            this.cancelbutton = new System.Windows.Forms.Button();
            this.replybutton = new System.Windows.Forms.Button();
            this.wrtdt = new System.Windows.Forms.TextBox();
            this.wrtusrid = new System.Windows.Forms.TextBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.label18 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.label17 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.label16 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.closebtn = new System.Windows.Forms.RadioButton();
            this.openbtn = new System.Windows.Forms.RadioButton();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // enddt
            // 
            this.enddt.Location = new System.Drawing.Point(1299, 259);
            this.enddt.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.enddt.Name = "enddt";
            this.enddt.Size = new System.Drawing.Size(138, 25);
            this.enddt.TabIndex = 321;
            this.enddt.Enter += new System.EventHandler(this.placeholder);
            this.enddt.Leave += new System.EventHandler(this.leaveplaceholder);
            // 
            // startdt
            // 
            this.startdt.Location = new System.Drawing.Point(1299, 170);
            this.startdt.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.startdt.Name = "startdt";
            this.startdt.Size = new System.Drawing.Size(138, 25);
            this.startdt.TabIndex = 319;
            this.startdt.Enter += new System.EventHandler(this.placeholder);
            this.startdt.Leave += new System.EventHandler(this.leaveplaceholder);
            // 
            // reqnm
            // 
            this.reqnm.Location = new System.Drawing.Point(899, 122);
            this.reqnm.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.reqnm.Name = "reqnm";
            this.reqnm.ReadOnly = true;
            this.reqnm.Size = new System.Drawing.Size(116, 25);
            this.reqnm.TabIndex = 317;
            // 
            // dataGridView2
            // 
            this.dataGridView2.AllowUserToAddRows = false;
            this.dataGridView2.AllowUserToResizeColumns = false;
            this.dataGridView2.AllowUserToResizeRows = false;
            this.dataGridView2.BackgroundColor = System.Drawing.SystemColors.ButtonHighlight;
            this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView2.Location = new System.Drawing.Point(12, 280);
            this.dataGridView2.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.ReadOnly = true;
            this.dataGridView2.RowHeadersVisible = false;
            this.dataGridView2.RowHeadersWidth = 51;
            this.dataGridView2.RowTemplate.Height = 27;
            this.dataGridView2.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView2.Size = new System.Drawing.Size(752, 416);
            this.dataGridView2.TabIndex = 315;
            // 
            // taskdetid
            // 
            this.taskdetid.Location = new System.Drawing.Point(900, 78);
            this.taskdetid.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.taskdetid.Name = "taskdetid";
            this.taskdetid.Size = new System.Drawing.Size(115, 25);
            this.taskdetid.TabIndex = 313;
            // 
            // taskid
            // 
            this.taskid.Location = new System.Drawing.Point(899, 35);
            this.taskid.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.taskid.Name = "taskid";
            this.taskid.ReadOnly = true;
            this.taskid.Size = new System.Drawing.Size(115, 25);
            this.taskid.TabIndex = 310;
            // 
            // method
            // 
            this.method.FormattingEnabled = true;
            this.method.Location = new System.Drawing.Point(1293, 35);
            this.method.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.method.Name = "method";
            this.method.Size = new System.Drawing.Size(149, 23);
            this.method.TabIndex = 308;
            // 
            // state
            // 
            this.state.FormattingEnabled = true;
            this.state.Location = new System.Drawing.Point(1299, 78);
            this.state.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.state.Name = "state";
            this.state.Size = new System.Drawing.Size(138, 23);
            this.state.TabIndex = 307;
            // 
            // turnatm
            // 
            this.turnatm.Location = new System.Drawing.Point(898, 211);
            this.turnatm.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.turnatm.Name = "turnatm";
            this.turnatm.Size = new System.Drawing.Size(117, 25);
            this.turnatm.TabIndex = 305;
            this.turnatm.Enter += new System.EventHandler(this.placeholder);
            this.turnatm.Leave += new System.EventHandler(this.leaveplaceholder);
            // 
            // endtm
            // 
            this.endtm.Location = new System.Drawing.Point(1299, 214);
            this.endtm.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.endtm.Name = "endtm";
            this.endtm.Size = new System.Drawing.Size(138, 25);
            this.endtm.TabIndex = 302;
            this.endtm.Enter += new System.EventHandler(this.placeholder);
            this.endtm.Leave += new System.EventHandler(this.leaveplaceholder);
            // 
            // starttm
            // 
            this.starttm.Location = new System.Drawing.Point(1299, 125);
            this.starttm.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.starttm.Name = "starttm";
            this.starttm.Size = new System.Drawing.Size(138, 25);
            this.starttm.TabIndex = 301;
            this.starttm.Enter += new System.EventHandler(this.placeholder);
            this.starttm.Leave += new System.EventHandler(this.leaveplaceholder);
            // 
            // usrid
            // 
            this.usrid.Location = new System.Drawing.Point(899, 170);
            this.usrid.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.usrid.Name = "usrid";
            this.usrid.Size = new System.Drawing.Size(116, 25);
            this.usrid.TabIndex = 295;
            // 
            // resultinfo
            // 
            this.resultinfo.Location = new System.Drawing.Point(899, 375);
            this.resultinfo.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.resultinfo.Multiline = true;
            this.resultinfo.Name = "resultinfo";
            this.resultinfo.Size = new System.Drawing.Size(541, 321);
            this.resultinfo.TabIndex = 292;
            // 
            // head
            // 
            this.head.Location = new System.Drawing.Point(899, 340);
            this.head.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.head.Name = "head";
            this.head.Size = new System.Drawing.Size(541, 25);
            this.head.TabIndex = 291;
            // 
            // cancelbutton
            // 
            this.cancelbutton.Location = new System.Drawing.Point(1351, 700);
            this.cancelbutton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.cancelbutton.Name = "cancelbutton";
            this.cancelbutton.Size = new System.Drawing.Size(86, 34);
            this.cancelbutton.TabIndex = 289;
            this.cancelbutton.Text = "취소";
            this.cancelbutton.UseVisualStyleBackColor = true;
            this.cancelbutton.Click += new System.EventHandler(this.cancelbutton_Click);
            // 
            // replybutton
            // 
            this.replybutton.Location = new System.Drawing.Point(1259, 700);
            this.replybutton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.replybutton.Name = "replybutton";
            this.replybutton.Size = new System.Drawing.Size(86, 34);
            this.replybutton.TabIndex = 288;
            this.replybutton.Text = "등록";
            this.replybutton.UseVisualStyleBackColor = true;
            this.replybutton.Click += new System.EventHandler(this.replybutton_Click);
            // 
            // wrtdt
            // 
            this.wrtdt.Location = new System.Drawing.Point(900, 258);
            this.wrtdt.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.wrtdt.Name = "wrtdt";
            this.wrtdt.ReadOnly = true;
            this.wrtdt.Size = new System.Drawing.Size(115, 25);
            this.wrtdt.TabIndex = 287;
            // 
            // wrtusrid
            // 
            this.wrtusrid.Location = new System.Drawing.Point(898, 304);
            this.wrtusrid.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.wrtusrid.Name = "wrtusrid";
            this.wrtusrid.ReadOnly = true;
            this.wrtusrid.Size = new System.Drawing.Size(117, 25);
            this.wrtusrid.TabIndex = 286;
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToResizeColumns = false;
            this.dataGridView1.AllowUserToResizeRows = false;
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView1.BackgroundColor = System.Drawing.SystemColors.ButtonHighlight;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.GridColor = System.Drawing.SystemColors.ButtonShadow;
            this.dataGridView1.Location = new System.Drawing.Point(12, 41);
            this.dataGridView1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowHeadersVisible = false;
            this.dataGridView1.RowHeadersWidth = 51;
            this.dataGridView1.RowTemplate.Height = 27;
            this.dataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView1.Size = new System.Drawing.Size(755, 201);
            this.dataGridView1.TabIndex = 283;
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            // 
            // label18
            // 
            this.label18.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.label18.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.label18.Location = new System.Drawing.Point(805, 32);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(87, 28);
            this.label18.TabIndex = 322;
            this.label18.Text = "태스크ID";
            this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.label18.Click += new System.EventHandler(this.label18_Click);
            // 
            // label5
            // 
            this.label5.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.label5.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.label5.Location = new System.Drawing.Point(1173, 35);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(87, 28);
            this.label5.TabIndex = 323;
            this.label5.Text = "작업방법";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.label3.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.label3.Location = new System.Drawing.Point(1176, 78);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(87, 28);
            this.label3.TabIndex = 324;
            this.label3.Text = "작업상태";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label8
            // 
            this.label8.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.label8.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.label8.Location = new System.Drawing.Point(1176, 125);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(104, 28);
            this.label8.TabIndex = 325;
            this.label8.Text = "작업시작시간";
            this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label6
            // 
            this.label6.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.label6.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.label6.Location = new System.Drawing.Point(1176, 214);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(104, 28);
            this.label6.TabIndex = 326;
            this.label6.Text = "작업종료시간";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.label2.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.label2.Location = new System.Drawing.Point(773, 78);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(120, 28);
            this.label2.TabIndex = 327;
            this.label2.Text = "상세태스크ID";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label11
            // 
            this.label11.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.label11.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.label11.Location = new System.Drawing.Point(805, 121);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(87, 28);
            this.label11.TabIndex = 328;
            this.label11.Text = "요청자명";
            this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label14
            // 
            this.label14.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.label14.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.label14.Location = new System.Drawing.Point(805, 170);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(87, 28);
            this.label14.TabIndex = 329;
            this.label14.Text = "작업자명";
            this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label9
            // 
            this.label9.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.label9.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.label9.Location = new System.Drawing.Point(1176, 170);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(104, 28);
            this.label9.TabIndex = 330;
            this.label9.Text = "작업시작일자";
            this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label15
            // 
            this.label15.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.label15.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.label15.Location = new System.Drawing.Point(1176, 255);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(104, 28);
            this.label15.TabIndex = 331;
            this.label15.Text = "작업종료일자";
            this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label17
            // 
            this.label17.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.label17.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.label17.Location = new System.Drawing.Point(805, 214);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(87, 28);
            this.label17.TabIndex = 332;
            this.label17.Text = "소요시간";
            this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label10
            // 
            this.label10.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.label10.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.label10.Location = new System.Drawing.Point(806, 256);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(87, 28);
            this.label10.TabIndex = 333;
            this.label10.Text = "작성일자";
            this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label12
            // 
            this.label12.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.label12.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.label12.Location = new System.Drawing.Point(805, 300);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(87, 28);
            this.label12.TabIndex = 334;
            this.label12.Text = "작성자 ID";
            this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label16
            // 
            this.label16.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.label16.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.label16.Location = new System.Drawing.Point(805, 340);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(88, 28);
            this.label16.TabIndex = 335;
            this.label16.Text = "제목";
            this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label7
            // 
            this.label7.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.label7.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.label7.Location = new System.Drawing.Point(805, 384);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(87, 28);
            this.label7.TabIndex = 336;
            this.label7.Text = "작업결과";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label4
            // 
            this.label4.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.label4.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.label4.Location = new System.Drawing.Point(9, 11);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(87, 28);
            this.label4.TabIndex = 337;
            this.label4.Text = "태스크정보";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.label1.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.label1.Location = new System.Drawing.Point(9, 250);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(117, 28);
            this.label1.TabIndex = 338;
            this.label1.Text = "작업결과정보";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label13
            // 
            this.label13.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.label13.Font = new System.Drawing.Font("굴림", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.label13.Location = new System.Drawing.Point(5, 14);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(84, 25);
            this.label13.TabIndex = 339;
            this.label13.Text = "공개방식";
            this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.closebtn);
            this.groupBox1.Controls.Add(this.openbtn);
            this.groupBox1.Controls.Add(this.label13);
            this.groupBox1.Location = new System.Drawing.Point(1171, 286);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(266, 46);
            this.groupBox1.TabIndex = 340;
            this.groupBox1.TabStop = false;
            // 
            // closebtn
            // 
            this.closebtn.AutoSize = true;
            this.closebtn.Location = new System.Drawing.Point(160, 20);
            this.closebtn.Name = "closebtn";
            this.closebtn.Size = new System.Drawing.Size(88, 19);
            this.closebtn.TabIndex = 341;
            this.closebtn.TabStop = true;
            this.closebtn.Text = "공개금지";
            this.closebtn.UseVisualStyleBackColor = true;
            // 
            // openbtn
            // 
            this.openbtn.AutoSize = true;
            this.openbtn.Location = new System.Drawing.Point(96, 20);
            this.openbtn.Name = "openbtn";
            this.openbtn.Size = new System.Drawing.Size(58, 19);
            this.openbtn.TabIndex = 340;
            this.openbtn.TabStop = true;
            this.openbtn.Text = "공개";
            this.openbtn.UseVisualStyleBackColor = true;
            // 
            // Form10
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(1470, 743);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label16);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.label17);
            this.Controls.Add(this.label15);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label14);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label18);
            this.Controls.Add(this.enddt);
            this.Controls.Add(this.startdt);
            this.Controls.Add(this.reqnm);
            this.Controls.Add(this.dataGridView2);
            this.Controls.Add(this.taskdetid);
            this.Controls.Add(this.taskid);
            this.Controls.Add(this.method);
            this.Controls.Add(this.state);
            this.Controls.Add(this.turnatm);
            this.Controls.Add(this.endtm);
            this.Controls.Add(this.starttm);
            this.Controls.Add(this.usrid);
            this.Controls.Add(this.resultinfo);
            this.Controls.Add(this.head);
            this.Controls.Add(this.cancelbutton);
            this.Controls.Add(this.replybutton);
            this.Controls.Add(this.wrtdt);
            this.Controls.Add(this.wrtusrid);
            this.Controls.Add(this.dataGridView1);
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "Form10";
            this.Text = "태스크 상세 등록";
            this.Load += new System.EventHandler(this.Form10_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox enddt;
        private System.Windows.Forms.TextBox startdt;
        private System.Windows.Forms.TextBox reqnm;
        private System.Windows.Forms.DataGridView dataGridView2;
        private System.Windows.Forms.TextBox taskdetid;
        private System.Windows.Forms.TextBox taskid;
        private System.Windows.Forms.ComboBox method;
        private System.Windows.Forms.ComboBox state;
        private System.Windows.Forms.TextBox turnatm;
        private System.Windows.Forms.TextBox endtm;
        private System.Windows.Forms.TextBox starttm;
        private System.Windows.Forms.TextBox usrid;
        private System.Windows.Forms.TextBox resultinfo;
        private System.Windows.Forms.TextBox head;
        private System.Windows.Forms.Button cancelbutton;
        private System.Windows.Forms.Button replybutton;
        private System.Windows.Forms.TextBox wrtdt;
        private System.Windows.Forms.TextBox wrtusrid;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton closebtn;
        private System.Windows.Forms.RadioButton openbtn;
    }
}