namespace CheckedExcel
{
    partial class Form1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.btStart = new System.Windows.Forms.Button();
            this.tbPath = new System.Windows.Forms.TextBox();
            this.btDialog = new System.Windows.Forms.Button();
            this.tBinfo = new System.Windows.Forms.TextBox();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.rBizg = new System.Windows.Forms.RadioButton();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.rBimport = new System.Windows.Forms.RadioButton();
            this.rBformat = new System.Windows.Forms.RadioButton();
            this.rBpot = new System.Windows.Forms.RadioButton();
            this.button1 = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.lbInfo = new System.Windows.Forms.Label();
            this.folderBrowserDialog2 = new System.Windows.Forms.FolderBrowserDialog();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // btStart
            // 
            this.btStart.Location = new System.Drawing.Point(480, 244);
            this.btStart.Name = "btStart";
            this.btStart.Size = new System.Drawing.Size(140, 35);
            this.btStart.TabIndex = 0;
            this.btStart.Text = "Начать проверку";
            this.btStart.UseVisualStyleBackColor = true;
            this.btStart.Click += new System.EventHandler(this.btStart_Click);
            // 
            // tbPath
            // 
            this.tbPath.Location = new System.Drawing.Point(12, 12);
            this.tbPath.Multiline = true;
            this.tbPath.Name = "tbPath";
            this.tbPath.ReadOnly = true;
            this.tbPath.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.tbPath.Size = new System.Drawing.Size(308, 226);
            this.tbPath.TabIndex = 2;
            // 
            // btDialog
            // 
            this.btDialog.Location = new System.Drawing.Point(334, 244);
            this.btDialog.Name = "btDialog";
            this.btDialog.Size = new System.Drawing.Size(140, 35);
            this.btDialog.TabIndex = 3;
            this.btDialog.Text = "Задать путь к файлам Excel";
            this.btDialog.UseVisualStyleBackColor = true;
            this.btDialog.Click += new System.EventHandler(this.btDialog_Click);
            // 
            // tBinfo
            // 
            this.tBinfo.Location = new System.Drawing.Point(326, 12);
            this.tBinfo.Multiline = true;
            this.tBinfo.Name = "tBinfo";
            this.tBinfo.ReadOnly = true;
            this.tBinfo.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.tBinfo.Size = new System.Drawing.Size(440, 226);
            this.tBinfo.TabIndex = 4;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(334, 285);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(432, 23);
            this.progressBar1.Step = 1;
            this.progressBar1.TabIndex = 5;
            // 
            // rBizg
            // 
            this.rBizg.AutoSize = true;
            this.rBizg.Location = new System.Drawing.Point(16, 43);
            this.rBizg.Name = "rBizg";
            this.rBizg.Size = new System.Drawing.Size(113, 17);
            this.rBizg.TabIndex = 6;
            this.rBizg.TabStop = true;
            this.rBizg.Text = "По изготовителю";
            this.rBizg.UseVisualStyleBackColor = true;
            this.rBizg.CheckedChanged += new System.EventHandler(this.rBizg_CheckedChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.rBimport);
            this.groupBox1.Controls.Add(this.rBformat);
            this.groupBox1.Controls.Add(this.rBpot);
            this.groupBox1.Controls.Add(this.rBizg);
            this.groupBox1.Location = new System.Drawing.Point(12, 244);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(308, 68);
            this.groupBox1.TabIndex = 7;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Режим";
            // 
            // rBimport
            // 
            this.rBimport.Location = new System.Drawing.Point(136, 42);
            this.rBimport.Name = "rBimport";
            this.rBimport.Size = new System.Drawing.Size(120, 19);
            this.rBimport.TabIndex = 12;
            this.rBimport.TabStop = true;
            this.rBimport.Text = "Импорт в PDF";
            this.rBimport.UseVisualStyleBackColor = true;
            this.rBimport.CheckedChanged += new System.EventHandler(this.rBimport_CheckedChanged);
            // 
            // rBformat
            // 
            this.rBformat.Location = new System.Drawing.Point(136, 10);
            this.rBformat.Name = "rBformat";
            this.rBformat.Size = new System.Drawing.Size(166, 37);
            this.rBformat.TabIndex = 11;
            this.rBformat.TabStop = true;
            this.rBformat.Text = "Исходное форматирование";
            this.rBformat.UseVisualStyleBackColor = true;
            this.rBformat.CheckedChanged += new System.EventHandler(this.rBformat_CheckedChanged);
            // 
            // rBpot
            // 
            this.rBpot.AutoSize = true;
            this.rBpot.Location = new System.Drawing.Point(16, 20);
            this.rBpot.Name = "rBpot";
            this.rBpot.Size = new System.Drawing.Size(108, 17);
            this.rBpot.TabIndex = 7;
            this.rBpot.TabStop = true;
            this.rBpot.Text = "По потребителю";
            this.rBpot.UseVisualStyleBackColor = true;
            this.rBpot.CheckedChanged += new System.EventHandler(this.rBpot_CheckedChanged);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(626, 244);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(140, 35);
            this.button1.TabIndex = 8;
            this.button1.Text = "Выгрузка ошибок в Excel\r\n\r\n";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1,
            this.Column2,
            this.Column3,
            this.Column4});
            this.dataGridView1.Location = new System.Drawing.Point(12, 318);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.Size = new System.Drawing.Size(754, 244);
            this.dataGridView1.TabIndex = 9;
            // 
            // Column1
            // 
            this.Column1.HeaderText = "Файл";
            this.Column1.Name = "Column1";
            this.Column1.ReadOnly = true;
            this.Column1.Width = 300;
            // 
            // Column2
            // 
            this.Column2.HeaderText = "Строка";
            this.Column2.Name = "Column2";
            this.Column2.ReadOnly = true;
            this.Column2.Width = 50;
            // 
            // Column3
            // 
            this.Column3.HeaderText = "Столбец";
            this.Column3.Name = "Column3";
            this.Column3.ReadOnly = true;
            this.Column3.Width = 60;
            // 
            // Column4
            // 
            this.Column4.HeaderText = "Сообщение";
            this.Column4.Name = "Column4";
            this.Column4.ReadOnly = true;
            this.Column4.Width = 450;
            // 
            // lbInfo
            // 
            this.lbInfo.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lbInfo.AutoSize = true;
            this.lbInfo.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.lbInfo.Location = new System.Drawing.Point(383, 565);
            this.lbInfo.Name = "lbInfo";
            this.lbInfo.Size = new System.Drawing.Size(0, 15);
            this.lbInfo.TabIndex = 10;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(778, 589);
            this.Controls.Add(this.lbInfo);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.tBinfo);
            this.Controls.Add(this.btDialog);
            this.Controls.Add(this.tbPath);
            this.Controls.Add(this.btStart);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Инвентаризация 2020";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.TextBox tbPath;
        private System.Windows.Forms.Button btDialog;
        private System.Windows.Forms.TextBox tBinfo;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.RadioButton rBizg;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton rBpot;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column2;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column3;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column4;
        public System.Windows.Forms.Button btStart;
        private System.Windows.Forms.Label lbInfo;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog2;
        private System.Windows.Forms.RadioButton rBimport;
        private System.Windows.Forms.RadioButton rBformat;
    }
}

