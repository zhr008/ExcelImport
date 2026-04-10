namespace ExcelImport;

partial class Form1
{
    /// <summary>
    ///  Required designer variable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;
    private DataGridView dgvTemplates = null!;
    private CheckBox chkStartWithWindows = null!;
    private Button btnSave = null!;
    private Button btnAdd = null!;
    private Button btnDelete = null!;
    private Button btnSelectFolder = null!;
    private Button btnSelect = null!;
    private Button btnRunNow = null!;
    private Button btnRefreshLog = null!;
    private Button btnClearLog = null!;
    private TextBox txtLogs = null!;

    /// <summary>
    ///  Clean up any resources being used.
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
    ///  Required method for Designer support - do not modify
    ///  the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent()
    {
        System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
        dgvTemplates = new DataGridView();
        dataGridViewCheckBoxColumn1 = new DataGridViewCheckBoxColumn();
        dataGridViewTextBoxColumn1 = new DataGridViewTextBoxColumn();
        dataGridViewTextBoxColumn2 = new DataGridViewTextBoxColumn();
        dataGridViewCheckBoxColumn2 = new DataGridViewCheckBoxColumn();
        dataGridViewTextBoxColumn3 = new DataGridViewTextBoxColumn();
        dataGridViewTextBoxColumn4 = new DataGridViewTextBoxColumn();
        dataGridViewTextBoxColumn5 = new DataGridViewTextBoxColumn();
        dataGridViewTextBoxColumn6 = new DataGridViewTextBoxColumn();
        chkStartWithWindows = new CheckBox();
        btnSave = new Button();
        btnAdd = new Button();
        btnDelete = new Button();
        btnSelectFolder = new Button();
        btnSelect = new Button();
        btnRunNow = new Button();
        btnRefreshLog = new Button();
        btnClearLog = new Button();
        txtLogs = new TextBox();
        ((System.ComponentModel.ISupportInitialize)dgvTemplates).BeginInit();
        SuspendLayout();
        // 
        // dgvTemplates
        // 
        dgvTemplates.AllowUserToAddRows = false;
        dgvTemplates.AllowUserToDeleteRows = false;
        dgvTemplates.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        dgvTemplates.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        dgvTemplates.Columns.AddRange(new DataGridViewColumn[] { dataGridViewCheckBoxColumn1, dataGridViewTextBoxColumn1, dataGridViewTextBoxColumn2, dataGridViewCheckBoxColumn2, dataGridViewTextBoxColumn3, dataGridViewTextBoxColumn4, dataGridViewTextBoxColumn5, dataGridViewTextBoxColumn6 });
        dgvTemplates.Location = new Point(8, 11);
        dgvTemplates.Margin = new Padding(2);
        dgvTemplates.Name = "dgvTemplates";
        dgvTemplates.RowHeadersWidth = 62;
        dgvTemplates.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        dgvTemplates.Size = new Size(887, 255);
        dgvTemplates.TabIndex = 0;
        // 
        // dataGridViewCheckBoxColumn1
        // 
        dataGridViewCheckBoxColumn1.DataPropertyName = "Enabled";
        dataGridViewCheckBoxColumn1.HeaderText = "启用";
        dataGridViewCheckBoxColumn1.Name = "dataGridViewCheckBoxColumn1";
        dataGridViewCheckBoxColumn1.Width = 60;
        // 
        // dataGridViewTextBoxColumn1
        // 
        dataGridViewTextBoxColumn1.DataPropertyName = "Name";
        dataGridViewTextBoxColumn1.HeaderText = "任务名称";
        dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
        dataGridViewTextBoxColumn1.Width = 120;
        // 
        // dataGridViewTextBoxColumn2
        // 
        dataGridViewTextBoxColumn2.DataPropertyName = "WatchPath";
        dataGridViewTextBoxColumn2.HeaderText = "监控目录";
        dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
        dataGridViewTextBoxColumn2.Width = 180;
        // 
        // dataGridViewCheckBoxColumn2
        // 
        dataGridViewCheckBoxColumn2.DataPropertyName = "IncludeSubdirectories";
        dataGridViewCheckBoxColumn2.HeaderText = "含子目录";
        dataGridViewCheckBoxColumn2.Name = "dataGridViewCheckBoxColumn2";
        dataGridViewCheckBoxColumn2.Width = 80;
        // 
        // dataGridViewTextBoxColumn3
        // 
        dataGridViewTextBoxColumn3.DataPropertyName = "IntervalMinutes";
        dataGridViewTextBoxColumn3.HeaderText = "间隔(分钟)";
        dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
        dataGridViewTextBoxColumn3.Width = 90;
        // 
        // dataGridViewTextBoxColumn4
        // 
        dataGridViewTextBoxColumn4.DataPropertyName = "TemplateFile";
        dataGridViewTextBoxColumn4.HeaderText = "模板文件";
        dataGridViewTextBoxColumn4.Name = "dataGridViewTextBoxColumn4";
        dataGridViewTextBoxColumn4.Width = 140;
        // 
        // dataGridViewTextBoxColumn5
        // 
        dataGridViewTextBoxColumn5.DataPropertyName = "TargetTable";
        dataGridViewTextBoxColumn5.HeaderText = "目标表";
        dataGridViewTextBoxColumn5.Name = "dataGridViewTextBoxColumn5";
        // 
        // dataGridViewTextBoxColumn6
        // 
        dataGridViewTextBoxColumn6.DataPropertyName = "FilePattern";
        dataGridViewTextBoxColumn6.HeaderText = "文件匹配";
        dataGridViewTextBoxColumn6.Name = "dataGridViewTextBoxColumn6";
        dataGridViewTextBoxColumn6.Width = 80;
        // 
        // chkStartWithWindows
        // 
        chkStartWithWindows.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        chkStartWithWindows.AutoSize = true;
        chkStartWithWindows.Location = new Point(809, 278);
        chkStartWithWindows.Margin = new Padding(2);
        chkStartWithWindows.Name = "chkStartWithWindows";
        chkStartWithWindows.Size = new Size(75, 21);
        chkStartWithWindows.TabIndex = 1;
        chkStartWithWindows.Text = "开机启动";
        chkStartWithWindows.UseVisualStyleBackColor = true;
        // 
        // btnSave
        // 
        btnSave.Location = new Point(10, 275);
        btnSave.Margin = new Padding(2);
        btnSave.Name = "btnSave";
        btnSave.Size = new Size(78, 24);
        btnSave.TabIndex = 3;
        btnSave.Text = "保存配置";
        btnSave.UseVisualStyleBackColor = true;
        btnSave.Click += btnSave_Click;
        // 
        // btnAdd
        // 
        btnAdd.Location = new Point(93, 275);
        btnAdd.Margin = new Padding(2);
        btnAdd.Name = "btnAdd";
        btnAdd.Size = new Size(78, 24);
        btnAdd.TabIndex = 4;
        btnAdd.Text = "新增任务";
        btnAdd.UseVisualStyleBackColor = true;
        btnAdd.Click += btnAdd_Click;
        // 
        // btnDelete
        // 
        btnDelete.Location = new Point(176, 275);
        btnDelete.Margin = new Padding(2);
        btnDelete.Name = "btnDelete";
        btnDelete.Size = new Size(78, 24);
        btnDelete.TabIndex = 5;
        btnDelete.Text = "删除任务";
        btnDelete.UseVisualStyleBackColor = true;
        btnDelete.Click += btnDelete_Click;
        // 
        // btnSelectFolder
        // 
        btnSelectFolder.Location = new Point(258, 275);
        btnSelectFolder.Margin = new Padding(2);
        btnSelectFolder.Name = "btnSelectFolder";
        btnSelectFolder.Size = new Size(78, 24);
        btnSelectFolder.TabIndex = 6;
        btnSelectFolder.Text = "选择目录";
        btnSelectFolder.UseVisualStyleBackColor = true;
        btnSelectFolder.Click += btnSelectFolder_Click;
        // 
        // btnSelect
        // 
        btnSelect.Location = new Point(341, 275);
        btnSelect.Margin = new Padding(2);
        btnSelect.Name = "btnSelect";
        btnSelect.Size = new Size(98, 24);
        btnSelect.TabIndex = 7;
        btnSelect.Text = "选择模板";
        btnSelect.UseVisualStyleBackColor = true;
        btnSelect.Click += btnSelect_Click;
        // 
        // btnRunNow
        // 
        btnRunNow.Location = new Point(443, 275);
        btnRunNow.Margin = new Padding(2);
        btnRunNow.Name = "btnRunNow";
        btnRunNow.Size = new Size(78, 24);
        btnRunNow.TabIndex = 8;
        btnRunNow.Text = "立即执行";
        btnRunNow.UseVisualStyleBackColor = true;
        btnRunNow.Click += btnRunNow_Click;
        // 
        // btnRefreshLog
        // 
        btnRefreshLog.Location = new Point(526, 275);
        btnRefreshLog.Margin = new Padding(2);
        btnRefreshLog.Name = "btnRefreshLog";
        btnRefreshLog.Size = new Size(78, 24);
        btnRefreshLog.TabIndex = 9;
        btnRefreshLog.Text = "刷新日志";
        btnRefreshLog.UseVisualStyleBackColor = true;
        btnRefreshLog.Click += btnRefreshLog_Click;
        // 
        // btnClearLog
        // 
        btnClearLog.Location = new Point(608, 275);
        btnClearLog.Margin = new Padding(2);
        btnClearLog.Name = "btnClearLog";
        btnClearLog.Size = new Size(78, 24);
        btnClearLog.TabIndex = 10;
        btnClearLog.Text = "清空日志";
        btnClearLog.UseVisualStyleBackColor = true;
        btnClearLog.Click += btnClearLog_Click;
        // 
        // txtLogs
        // 
        txtLogs.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        txtLogs.Location = new Point(8, 313);
        txtLogs.Margin = new Padding(2);
        txtLogs.Multiline = true;
        txtLogs.Name = "txtLogs";
        txtLogs.ReadOnly = true;
        txtLogs.ScrollBars = ScrollBars.Vertical;
        txtLogs.Size = new Size(888, 173);
        txtLogs.TabIndex = 10;
        // 
        // Form1
        // 
        AutoScaleDimensions = new SizeF(7F, 17F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(904, 501);
        Controls.Add(txtLogs);
        Controls.Add(btnClearLog);
        Controls.Add(btnRefreshLog);
        Controls.Add(btnRunNow);
        Controls.Add(btnSelect);
        Controls.Add(btnSelectFolder);
        Controls.Add(btnDelete);
        Controls.Add(btnAdd);
        Controls.Add(btnSave);
        Controls.Add(chkStartWithWindows);
        Controls.Add(dgvTemplates);
        Icon = (Icon)resources.GetObject("$this.Icon");
        Margin = new Padding(2);
        MinimumSize = new Size(845, 507);
        Name = "Form1";
        StartPosition = FormStartPosition.CenterScreen;
        Text = "Excel 导入工具";
        ((System.ComponentModel.ISupportInitialize)dgvTemplates).EndInit();
        ResumeLayout(false);
        PerformLayout();
    }

    #endregion

    private DataGridViewCheckBoxColumn dataGridViewCheckBoxColumn1;
    private DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
    private DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
    private DataGridViewCheckBoxColumn dataGridViewCheckBoxColumn2;
    private DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
    private DataGridViewTextBoxColumn dataGridViewTextBoxColumn4;
    private DataGridViewTextBoxColumn dataGridViewTextBoxColumn5;
    private DataGridViewTextBoxColumn dataGridViewTextBoxColumn6;
}
