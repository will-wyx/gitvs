namespace ExpressMaster
{
    partial class FormMain
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle8 = new System.Windows.Forms.DataGridViewCellStyle();
            this.cbxMain = new System.Windows.Forms.ComboBox();
            this.dgvMain = new System.Windows.Forms.DataGridView();
            this.colKey = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colFirstWeight = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colFirstAmount = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colFirstAmountB = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colOtherAmount = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cmsMain = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.tsmiRemove = new System.Windows.Forms.ToolStripMenuItem();
            this.ssMain = new System.Windows.Forms.StatusStrip();
            this.tsslInfo = new System.Windows.Forms.ToolStripStatusLabel();
            this.btnExport = new System.Windows.Forms.Button();
            this.btnAddRow = new System.Windows.Forms.Button();
            this.cbxMinor = new System.Windows.Forms.ComboBox();
            this.btnAddItem = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgvMain)).BeginInit();
            this.cmsMain.SuspendLayout();
            this.SuspendLayout();
            // 
            // cbxMain
            // 
            this.cbxMain.FormattingEnabled = true;
            this.cbxMain.Location = new System.Drawing.Point(12, 7);
            this.cbxMain.Name = "cbxMain";
            this.cbxMain.Size = new System.Drawing.Size(186, 20);
            this.cbxMain.TabIndex = 0;
            this.cbxMain.SelectedIndexChanged += new System.EventHandler(this.cbxMain_SelectedIndexChanged);
            // 
            // dgvMain
            // 
            this.dgvMain.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvMain.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvMain.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colKey,
            this.colFirstWeight,
            this.colFirstAmount,
            this.colFirstAmountB,
            this.colOtherAmount});
            this.dgvMain.ContextMenuStrip = this.cmsMain;
            this.dgvMain.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.dgvMain.Location = new System.Drawing.Point(12, 62);
            this.dgvMain.Name = "dgvMain";
            this.dgvMain.RowTemplate.Height = 23;
            this.dgvMain.Size = new System.Drawing.Size(659, 380);
            this.dgvMain.TabIndex = 1;
            // 
            // colKey
            // 
            this.colKey.DataPropertyName = "Key";
            this.colKey.FillWeight = 86.88244F;
            this.colKey.HeaderText = "地区关键字";
            this.colKey.Name = "colKey";
            // 
            // colFirstWeight
            // 
            this.colFirstWeight.DataPropertyName = "FirstWeight";
            dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle5.Format = "N2";
            dataGridViewCellStyle5.NullValue = "0";
            this.colFirstWeight.DefaultCellStyle = dataGridViewCellStyle5;
            this.colFirstWeight.FillWeight = 86.88244F;
            this.colFirstWeight.HeaderText = "首重重量";
            this.colFirstWeight.Name = "colFirstWeight";
            // 
            // colFirstAmount
            // 
            this.colFirstAmount.DataPropertyName = "FirstAmount";
            dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle6.Format = "C2";
            dataGridViewCellStyle6.NullValue = "0";
            this.colFirstAmount.DefaultCellStyle = dataGridViewCellStyle6;
            this.colFirstAmount.FillWeight = 121.3103F;
            this.colFirstAmount.HeaderText = "小件首重金额";
            this.colFirstAmount.Name = "colFirstAmount";
            // 
            // colFirstAmountB
            // 
            this.colFirstAmountB.DataPropertyName = "FirstAmountB";
            dataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle7.Format = "C2";
            dataGridViewCellStyle7.NullValue = "0";
            this.colFirstAmountB.DefaultCellStyle = dataGridViewCellStyle7;
            this.colFirstAmountB.FillWeight = 121.3103F;
            this.colFirstAmountB.HeaderText = "大件首重金额";
            this.colFirstAmountB.Name = "colFirstAmountB";
            // 
            // colOtherAmount
            // 
            this.colOtherAmount.DataPropertyName = "OtherAmount";
            dataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle8.Format = "C2";
            dataGridViewCellStyle8.NullValue = "0";
            this.colOtherAmount.DefaultCellStyle = dataGridViewCellStyle8;
            this.colOtherAmount.FillWeight = 86.88244F;
            this.colOtherAmount.HeaderText = "续重金额";
            this.colOtherAmount.Name = "colOtherAmount";
            // 
            // cmsMain
            // 
            this.cmsMain.ImageScalingSize = new System.Drawing.Size(24, 24);
            this.cmsMain.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmiRemove});
            this.cmsMain.Name = "cmsMain";
            this.cmsMain.Size = new System.Drawing.Size(101, 26);
            this.cmsMain.Opening += new System.ComponentModel.CancelEventHandler(this.cmsMain_Opening);
            // 
            // tsmiRemove
            // 
            this.tsmiRemove.Name = "tsmiRemove";
            this.tsmiRemove.Size = new System.Drawing.Size(100, 22);
            this.tsmiRemove.Text = "删除";
            this.tsmiRemove.Click += new System.EventHandler(this.tsmiRemove_Click);
            // 
            // ssMain
            // 
            this.ssMain.ImageScalingSize = new System.Drawing.Size(24, 24);
            this.ssMain.Location = new System.Drawing.Point(0, 445);
            this.ssMain.Name = "ssMain";
            this.ssMain.Size = new System.Drawing.Size(683, 22);
            this.ssMain.TabIndex = 2;
            this.ssMain.Text = "statusStrip1";
            // 
            // tsslInfo
            // 
            this.tsslInfo.Name = "tsslInfo";
            this.tsslInfo.Size = new System.Drawing.Size(52, 17);
            this.tsslInfo.Text = "Normal";
            // 
            // btnExport
            // 
            this.btnExport.Location = new System.Drawing.Point(590, 7);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(81, 23);
            this.btnExport.TabIndex = 3;
            this.btnExport.Text = "导出";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // btnAddRow
            // 
            this.btnAddRow.Location = new System.Drawing.Point(12, 33);
            this.btnAddRow.Name = "btnAddRow";
            this.btnAddRow.Size = new System.Drawing.Size(75, 23);
            this.btnAddRow.TabIndex = 4;
            this.btnAddRow.Text = "加入行";
            this.btnAddRow.UseVisualStyleBackColor = true;
            this.btnAddRow.Click += new System.EventHandler(this.btnAddRow_Click);
            // 
            // cbxMinor
            // 
            this.cbxMinor.FormattingEnabled = true;
            this.cbxMinor.Location = new System.Drawing.Point(204, 7);
            this.cbxMinor.Name = "cbxMinor";
            this.cbxMinor.Size = new System.Drawing.Size(199, 20);
            this.cbxMinor.TabIndex = 5;
            this.cbxMinor.SelectedIndexChanged += new System.EventHandler(this.cbxMinor_SelectedIndexChanged);
            // 
            // btnAddItem
            // 
            this.btnAddItem.Location = new System.Drawing.Point(409, 6);
            this.btnAddItem.Name = "btnAddItem";
            this.btnAddItem.Size = new System.Drawing.Size(75, 23);
            this.btnAddItem.TabIndex = 6;
            this.btnAddItem.Text = "加项";
            this.btnAddItem.UseVisualStyleBackColor = true;
            this.btnAddItem.Click += new System.EventHandler(this.btnAddItem_Click);
            // 
            // FormMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(683, 467);
            this.Controls.Add(this.btnAddItem);
            this.Controls.Add(this.cbxMinor);
            this.Controls.Add(this.btnAddRow);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.ssMain);
            this.Controls.Add(this.dgvMain);
            this.Controls.Add(this.cbxMain);
            this.Name = "FormMain";
            this.Text = "Main";
            this.Load += new System.EventHandler(this.FormMain_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvMain)).EndInit();
            this.cmsMain.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox cbxMain;
        private System.Windows.Forms.DataGridView dgvMain;
        private System.Windows.Forms.StatusStrip ssMain;
        private System.Windows.Forms.ToolStripStatusLabel tsslInfo;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.Button btnAddRow;
        private System.Windows.Forms.ContextMenuStrip cmsMain;
        private System.Windows.Forms.ToolStripMenuItem tsmiRemove;
        private System.Windows.Forms.ComboBox cbxMinor;
        private System.Windows.Forms.DataGridViewTextBoxColumn colKey;
        private System.Windows.Forms.DataGridViewTextBoxColumn colFirstWeight;
        private System.Windows.Forms.DataGridViewTextBoxColumn colFirstAmount;
        private System.Windows.Forms.DataGridViewTextBoxColumn colFirstAmountB;
        private System.Windows.Forms.DataGridViewTextBoxColumn colOtherAmount;
        private System.Windows.Forms.Button btnAddItem;
    }
}

