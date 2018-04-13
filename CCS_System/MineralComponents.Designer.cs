namespace CCS_System
{
    partial class MineralComponents
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
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.仓号 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.MineralName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Cu = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Fe = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.S = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.SiO2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CaO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.MgO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Al2O3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Co = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.AllowUserToResizeColumns = false;
            this.dataGridView1.AllowUserToResizeRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.仓号,
            this.MineralName,
            this.Cu,
            this.Fe,
            this.S,
            this.SiO2,
            this.CaO,
            this.MgO,
            this.Al2O3,
            this.Co});
            this.dataGridView1.Location = new System.Drawing.Point(55, 52);
            this.dataGridView1.Margin = new System.Windows.Forms.Padding(4);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 23;
            this.dataGridView1.Size = new System.Drawing.Size(991, 408);
            this.dataGridView1.TabIndex = 1;
            // 
            // 仓号
            // 
            this.仓号.HeaderText = "仓号";
            this.仓号.Name = "仓号";
            this.仓号.ReadOnly = true;
            this.仓号.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.仓号.Width = 50;
            // 
            // MineralName
            // 
            this.MineralName.HeaderText = "精矿名称";
            this.MineralName.Name = "MineralName";
            this.MineralName.ReadOnly = true;
            this.MineralName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.MineralName.Width = 80;
            // 
            // Cu
            // 
            this.Cu.HeaderText = "Cu";
            this.Cu.Name = "Cu";
            this.Cu.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.Cu.Width = 70;
            // 
            // Fe
            // 
            this.Fe.HeaderText = "Fe";
            this.Fe.Name = "Fe";
            this.Fe.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.Fe.Width = 70;
            // 
            // S
            // 
            this.S.HeaderText = "S";
            this.S.Name = "S";
            this.S.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.S.Width = 70;
            // 
            // SiO2
            // 
            this.SiO2.HeaderText = "SiO2";
            this.SiO2.Name = "SiO2";
            this.SiO2.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.SiO2.Width = 70;
            // 
            // CaO
            // 
            this.CaO.HeaderText = "CaO";
            this.CaO.Name = "CaO";
            this.CaO.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.CaO.Width = 70;
            // 
            // MgO
            // 
            this.MgO.HeaderText = "MgO";
            this.MgO.Name = "MgO";
            this.MgO.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.MgO.Width = 70;
            // 
            // Al2O3
            // 
            this.Al2O3.HeaderText = "Al2O3";
            this.Al2O3.Name = "Al2O3";
            this.Al2O3.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.Al2O3.Width = 70;
            // 
            // Co
            // 
            this.Co.HeaderText = "Co";
            this.Co.Name = "Co";
            this.Co.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.Co.Width = 70;
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("宋体", 11F);
            this.button1.Location = new System.Drawing.Point(561, 497);
            this.button1.Margin = new System.Windows.Forms.Padding(4);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(133, 50);
            this.button1.TabIndex = 2;
            this.button1.Text = "保存数据";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // mine_recommend
            // 
            this.button2.Font = new System.Drawing.Font("宋体", 11F);
            this.button2.Location = new System.Drawing.Point(378, 497);
            this.button2.Margin = new System.Windows.Forms.Padding(4);
            this.button2.Name = "mine_recommend";
            this.button2.Size = new System.Drawing.Size(133, 50);
            this.button2.TabIndex = 3;
            this.button2.Text = "一键填写";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // MineralComponents
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1112, 588);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.dataGridView1);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "MineralComponents";
            this.Text = "MineralComponents";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.DataGridViewTextBoxColumn 仓号;
        private System.Windows.Forms.DataGridViewTextBoxColumn MineralName;
        private System.Windows.Forms.DataGridViewTextBoxColumn Cu;
        private System.Windows.Forms.DataGridViewTextBoxColumn Fe;
        private System.Windows.Forms.DataGridViewTextBoxColumn S;
        private System.Windows.Forms.DataGridViewTextBoxColumn SiO2;
        private System.Windows.Forms.DataGridViewTextBoxColumn CaO;
        private System.Windows.Forms.DataGridViewTextBoxColumn MgO;
        private System.Windows.Forms.DataGridViewTextBoxColumn Al2O3;
        private System.Windows.Forms.DataGridViewTextBoxColumn Co;
        private System.Windows.Forms.Button button2;
    }
}