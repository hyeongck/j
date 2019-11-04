namespace TestApplication
{
    partial class Merge_Form
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
            this.components = new System.ComponentModel.Container();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.button1 = new System.Windows.Forms.Button();
            this.treeView1 = new System.Windows.Forms.TreeView();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.button4 = new System.Windows.Forms.Button();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.button5 = new System.Windows.Forms.Button();
            this.listBox2 = new System.Windows.Forms.ListBox();
            this.button6 = new System.Windows.Forms.Button();
            this.button7 = new System.Windows.Forms.Button();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.button3 = new System.Windows.Forms.Button();
            this.dataGridView3 = new System.Windows.Forms.DataGridView();
            this.button8 = new System.Windows.Forms.Button();
            this.checkBox2 = new System.Windows.Forms.CheckBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView3)).BeginInit();
            this.SuspendLayout();
            // 
            // listBox1
            // 
            this.listBox1.AllowDrop = true;
            this.listBox1.FormattingEnabled = true;
            this.listBox1.ItemHeight = 29;
            this.listBox1.Location = new System.Drawing.Point(21, 48);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(874, 265);
            this.listBox1.TabIndex = 0;
            this.listBox1.DragDrop += new System.Windows.Forms.DragEventHandler(this.listBox1_DragDrop);
            this.listBox1.DragEnter += new System.Windows.Forms.DragEventHandler(this.listBox1_DragEnter);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(1125, 261);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(152, 80);
            this.button1.TabIndex = 1;
            this.button1.Text = "Unzip";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // treeView1
            // 
            this.treeView1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.treeView1.CheckBoxes = true;
            this.treeView1.Location = new System.Drawing.Point(32, 35);
            this.treeView1.Name = "treeView1";
            this.treeView1.Size = new System.Drawing.Size(464, 369);
            this.treeView1.TabIndex = 3;
            this.treeView1.AfterCheck += new System.Windows.Forms.TreeViewEventHandler(this.treeView1_AfterCheck);
            this.treeView1.NodeMouseClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.treeView1_NodeMouseClick);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.listBox1);
            this.groupBox1.Location = new System.Drawing.Point(10, 11);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(930, 331);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Data";
            // 
            // groupBox2
            // 
            this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.groupBox2.Controls.Add(this.treeView1);
            this.groupBox2.Location = new System.Drawing.Point(10, 358);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(528, 433);
            this.groupBox2.TabIndex = 5;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Lot";
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.ImageScalingSize = new System.Drawing.Size(40, 40);
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(61, 4);
            // 
            // dataGridView1
            // 
            this.dataGridView1.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCells;
            this.dataGridView1.BackgroundColor = System.Drawing.SystemColors.ButtonFace;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(962, 558);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 40;
            this.dataGridView1.Size = new System.Drawing.Size(399, 277);
            this.dataGridView1.TabIndex = 6;
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(967, 358);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(152, 81);
            this.button4.TabIndex = 8;
            this.button4.Text = "Export CSV";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(967, 117);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(251, 33);
            this.checkBox1.TabIndex = 9;
            this.checkBox1.Text = "Subsequent writing";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(967, 261);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(152, 81);
            this.button5.TabIndex = 10;
            this.button5.Text = "Open DB";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // listBox2
            // 
            this.listBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.listBox2.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.listBox2.FormattingEnabled = true;
            this.listBox2.IntegralHeight = false;
            this.listBox2.ItemHeight = 10;
            this.listBox2.Location = new System.Drawing.Point(558, 376);
            this.listBox2.Name = "listBox2";
            this.listBox2.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.listBox2.Size = new System.Drawing.Size(385, 275);
            this.listBox2.TabIndex = 11;
            this.listBox2.MouseClick += new System.Windows.Forms.MouseEventHandler(this.listBox2_MouseClick);
            this.listBox2.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.listBox2_DrawItem);
            this.listBox2.MouseDown += new System.Windows.Forms.MouseEventHandler(this.listBox2_MouseDown);
            // 
            // button6
            // 
            this.button6.Location = new System.Drawing.Point(967, 23);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(310, 81);
            this.button6.TabIndex = 12;
            this.button6.Text = "Select Forder and Merge";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.button6_Click);
            // 
            // button7
            // 
            this.button7.Location = new System.Drawing.Point(967, 454);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(152, 83);
            this.button7.TabIndex = 13;
            this.button7.Text = "Delete Data";
            this.button7.UseVisualStyleBackColor = true;
            this.button7.Click += new System.EventHandler(this.button7_Click);
            // 
            // dataGridView2
            // 
            this.dataGridView2.AllowUserToAddRows = false;
            this.dataGridView2.AllowUserToDeleteRows = false;
            this.dataGridView2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView2.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView2.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCells;
            this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView2.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.dataGridView2.Location = new System.Drawing.Point(1364, 358);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.RowTemplate.Height = 37;
            this.dataGridView2.Size = new System.Drawing.Size(505, 308);
            this.dataGridView2.TabIndex = 14;
            this.dataGridView2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dataGridView2_KeyDown);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(1125, 454);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(152, 83);
            this.button3.TabIndex = 15;
            this.button3.Text = "Edit data";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click_1);
            // 
            // dataGridView3
            // 
            this.dataGridView3.AllowUserToDeleteRows = false;
            this.dataGridView3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView3.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dataGridView3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView3.Location = new System.Drawing.Point(1363, 22);
            this.dataGridView3.Name = "dataGridView3";
            this.dataGridView3.RowTemplate.Height = 37;
            this.dataGridView3.Size = new System.Drawing.Size(506, 330);
            this.dataGridView3.TabIndex = 16;
            this.dataGridView3.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView3_CellContentClick);
            // 
            // button8
            // 
            this.button8.Location = new System.Drawing.Point(967, 156);
            this.button8.Name = "button8";
            this.button8.Size = new System.Drawing.Size(137, 81);
            this.button8.TabIndex = 17;
            this.button8.Text = "Execute";
            this.button8.UseVisualStyleBackColor = true;
            this.button8.Click += new System.EventHandler(this.button8_Click);
            // 
            // checkBox2
            // 
            this.checkBox2.AutoSize = true;
            this.checkBox2.Location = new System.Drawing.Point(1224, 117);
            this.checkBox2.Name = "checkBox2";
            this.checkBox2.Size = new System.Drawing.Size(58, 33);
            this.checkBox2.TabIndex = 18;
            this.checkBox2.Text = "1";
            this.checkBox2.UseVisualStyleBackColor = true;
            // 
            // Merge_Form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(14F, 29F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(2221, 1080);
            this.Controls.Add(this.checkBox2);
            this.Controls.Add(this.button8);
            this.Controls.Add(this.dataGridView3);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.dataGridView2);
            this.Controls.Add(this.button7);
            this.Controls.Add(this.button6);
            this.Controls.Add(this.listBox2);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.button1);
            this.Name = "Merge_Form";
            this.Text = "Merge_Form";
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView3)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TreeView treeView1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.ListBox listBox2;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.Button button7;
        private System.Windows.Forms.DataGridView dataGridView2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.DataGridView dataGridView3;
        private System.Windows.Forms.Button button8;
        private System.Windows.Forms.CheckBox checkBox2;
    }
}