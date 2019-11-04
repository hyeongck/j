namespace TestApplication
{
    partial class Spec_Gen_Form
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
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(1390, 58);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(252, 73);
            this.button1.TabIndex = 0;
            this.button1.Text = "Add SpecNumber";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(1664, 58);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(252, 72);
            this.button2.TabIndex = 1;
            this.button2.Text = "Analysis";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // textBox1
            // 
            this.textBox1.AllowDrop = true;
            this.textBox1.Location = new System.Drawing.Point(18, 38);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(1265, 66);
            this.textBox1.TabIndex = 2;
            this.textBox1.DragDrop += new System.Windows.Forms.DragEventHandler(this.textBox1_DragDrop);
            this.textBox1.DragEnter += new System.Windows.Forms.DragEventHandler(this.textBox1_DragEnter);
            // 
            // textBox2
            // 
            this.textBox2.AllowDrop = true;
            this.textBox2.Location = new System.Drawing.Point(18, 34);
            this.textBox2.Multiline = true;
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(1265, 72);
            this.textBox2.TabIndex = 3;
            this.textBox2.DragDrop += new System.Windows.Forms.DragEventHandler(this.textBox2_DragDrop);
            this.textBox2.DragEnter += new System.Windows.Forms.DragEventHandler(this.textBox2_DragEnter);
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(24, 34);
            this.textBox3.Multiline = true;
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(361, 72);
            this.textBox3.TabIndex = 4;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.textBox1);
            this.groupBox1.Location = new System.Drawing.Point(55, 27);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(1308, 126);
            this.groupBox1.TabIndex = 16;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "NPI Test Plan";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.textBox2);
            this.groupBox2.Location = new System.Drawing.Point(55, 192);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(1308, 126);
            this.groupBox2.TabIndex = 17;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Spec Sheet";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.textBox3);
            this.groupBox3.Location = new System.Drawing.Point(1390, 192);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(413, 126);
            this.groupBox3.TabIndex = 18;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Pass word";
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.textBox4);
            this.groupBox4.Controls.Add(this.groupBox3);
            this.groupBox4.Controls.Add(this.button2);
            this.groupBox4.Controls.Add(this.groupBox2);
            this.groupBox4.Controls.Add(this.groupBox1);
            this.groupBox4.Controls.Add(this.button1);
            this.groupBox4.Location = new System.Drawing.Point(65, 29);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(1951, 572);
            this.groupBox4.TabIndex = 19;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "groupBox4";
            // 
            // textBox4
            // 
            this.textBox4.AllowDrop = true;
            this.textBox4.Location = new System.Drawing.Point(73, 362);
            this.textBox4.Multiline = true;
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(1265, 72);
            this.textBox4.TabIndex = 19;
            this.textBox4.DragDrop += new System.Windows.Forms.DragEventHandler(this.textBox4_DragDrop_1);
            this.textBox4.DragEnter += new System.Windows.Forms.DragEventHandler(this.textBox4_DragEnter_1);
            // 
            // Spec_Gen_Form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(14F, 29F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(2481, 920);
            this.Controls.Add(this.groupBox4);
            this.Name = "Spec_Gen_Form";
            this.Text = "Spec_Gen_Form";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.TextBox textBox4;
    }
}