namespace TestApplication
{
    partial class MakeSpec_Form
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
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabControl2 = new System.Windows.Forms.TabControl();
            this.tabControl3 = new System.Windows.Forms.TabControl();
            this.button2 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(3022, 460);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(159, 79);
            this.button1.TabIndex = 2;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Location = new System.Drawing.Point(12, 27);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(3004, 1714);
            this.tabControl1.TabIndex = 3;
            this.tabControl1.SelectedIndexChanged += new System.EventHandler(this.tabControl1_SelectedIndexChanged);
            // 
            // tabControl2
            // 
            this.tabControl2.Location = new System.Drawing.Point(3022, 562);
            this.tabControl2.Name = "tabControl2";
            this.tabControl2.SelectedIndex = 0;
            this.tabControl2.Size = new System.Drawing.Size(792, 1179);
            this.tabControl2.TabIndex = 4;
            this.tabControl2.SelectedIndexChanged += new System.EventHandler(this.tabControl2_SelectedIndexChanged);
            // 
            // tabControl3
            // 
            this.tabControl3.Location = new System.Drawing.Point(3022, 27);
            this.tabControl3.Name = "tabControl3";
            this.tabControl3.SelectedIndex = 0;
            this.tabControl3.Size = new System.Drawing.Size(551, 412);
            this.tabControl3.TabIndex = 5;
            this.tabControl3.SelectedIndexChanged += new System.EventHandler(this.tabControl3_SelectedIndexChanged);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(3251, 460);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(159, 79);
            this.button2.TabIndex = 6;
            this.button2.Text = "button2";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // MakeSpec_Form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(14F, 29F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(3831, 1749);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.tabControl3);
            this.Controls.Add(this.tabControl2);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.button1);
            this.Name = "MakeSpec_Form";
            this.Text = "MakeSpec_Form";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MakeSpec_Form_FormClosing);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button button1;
        public System.Windows.Forms.TabControl tabControl1;
        public System.Windows.Forms.TabControl tabControl2;
        public System.Windows.Forms.TabControl tabControl3;
        private System.Windows.Forms.Button button2;
    }
}