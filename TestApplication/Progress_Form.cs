using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TestApplication
{
    public partial class Progress_Form : Form
    {
        int Before_File = 0;
        long TotalCount;
        int DB_Count;

        public Progress_Form(string Type, int DB_Count)
        {

            if (Type == "MERGE")
            {
                #region

                this.label1 = new System.Windows.Forms.Label();
                this.label2 = new System.Windows.Forms.Label();
                this.label3 = new System.Windows.Forms.Label();
                this.progressBar1 = new System.Windows.Forms.ProgressBar();
                this.label4 = new System.Windows.Forms.Label();
                this.SuspendLayout();
                // 
                // label1
                // 
                this.label1.AutoSize = true;
                this.label1.Location = new System.Drawing.Point(43, 24);
                this.label1.Name = "label1";
                this.label1.Size = new System.Drawing.Size(105, 29);
                this.label1.TabIndex = 0;
                this.label1.Text = "Inserting";
                // 
                // label2
                // 
                this.label2.AutoSize = true;
                this.label2.Location = new System.Drawing.Point(312, 26);
                this.label2.Name = "label2";
                this.label2.Size = new System.Drawing.Size(74, 29);
                this.label2.TabIndex = 1;
                this.label2.Text = "Rows";
                // 
                // label3
                // 
                this.label3.AutoSize = true;
                this.label3.Location = new System.Drawing.Point(187, 26);
                this.label3.Name = "label3";
                this.label3.Size = new System.Drawing.Size(0, 29);
                this.label3.TabIndex = 2;
                // 
                // progressBar1
                // 
                this.progressBar1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                | System.Windows.Forms.AnchorStyles.Left)
                | System.Windows.Forms.AnchorStyles.Right)));
                this.progressBar1.Location = new System.Drawing.Point(48, 70);
                this.progressBar1.Name = "progressBar1";
                this.progressBar1.Size = new System.Drawing.Size(619, 70);
                this.progressBar1.TabIndex = 3;
                // 
                // label4
                // 
                this.label4.AutoSize = true;
                this.label4.Location = new System.Drawing.Point(307, 90);
                this.label4.Name = "label4";
                this.label4.Size = new System.Drawing.Size(79, 29);
                this.label4.TabIndex = 4;
                this.label4.Text = "label4";
                // 
                // Progress_Form
                // 
                this.AutoScaleDimensions = new System.Drawing.SizeF(14F, 29F);
                this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
                this.AutoScroll = true;
                this.AutoSize = true;
                this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
                this.ClientSize = new System.Drawing.Size(705, 220);
                this.Controls.Add(this.label4);
                this.Controls.Add(this.progressBar1);
                this.Controls.Add(this.label3);
                this.Controls.Add(this.label2);
                this.Controls.Add(this.label1);
                this.Name = "Progress_Form";
                this.Text = "Insert_Count_Form";
                this.ResumeLayout(false);
                this.PerformLayout();

                #endregion
            }
            else if (Type == "UNZIP")
            {
                #region
                this.progressBar2 = new System.Windows.Forms.ProgressBar();

                this.label5 = new System.Windows.Forms.Label();
                this.SuspendLayout();
                // 
                // progressBar2
                // 
                this.progressBar2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                                        | System.Windows.Forms.AnchorStyles.Left)
                                        | System.Windows.Forms.AnchorStyles.Right)));
                this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
                this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
                this.progressBar2.Location = new System.Drawing.Point(40, 63);
                this.progressBar2.Name = "progressBar2";
                this.progressBar2.Size = new System.Drawing.Size(619, 70);
                this.progressBar2.TabIndex = 0;
                // 
                // label5
                // 
                this.label5.AutoSize = true;
                this.label5.Location = new System.Drawing.Point(321, 98);
                this.label5.Name = "label5";
                this.label5.Size = new System.Drawing.Size(79, 29);
                this.label5.TabIndex = 4;
                this.label5.Text = "";

                // 
                // Progress_Form
                // 
                this.AutoScaleDimensions = new System.Drawing.SizeF(14F, 29F);
                this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
                this.AutoScroll = true;
                this.AutoSize = true;
                this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
                this.ClientSize = new System.Drawing.Size(705, 220);
                this.Controls.Add(this.label5);
                this.Controls.Add(this.progressBar2);
                this.Name = "Progress_Form";
                this.ResumeLayout(false);
                this.PerformLayout();
                #endregion

            }
            else if (Type == "YIELD")
            {
                #region

                this.progressBar3 = new ProgressBar[DB_Count];
                for (int i = 0; i < DB_Count; i++)
                {
                    this.progressBar3[i] = new ProgressBar();
                }


                this.SuspendLayout();

                // 
                // progressBar1
                // 

                int Height = 70;
                for (int i = 0; i < DB_Count; i++)
                {
                    //this.progressBar3[i].Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                    //                            | System.Windows.Forms.AnchorStyles.Left)
                    //                            | System.Windows.Forms.AnchorStyles.Right)));
                    this.progressBar3[i].Location = new System.Drawing.Point(48, 70 * i + 5);
                    this.progressBar3[i].Name = "progressBar1";
                    this.progressBar3[i].Size = new System.Drawing.Size(619, 70);
                    this.progressBar3[i].TabIndex = 3;

                }

                // 
                // label4
                // 
                //this.label4.AutoSize = true;
                //this.label4.Location = new System.Drawing.Point(307, 90);
                //this.label4.Name = "label4";
                //this.label4.Size = new System.Drawing.Size(79, 29);
                //this.label4.TabIndex = 4;
                //this.label4.Text = "label4";
                // 
                // Progress_Form
                // 
                this.AutoScaleDimensions = new System.Drawing.SizeF(14F, 29F);
                this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
                this.AutoScroll = true;
                this.AutoSize = true;
                this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
                this.ClientSize = new System.Drawing.Size(705, 220);
                //  this.Controls.Add(this.label4);
                for (int i = 0; i < DB_Count; i++)
                {
                    this.Controls.Add(this.progressBar3[i]);
                }


                this.Name = "Progress_Form";
                this.Text = "Reading_Count_Form";
                this.ResumeLayout(false);
                this.PerformLayout();

                #endregion
            }
            //  InitializeComponent();
        }

        public void Merge_Init(int FileCount)
        {
            progressBar1.Minimum = 0;
            progressBar1.Maximum = FileCount;
            progressBar1.Step = 1;
            progressBar1.Style = ProgressBarStyle.Marquee;

            TotalCount = FileCount;

            label1.Show();
            label2.Show();
            label3.Show();
            label4.Text = "(0/ " + TotalCount + ")";

            progressBar1.Enabled = true;
        }
        public void Merge_Print(long Count, int CurrentFile)
        {
            label3.Text = Convert.ToString(Count);
            label3.Update();
            if (Before_File != CurrentFile)
            {
                progressBar1.Value = CurrentFile;
                Before_File = CurrentFile;
                label4.Text = "(" + CurrentFile + "/ " + TotalCount + ")";
            }
        }

        public void Merge_Unzip_Init(int FileCount)
        {
            progressBar2.Enabled = true;
            progressBar2.Minimum = 0;
            progressBar2.Maximum = FileCount;
            progressBar2.Step = 1;
            progressBar2.Style = ProgressBarStyle.Continuous;

            TotalCount = FileCount;

            label5.Text = "(0/ " + TotalCount + ")";
            label5.Show();

            label5.Update();


        }
        public void Merge_Unzip_Print(int CurrentFile)
        {
            if (Before_File != CurrentFile)
            {
                progressBar2.Value = CurrentFile;
                progressBar2.PerformStep();
                Before_File = CurrentFile;
                label5.Text = "(" + CurrentFile + "/ " + TotalCount + ")";
                label5.Update();
            }
        }

        public void Yield_Init(long FileCount, int DB_Count, long Sample_Count)
        {

            for (int i = 0; i < DB_Count; i++)
            {
                progressBar3[i].Enabled = true;

                progressBar3[i].Minimum = 0;
                progressBar3[i].Maximum = Convert.ToInt16(Sample_Count) / 100;
                progressBar3[i].Step = 100;
                progressBar3[i].Style = ProgressBarStyle.Continuous;

            }



        }
        public void Yield_Print(long Count, int DB_Count, int[] SampleCount)
        {
            for (int i = 0; i < DB_Count; i++)
            {
                progressBar3[i].Value = SampleCount[i];
                progressBar3[i].PerformStep();
            }

        }


    }
}
