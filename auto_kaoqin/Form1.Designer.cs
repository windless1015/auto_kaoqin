namespace auto_kaoqin
{
    partial class main_Form
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
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.dateTimePicker_start = new System.Windows.Forms.DateTimePicker();
            this.btn_add_kaoqin = new System.Windows.Forms.Button();
            this.btn_add_shenpi = new System.Windows.Forms.Button();
            this.dateTimePicker_end = new System.Windows.Forms.DateTimePicker();
            this.lab_start = new System.Windows.Forms.Label();
            this.lab_end = new System.Windows.Forms.Label();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.btn_generate_final = new System.Windows.Forms.Button();
            this.btn_gen_acc_chidao = new System.Windows.Forms.Button();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // dateTimePicker_start
            // 
            this.dateTimePicker_start.Location = new System.Drawing.Point(88, 12);
            this.dateTimePicker_start.Name = "dateTimePicker_start";
            this.dateTimePicker_start.Size = new System.Drawing.Size(126, 21);
            this.dateTimePicker_start.TabIndex = 0;
            // 
            // btn_add_kaoqin
            // 
            this.btn_add_kaoqin.Location = new System.Drawing.Point(234, 12);
            this.btn_add_kaoqin.Name = "btn_add_kaoqin";
            this.btn_add_kaoqin.Size = new System.Drawing.Size(125, 21);
            this.btn_add_kaoqin.TabIndex = 1;
            this.btn_add_kaoqin.Text = "加载考勤excel文件";
            this.btn_add_kaoqin.UseVisualStyleBackColor = true;
            this.btn_add_kaoqin.Click += new System.EventHandler(this.btn_add_kaoqin_Click);
            // 
            // btn_add_shenpi
            // 
            this.btn_add_shenpi.Location = new System.Drawing.Point(234, 68);
            this.btn_add_shenpi.Name = "btn_add_shenpi";
            this.btn_add_shenpi.Size = new System.Drawing.Size(125, 21);
            this.btn_add_shenpi.TabIndex = 1;
            this.btn_add_shenpi.Text = "加载审批excel文件";
            this.btn_add_shenpi.UseVisualStyleBackColor = true;
            this.btn_add_shenpi.Click += new System.EventHandler(this.btn_add_shenpi_Click);
            // 
            // dateTimePicker_end
            // 
            this.dateTimePicker_end.Location = new System.Drawing.Point(88, 66);
            this.dateTimePicker_end.Name = "dateTimePicker_end";
            this.dateTimePicker_end.Size = new System.Drawing.Size(126, 21);
            this.dateTimePicker_end.TabIndex = 0;
            // 
            // lab_start
            // 
            this.lab_start.AutoSize = true;
            this.lab_start.Location = new System.Drawing.Point(12, 18);
            this.lab_start.Name = "lab_start";
            this.lab_start.Size = new System.Drawing.Size(59, 12);
            this.lab_start.TabIndex = 2;
            this.lab_start.Text = "起始日期:";
            // 
            // lab_end
            // 
            this.lab_end.AutoSize = true;
            this.lab_end.Location = new System.Drawing.Point(12, 66);
            this.lab_end.Name = "lab_end";
            this.lab_end.Size = new System.Drawing.Size(59, 12);
            this.lab_end.TabIndex = 2;
            this.lab_end.Text = "结束日期:";
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1});
            this.statusStrip1.Location = new System.Drawing.Point(0, 199);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(896, 22);
            this.statusStrip1.TabIndex = 5;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(56, 17);
            this.toolStripStatusLabel1.Text = "当前状态";
            // 
            // btn_generate_final
            // 
            this.btn_generate_final.Location = new System.Drawing.Point(517, 13);
            this.btn_generate_final.Name = "btn_generate_final";
            this.btn_generate_final.Size = new System.Drawing.Size(148, 23);
            this.btn_generate_final.TabIndex = 7;
            this.btn_generate_final.Text = "生成考勤矩阵excel";
            this.btn_generate_final.UseVisualStyleBackColor = true;
            this.btn_generate_final.Click += new System.EventHandler(this.btn_generate_final_Click);
            // 
            // btn_gen_acc_chidao
            // 
            this.btn_gen_acc_chidao.Location = new System.Drawing.Point(517, 65);
            this.btn_gen_acc_chidao.Name = "btn_gen_acc_chidao";
            this.btn_gen_acc_chidao.Size = new System.Drawing.Size(148, 23);
            this.btn_gen_acc_chidao.TabIndex = 8;
            this.btn_gen_acc_chidao.Text = "生成每月累积迟到excel";
            this.btn_gen_acc_chidao.UseVisualStyleBackColor = true;
            this.btn_gen_acc_chidao.Click += new System.EventHandler(this.btn_gen_acc_chidao_Click);
            // 
            // main_Form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(896, 221);
            this.Controls.Add(this.btn_gen_acc_chidao);
            this.Controls.Add(this.btn_generate_final);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.lab_end);
            this.Controls.Add(this.lab_start);
            this.Controls.Add(this.btn_add_shenpi);
            this.Controls.Add(this.btn_add_kaoqin);
            this.Controls.Add(this.dateTimePicker_end);
            this.Controls.Add(this.dateTimePicker_start);
            this.Name = "main_Form";
            this.Text = "自动化考勤";
            this.Load += new System.EventHandler(this.main_Form_Load);
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DateTimePicker dateTimePicker_start;
        private System.Windows.Forms.Button btn_add_kaoqin;
        private System.Windows.Forms.Button btn_add_shenpi;
        private System.Windows.Forms.DateTimePicker dateTimePicker_end;
        private System.Windows.Forms.Label lab_start;
        private System.Windows.Forms.Label lab_end;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.Button btn_generate_final;
        private System.Windows.Forms.Button btn_gen_acc_chidao;
    }
}

