namespace MSOfficeExportApp
{
    partial class MainForm
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
            this.exportBtn = new System.Windows.Forms.Button();
            this.textBox = new System.Windows.Forms.TextBox();
            this.l_directory = new System.Windows.Forms.Label();
            this.choose = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // exportBtn
            // 
            this.exportBtn.Location = new System.Drawing.Point(606, 380);
            this.exportBtn.Name = "exportBtn";
            this.exportBtn.Size = new System.Drawing.Size(105, 48);
            this.exportBtn.TabIndex = 0;
            this.exportBtn.Text = "Export";
            this.exportBtn.UseVisualStyleBackColor = true;
            this.exportBtn.Click += new System.EventHandler(this.Export_Click);
            // 
            // textBox
            // 
            this.textBox.Location = new System.Drawing.Point(37, 32);
            this.textBox.Multiline = true;
            this.textBox.Name = "textBox";
            this.textBox.Size = new System.Drawing.Size(674, 329);
            this.textBox.TabIndex = 1;
            // 
            // l_directory
            // 
            this.l_directory.AutoSize = true;
            this.l_directory.Location = new System.Drawing.Point(34, 9);
            this.l_directory.Name = "l_directory";
            this.l_directory.Size = new System.Drawing.Size(0, 17);
            this.l_directory.TabIndex = 2;
            // 
            // choose
            // 
            this.choose.Location = new System.Drawing.Point(37, 380);
            this.choose.Name = "choose";
            this.choose.Size = new System.Drawing.Size(105, 48);
            this.choose.TabIndex = 3;
            this.choose.Text = "Choose Directory";
            this.choose.UseVisualStyleBackColor = true;
            this.choose.Click += new System.EventHandler(this.Choose_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(742, 450);
            this.Controls.Add(this.choose);
            this.Controls.Add(this.l_directory);
            this.Controls.Add(this.textBox);
            this.Controls.Add(this.exportBtn);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "MainForm";
            this.Text = " ";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button exportBtn;
        private System.Windows.Forms.TextBox textBox;
        private System.Windows.Forms.Label l_directory;
        private System.Windows.Forms.Button choose;
    }
}

