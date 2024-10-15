namespace Urdu_library5
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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
            button1 = new Button();
            progressBar1 = new ProgressBar();
            label1 = new Label();
            button2 = new Button();
            SuspendLayout();
            // 
            // button1
            // 
            button1.Font = new Font("Segoe UI", 20F);
            button1.Location = new Point(279, 86);
            button1.Name = "button1";
            button1.Size = new Size(254, 60);
            button1.TabIndex = 0;
            button1.Text = "Faylni tanlash";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // progressBar1
            // 
            progressBar1.Location = new Point(279, 178);
            progressBar1.Name = "progressBar1";
            progressBar1.Size = new Size(254, 43);
            progressBar1.Style = ProgressBarStyle.Marquee;
            progressBar1.TabIndex = 1;
            progressBar1.Value = 10;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("Segoe UI", 20F);
            label1.Location = new Point(86, 178);
            label1.Name = "label1";
            label1.Size = new Size(187, 37);
            label1.TabIndex = 2;
            label1.Text = "Yuborilmoqda";
            // 
            // button2
            // 
            button2.Font = new Font("Segoe UI", 15F);
            button2.Location = new Point(665, 393);
            button2.Name = "button2";
            button2.Size = new Size(123, 45);
            button2.TabIndex = 3;
            button2.Text = "Bog'lanish";
            button2.UseVisualStyleBackColor = true;
            button2.Click += button2_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(button2);
            Controls.Add(label1);
            Controls.Add(progressBar1);
            Controls.Add(button1);
            Name = "Form1";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Urdu elektron kutubxona";
            Load += Form1_Load;
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button button1;
        private ProgressBar progressBar1;
        private Label label1;
        private Button button2;
    }
}
