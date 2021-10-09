namespace ScoreConverter
{
    partial class SelectFileForm
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
            this.label1 = new System.Windows.Forms.Label();
            this.SourceWorkbook = new System.Windows.Forms.ComboBox();
            this.SourceWorksheet = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.TargetWorkbook = new System.Windows.Forms.ComboBox();
            this.ValidateButton = new System.Windows.Forms.Button();
            this.ExecuteButton = new System.Windows.Forms.Button();
            this.CreateSourceButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Malgun Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.label1.Location = new System.Drawing.Point(37, 28);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(61, 21);
            this.label1.TabIndex = 0;
            this.label1.Text = "Source";
            // 
            // SourceWorkbook
            // 
            this.SourceWorkbook.Font = new System.Drawing.Font("Malgun Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.SourceWorkbook.FormattingEnabled = true;
            this.SourceWorkbook.Location = new System.Drawing.Point(41, 52);
            this.SourceWorkbook.Name = "SourceWorkbook";
            this.SourceWorkbook.Size = new System.Drawing.Size(461, 29);
            this.SourceWorkbook.TabIndex = 1;
            this.SourceWorkbook.SelectedIndexChanged += new System.EventHandler(this.SourceWorkbook_SelectedIndexChanged);
            // 
            // SourceWorksheet
            // 
            this.SourceWorksheet.Font = new System.Drawing.Font("Malgun Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.SourceWorksheet.FormattingEnabled = true;
            this.SourceWorksheet.Location = new System.Drawing.Point(41, 87);
            this.SourceWorksheet.Name = "SourceWorksheet";
            this.SourceWorksheet.Size = new System.Drawing.Size(461, 29);
            this.SourceWorksheet.TabIndex = 2;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Malgun Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.label2.Location = new System.Drawing.Point(40, 173);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(58, 21);
            this.label2.TabIndex = 3;
            this.label2.Text = "Target";
            // 
            // TargetWorkbook
            // 
            this.TargetWorkbook.Font = new System.Drawing.Font("Malgun Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.TargetWorkbook.FormattingEnabled = true;
            this.TargetWorkbook.Location = new System.Drawing.Point(41, 197);
            this.TargetWorkbook.Name = "TargetWorkbook";
            this.TargetWorkbook.Size = new System.Drawing.Size(461, 29);
            this.TargetWorkbook.TabIndex = 4;
            // 
            // ValidateButton
            // 
            this.ValidateButton.Font = new System.Drawing.Font("Malgun Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.ValidateButton.Location = new System.Drawing.Point(169, 250);
            this.ValidateButton.Name = "ValidateButton";
            this.ValidateButton.Size = new System.Drawing.Size(163, 35);
            this.ValidateButton.TabIndex = 6;
            this.ValidateButton.Text = "Validate";
            this.ValidateButton.UseVisualStyleBackColor = true;
            this.ValidateButton.Click += new System.EventHandler(this.ValidateButton_Click);
            // 
            // ExecuteButton
            // 
            this.ExecuteButton.Font = new System.Drawing.Font("Malgun Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.ExecuteButton.Location = new System.Drawing.Point(339, 250);
            this.ExecuteButton.Name = "ExecuteButton";
            this.ExecuteButton.Size = new System.Drawing.Size(163, 35);
            this.ExecuteButton.TabIndex = 7;
            this.ExecuteButton.Text = "Execute";
            this.ExecuteButton.UseVisualStyleBackColor = true;
            this.ExecuteButton.Click += new System.EventHandler(this.ExecuteButton_Click);
            // 
            // CreateSourceButton
            // 
            this.CreateSourceButton.Font = new System.Drawing.Font("Malgun Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.CreateSourceButton.Location = new System.Drawing.Point(41, 250);
            this.CreateSourceButton.Name = "CreateSourceButton";
            this.CreateSourceButton.Size = new System.Drawing.Size(122, 35);
            this.CreateSourceButton.TabIndex = 5;
            this.CreateSourceButton.Text = "Create";
            this.CreateSourceButton.UseVisualStyleBackColor = true;
            this.CreateSourceButton.Click += new System.EventHandler(this.CreateSourceButton_Click);
            // 
            // SelectFileForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(550, 343);
            this.Controls.Add(this.CreateSourceButton);
            this.Controls.Add(this.ExecuteButton);
            this.Controls.Add(this.ValidateButton);
            this.Controls.Add(this.TargetWorkbook);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.SourceWorksheet);
            this.Controls.Add(this.SourceWorkbook);
            this.Controls.Add(this.label1);
            this.Name = "SelectFileForm";
            this.Text = "SelectFileForm";
            this.Load += new System.EventHandler(this.SelectFileForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox SourceWorkbook;
        private System.Windows.Forms.ComboBox SourceWorksheet;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox TargetWorkbook;
        private System.Windows.Forms.Button ValidateButton;
        private System.Windows.Forms.Button ExecuteButton;
        private System.Windows.Forms.Button CreateSourceButton;
    }
}