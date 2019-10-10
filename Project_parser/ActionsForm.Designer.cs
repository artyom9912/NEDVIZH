namespace Project_parser
{
    partial class ActionsForm
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
            this.exportButton = new System.Windows.Forms.Button();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.parseButton = new System.Windows.Forms.Button();
            this.exploreButton = new System.Windows.Forms.Button();
            this.uploadButton = new System.Windows.Forms.Button();
            this.stopButton = new System.Windows.Forms.Button();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.statusLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.fromBox = new System.Windows.Forms.NumericUpDown();
            this.toBox = new System.Windows.Forms.NumericUpDown();
            this.label3 = new System.Windows.Forms.Label();
            this.FixButton = new System.Windows.Forms.Button();
            this.statusStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.fromBox)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.toBox)).BeginInit();
            this.SuspendLayout();
            // 
            // exportButton
            // 
            this.exportButton.Location = new System.Drawing.Point(10, 90);
            this.exportButton.Name = "exportButton";
            this.exportButton.Size = new System.Drawing.Size(200, 23);
            this.exportButton.TabIndex = 15;
            this.exportButton.Text = "Выгрузить в Excel";
            this.exportButton.UseVisualStyleBackColor = true;
            this.exportButton.Click += new System.EventHandler(this.ExportButton_Click);
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(10, 12);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(321, 20);
            this.textBox2.TabIndex = 10;
            this.textBox2.TextChanged += new System.EventHandler(this.textBox2_TextChanged);
            // 
            // parseButton
            // 
            this.parseButton.Location = new System.Drawing.Point(216, 64);
            this.parseButton.Name = "parseButton";
            this.parseButton.Size = new System.Drawing.Size(200, 23);
            this.parseButton.TabIndex = 14;
            this.parseButton.Text = "Парсинг";
            this.parseButton.UseVisualStyleBackColor = true;
            this.parseButton.Click += new System.EventHandler(this.ParseButton_Click);
            // 
            // exploreButton
            // 
            this.exploreButton.Location = new System.Drawing.Point(337, 12);
            this.exploreButton.Name = "exploreButton";
            this.exploreButton.Size = new System.Drawing.Size(75, 20);
            this.exploreButton.TabIndex = 11;
            this.exploreButton.Text = "Обзор...";
            this.exploreButton.UseVisualStyleBackColor = true;
            this.exploreButton.Click += new System.EventHandler(this.ExploreButton_Click);
            // 
            // uploadButton
            // 
            this.uploadButton.Location = new System.Drawing.Point(10, 64);
            this.uploadButton.Name = "uploadButton";
            this.uploadButton.Size = new System.Drawing.Size(200, 23);
            this.uploadButton.TabIndex = 13;
            this.uploadButton.Text = "Скачать объявления";
            this.uploadButton.UseVisualStyleBackColor = true;
            this.uploadButton.Click += new System.EventHandler(this.UploadButton_Click);
            // 
            // stopButton
            // 
            this.stopButton.Enabled = false;
            this.stopButton.Location = new System.Drawing.Point(216, 90);
            this.stopButton.Name = "stopButton";
            this.stopButton.Size = new System.Drawing.Size(200, 23);
            this.stopButton.TabIndex = 16;
            this.stopButton.Text = "Остановить";
            this.stopButton.UseVisualStyleBackColor = true;
            this.stopButton.Click += new System.EventHandler(this.StopButton_Click);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.statusLabel});
            this.statusStrip1.Location = new System.Drawing.Point(0, 121);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(424, 22);
            this.statusStrip1.SizingGrip = false;
            this.statusStrip1.TabIndex = 17;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // statusLabel
            // 
            this.statusLabel.Name = "statusLabel";
            this.statusLabel.Size = new System.Drawing.Size(73, 17);
            this.statusLabel.Text = "Ожидание...";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(10, 41);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(23, 13);
            this.label1.TabIndex = 20;
            this.label1.Text = "От:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(110, 40);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(25, 13);
            this.label2.TabIndex = 21;
            this.label2.Text = "До:";
            // 
            // fromBox
            // 
            this.fromBox.Location = new System.Drawing.Point(33, 38);
            this.fromBox.Maximum = new decimal(new int[] {
            999999999,
            0,
            0,
            0});
            this.fromBox.Name = "fromBox";
            this.fromBox.Size = new System.Drawing.Size(73, 20);
            this.fromBox.TabIndex = 22;
            this.fromBox.ValueChanged += new System.EventHandler(this.FromBox_ValueChanged);
            // 
            // toBox
            // 
            this.toBox.Location = new System.Drawing.Point(134, 38);
            this.toBox.Maximum = new decimal(new int[] {
            999999999,
            0,
            0,
            0});
            this.toBox.Name = "toBox";
            this.toBox.Size = new System.Drawing.Size(73, 20);
            this.toBox.TabIndex = 23;
            this.toBox.ValueChanged += new System.EventHandler(this.ToBox_ValueChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(213, 41);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(44, 13);
            this.label3.TabIndex = 24;
            this.label3.Text = "Кол-во:";
            // 
            // FixButton
            // 
            this.FixButton.Location = new System.Drawing.Point(337, 36);
            this.FixButton.Name = "FixButton";
            this.FixButton.Size = new System.Drawing.Size(75, 23);
            this.FixButton.TabIndex = 25;
            this.FixButton.Text = "Исправить";
            this.FixButton.UseVisualStyleBackColor = true;
            this.FixButton.Click += new System.EventHandler(this.FixButton_Click);
            // 
            // ActionsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(424, 143);
            this.Controls.Add(this.FixButton);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.toBox);
            this.Controls.Add(this.fromBox);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.stopButton);
            this.Controls.Add(this.exportButton);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.uploadButton);
            this.Controls.Add(this.parseButton);
            this.Controls.Add(this.exploreButton);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label2);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.Name = "ActionsForm";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.ActionsForm_FormClosed);
            this.Load += new System.EventHandler(this.ActionsForm_Load);
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.fromBox)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.toBox)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button exportButton;
        private System.Windows.Forms.Button parseButton;
        private System.Windows.Forms.Button exploreButton;
        private System.Windows.Forms.Button uploadButton;
        private System.Windows.Forms.Button stopButton;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel statusLabel;
        public System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.NumericUpDown fromBox;
        private System.Windows.Forms.NumericUpDown toBox;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button FixButton;
    }
}