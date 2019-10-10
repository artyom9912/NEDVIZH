namespace Project_parser
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
            this.label1 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.startButton = new System.Windows.Forms.Button();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.countLabel = new System.Windows.Forms.Label();
            this.guideButton = new System.Windows.Forms.Button();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.statusLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.button3 = new System.Windows.Forms.Button();
            this.archiveButton = new System.Windows.Forms.Button();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.CommentBox = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(14, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(73, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Адрес сайта:";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(17, 27);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(504, 20);
            this.textBox1.TabIndex = 1;
            this.textBox1.Text = "https://www.farpost.ru/oktyabrskii-r-n/realty/sell_business_realty/";
            // 
            // startButton
            // 
            this.startButton.Location = new System.Drawing.Point(18, 298);
            this.startButton.Name = "startButton";
            this.startButton.Size = new System.Drawing.Size(75, 23);
            this.startButton.TabIndex = 2;
            this.startButton.Text = "Начать";
            this.startButton.UseVisualStyleBackColor = true;
            this.startButton.Click += new System.EventHandler(this.StartButton_Click);
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.Location = new System.Drawing.Point(17, 53);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(581, 199);
            this.listBox1.TabIndex = 3;
            // 
            // countLabel
            // 
            this.countLabel.AutoSize = true;
            this.countLabel.Location = new System.Drawing.Point(481, 303);
            this.countLabel.Name = "countLabel";
            this.countLabel.Size = new System.Drawing.Size(117, 13);
            this.countLabel.TabIndex = 4;
            this.countLabel.Text = "Кол-во объявлений: 0";
            // 
            // guideButton
            // 
            this.guideButton.Location = new System.Drawing.Point(197, 298);
            this.guideButton.Name = "guideButton";
            this.guideButton.Size = new System.Drawing.Size(86, 23);
            this.guideButton.TabIndex = 6;
            this.guideButton.Text = "Руководство";
            this.guideButton.UseVisualStyleBackColor = true;
            this.guideButton.Click += new System.EventHandler(this.GuideButton_Click);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.statusLabel});
            this.statusStrip1.Location = new System.Drawing.Point(0, 342);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(615, 22);
            this.statusStrip1.SizingGrip = false;
            this.statusStrip1.TabIndex = 7;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // statusLabel
            // 
            this.statusLabel.Name = "statusLabel";
            this.statusLabel.Size = new System.Drawing.Size(0, 17);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(99, 298);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(92, 23);
            this.button3.TabIndex = 13;
            this.button3.Text = "Управление...";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.Button3_Click);
            // 
            // archiveButton
            // 
            this.archiveButton.Location = new System.Drawing.Point(523, 27);
            this.archiveButton.Name = "archiveButton";
            this.archiveButton.Size = new System.Drawing.Size(75, 20);
            this.archiveButton.TabIndex = 14;
            this.archiveButton.Text = "Архив";
            this.archiveButton.UseVisualStyleBackColor = true;
            this.archiveButton.Click += new System.EventHandler(this.ArchiveButton_Click);
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(289, 302);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(195, 17);
            this.checkBox1.TabIndex = 15;
            this.checkBox1.Text = "Учитывать город для дубликатов";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // CommentBox
            // 
            this.CommentBox.Location = new System.Drawing.Point(18, 272);
            this.CommentBox.Name = "CommentBox";
            this.CommentBox.Size = new System.Drawing.Size(580, 20);
            this.CommentBox.TabIndex = 16;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(16, 255);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(77, 13);
            this.label2.TabIndex = 17;
            this.label2.Text = "Комментарий";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(615, 364);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.CommentBox);
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.archiveButton);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.guideButton);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.countLabel);
            this.Controls.Add(this.startButton);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "MainForm";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Модуль Парсер";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button startButton;
        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.Label countLabel;
        private System.Windows.Forms.Button guideButton;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel statusLabel;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button archiveButton;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.TextBox CommentBox;
        private System.Windows.Forms.Label label2;
    }
}

