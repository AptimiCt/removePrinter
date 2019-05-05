namespace remove_printer
{
    partial class Form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.ReadFileAndRemove = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // ReadFileAndRemove
            // 
            this.ReadFileAndRemove.AutoSize = true;
            this.ReadFileAndRemove.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.ReadFileAndRemove.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Red;
            this.ReadFileAndRemove.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.ReadFileAndRemove.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.125F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ReadFileAndRemove.Location = new System.Drawing.Point(182, 74);
            this.ReadFileAndRemove.Name = "ReadFileAndRemove";
            this.ReadFileAndRemove.Size = new System.Drawing.Size(242, 43);
            this.ReadFileAndRemove.TabIndex = 9;
            this.ReadFileAndRemove.Text = "ReadAndRemove";
            this.ReadFileAndRemove.UseVisualStyleBackColor = true;
            this.ReadFileAndRemove.Click += new System.EventHandler(this.ReadFileAndRemove_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(556, 196);
            this.Controls.Add(this.ReadFileAndRemove);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "RemovePrinter";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button ReadFileAndRemove;
    }
}

