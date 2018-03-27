namespace HandleXls
{
    sealed partial class Form1
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
            this.ShowTab = new System.Windows.Forms.TabControl();
            this.LoadXls = new System.Windows.Forms.Button();
            this.dataSet1 = new System.Data.DataSet();
            this.Handler = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataSet1)).BeginInit();
            this.SuspendLayout();
            // 
            // ShowTab
            // 
            this.ShowTab.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.ShowTab.Location = new System.Drawing.Point(141, 12);
            this.ShowTab.Name = "ShowTab";
            this.ShowTab.SelectedIndex = 0;
            this.ShowTab.Size = new System.Drawing.Size(647, 414);
            this.ShowTab.TabIndex = 0;
            // 
            // LoadXls
            // 
            this.LoadXls.Location = new System.Drawing.Point(30, 39);
            this.LoadXls.Name = "LoadXls";
            this.LoadXls.Size = new System.Drawing.Size(75, 23);
            this.LoadXls.TabIndex = 1;
            this.LoadXls.Text = "加载表格";
            this.LoadXls.UseVisualStyleBackColor = true;
            // 
            // dataSet1
            // 
            this.dataSet1.DataSetName = "NewDataSet";
            // 
            // Handler
            // 
            this.Handler.Location = new System.Drawing.Point(30, 69);
            this.Handler.Name = "Handler";
            this.Handler.Size = new System.Drawing.Size(75, 23);
            this.Handler.TabIndex = 2;
            this.Handler.Text = "处理";
            this.Handler.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 438);
            this.Controls.Add(this.Handler);
            this.Controls.Add(this.LoadXls);
            this.Controls.Add(this.ShowTab);
            this.Name = "Form1";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.dataSet1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl ShowTab;
        private System.Windows.Forms.Button LoadXls;
        private System.Data.DataSet dataSet1;
        private System.Windows.Forms.Button Handler;
    }
}