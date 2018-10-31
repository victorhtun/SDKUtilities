namespace AccessMatrix
{
    partial class Test
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
            this.btnConnect = new System.Windows.Forms.Button();
            this.btnCreateUDT = new System.Windows.Forms.Button();
            this.btnCreateQueries = new System.Windows.Forms.Button();
            this.btnAddUDVs = new System.Windows.Forms.Button();
            this.btnConnectUI = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnConnect
            // 
            this.btnConnect.Location = new System.Drawing.Point(16, 54);
            this.btnConnect.Margin = new System.Windows.Forms.Padding(4);
            this.btnConnect.Name = "btnConnect";
            this.btnConnect.Size = new System.Drawing.Size(100, 28);
            this.btnConnect.TabIndex = 0;
            this.btnConnect.Text = "Connect";
            this.btnConnect.UseVisualStyleBackColor = true;
            this.btnConnect.Click += new System.EventHandler(this.btnConnect_Click);
            // 
            // btnCreateUDT
            // 
            this.btnCreateUDT.Location = new System.Drawing.Point(16, 113);
            this.btnCreateUDT.Margin = new System.Windows.Forms.Padding(4);
            this.btnCreateUDT.Name = "btnCreateUDT";
            this.btnCreateUDT.Size = new System.Drawing.Size(100, 28);
            this.btnCreateUDT.TabIndex = 1;
            this.btnCreateUDT.Text = "Create UDT";
            this.btnCreateUDT.UseVisualStyleBackColor = true;
            this.btnCreateUDT.Click += new System.EventHandler(this.btnCreateUDT_Click);
            // 
            // btnCreateQueries
            // 
            this.btnCreateQueries.Location = new System.Drawing.Point(16, 164);
            this.btnCreateQueries.Margin = new System.Windows.Forms.Padding(4);
            this.btnCreateQueries.Name = "btnCreateQueries";
            this.btnCreateQueries.Size = new System.Drawing.Size(100, 47);
            this.btnCreateQueries.TabIndex = 2;
            this.btnCreateQueries.Text = "Create Queries";
            this.btnCreateQueries.UseVisualStyleBackColor = true;
            this.btnCreateQueries.Click += new System.EventHandler(this.btnCreateQueries_Click);
            // 
            // btnAddUDVs
            // 
            this.btnAddUDVs.Location = new System.Drawing.Point(16, 234);
            this.btnAddUDVs.Margin = new System.Windows.Forms.Padding(4);
            this.btnAddUDVs.Name = "btnAddUDVs";
            this.btnAddUDVs.Size = new System.Drawing.Size(100, 28);
            this.btnAddUDVs.TabIndex = 3;
            this.btnAddUDVs.Text = "Add UDVs";
            this.btnAddUDVs.UseVisualStyleBackColor = true;
            this.btnAddUDVs.Click += new System.EventHandler(this.btnAddUDVs_Click);
            // 
            // btnConnectUI
            // 
            this.btnConnectUI.Location = new System.Drawing.Point(16, 284);
            this.btnConnectUI.Name = "btnConnectUI";
            this.btnConnectUI.Size = new System.Drawing.Size(100, 31);
            this.btnConnectUI.TabIndex = 4;
            this.btnConnectUI.Text = "Connect UI";
            this.btnConnectUI.UseVisualStyleBackColor = true;
            this.btnConnectUI.Click += new System.EventHandler(this.btnConnectUI_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(16, 337);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(100, 33);
            this.button1.TabIndex = 5;
            this.button1.Text = "Connect DI";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // Test
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(799, 577);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.btnConnectUI);
            this.Controls.Add(this.btnAddUDVs);
            this.Controls.Add(this.btnCreateQueries);
            this.Controls.Add(this.btnCreateUDT);
            this.Controls.Add(this.btnConnect);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "Test";
            this.Text = "Test";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnConnect;
        private System.Windows.Forms.Button btnCreateUDT;
        private System.Windows.Forms.Button btnCreateQueries;
        private System.Windows.Forms.Button btnAddUDVs;
        private System.Windows.Forms.Button btnConnectUI;
        private System.Windows.Forms.Button button1;
    }
}