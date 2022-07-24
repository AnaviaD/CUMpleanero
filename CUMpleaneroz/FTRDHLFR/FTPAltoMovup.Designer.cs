namespace FTPALTOMOVUP
{
    partial class FTPAltoMovup
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de Windows Forms

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FTPAltoMovup));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.ltb_log = new System.Windows.Forms.ListBox();
            this.pgb_Progreso = new System.Windows.Forms.ProgressBar();
            this.btn_actuaizar = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.tmr_Actualiza = new System.Windows.Forms.Timer(this.components);
            this.tmr_manual = new System.Windows.Forms.Timer(this.components);
            this.tmr_envio = new System.Windows.Forms.Timer(this.components);
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.ltb_log);
            this.groupBox1.Controls.Add(this.pgb_Progreso);
            this.groupBox1.Controls.Add(this.btn_actuaizar);
            this.groupBox1.Controls.Add(this.pictureBox1);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(538, 203);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            // 
            // ltb_log
            // 
            this.ltb_log.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ltb_log.FormattingEnabled = true;
            this.ltb_log.HorizontalScrollbar = true;
            this.ltb_log.ItemHeight = 14;
            this.ltb_log.Location = new System.Drawing.Point(168, 19);
            this.ltb_log.Name = "ltb_log";
            this.ltb_log.ScrollAlwaysVisible = true;
            this.ltb_log.Size = new System.Drawing.Size(361, 172);
            this.ltb_log.TabIndex = 7;
            // 
            // pgb_Progreso
            // 
            this.pgb_Progreso.Location = new System.Drawing.Point(6, 181);
            this.pgb_Progreso.Name = "pgb_Progreso";
            this.pgb_Progreso.Size = new System.Drawing.Size(154, 10);
            this.pgb_Progreso.TabIndex = 6;
            // 
            // btn_actuaizar
            // 
            this.btn_actuaizar.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_actuaizar.Image = ((System.Drawing.Image)(resources.GetObject("btn_actuaizar.Image")));
            this.btn_actuaizar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btn_actuaizar.Location = new System.Drawing.Point(31, 146);
            this.btn_actuaizar.Name = "btn_actuaizar";
            this.btn_actuaizar.Size = new System.Drawing.Size(97, 29);
            this.btn_actuaizar.TabIndex = 5;
            this.btn_actuaizar.Text = "Actualizar";
            this.btn_actuaizar.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btn_actuaizar.UseVisualStyleBackColor = true;
            this.btn_actuaizar.Click += new System.EventHandler(this.btn_actuaizar_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(6, 19);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(156, 122);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 2;
            this.pictureBox1.TabStop = false;
            // 
            // tmr_Actualiza
            // 
            this.tmr_Actualiza.Enabled = true;
            this.tmr_Actualiza.Interval = 50000;
            // 
            // tmr_manual
            // 
            this.tmr_manual.Enabled = true;
            this.tmr_manual.Interval = 4000;
            this.tmr_manual.Tick += new System.EventHandler(this.tmr_manual_Tick);
            // 
            // tmr_envio
            // 
            this.tmr_envio.Enabled = true;
            this.tmr_envio.Interval = 50000;
            // 
            // FTPAltoMovup
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(558, 221);
            this.Controls.Add(this.groupBox1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FTPAltoMovup";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "STP FTR";
            this.Load += new System.EventHandler(this.SFTP_Load);
            this.groupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ListBox ltb_log;
        private System.Windows.Forms.ProgressBar pgb_Progreso;
        private System.Windows.Forms.Button btn_actuaizar;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Timer tmr_Actualiza;
        private System.Windows.Forms.Timer tmr_manual;
        private System.Windows.Forms.Timer tmr_envio;
    }
}

