namespace removename
{
    partial class Form1
    {
        /// <summary>
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 디자이너에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다. 
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.axHwpCtrl1 = new AxHWPCONTROLLib.AxHwpCtrl();
            this.axHwpCtrl2 = new AxHWPCONTROLLib.AxHwpCtrl();
            this.axHwpCtrl3 = new AxHWPCONTROLLib.AxHwpCtrl();
            ((System.ComponentModel.ISupportInitialize)(this.axHwpCtrl1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.axHwpCtrl2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.axHwpCtrl3)).BeginInit();
            this.SuspendLayout();
            // 
            // axHwpCtrl1
            // 
            this.axHwpCtrl1.Enabled = true;
            this.axHwpCtrl1.Location = new System.Drawing.Point(0, 0);
            this.axHwpCtrl1.Name = "axHwpCtrl1";
            this.axHwpCtrl1.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("axHwpCtrl1.OcxState")));
            this.axHwpCtrl1.Size = new System.Drawing.Size(100, 50);
            this.axHwpCtrl1.TabIndex = 0;
            // 
            // axHwpCtrl2
            // 
            this.axHwpCtrl2.Enabled = true;
            this.axHwpCtrl2.Location = new System.Drawing.Point(115, 0);
            this.axHwpCtrl2.Name = "axHwpCtrl2";
            this.axHwpCtrl2.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("axHwpCtrl2.OcxState")));
            this.axHwpCtrl2.Size = new System.Drawing.Size(100, 50);
            this.axHwpCtrl2.TabIndex = 1;
            // 
            // axHwpCtrl3
            // 
            this.axHwpCtrl3.Enabled = true;
            this.axHwpCtrl3.Location = new System.Drawing.Point(235, 0);
            this.axHwpCtrl3.Name = "axHwpCtrl3";
            this.axHwpCtrl3.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("axHwpCtrl3.OcxState")));
            this.axHwpCtrl3.Size = new System.Drawing.Size(100, 50);
            this.axHwpCtrl3.TabIndex = 2;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(709, 468);
            this.Controls.Add(this.axHwpCtrl3);
            this.Controls.Add(this.axHwpCtrl2);
            this.Controls.Add(this.axHwpCtrl1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.axHwpCtrl1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.axHwpCtrl2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.axHwpCtrl3)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private AxHWPCONTROLLib.AxHwpCtrl axHwpCtrl1;
        private AxHWPCONTROLLib.AxHwpCtrl axHwpCtrl2;
        private AxHWPCONTROLLib.AxHwpCtrl axHwpCtrl3;
    }
}

