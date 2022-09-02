
namespace VSTO_Sample
{
    partial class Main
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

        #region 구성 요소 디자이너에서 생성한 코드

        /// <summary> 
        /// 디자이너 지원에 필요한 메서드입니다. 
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnDataLoad = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnDataLoad
            // 
            this.btnDataLoad.Location = new System.Drawing.Point(36, 36);
            this.btnDataLoad.Name = "btnDataLoad";
            this.btnDataLoad.Size = new System.Drawing.Size(336, 90);
            this.btnDataLoad.TabIndex = 0;
            this.btnDataLoad.Text = "Data Load";
            this.btnDataLoad.UseVisualStyleBackColor = true;
            this.btnDataLoad.Click += new System.EventHandler(this.btnDataLoad_Click);
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(36, 201);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(336, 90);
            this.btnSave.TabIndex = 1;
            this.btnSave.Text = "Save";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnDataLoad);
            this.Name = "Main";
            this.Size = new System.Drawing.Size(428, 653);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnDataLoad;
        private System.Windows.Forms.Button btnSave;
    }
}
