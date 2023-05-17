namespace InvaSAP
{
    partial class FormSAP
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
            panelSAP = new Panel();
            SuspendLayout();
            // 
            // panelSAP
            // 
            panelSAP.Location = new Point(38, 50);
            panelSAP.Name = "panelSAP";
            panelSAP.Size = new Size(599, 478);
            panelSAP.TabIndex = 0;
            // 
            // FormSAP
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(689, 569);
            Controls.Add(panelSAP);
            Name = "FormSAP";
            Text = "FormSAP";
            ResumeLayout(false);
        }

        #endregion

        private Panel panelSAP;
    }
}