namespace WordOpenXml_cs
{
    partial class Form1
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.cmdCreateNewEmptyDocument = new System.Windows.Forms.Button();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.cmdCreateNewDocumentWithSimpleParagraph = new System.Windows.Forms.Button();
            this.cmdCreateNewDocumentWithUnoderedList = new System.Windows.Forms.Button();
            this.cmdCreateNewDocumentWithMultipleParagraphAndImage = new System.Windows.Forms.Button();
            this.cmdCreateNewDocumentWithTableFromArray = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // cmdCreateNewEmptyDocument
            // 
            this.cmdCreateNewEmptyDocument.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.cmdCreateNewEmptyDocument.ImageKey = "(none)";
            this.cmdCreateNewEmptyDocument.ImageList = this.imageList1;
            this.cmdCreateNewEmptyDocument.Location = new System.Drawing.Point(10, 14);
            this.cmdCreateNewEmptyDocument.Name = "cmdCreateNewEmptyDocument";
            this.cmdCreateNewEmptyDocument.Size = new System.Drawing.Size(179, 23);
            this.cmdCreateNewEmptyDocument.TabIndex = 0;
            this.cmdCreateNewEmptyDocument.Text = "Empty document";
            this.cmdCreateNewEmptyDocument.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.cmdCreateNewEmptyDocument.UseVisualStyleBackColor = true;
            this.cmdCreateNewEmptyDocument.Click += new System.EventHandler(this.cmdCreateNewEmptyDocument_Click);
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "Success");
            this.imageList1.Images.SetKeyName(1, "Failed");
            // 
            // cmdCreateNewDocumentWithSimpleParagraph
            // 
            this.cmdCreateNewDocumentWithSimpleParagraph.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.cmdCreateNewDocumentWithSimpleParagraph.ImageKey = "(none)";
            this.cmdCreateNewDocumentWithSimpleParagraph.ImageList = this.imageList1;
            this.cmdCreateNewDocumentWithSimpleParagraph.Location = new System.Drawing.Point(10, 43);
            this.cmdCreateNewDocumentWithSimpleParagraph.Name = "cmdCreateNewDocumentWithSimpleParagraph";
            this.cmdCreateNewDocumentWithSimpleParagraph.Size = new System.Drawing.Size(179, 23);
            this.cmdCreateNewDocumentWithSimpleParagraph.TabIndex = 1;
            this.cmdCreateNewDocumentWithSimpleParagraph.Text = "Simple paragraph";
            this.cmdCreateNewDocumentWithSimpleParagraph.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.cmdCreateNewDocumentWithSimpleParagraph.UseVisualStyleBackColor = true;
            this.cmdCreateNewDocumentWithSimpleParagraph.Click += new System.EventHandler(this.cmdCreateNewDocumentWithSimpleParagraph_Click);
            // 
            // cmdCreateNewDocumentWithUnoderedList
            // 
            this.cmdCreateNewDocumentWithUnoderedList.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.cmdCreateNewDocumentWithUnoderedList.ImageKey = "(none)";
            this.cmdCreateNewDocumentWithUnoderedList.ImageList = this.imageList1;
            this.cmdCreateNewDocumentWithUnoderedList.Location = new System.Drawing.Point(10, 72);
            this.cmdCreateNewDocumentWithUnoderedList.Name = "cmdCreateNewDocumentWithUnoderedList";
            this.cmdCreateNewDocumentWithUnoderedList.Size = new System.Drawing.Size(179, 23);
            this.cmdCreateNewDocumentWithUnoderedList.TabIndex = 2;
            this.cmdCreateNewDocumentWithUnoderedList.Text = "Unordered list (document writer)";
            this.cmdCreateNewDocumentWithUnoderedList.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.cmdCreateNewDocumentWithUnoderedList.UseVisualStyleBackColor = true;
            this.cmdCreateNewDocumentWithUnoderedList.Click += new System.EventHandler(this.cmdCreateNewDocumentWithUnoderedList_Click);
            // 
            // cmdCreateNewDocumentWithMultipleParagraphAndImage
            // 
            this.cmdCreateNewDocumentWithMultipleParagraphAndImage.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.cmdCreateNewDocumentWithMultipleParagraphAndImage.ImageKey = "(none)";
            this.cmdCreateNewDocumentWithMultipleParagraphAndImage.ImageList = this.imageList1;
            this.cmdCreateNewDocumentWithMultipleParagraphAndImage.Location = new System.Drawing.Point(10, 101);
            this.cmdCreateNewDocumentWithMultipleParagraphAndImage.Name = "cmdCreateNewDocumentWithMultipleParagraphAndImage";
            this.cmdCreateNewDocumentWithMultipleParagraphAndImage.Size = new System.Drawing.Size(179, 23);
            this.cmdCreateNewDocumentWithMultipleParagraphAndImage.TabIndex = 3;
            this.cmdCreateNewDocumentWithMultipleParagraphAndImage.Text = "Paragraphs and image";
            this.cmdCreateNewDocumentWithMultipleParagraphAndImage.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.cmdCreateNewDocumentWithMultipleParagraphAndImage.UseVisualStyleBackColor = true;
            this.cmdCreateNewDocumentWithMultipleParagraphAndImage.Click += new System.EventHandler(this.cmdCreateNewDocumentWithMultipleParagraphAndImage_Click);
            // 
            // cmdCreateNewDocumentWithTableFromArray
            // 
            this.cmdCreateNewDocumentWithTableFromArray.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.cmdCreateNewDocumentWithTableFromArray.ImageKey = "(none)";
            this.cmdCreateNewDocumentWithTableFromArray.ImageList = this.imageList1;
            this.cmdCreateNewDocumentWithTableFromArray.Location = new System.Drawing.Point(10, 130);
            this.cmdCreateNewDocumentWithTableFromArray.Name = "cmdCreateNewDocumentWithTableFromArray";
            this.cmdCreateNewDocumentWithTableFromArray.Size = new System.Drawing.Size(179, 23);
            this.cmdCreateNewDocumentWithTableFromArray.TabIndex = 4;
            this.cmdCreateNewDocumentWithTableFromArray.Text = "Table from array";
            this.cmdCreateNewDocumentWithTableFromArray.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.cmdCreateNewDocumentWithTableFromArray.UseVisualStyleBackColor = true;
            this.cmdCreateNewDocumentWithTableFromArray.Click += new System.EventHandler(this.cmdCreateNewDocumentWithTableFromArray_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(203, 173);
            this.Controls.Add(this.cmdCreateNewDocumentWithTableFromArray);
            this.Controls.Add(this.cmdCreateNewDocumentWithMultipleParagraphAndImage);
            this.Controls.Add(this.cmdCreateNewDocumentWithUnoderedList);
            this.Controls.Add(this.cmdCreateNewDocumentWithSimpleParagraph);
            this.Controls.Add(this.cmdCreateNewEmptyDocument);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Word code samples";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button cmdCreateNewEmptyDocument;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.Button cmdCreateNewDocumentWithSimpleParagraph;
        private System.Windows.Forms.Button cmdCreateNewDocumentWithUnoderedList;
        private System.Windows.Forms.Button cmdCreateNewDocumentWithMultipleParagraphAndImage;
        private System.Windows.Forms.Button cmdCreateNewDocumentWithTableFromArray;
    }
}

