namespace WordAddIn
{
    partial class WordRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public WordRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.rbimage = this.Factory.CreateRibbonGroup();
            this.btnimage = this.Factory.CreateRibbonButton();
            this.btnAddImage = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.rbimage.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.rbimage);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // rbimage
            // 
            this.rbimage.Items.Add(this.btnimage);
            this.rbimage.Items.Add(this.btnAddImage);
            this.rbimage.Items.Add(this.button2);
            this.rbimage.Name = "rbimage";
            // 
            // btnimage
            // 
            this.btnimage.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnimage.Image = global::WordAddIn.Properties.Resources.pdf1;
            this.btnimage.ImageName = "Export to PDF";
            this.btnimage.Label = "Export Pdf";
            this.btnimage.Name = "btnimage";
            this.btnimage.ShowImage = true;
            this.btnimage.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // btnAddImage
            // 
            this.btnAddImage.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnAddImage.Image = global::WordAddIn.Properties.Resources.canon;
            this.btnAddImage.Label = "Add Image";
            this.btnAddImage.Name = "btnAddImage";
            this.btnAddImage.ShowImage = true;
            this.btnAddImage.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddImage_Click);
            // 
            // button2
            // 
            this.button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button2.Image = global::WordAddIn.Properties.Resources.table;
            this.button2.Label = "Add Table";
            this.button2.Name = "button2";
            this.button2.ShowImage = true;
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // WordRibbon
            // 
            this.Name = "WordRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.WordRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.rbimage.ResumeLayout(false);
            this.rbimage.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup rbimage;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnimage;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddImage;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
    }

    partial class ThisRibbonCollection
    {
        internal WordRibbon WordRibbon
        {
            get { return this.GetRibbon<WordRibbon>(); }
        }
    }
}
