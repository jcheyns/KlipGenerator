namespace KlipGenerator
{
    partial class RibbonKlip : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonKlip()
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
            this.grpKlipGenerator = this.Factory.CreateRibbonGroup();
            this.btnGenerate = this.Factory.CreateRibbonButton();
            this.btnGenerateFromSelection = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.grpKlipGenerator.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.grpKlipGenerator);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // grpKlipGenerator
            // 
            this.grpKlipGenerator.Items.Add(this.btnGenerate);
            this.grpKlipGenerator.Items.Add(this.btnGenerateFromSelection);
            this.grpKlipGenerator.Label = "KlipGen";
            this.grpKlipGenerator.Name = "grpKlipGenerator";
            // 
            // btnGenerate
            // 
            this.btnGenerate.Label = "Generate All";
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnGenerate_Click);
            // 
            // btnGenerateFromSelection
            // 
            this.btnGenerateFromSelection.Label = "Generate Selected";
            this.btnGenerateFromSelection.Name = "btnGenerateFromSelection";
            this.btnGenerateFromSelection.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnGenerateFromSelection_Click);
            // 
            // RibbonKlip
            // 
            this.Name = "RibbonKlip";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.grpKlipGenerator.ResumeLayout(false);
            this.grpKlipGenerator.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpKlipGenerator;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGenerate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGenerateFromSelection;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonKlip Ribbon1
        {
            get { return this.GetRibbon<RibbonKlip>(); }
        }
    }
}
