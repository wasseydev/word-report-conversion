namespace WordReportingTool
{
    partial class WordReportingRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public WordReportingRibbon()
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
            this.ConversionTab = this.Factory.CreateRibbonTab();
            this.conversionGroup = this.Factory.CreateRibbonGroup();
            this.ConversionTab.SuspendLayout();
            this.SuspendLayout();
            // 
            // ConversionTab
            // 
            this.ConversionTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.ConversionTab.Groups.Add(this.conversionGroup);
            this.ConversionTab.Label = "Conversion";
            this.ConversionTab.Name = "ConversionTab";
            // 
            // conversionGroup
            // 
            this.conversionGroup.Label = "Conversion options";
            this.conversionGroup.Name = "conversionGroup";
            // 
            // WordReportingRibbon
            // 
            this.Name = "WordReportingRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.ConversionTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.WordReportingRibbon_Load);
            this.ConversionTab.ResumeLayout(false);
            this.ConversionTab.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab ConversionTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup conversionGroup;
    }

    partial class ThisRibbonCollection
    {
        internal WordReportingRibbon WordReportingRibbon
        {
            get { return this.GetRibbon<WordReportingRibbon>(); }
        }
    }
}
