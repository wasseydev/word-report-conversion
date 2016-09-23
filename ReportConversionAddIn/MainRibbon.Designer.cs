namespace WordReportingTool
{
    partial class MainRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MainRibbon()
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
            this.CreateJRXML = this.Factory.CreateRibbonButton();
            this.ConversionTab.SuspendLayout();
            this.conversionGroup.SuspendLayout();
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
            this.conversionGroup.Items.Add(this.CreateJRXML);
            this.conversionGroup.Label = "Conversion options";
            this.conversionGroup.Name = "conversionGroup";
            // 
            // CreateJRXML
            // 
            this.CreateJRXML.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.CreateJRXML.Image = global::WordReportingTool.Properties.Resources.optionJrxml;
            this.CreateJRXML.Label = "Jasper Reports";
            this.CreateJRXML.Name = "CreateJRXML";
            this.CreateJRXML.ScreenTip = "Create a Jasper Reports JRXML file";
            this.CreateJRXML.ShowImage = true;
            this.CreateJRXML.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CreateJRXML_Click);
            // 
            // WordReportingRibbon
            // 
            this.Name = "WordReportingRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.ConversionTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MainRibbon_Load);
            this.ConversionTab.ResumeLayout(false);
            this.ConversionTab.PerformLayout();
            this.conversionGroup.ResumeLayout(false);
            this.conversionGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab ConversionTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup conversionGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CreateJRXML;
    }

    partial class ThisRibbonCollection
    {
        internal MainRibbon MainRibbon
        {
            get { return this.GetRibbon<MainRibbon>(); }
        }
    }
}
