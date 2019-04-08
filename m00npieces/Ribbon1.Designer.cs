namespace m00npieces
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
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
            this.group1 = this.Factory.CreateRibbonGroup();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btnSwap = this.Factory.CreateRibbonButton();
            this.btnMatchSize = this.Factory.CreateRibbonButton();
            this.btnTL = this.Factory.CreateRibbonToggleButton();
            this.btnML = this.Factory.CreateRibbonToggleButton();
            this.btnBL = this.Factory.CreateRibbonToggleButton();
            this.btnTC = this.Factory.CreateRibbonToggleButton();
            this.btnMC = this.Factory.CreateRibbonToggleButton();
            this.btnBC = this.Factory.CreateRibbonToggleButton();
            this.btnTR = this.Factory.CreateRibbonToggleButton();
            this.btnMR = this.Factory.CreateRibbonToggleButton();
            this.btnBR = this.Factory.CreateRibbonToggleButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "m00npieces";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnSwap);
            this.group1.Items.Add(this.btnMatchSize);
            this.group1.Label = "Swap";
            this.group1.Name = "group1";
            // 
            // group2
            // 
            this.group2.Items.Add(this.btnTL);
            this.group2.Items.Add(this.btnML);
            this.group2.Items.Add(this.btnBL);
            this.group2.Items.Add(this.btnTC);
            this.group2.Items.Add(this.btnMC);
            this.group2.Items.Add(this.btnBC);
            this.group2.Items.Add(this.btnTR);
            this.group2.Items.Add(this.btnMR);
            this.group2.Items.Add(this.btnBR);
            this.group2.Label = "Anchor";
            this.group2.Name = "group2";
            // 
            // btnSwap
            // 
            this.btnSwap.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSwap.Image = global::m00npieces.Properties.Resources.swap__1_;
            this.btnSwap.Label = "교체";
            this.btnSwap.Name = "btnSwap";
            this.btnSwap.ShowImage = true;
            this.btnSwap.SuperTip = "2개를 선택해야 합니다.";
            this.btnSwap.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSwap_Clicked);
            // 
            // btnMatchSize
            // 
            this.btnMatchSize.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnMatchSize.Image = global::m00npieces.Properties.Resources.matchsize;
            this.btnMatchSize.Label = "크기맞춤";
            this.btnMatchSize.Name = "btnMatchSize";
            this.btnMatchSize.ShowImage = true;
            this.btnMatchSize.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnMatchSize_Click);
            // 
            // btnTL
            // 
            this.btnTL.Label = "○";
            this.btnTL.Name = "btnTL";
            // 
            // btnML
            // 
            this.btnML.Label = "○";
            this.btnML.Name = "btnML";
            // 
            // btnBL
            // 
            this.btnBL.Label = "○";
            this.btnBL.Name = "btnBL";
            // 
            // btnTC
            // 
            this.btnTC.Label = "○";
            this.btnTC.Name = "btnTC";
            // 
            // btnMC
            // 
            this.btnMC.Label = "○";
            this.btnMC.Name = "btnMC";
            // 
            // btnBC
            // 
            this.btnBC.Label = "○";
            this.btnBC.Name = "btnBC";
            // 
            // btnTR
            // 
            this.btnTR.Label = "○";
            this.btnTR.Name = "btnTR";
            // 
            // btnMR
            // 
            this.btnMR.Label = "○";
            this.btnMR.Name = "btnMR";
            // 
            // btnBR
            // 
            this.btnBR.Label = "○";
            this.btnBR.Name = "btnBR";
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSwap;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMatchSize;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnML;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnTL;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnBL;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnTC;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnMC;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnBC;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnTR;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnMR;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnBR;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
