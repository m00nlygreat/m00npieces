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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.edtGoToSlide = this.Factory.CreateRibbonEditBox();
            this.ebxName = this.Factory.CreateRibbonEditBox();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btnTL = this.Factory.CreateRibbonToggleButton();
            this.btnML = this.Factory.CreateRibbonToggleButton();
            this.btnBL = this.Factory.CreateRibbonToggleButton();
            this.btnTC = this.Factory.CreateRibbonToggleButton();
            this.btnMC = this.Factory.CreateRibbonToggleButton();
            this.btnBC = this.Factory.CreateRibbonToggleButton();
            this.btnTR = this.Factory.CreateRibbonToggleButton();
            this.btnMR = this.Factory.CreateRibbonToggleButton();
            this.btnBR = this.Factory.CreateRibbonToggleButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.btnSwap = this.Factory.CreateRibbonButton();
            this.btnMatchSize = this.Factory.CreateRibbonButton();
            this.btn_Expand = this.Factory.CreateRibbonButton();
            this.btnGather = this.Factory.CreateRibbonButton();
            this.btnSync = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.btnAdjoinHorizontal = this.Factory.CreateRibbonButton();
            this.btnAdjoinVertical = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.btnFontAntiAlias = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group4.SuspendLayout();
            this.group3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Groups.Add(this.group3);
            this.tab1.KeyTip = "X";
            this.tab1.Label = "m00npieces";
            this.tab1.Name = "tab1";
            this.tab1.Position = this.Factory.RibbonPosition.BeforeOfficeId("TabHome");
            // 
            // group1
            // 
            this.group1.Items.Add(this.edtGoToSlide);
            this.group1.Items.Add(this.ebxName);
            this.group1.Label = "General";
            this.group1.Name = "group1";
            // 
            // edtGoToSlide
            // 
            this.edtGoToSlide.KeyTip = "X";
            this.edtGoToSlide.Label = "Go to";
            this.edtGoToSlide.Name = "edtGoToSlide";
            this.edtGoToSlide.SizeString = "9999";
            this.edtGoToSlide.Text = null;
            this.edtGoToSlide.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.EdtGoToSlide_changed);
            // 
            // ebxName
            // 
            this.ebxName.Label = "이름";
            this.ebxName.Name = "ebxName";
            this.ebxName.SizeString = "12345678901234";
            this.ebxName.Text = null;
            this.ebxName.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.EbxName_TextChanged);
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
            this.group2.KeyTip = "S";
            this.group2.Label = "Anchor";
            this.group2.Name = "group2";
            // 
            // btnTL
            // 
            this.btnTL.Checked = true;
            this.btnTL.KeyTip = "1";
            this.btnTL.Label = "◇";
            this.btnTL.Name = "btnTL";
            this.btnTL.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnTL_Click);
            // 
            // btnML
            // 
            this.btnML.KeyTip = "4";
            this.btnML.Label = "◇";
            this.btnML.Name = "btnML";
            this.btnML.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnML_Click);
            // 
            // btnBL
            // 
            this.btnBL.KeyTip = "7";
            this.btnBL.Label = "◇";
            this.btnBL.Name = "btnBL";
            this.btnBL.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnBL_Click);
            // 
            // btnTC
            // 
            this.btnTC.KeyTip = "2";
            this.btnTC.Label = "◇";
            this.btnTC.Name = "btnTC";
            this.btnTC.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnTC_Click);
            // 
            // btnMC
            // 
            this.btnMC.KeyTip = "5";
            this.btnMC.Label = "◇";
            this.btnMC.Name = "btnMC";
            this.btnMC.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnMC_Click);
            // 
            // btnBC
            // 
            this.btnBC.KeyTip = "8";
            this.btnBC.Label = "◇";
            this.btnBC.Name = "btnBC";
            this.btnBC.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnBC_Click);
            // 
            // btnTR
            // 
            this.btnTR.KeyTip = "3";
            this.btnTR.Label = "◇";
            this.btnTR.Name = "btnTR";
            this.btnTR.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnTR_Click);
            // 
            // btnMR
            // 
            this.btnMR.KeyTip = "6";
            this.btnMR.Label = "◇";
            this.btnMR.Name = "btnMR";
            this.btnMR.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnMR_Click);
            // 
            // btnBR
            // 
            this.btnBR.KeyTip = "9";
            this.btnBR.Label = "◇";
            this.btnBR.Name = "btnBR";
            this.btnBR.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnBR_Click);
            // 
            // group4
            // 
            this.group4.Items.Add(this.btnSwap);
            this.group4.Items.Add(this.btnMatchSize);
            this.group4.Items.Add(this.btn_Expand);
            this.group4.Items.Add(this.btnGather);
            this.group4.Items.Add(this.btnSync);
            this.group4.Items.Add(this.separator1);
            this.group4.Items.Add(this.btnAdjoinHorizontal);
            this.group4.Items.Add(this.btnAdjoinVertical);
            this.group4.Label = "Shape";
            this.group4.Name = "group4";
            // 
            // btnSwap
            // 
            this.btnSwap.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSwap.Image = global::m00npieces.Properties.Resources.swap;
            this.btnSwap.KeyTip = "C";
            this.btnSwap.Label = "교체";
            this.btnSwap.Name = "btnSwap";
            this.btnSwap.ShowImage = true;
            this.btnSwap.SuperTip = "2개를 선택해야 합니다.";
            this.btnSwap.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSwap_Clicked);
            // 
            // btnMatchSize
            // 
            this.btnMatchSize.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnMatchSize.Image = global::m00npieces.Properties.Resources.expand;
            this.btnMatchSize.KeyTip = "A";
            this.btnMatchSize.Label = "크기맞춤";
            this.btnMatchSize.Name = "btnMatchSize";
            this.btnMatchSize.ShowImage = true;
            this.btnMatchSize.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnMatchSize_Click);
            // 
            // btn_Expand
            // 
            this.btn_Expand.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_Expand.Image = ((System.Drawing.Image)(resources.GetObject("btn_Expand.Image")));
            this.btn_Expand.KeyTip = "E";
            this.btn_Expand.Label = "늘이기";
            this.btn_Expand.Name = "btn_Expand";
            this.btn_Expand.ShowImage = true;
            this.btn_Expand.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Btn_Expand_Click);
            // 
            // btnGather
            // 
            this.btnGather.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnGather.Image = global::m00npieces.Properties.Resources.gather;
            this.btnGather.Label = "모으기";
            this.btnGather.Name = "btnGather";
            this.btnGather.ShowImage = true;
            this.btnGather.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnGather_Click);
            // 
            // btnSync
            // 
            this.btnSync.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSync.Label = "동기화";
            this.btnSync.Name = "btnSync";
            this.btnSync.ShowImage = true;
            this.btnSync.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnSync_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // btnAdjoinHorizontal
            // 
            this.btnAdjoinHorizontal.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnAdjoinHorizontal.Image = global::m00npieces.Properties.Resources.adjoinhorizontal;
            this.btnAdjoinHorizontal.KeyTip = "H";
            this.btnAdjoinHorizontal.Label = "가로로 붙이기";
            this.btnAdjoinHorizontal.Name = "btnAdjoinHorizontal";
            this.btnAdjoinHorizontal.ShowImage = true;
            this.btnAdjoinHorizontal.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAdjoinHorizontal_Clicked);
            // 
            // btnAdjoinVertical
            // 
            this.btnAdjoinVertical.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnAdjoinVertical.Image = global::m00npieces.Properties.Resources.adjoinvertical;
            this.btnAdjoinVertical.KeyTip = "V";
            this.btnAdjoinVertical.Label = "세로로 붙이기";
            this.btnAdjoinVertical.Name = "btnAdjoinVertical";
            this.btnAdjoinVertical.ShowImage = true;
            this.btnAdjoinVertical.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnAdjoinVertical_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.btnFontAntiAlias);
            this.group3.Label = "et cetra";
            this.group3.Name = "group3";
            // 
            // btnFontAntiAlias
            // 
            this.btnFontAntiAlias.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnFontAntiAlias.Image = global::m00npieces.Properties.Resources.glit;
            this.btnFontAntiAlias.Label = "글씨를 예쁘게";
            this.btnFontAntiAlias.Name = "btnFontAntiAlias";
            this.btnFontAntiAlias.ShowImage = true;
            this.btnFontAntiAlias.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFontAntiAlias_Clicked);
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
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSwap;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnML;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnTL;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnBL;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnTC;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnMC;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnBC;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnMR;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnTR;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnBR;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFontAntiAlias;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox edtGoToSlide;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAdjoinHorizontal;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAdjoinVertical;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_Expand;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGather;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox ebxName;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSync;
        public Microsoft.Office.Tools.Ribbon.RibbonButton btnMatchSize;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
