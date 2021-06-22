
namespace AttachmentSaver
{
    partial class _Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public _Ribbon()
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
            this.autoCheck = this.Factory.CreateRibbonToggleButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.runBtn = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.manageProfilesBtn = this.Factory.CreateRibbonButton();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.setLastRunBtn = this.Factory.CreateRibbonButton();
            this.viewLogBtn = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "Bijlage-import";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.autoCheck);
            this.group1.Items.Add(this.separator1);
            this.group1.Items.Add(this.runBtn);
            this.group1.Items.Add(this.separator2);
            this.group1.Items.Add(this.manageProfilesBtn);
            this.group1.Items.Add(this.separator3);
            this.group1.Items.Add(this.setLastRunBtn);
            this.group1.Items.Add(this.viewLogBtn);
            this.group1.Name = "group1";
            // 
            // autoCheck
            // 
            this.autoCheck.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.autoCheck.Image = global::AttachmentSaver.Properties.Resources.switchOff;
            this.autoCheck.Label = "Automatisch uitvoeren";
            this.autoCheck.Name = "autoCheck";
            this.autoCheck.ShowImage = true;
            this.autoCheck.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AutoCheck_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // runBtn
            // 
            this.runBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.runBtn.Image = global::AttachmentSaver.Properties.Resources.play;
            this.runBtn.Label = "Importeer bijlagen";
            this.runBtn.Name = "runBtn";
            this.runBtn.ShowImage = true;
            this.runBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RunBtn_Click);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // manageProfilesBtn
            // 
            this.manageProfilesBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.manageProfilesBtn.Image = global::AttachmentSaver.Properties.Resources.settings;
            this.manageProfilesBtn.Label = "Opslagprofielen...";
            this.manageProfilesBtn.Name = "manageProfilesBtn";
            this.manageProfilesBtn.ShowImage = true;
            this.manageProfilesBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ManageProfiles_Click);
            // 
            // separator3
            // 
            this.separator3.Name = "separator3";
            // 
            // setLastRunBtn
            // 
            this.setLastRunBtn.Label = "Laatst geïmporteerd om: -";
            this.setLastRunBtn.Name = "setLastRunBtn";
            this.setLastRunBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.setLastRunBtn_Click);
            // 
            // viewLogBtn
            // 
            this.viewLogBtn.Label = "Bekijk Log";
            this.viewLogBtn.Name = "viewLogBtn";
            this.viewLogBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ViewLogBtn_Click);
            // 
            // _Ribbon
            // 
            this.Name = "_Ribbon";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this._Ribbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton runBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton manageProfilesBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton autoCheck;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton viewLogBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton setLastRunBtn;
    }

    partial class ThisRibbonCollection
    {
        internal _Ribbon _Ribbon
        {
            get { return this.GetRibbon<_Ribbon>(); }
        }
    }
}
