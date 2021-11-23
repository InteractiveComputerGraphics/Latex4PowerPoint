namespace Latex4PowerPoint
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
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

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tabLatex = this.Factory.CreateRibbonTab();
            this.groupLatex = this.Factory.CreateRibbonGroup();
            this.buttonLatex = this.Factory.CreateRibbonButton();
            this.buttonNewEquation = this.Factory.CreateRibbonButton();
            this.buttonNewEqnArray = this.Factory.CreateRibbonButton();
            this.groupSettings = this.Factory.CreateRibbonGroup();
            this.checkBoxCursor = this.Factory.CreateRibbonCheckBox();
            this.buttonMiktex = this.Factory.CreateRibbonButton();
            this.buttonInvert = this.Factory.CreateRibbonButton();
            this.tabLatex.SuspendLayout();
            this.groupLatex.SuspendLayout();
            this.groupSettings.SuspendLayout();
            // 
            // tabLatex
            // 
            this.tabLatex.Groups.Add(this.groupLatex);
            this.tabLatex.Groups.Add(this.groupSettings);
            this.tabLatex.KeyTip = "X";
            this.tabLatex.Label = "Latex";
            this.tabLatex.Name = "tabLatex";
            // 
            // groupLatex
            // 
            this.groupLatex.Items.Add(this.buttonLatex);
            this.groupLatex.Items.Add(this.buttonNewEquation);
            this.groupLatex.Items.Add(this.buttonNewEqnArray);
            this.groupLatex.Label = "Latex";
            this.groupLatex.Name = "groupLatex";
            // 
            // buttonLatex
            // 
            this.buttonLatex.KeyTip = "X";
            this.buttonLatex.Label = "Latex";
            this.buttonLatex.Name = "buttonLatex";
            this.buttonLatex.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonLatex_Click);
            // 
            // buttonNewEquation
            // 
            this.buttonNewEquation.Label = "New equation";
            this.buttonNewEquation.Name = "buttonNewEquation";
            this.buttonNewEquation.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonNewEquation_Click);
            // 
            // buttonNewEqnArray
            // 
            this.buttonNewEqnArray.Label = "New equation array";
            this.buttonNewEqnArray.Name = "buttonNewEqnArray";
            this.buttonNewEqnArray.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonNewEqnArray_Click);
            // 
            // groupSettings
            // 
            this.groupSettings.Items.Add(this.checkBoxCursor);
            this.groupSettings.Items.Add(this.buttonMiktex);
            this.groupSettings.Items.Add(this.buttonInvert);
            this.groupSettings.Label = "Settings";
            this.groupSettings.Name = "groupSettings";
            // 
            // checkBoxCursor
            // 
            this.checkBoxCursor.Label = "Insert at cursor";
            this.checkBoxCursor.Name = "checkBoxCursor";
            this.checkBoxCursor.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBoxCursor_Click);
            // 
            // buttonMiktex
            // 
            this.buttonMiktex.Label = "Set MiKTeX location";
            this.buttonMiktex.Name = "buttonMiktex";
            this.buttonMiktex.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonMiktex_Click);
            // 
            // buttonInvert
            // 
            this.buttonInvert.Label = "Invert images";
            this.buttonInvert.Name = "buttonInvert";
            this.buttonInvert.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonInvert_Click);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tabLatex);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tabLatex.ResumeLayout(false);
            this.tabLatex.PerformLayout();
            this.groupLatex.ResumeLayout(false);
            this.groupLatex.PerformLayout();
            this.groupSettings.ResumeLayout(false);
            this.groupSettings.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabLatex;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupLatex;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonLatex;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonNewEquation;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonNewEqnArray;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonInvert;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBoxCursor;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonMiktex;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon1
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
