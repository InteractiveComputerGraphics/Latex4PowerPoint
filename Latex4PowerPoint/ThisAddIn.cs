using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace Latex4PowerPoint
{
    public partial class ThisAddIn
    {
        static private ThisAddIn m_current;
        static public ThisAddIn Current
        {
            get { return m_current; }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            m_current = this;

            AddinUtilities.initFonts();
            AddinUtilities.copyLatexTemplate("LatexTemplate.txt", Properties.Resources.LatexTemplate);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
