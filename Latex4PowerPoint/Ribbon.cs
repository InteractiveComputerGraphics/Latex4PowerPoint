using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Office = Microsoft.Office.Core;
using Tools = Microsoft.Office.Tools;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Windows.Forms;
using System.IO;

namespace Latex4PowerPoint
{
    public partial class Ribbon
    {
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        private LatexDialog m_dialog;

        public Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
            m_dialog = new LatexDialog();

            SettingsManager mgr = SettingsManager.getCurrent();
            checkBoxCursor.Checked = mgr.SettingsData.insertAtCursor;
        }

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void buttonLatex_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Selection sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            bool found = false;
            if (sel.Type == Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                foreach (PowerPoint.Shape s in sel.ShapeRange)
                {
                    if (editLatexObject(s))
                        found = true;
                }
            }
            if (!found)
                createLatexObject(null, "Create latex object");
        }


        public void createLatexObject(string template, string title)
        {
            // get current slide
            int slideId = ThisAddIn.Current.Application.ActiveWindow.Selection.SlideRange.SlideID;
            Microsoft.Office.Interop.PowerPoint.Slide slide = ThisAddIn.Current.Application.ActivePresentation.Slides.FindBySlideID(slideId);

            List<string> args = new List<string>();
            bool useTemplate = template != null;

            m_dialog.init(template, useTemplate, title);

            // We cannot use the usual dialog result, since it is not correct, when hiding the window
            IntPtr hwnd = GetForegroundWindow();
            m_dialog.ShowDialog();
            SetForegroundWindow(hwnd);
        }

        public bool editLatexObject(PowerPoint.Shape s)
        {
            LatexEquation eq = ShapeTags.getLatexEquation(s);
            if (eq != null)
            {
                m_dialog.init(eq, "Edit latex object");

                // We cannot use the usual dialog result, since it is not correct, when hiding the window
                IntPtr hwnd = GetForegroundWindow();
                m_dialog.ShowDialog();
                SetForegroundWindow(hwnd);
                if (m_dialog.Result == DialogResult.OK)
                {
                    Microsoft.Office.Interop.PowerPoint.Shape latexObj = m_dialog.LatexEquation.m_shape;
                    if (latexObj != null)
                    {
                        float left = s.Left;
                        float top = s.Top;

                        AddinUtilities.scaleEditedLatexObject(s, latexObj, m_dialog.FontSize);

                        // Copy values of old shape
                        latexObj.Left = left;
                        latexObj.Top = top;
                        latexObj.Rotation = s.Rotation;

                        // Remove old shape
                        s.Delete();
                    }
                }
                return true;
            }
            return false;
        }


        private void buttonNewEquation_Click(object sender, RibbonControlEventArgs e)
        {
            string template = "\\begin{equation*}\r\n\t<Enter latex code>\r\n\\end{equation*}\r\n";
            createLatexObject(template, "Create latex equation");
        }

        private void buttonNewEqnArray_Click(object sender, RibbonControlEventArgs e)
        {
            string template = "\\begin{eqnarray*}\r\n\t<Enter latex code>\r\n\\end{eqnarray*}\r\n";
            createLatexObject(template, "Create latex equation array");
        }

        private void checkBoxCursor_Click(object sender, RibbonControlEventArgs e)
        {
            SettingsManager mgr = SettingsManager.getCurrent();
            mgr.SettingsData.insertAtCursor = checkBoxCursor.Checked;
            mgr.saveSettings();
        }

        private void buttonMiktex_Click(object sender, RibbonControlEventArgs e)
        {
            AddinUtilities.changeOptions();
        }
    }
}
