﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;

namespace Latex4PowerPoint
{    
    public partial class LatexDialog : Form
    {
        private string m_textColor;
        public string TextColor
        {
            get { return m_textColor; }
        }

        public string FontSize 
        {
            get { return comboBoxFontSize.Text; }
        }

        public string LatexFont
        {
            get { return comboBoxFont.Text; }
        }

        public string FontSeries
        {
            get { return comboBoxSeries.Text; }
        }

        public string FontShape
        {
            get { return comboBoxShape.Text; }
        }

        public string DPI
        {
            get { return comboBoxDPI.Text; }
        }

        public string LatexCode
        {
            get { return m_scintilla.Text; }
        }

        private DialogResult m_result;
        public System.Windows.Forms.DialogResult Result
        {
            get { return m_result; }
        }

        private LatexEquation m_latexEquation;
        public Latex4PowerPoint.LatexEquation LatexEquation
        {
            get { return m_latexEquation; }
            set { m_latexEquation = value; }
        }
        

        private bool m_finishedSuccessfully;

        private ScintillaNET.Scintilla m_scintilla;
        private TextRange m_textRange;

        public LatexDialog()
        {
            InitializeComponent();

            // Be sure, there is a language.xml file
            AddinUtilities.copyLanguageFile();

            this.SuspendLayout();
            m_scintilla = new ScintillaNET.Scintilla();
            this.groupBoxLatex.Controls.Add(m_scintilla);
            m_scintilla.Dock = DockStyle.Fill;
            m_scintilla.Margins[0].Width = 20;
            m_scintilla.ConfigurationManager.CustomLocation = AddinUtilities.getAppDataLocation() + "\\Language.xml";
            m_scintilla.ConfigurationManager.Language = "mytex";
            m_scintilla.IsBraceMatching = true;
            m_scintilla.TabIndex = 0;
            m_scintilla.AutoComplete.DropRestOfWord = true;
            this.ResumeLayout(false);
            m_scintilla.Focus();
            m_scintilla.KeyDown += new KeyEventHandler(m_scintilla_KeyDown);

            createFontEntries();

            m_finishedSuccessfully = false;
            this.FormClosing += new FormClosingEventHandler(LatexDialog_FormClosing);
        }

        void m_scintilla_KeyDown(object sender, KeyEventArgs e)
        {
            if ((e.KeyCode == Keys.Enter) && (e.Modifiers == Keys.Control))
            {
                buttonOk.PerformClick();
            }
        }

        private void init(string title)
        {
            this.Text = title;
            m_finishedSuccessfully = false;

            m_scintilla.Text = "";
            m_scintilla.UndoRedo.EmptyUndoBuffer();

            SettingsManager mgr = SettingsManager.getCurrent();
            if (!comboBoxFontSize.Items.Contains(mgr.SettingsData.fontSize))
                comboBoxFontSize.Items.Add(mgr.SettingsData.fontSize);
            comboBoxFontSize.SelectedItem = mgr.SettingsData.fontSize;
            if (!comboBoxDPI.Items.Contains(mgr.SettingsData.dpi))
                comboBoxDPI.Items.Add(mgr.SettingsData.dpi);
            comboBoxDPI.Text = mgr.SettingsData.dpi;
            comboBoxFont.Text = mgr.SettingsData.font;
            comboBoxSeries.Text = mgr.SettingsData.fontSeries;
            comboBoxShape.Text = mgr.SettingsData.fontShape;
            buttonColor.BackColor = AddinUtilities.stringToColor(mgr.SettingsData.textColor);
            m_textColor = mgr.SettingsData.textColor;
            pictureBoxPreview.BackColor = Color.White;
            panel1.BackColor = Color.White;

            getCurrentTextRange();
        }

        public void init(LatexEquation eq, string title)
        {
            init(title);
            if (eq != null)
            {
                m_scintilla.Text = eq.m_code;
                m_scintilla.Selection.SelectAll();

                string fontSizeStr = eq.m_fontSize.ToString();
                if (!comboBoxFontSize.Items.Contains(fontSizeStr))
                    comboBoxFontSize.Items.Add(fontSizeStr);
                comboBoxFontSize.Text = fontSizeStr;

                string dpiStr = eq.m_dpi.ToString();
                if (!comboBoxDPI.Items.Contains(dpiStr))
                    comboBoxDPI.Items.Add(dpiStr);
                comboBoxDPI.Text = eq.m_dpi.ToString();

                comboBoxFont.Text = eq.m_font.fontName;
                comboBoxSeries.Text = eq.m_fontSeries.fontSeries;
                comboBoxShape.Text = eq.m_fontShape.fontShape;

                try
                {
                    buttonColor.BackColor = AddinUtilities.stringToColor(eq.m_color);
                    m_textColor = eq.m_color;
                }
                catch
                {
                }
            }
        }

        public void init(string template, bool useTemplate, string title)
        {
            init(title);
            if (template != null)
            {
                m_scintilla.Text = template;
                if (useTemplate)
                {
                    int index = m_scintilla.Text.IndexOf("<Enter latex code>", 0);
                    //m_scintilla.Select(index, 18);
                    m_scintilla.Selection.Start = index;
                    m_scintilla.Selection.End = index + 18;
                }
                m_scintilla.Selection.SelectAll();
            }
            getCurrentTextRange();
        }

        private void getCurrentTextRange()
        {
            m_textRange = null;
            SettingsManager mgr = SettingsManager.getCurrent();
            if (mgr.SettingsData.insertAtCursor)
            {
                if (ThisAddIn.Current.Application.ActiveWindow.Selection.Type == Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionText)
                {
                    int currentCursorPositionInCharacters = ThisAddIn.Current.Application.ActiveWindow.Selection.TextRange.Start;
                    if (currentCursorPositionInCharacters > 0)
                    {
                        m_textRange = ThisAddIn.Current.Application.ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Characters(currentCursorPositionInCharacters - 1, 1);
                    }
                }
            }
        }

        private void createFontEntries()
        {
            AddinUtilities.initFonts();
            comboBoxFont.Items.AddRange(AddinUtilities.LatexFonts.ToArray());
            comboBoxSeries.Items.AddRange(AddinUtilities.LatexFontSeries.ToArray());
            comboBoxShape.Items.AddRange(AddinUtilities.LatexFontShapes.ToArray());
        }

        void LatexDialog_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            if ((this.DialogResult == DialogResult.OK) && (!m_finishedSuccessfully))
                return;
            m_result = this.DialogResult;

            // Set the focus => next time we open the dialog it is correct
            m_scintilla.Focus();
            this.Hide();

            if (pictureBoxPreview.Image != null)
                pictureBoxPreview.Image.Dispose();
            pictureBoxPreview.Image = null;
        }
      

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void buttonOk_Click(object sender, EventArgs e)
        {
            generateEquation(true, false);
        }

        private bool generateEquation(bool createShape, bool useSystemDPI)
        {
            // Check paths
            SettingsManager mgr = SettingsManager.getCurrent();

            // Check font size
            string fontSize = comboBoxFontSize.Text;
            float size = 12;
            try
            {
                size = Convert.ToSingle(fontSize);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Font size exception: \n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            // Check Dpi
            float[] systemDPI = AddinUtilities.getSystemDPI();
            float dpiValue;
            if (useSystemDPI)
                dpiValue = systemDPI[0];
            else
                dpiValue = AddinUtilities.getFloat(comboBoxDPI.Text, systemDPI[0]);
            
            mgr.SettingsData.fontSize = comboBoxFontSize.Text;
            mgr.SettingsData.font = comboBoxFont.Text;
            mgr.SettingsData.fontSeries = comboBoxSeries.Text;
            mgr.SettingsData.fontShape = comboBoxShape.Text;
            mgr.SettingsData.textColor = m_textColor;
            mgr.SettingsData.dpi = comboBoxDPI.Text;
            mgr.saveSettings();

            m_latexEquation = new LatexEquation(m_scintilla.Text, size, dpiValue, m_textColor, (LatexFont)comboBoxFont.SelectedItem,
                                                      (LatexFontSeries)comboBoxSeries.SelectedItem,
                                                      (LatexFontShape)comboBoxShape.SelectedItem);


            m_finishedSuccessfully = AddinUtilities.createLatexPng(m_latexEquation);
            if (createShape)
            {
                AddinUtilities.createLatexShape("teximport.png", m_latexEquation);
                m_latexEquation.m_shape.Select(Microsoft.Office.Core.MsoTriState.msoTrue);
                AddinUtilities.centerLatexShape(m_latexEquation);
                alignLatexShape(m_latexEquation);
            }
            return m_finishedSuccessfully;
        }

        private void alignLatexShape(LatexEquation equation)
        {
            if (m_textRange != null)
            {
                //System.Drawing.FontFamily family = AddinUtilities.getFontFamily(m_textRange);
                ////float height = (float)(m_textRange.BoundHeight * ((float)family.GetCellAscent(System.Drawing.FontStyle.Regular) / (float)family.GetLineSpacing(System.Drawing.FontStyle.Regular)));
                //float lineSpacing = m_textRange.Font.Size*(float)family.GetLineSpacing(System.Drawing.FontStyle.Regular) / (float)family.GetEmHeight(System.Drawing.FontStyle.Regular);
                //float spaceBefore = 0.0f;
                //if (m_textRange.BoundHeight > lineSpacing)
                //    spaceBefore = m_textRange.ParagraphFormat.SpaceBefore * lineSpacing;
                //float height = spaceBefore + (m_textRange.Font.Size * (float)family.GetCellAscent(System.Drawing.FontStyle.Regular)) / (float)family.GetEmHeight(System.Drawing.FontStyle.Regular);
                //float offset = (float)equation.m_offset / (float)equation.m_imageHeight;

                //// Test
                //int slideId = ThisAddIn.Current.Application.ActiveWindow.Selection.SlideRange.SlideID;
                //Microsoft.Office.Interop.PowerPoint.Slide slide = ThisAddIn.Current.Application.ActivePresentation.Slides.FindBySlideID(slideId);
                //slide.Shapes.AddLine(m_textRange.BoundLeft, m_textRange.BoundTop, m_textRange.BoundLeft + m_textRange.BoundWidth, m_textRange.BoundTop + m_textRange.BoundHeight);
                //// Test                

                if (equation.m_shape != null)
                {
                    equation.m_shape.Left = m_textRange.BoundLeft + m_textRange.BoundWidth;
                    equation.m_shape.Top = m_textRange.BoundTop + m_textRange.BoundHeight - equation.m_shape.Height; // - offset * equation.m_shape.Height;
                }
            }
        }

        private void changeOptionsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AddinUtilities.changeOptions();
        }

        private void buttonColor_Click(object sender, EventArgs e)
        {
            ColorDialog dialog = new ColorDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                buttonColor.BackColor = dialog.Color;
                Color col = buttonColor.BackColor;
                float r = (float)col.R / 255.0f;
                float g = (float)col.G / 255.0f;
                float b = (float)col.B / 255.0f;
                string rStr = r.ToString().Replace(',', '.');
                string gStr = g.ToString().Replace(',', '.');
                string bStr = b.ToString().Replace(',', '.');
                m_textColor = rStr + "," + gStr + "," + bStr;

            }
        }

        private void findToolStripMenuItem_Click(object sender, EventArgs e)
        {
            m_scintilla.FindReplace.ShowFind();
        }

        private void replaceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            m_scintilla.FindReplace.ShowReplace();
        }

        private void incrementalSearchToolStripMenuItem_Click(object sender, EventArgs e)
        {
            m_scintilla.FindReplace.IncrementalSearcher.Show();
        }

        private void commentLineToolStripMenuItem_Click(object sender, EventArgs e)
        {
            m_scintilla.Commands.Execute(ScintillaNET.BindableCommand.LineComment);
        }

        private void uncommentLineToolStripMenuItem_Click(object sender, EventArgs e)
        {
            m_scintilla.Commands.Execute(ScintillaNET.BindableCommand.LineUncomment);
        }

        private void addSnippetToolStripMenuItem_Click(object sender, EventArgs e)
        {
            m_scintilla.Snippets.ShowSnippetList();
        }

        private void buttonPreview_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (pictureBoxPreview.Image != null)
                pictureBoxPreview.Image.Dispose();

            bool finishedSuccessfully = generateEquation(false, true);

            if (finishedSuccessfully)
            {
                string imageFile = Path.Combine(AddinUtilities.getAppDataLocation(), "teximport.png");
                using (FileStream stream = new FileStream(imageFile, FileMode.Open, FileAccess.Read))
                {
                    pictureBoxPreview.Image = Image.FromStream(stream);
                    int sum = buttonColor.BackColor.R + buttonColor.BackColor.G + buttonColor.BackColor.B;
                    //Color col = Color.FromArgb(255 - buttonColor.BackColor.R, 255 - buttonColor.BackColor.G, 255 - buttonColor.BackColor.B);
                    Color col = Color.White;
                    if (sum / 3 > 127)
                        col = Color.Black;
                    pictureBoxPreview.BackColor = col;
                    panel1.BackColor = col;
                    stream.Close();
                }
            }
            Cursor.Current = Cursors.Default;
        }

        private void panel1_VisibleChanged(object sender, EventArgs e)
        {
            if (m_scintilla.Text != "")
                buttonPreview.PerformClick();
        }

        private void openLatexTemplateToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string templateFileName = AddinUtilities.getAppDataLocation() + "\\LatexTemplate.txt";
            System.Diagnostics.Process.Start(templateFileName);
        }
 
    }

    public class LatexFont
    {
        public LatexFont(string fn, string lfn)
        {
            fontName = fn;
            latexFontName = lfn;
        }
        public string fontName;
        public string latexFontName;

        public override string ToString()
        {
            return fontName;
        }
    }

    public class LatexFontSeries
    {
        public LatexFontSeries(string fns, string lfns)
        {
            fontSeries = fns;
            latexFontSeries = lfns;
        }
        public string fontSeries;
        public string latexFontSeries;

        public override string ToString()
        {
            return fontSeries;
        }
    }

    public class LatexFontShape
    {
        public LatexFontShape(string fns, string lfns)
        {
            fontShape = fns;
            latexFontShape = lfns;
        }
        public string fontShape;
        public string latexFontShape;

        public override string ToString()
        {
            return fontShape;
        }
    }

}
