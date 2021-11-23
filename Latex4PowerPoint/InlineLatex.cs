using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;
using System.Windows.Forms;
using System.IO;

namespace Latex4PowerPoint
{
    public class InlineLatex
    {
        public static void createInlineEquations(Shape shape)
        {
            if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                Microsoft.Office.Interop.PowerPoint.TextFrame textFrame = shape.TextFrame;
                if (textFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    TextRange textRange = textFrame.TextRange;
                    createEquations(shape, textRange);
                }
            }
        }

        public static void alignLatexShape(TextRange range, Shape shape, float offset)
        {
            const char nl = (char)8232;
            const char nbsc = (char)160;
            range.Text = nbsc.ToString();
            float rangeHeight = range.BoundHeight;
            while (range.BoundWidth < shape.Width)
            {
                range.Text += nbsc;
                if (rangeHeight != range.BoundHeight)
                {
                    range.Text = range.Text.Remove(range.Text.Length-2, 2);
                    range.Text += nl.ToString();
                    break;
                }
            }

            System.Drawing.FontFamily family = AddinUtilities.getFontFamily(range);
            //float height = (float)(range.BoundHeight * ((float)family.GetCellAscent(System.Drawing.FontStyle.Regular) / (float)family.GetLineSpacing(System.Drawing.FontStyle.Regular)));
            float lineSpacing = range.Font.Size * (float)family.GetLineSpacing(System.Drawing.FontStyle.Regular) / (float)family.GetEmHeight(System.Drawing.FontStyle.Regular);
            float spaceBefore = 0.0f;
            if (range.BoundHeight > lineSpacing)
                spaceBefore = range.ParagraphFormat.SpaceBefore * lineSpacing;
            float height = spaceBefore + (range.Font.Size * (float)family.GetCellAscent(System.Drawing.FontStyle.Regular)) / (float)family.GetEmHeight(System.Drawing.FontStyle.Regular);

            shape.Left = range.BoundLeft;
            shape.Top = range.BoundTop + height - (1.0f - offset) * shape.Height;
        }

        static private void createEquations(Shape shape, TextRange range)
        {
            List<LatexEquation> equations = createEquationList(shape, range);
            createImages(equations);
            createAndAlignShapes(shape, range, equations);
        }

        static private List<LatexEquation> createEquationList(Shape shape, TextRange range)
        {
            int currentIndex = 0;
            bool foundEquation = true;
            List<LatexEquation> equations = new List<LatexEquation>();
            SettingsManager mgr = SettingsManager.getCurrent();

            // Check Dpi
            float[] systemDPI = AddinUtilities.getSystemDPI();
            float dpiValue = AddinUtilities.getFloat(mgr.SettingsData.dpi, systemDPI[0]);

            while (foundEquation)
            {
                int startLatex = range.Text.IndexOf("$$", currentIndex);
                int endLatex = -1;
                if (startLatex != -1)
                {
                    startLatex += 2;
                    currentIndex = startLatex;
                    endLatex = range.Text.IndexOf("$$", startLatex);
                    if (endLatex != -1)
                    {
                        currentIndex = endLatex + 2;
                        string latexCode = range.Text.Substring(startLatex - 1, endLatex - startLatex + 2);
                        float fontSize = range.Characters(startLatex + 1, 1).Font.Size;
                        int color = range.Characters(startLatex + 1, 1).Font.Color.RGB;
                        System.Drawing.Color col = System.Drawing.Color.FromArgb(color);

                        LatexEquation ie = new LatexEquation(
                                                latexCode,
                                                fontSize,
                                                dpiValue,
                                                AddinUtilities.colorToString(col),
                                                AddinUtilities.getLatexFont(mgr.SettingsData.font),
                                                AddinUtilities.getLatexFontSeries(mgr.SettingsData.fontSeries),
                                                AddinUtilities.getLatexFontShape(mgr.SettingsData.fontShape),
                                                true,
                                                shape.Id);
                        equations.Add(ie);
                    }
                    else
                        foundEquation = false;
                }
                else
                    foundEquation = false;
            }
            return equations;
        }

        static private void createImages(List<LatexEquation> equations)
        {
            string appPath = AddinUtilities.getAppDataLocation();
            Directory.SetCurrentDirectory(appPath);
            LatexFileGenerator.writeTexFile(appPath + "\\teximport.tex", equations);
            if (!AddinUtilities.executeMikTex())
                return;
            if (!AddinUtilities.executeDviPng(equations))
                return;
        }

        static private void createAndAlignShapes(Shape shape, TextRange range, List<LatexEquation> equations)
        {
            bool foundEquation = true;
            LatexEquation[] eq = equations.ToArray();
            int equationIndex = 0;
            int currentIndex = 0;
            while (foundEquation)
            {
                int startLatex = range.Text.IndexOf("$$", currentIndex);
                int endLatex = -1;
                if (startLatex != -1)
                {
                    startLatex += 2;
                    //currentIndex = startLatex;
                    endLatex = range.Text.IndexOf("$$", startLatex);
                    if (endLatex != -1)
                    {
                        //currentIndex = endLatex + 2;
                        TextRange currentRange = range.Characters(startLatex - 1, endLatex - startLatex + 4);

                        string filename = "teximport" + (equationIndex + 1).ToString() + ".png";
                        AddinUtilities.createLatexShape(filename, equations[equationIndex]);
                        AddinUtilities.centerLatexShape(equations[equationIndex]);
                        if (equations[equationIndex].m_shape != null)
                        {
                            alignLatexShape(currentRange, equations[equationIndex].m_shape, (float)equations[equationIndex].m_offset / (float)equations[equationIndex].m_imageHeight);

                            // Copy format
                            shape.PickUp();
                            equations[equationIndex].m_shape.Apply();
                            equationIndex++;
                        }
                    }
                    else
                        foundEquation = false;
                }
                else
                    foundEquation = false;
            }
        }
         
    }

    
}
