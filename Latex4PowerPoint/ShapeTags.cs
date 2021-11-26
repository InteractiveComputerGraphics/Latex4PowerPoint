using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;

namespace Latex4PowerPoint
{
    public class ShapeTags
    {
        public static void setShapeTags(LatexEquation equation)
        {
            equation.m_shape.Tags.Add("LatexCode", equation.m_code.Replace("\r\n", "\\r\\n"));
            equation.m_shape.Tags.Add("LatexFontSize", equation.m_fontSize.ToString());
            equation.m_shape.Tags.Add("LatexTextColor", equation.m_color);
            equation.m_shape.Tags.Add("LatexDPI", equation.m_dpi.ToString());
            equation.m_shape.Tags.Add("LatexFont", equation.m_font.fontName);
            equation.m_shape.Tags.Add("LatexFontSeries", equation.m_fontSeries.fontSeries);
            equation.m_shape.Tags.Add("LatexFontShape", equation.m_fontShape.fontShape);
            equation.m_shape.Tags.Add("LatexIsInline", equation.m_isInline.ToString());
            equation.m_shape.Tags.Add("LatexTextShapeId", equation.m_textShapeId.ToString());
            equation.m_shape.Tags.Add("LatexAddinVersion", AddinUtilities.getVersionString());
        }

        public static LatexEquation getLatexEquation(Shape s)
        {
            if ((getLatexCode(s) != null) && (getLatexCode(s) != ""))
            {
                return new LatexEquation(
                        getLatexCode(s),
                        getLatexFontSize(s),
                        getLatexDPI(s),
                        getLatexTextColor(s),
                        getLatexFont(s),
                        getLatexFontSeries(s),
                        getLatexFontShape(s),
                        getLatexIsInline(s));
            }
            return null;
        }

        public static string getLatexCode(Shape s)
        {
            string str = s.Tags["LatexCode"];
            if (str != null)
                str = str.Replace("\\r\\n", "\r\n");
            return str;
        }

        public static string getLatexTextColor(Shape s)
        {
            return s.Tags["LatexTextColor"];
        }

        public static LatexFont getLatexFont(Shape s)
        {
            string str = s.Tags["LatexFont"];
            if (str != null)
                return AddinUtilities.getLatexFont(str);
            return null;
        }

        public static LatexFontShape getLatexFontShape(Shape s)
        {
            string str = s.Tags["LatexFontShape"];
            if (str != null)
                return AddinUtilities.getLatexFontShape(str);
            return null;
        }

        public static LatexFontSeries getLatexFontSeries(Shape s)
        {
            string str = s.Tags["LatexFontSeries"];
            if (str != null)
                return AddinUtilities.getLatexFontSeries(str);
            return null;
        }

        public static string getLatex4PowerPointVersion(Shape s)
        {
            return s.Tags["LatexAddinVersion"];
        }

        public static float getLatexFontSize(Shape s)
        {
            string str = s.Tags["LatexFontSize"];
            float fontSizeValue = 12.0f;
            if (str != null)
                fontSizeValue = AddinUtilities.getFloat(str, 12.0f);
            return fontSizeValue;
        }

        public static bool getLatexIsInline(Shape s)
        {
            string str = s.Tags["LatexIsInline"];
            if ((str != null) && (str != ""))
            {
                bool value = false;
                if (str != null)
                    value = AddinUtilities.getBool(str, false);
                return value;
            }
            return false;
        }


        public static int getLatexTextShapeId(Shape s)
        {
            string str = s.Tags["LatexTextShapeId"];
            int value = -1;
            if ((str != null) && (str != ""))
                value = AddinUtilities.getInt(str, -1);
            return value;
        }

        public static float getLatexDPI(Shape s)
        {
            // Check Dpi
            float[] systemDPI = AddinUtilities.getSystemDPI();
            string str = s.Tags["LatexDPI"];
            float value = systemDPI[0];
            if (str != null)
                value = AddinUtilities.getFloat(str, 12.0f);
            return value;
        }
    }
          
}
