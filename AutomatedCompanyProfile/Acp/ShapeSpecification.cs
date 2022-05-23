using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Office = Microsoft.Office.Core;

namespace Acp
{
    internal class ShapeSpecification
    {
        public dynamic[] Raw = new dynamic[61];

        public Office.MsoTextOrientation TextFrameOrientation = Office.MsoTextOrientation.msoTextOrientationHorizontal; // [31]
        public Office.MsoVerticalAnchor TextFrameVerticalAnchor = Office.MsoVerticalAnchor.msoAnchorTop; // [32]
        public Office.MsoAutoSize TextFrameAutoSize = Office.MsoAutoSize.msoAutoSizeNone; // [33]
        public float TextFrameMarginLeft = 0; // 34;
        public float TextFrameMarginRight = 0; // [35];
        public float TextFrameMarginTop = 0; // [36];
        public float TextFrameMarginBottom = 0; // [37];
        public Office.MsoTriState TextFrameWordWrap = Office.MsoTriState.msoTrue; // [38];

        /*
        var textRange = textFrame.TextRange;
        if (textAdjust)
            textRange.Text = sldText;
        textRange.Font.Name = shpSpn[42].ToString();
        textRange.Font.Bold = (Office.MsoTriState) shpSpn[43];
        textRange.Font.Italic = (Office.MsoTriState) shpSpn[44];
        textRange.Font.Underline = (Office.MsoTriState) shpSpn[45];
        textRange.Font.Size = (float) shpSpn[46];
        //           textRange.Font.Color = shpSpn[47];
        textRange.Font.Shadow = (Office.MsoTriState) shpSpn[48];
        */
    }

}
