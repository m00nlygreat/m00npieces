using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Ribbon;


namespace m00npieces
{
    public partial class Ribbon1
    {
        
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void Button1_Click(object sender, RibbonControlEventArgs e)
        {
            //    Globals.ThisAddIn.Application.ActiveWindow.View.Slide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 0, 0, 10, 20); //현재 슬라이드에 사각형 삽입 ㅠㅠ 드디어 해냈따
            var newShape = Globals.ThisAddIn.Application.ActiveWindow.View.Slide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 0, 0, 10, 20).TextFrame.TextRange.Text = "Here is some test text";
            // newShape.Tags.Add("testKey", "testValue");
            foreach (var shape in Globals.ThisAddIn.Application.ActiveWindow.View.Slide.Shapes)
            {
                var tagValue = shape.Tags["testKey"];
                if (!string.IsNullOrEmpty(tagValue))
                {
                    // found it!
                    break;
                }
            }
        }

        private void Button2_Click(object sender, RibbonControlEventArgs e)
        {
            var sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            //string strTag = sel.ShapeRange[1].Type.ToString();

            float floFirstTop = sel.ShapeRange[1].Top; // 선택한 두 도형의 위치를 바꿈. (Top/Left)
            float floFirstLeft = sel.ShapeRange[1].Left;
            float floSecondTop = sel.ShapeRange[2].Top;
            float floSecondLeft = sel.ShapeRange[2].Left;
            sel.ShapeRange[1].Top = floSecondTop;
            sel.ShapeRange[1].Left = floSecondLeft;
            sel.ShapeRange[2].Top = floFirstTop;
            sel.ShapeRange[2].Left = floFirstLeft;

            //MessageBox.Show(sel.ShapeRange[2].Top.ToString() + " " + sel.ShapeRange[2].Left.ToString() + " " + sel.ShapeRange.Count.ToString());
            //MessageBox.Show(sel.TextFrame.TextRange.Text);
            //sel.TextFrame.TextRange.Text = strTag;
        }
    }
    //PowerPoint._Application myPPT = Globals.ThisAddIn.Application;
    //PowerPoint.Slide curSlide = myPPT.ActiveWindow.View.Slide;
}
 

