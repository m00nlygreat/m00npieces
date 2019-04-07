using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;
using System.Diagnostics;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;




namespace m00npieces
{
    public partial class Ribbon1
    {
        
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnSwap_Clicked(object sender, RibbonControlEventArgs e)
        {
            var sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;

            if (sel.ShapeRange.Count == 2) // 2개를 선택했을때만 작동.
            {
                float floFirstTop = sel.ShapeRange[1].Top; // 선택한 두 도형의 최초 위치를 변수에 담기
                float floFirstLeft = sel.ShapeRange[1].Left;
                float floSecondTop = sel.ShapeRange[2].Top;
                float floSecondLeft = sel.ShapeRange[2].Left;
                int intFirstZorder = WhereInSlide(sel.ShapeRange[1]); // 선택한 두 도형의 최초 순서를 변수에 담는다. 이름은 ZOrder이나 ZOrderPosition 속성과는 관련없는 커스텀 구현된 위치임.
                int intSecondZorder = WhereInSlide(sel.ShapeRange[2]);

                sel.ShapeRange[1].Top = floSecondTop; // 서로의 위치를 바꾸는 부분
                sel.ShapeRange[1].Left = floSecondLeft;
                sel.ShapeRange[2].Top = floFirstTop;
                sel.ShapeRange[2].Left = floFirstLeft;
                if (intFirstZorder < intSecondZorder) // First냐, Second냐의 차이는 클릭 순서에 따라 달라진다. 암튼 뭐가 더 위에 있느냐에 따라,  
                {
                    for (int i = intFirstZorder; i < intSecondZorder; i++)
                    {
                        sel.ShapeRange[1].ZOrder(MsoZOrderCmd.msoBringForward);
                    }
                    for (int i = intSecondZorder - 1; i > intFirstZorder; i--)
                    {
                        sel.ShapeRange[2].ZOrder(MsoZOrderCmd.msoSendBackward);
                    }

                }
                else
                {
                    for (int i = intFirstZorder; i > intSecondZorder; i--)
                    {
                        sel.ShapeRange[1].ZOrder(MsoZOrderCmd.msoSendBackward);
                    }
                    for (int i = intSecondZorder + 1; i < intFirstZorder; i++)
                    {
                        sel.ShapeRange[2].ZOrder(MsoZOrderCmd.msoBringForward);
                    }
                }
            }
            else
            {
                return;
            }

            //MessageBox.Show(sel.ShapeRange[2].Top.ToString() + " " + sel.ShapeRange[2].Left.ToString() + " " + sel.ShapeRange.Count.ToString());
            //MessageBox.Show(sel.TextFrame.TextRange.Text);
            //sel.TextFrame.TextRange.Text = strTag;
        }

        //private void Button1_Click(object sender, RibbonControlEventArgs e)
        //{
        //    var sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
        //    //MessageBox.Show(sel.ShapeRange.ZOrderPosition.ToString());
        //    MessageBox.Show(WhereInSlide(sel.ShapeRange[1]).ToString());
        //}

        //private void Button3_Click(object sender, RibbonControlEventArgs e)
        //{
        //    var sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
        //    WhereInSlide(sel.ShapeRange[1]);
        //    //var curSlide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
        //    //// MessageBox.Show(curSlide.Shapes.Count.ToString());
        //    //MessageBox.Show(curSlide.Shapes[1].ZOrderPosition.ToString());
        //}
        public int WhereInSlide(PowerPoint.Shape shape)
        {
            int intOrderInSlide = 0;
            var curSlide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            var shapes = curSlide.Shapes;
            //MessageBox.Show("슬라이드 내 도형 개수" + curSlide.Shapes.Count.ToString());
            //MessageBox.Show("현재 도형의 ZOrder" + shape.ZOrderPosition.ToString());

            for (int i = 1; i <= shapes.Count; i++)
            {
                if (shapes[i] == shape)
                {
                    intOrderInSlide = i;
                    break;
                }
            }
            return intOrderInSlide;
        }
    }
    //PowerPoint._Application myPPT = Globals.ThisAddIn.Application;
    //PowerPoint.Slide curSlide = myPPT.ActiveWindow.View.Slide;
}


