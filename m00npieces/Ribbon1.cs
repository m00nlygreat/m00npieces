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
        int intAnchorPoint;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            intAnchorPoint = btnAnchor_Clicked(0); // 초기 Anchor 설정
        }

        private void btnSwap_Clicked(object sender, RibbonControlEventArgs e)
        {
            var sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;

            if (sel.ShapeRange.Count == 2) // 2개를 선택했을때만 작동.
            {
                float firstTop;
                float firstLeft;
                float secondTop;
                float secondLeft;
                GetAnchored(sel.ShapeRange[1], sel.ShapeRange[2], out firstTop, out firstLeft);
                GetAnchored(sel.ShapeRange[2], sel.ShapeRange[1], out secondTop, out secondLeft);
                sel.ShapeRange[1].Top = secondTop; // 서로의 위치를 바꾸는 부분
                sel.ShapeRange[1].Left = secondLeft;
                sel.ShapeRange[2].Top = firstTop;
                sel.ShapeRange[2].Left = firstLeft;

                int intFirstZorder = WhereInSlide(sel.ShapeRange[1]); // 선택한 두 도형의 최초 순서를 변수에 담는다. 이름은 ZOrder이나 ZOrderPosition 속성과는 관련없는 커스텀 구현된 위치임.
                int intSecondZorder = WhereInSlide(sel.ShapeRange[2]);
                if (intFirstZorder < intSecondZorder) // First냐, Second냐의 차이는 클릭 순서에 따라 달라진다. 암튼 뭐가 더 위에 있느냐에 따라, 앞으로 보내거나 뒤로 보내야 하는 방향이 달라짐.
                {
                    for (int i = intFirstZorder; i < intSecondZorder; i++)
                    {
                        sel.ShapeRange[1].ZOrder(MsoZOrderCmd.msoBringForward);
                    }
                    for (int i = intSecondZorder - 1; i > intFirstZorder; i--) // 두번째 도형의 순서를 옮길 때는, 한 번 덜 가야 한다. 왜냐면 첫번째 도형이 이미 두 번째 도형의 위치까지 와 있어서 서로 순서가 교체되었으므로.
                    {
                        sel.ShapeRange[2].ZOrder(MsoZOrderCmd.msoSendBackward);
                    }

                }
                else // 같은 내용을 방향 바꿔서 보냄.
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
               
            }

        }
        public int WhereInSlide(PowerPoint.Shape shape) //ZOrderPosition이 아닌, 실제 앞으로 보내기/뒤로 보내기 시 작동하는 선택 도형의 슬라이드 내 레이어 위치를 구한다.
        {
            int intOrderInSlide = 0;
            var curSlide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            var shapes = curSlide.Shapes;

            for (int i = 1; i <= shapes.Count; i++) // 현재 슬라이드의 모든 도형과, 선택된 도형을 ==연산자로 비교해, 선택된 도형이 슬라이드 내에 몇 번 도형인지 알아냄.
            {
                if (shapes[i] == shape)
                {
                    intOrderInSlide = i;
                    break;
                }
            }
            return intOrderInSlide;
        }
        public int btnAnchor_Clicked(int whichclicked) // 앵커 버튼을 클릭했을 때, 버튼의 보이는 상태를 바꿈.
        {
            RibbonToggleButton[] btnsAnchor = new RibbonToggleButton[9] { btnTL, btnTC, btnTR, btnML, btnMC, btnMR, btnBL, btnBC, btnBR }; // 앵커를 설정하는 9개 버튼을 일단 배열에 담아봄.
            for (int i = 0; i <= btnsAnchor.Length - 1; i++) 
            {
                if (i == whichclicked)
                {
                    btnsAnchor[i].Label = "◆";
                }
                else
                {
                    btnsAnchor[i].Label = "◇";
                    btnsAnchor[i].Checked = false;
                }
            }
            return whichclicked + 1;
        }

        private void BtnMatchSize_Click(object sender, RibbonControlEventArgs e) // 사이즈 매치
        {
            var sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            for (int i=2;i <= sel.ShapeRange.Count; i++)
            {
                float secondTop;
                float secondLeft;
                MatchSizeGetAnchored(sel.ShapeRange[1], sel.ShapeRange[i], out secondTop, out secondLeft); // 앵커를 기준으로, 변경된 크기일때 옮겨야 하는 Top/Left값 산출
                sel.ShapeRange[i].Top = secondTop; // 크기변경은 먼저 위치를 옮기고 해야된다.
                sel.ShapeRange[i].Left = secondLeft;
                sel.ShapeRange[i].Width = sel.ShapeRange[1].Width; // 이걸로 잘 될까 싶지만 놀랍도록 잘 된다.
                sel.ShapeRange[i].Height = sel.ShapeRange[1].Height;
            }

        }
        public void GetAnchored(PowerPoint.Shape first, PowerPoint.Shape second, out float top, out float left) // 도형 2개를 근거로, 위치 변경 후 2번째 도형의 바뀌어야 하는 Top, Left값 산출
        {
            switch (intAnchorPoint)
            {
                case 1:
                    top = first.Top;
                    left = first.Left;
                    break;
                case 2:
                    top = first.Top;
                    left = first.Left + first.Width / 2 - second.Width / 2;
                    break;
                case 3:
                    top = first.Top;
                    left = first.Left + first.Width - second.Width;
                    break;
                case 4:
                    top = first.Top + first.Height / 2 - second.Height / 2;
                    left = first.Left;
                    break;
                case 5:
                    top = first.Top + first.Height / 2 - second.Height / 2;
                    left = first.Left + first.Width / 2 - second.Width / 2;
                    break;
                case 6:
                    top = first.Top + first.Height / 2 - second.Height / 2;
                    left = first.Left + first.Width - second.Width;
                    break;
                case 7:
                    top = first.Top + first.Height - second.Height;
                    left = first.Left;
                    break;
                case 8:
                    top = first.Top + first.Height - second.Height;
                    left = first.Left + first.Width / 2 - second.Width / 2;
                    break;
                case 9:
                    top = first.Top + first.Height - second.Height;
                    left = first.Left + first.Width - second.Width;
                    break;
                default:
                    top = first.Top;
                    left = first.Left;
                    break;
            }
        }
        public void MatchSizeGetAnchored(PowerPoint.Shape first, PowerPoint.Shape second, out float top, out float left) // 도형 2개를 근거로, 크기 변경 후 앵커 설정에 의해 바뀌어야 하는 2번째 도형의 위치값 산출
        {
            switch (intAnchorPoint)
            {
                default:
                    top = second.Top;
                    left = second.Left;
                    break;
                case 1:
                    top = second.Top;
                    left = second.Left;
                    break;
                case 2:
                    top = second.Top;
                    left = second.Left - (first.Width - second.Width) / 2;
                    break;
                case 3:
                    top = second.Top;
                    left = second.Left - (first.Width - second.Width);
                    break;
                case 4:
                    top = second.Top - (first.Height - second.Height) / 2;
                    left = second.Left;
                    break;
                case 5:
                    top = second.Top - (first.Height - second.Height) / 2;
                    left = second.Left - (first.Width - second.Width) / 2;

                    break;
                case 6:
                    top = second.Top - (first.Height - second.Height) / 2;
                    left = second.Left - (first.Width - second.Width);
                    break;
                case 7:
                    top = second.Top - (first.Height - second.Height);
                    left = second.Left;
                    break;
                case 8:
                    top = second.Top - (first.Height - second.Height);
                    left = second.Left - (first.Width - second.Width) / 2;
                    break;
                case 9:
                    top = second.Top - (first.Height - second.Height);
                    left = second.Left - (first.Width - second.Width);
                    break;
            }
        }
        #region
        private void BtnTL_Click(object sender, RibbonControlEventArgs e) // 앵커 버튼 누를 때마다, 함수 호출해서 값 설정
        {
            intAnchorPoint = btnAnchor_Clicked(0);
        }

        private void BtnTC_Click(object sender, RibbonControlEventArgs e)
        {
            intAnchorPoint = btnAnchor_Clicked(1);
        }

        private void BtnTR_Click(object sender, RibbonControlEventArgs e)
        {
            intAnchorPoint = btnAnchor_Clicked(2);
        }

        private void BtnML_Click(object sender, RibbonControlEventArgs e)
        {
            intAnchorPoint = btnAnchor_Clicked(3);
        }

        private void BtnMC_Click(object sender, RibbonControlEventArgs e)
        {
            intAnchorPoint = btnAnchor_Clicked(4);
        }

        private void BtnMR_Click(object sender, RibbonControlEventArgs e)
        {
            intAnchorPoint = btnAnchor_Clicked(5);
        }

        private void BtnBL_Click(object sender, RibbonControlEventArgs e)
        {
            intAnchorPoint = btnAnchor_Clicked(6);
        }

        private void BtnBC_Click(object sender, RibbonControlEventArgs e)
        {
            intAnchorPoint = btnAnchor_Clicked(7);
        }

        private void BtnBR_Click(object sender, RibbonControlEventArgs e)
        {
            intAnchorPoint = btnAnchor_Clicked(8);
        }
        #endregion

        private void Button1_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnFontAntiAlias_Clicked(object sender, RibbonControlEventArgs e)
        {
            var sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            //if (sel.ShapeRange.HasTextFrame==MsoTriState.msoTrue)
            //{
                sel.TextRange2.Font.Line.Visible = MsoTriState.msoTrue;
                sel.TextRange2.Font.Line.Transparency = 1;
            //}

            //sel.TextRange2.Font.Line.ForeColor.RGB = System.Drawing.Color.Black.ToArgb();
            //if (sel.ShapeRange.HasTable == MsoTriState.msoTrue)
            //{
            //    var tab = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange.Table;
            //    int col = sel.ShapeRange.Table.Columns.Count;
            //    int row = sel.ShapeRange.Table.Rows.Count;
            //    for (int i = 1; i <= row - 1; i++)
            //    {
            //        for (int j=1;j<= col - 1; j++)
            //        {
            //            tab.Cell(i, j).Shape.TextFrame2.TextRange.Font.Bold = MsoTriState.msoFalse;
            //            tab.Cell(i, j).Shape.TextFrame2.TextRange.Font.Line.Visible = MsoTriState.msoTrue;
            //            tab.Cell(i, j).Shape.TextFrame2.TextRange.Font.Line.Transparency = 1;
            //        }
            //    }
            //}
            //else
            //{
            //}
        }
    }
    //PowerPoint._Application myPPT = Globals.ThisAddIn.Application;
    //PowerPoint.Slide curSlide = myPPT.ActiveWindow.View.Slide;
}


