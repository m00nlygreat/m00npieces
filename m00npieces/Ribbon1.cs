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
        int intAnchor;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            btnAnchor_Clicked(0); // 초기 Anchor 설정
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
        public int btnAnchor_Clicked(int whichclicked) // 앵커 버튼을 클릭했을 때, 설정을 바꿈.
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

            if (sel.ShapeRange.Count == 2) // 2개를 선택했을때만 작동.
            {
                sel.ShapeRange[2].Width = sel.ShapeRange[1].Width; // 이걸로 잘 될까 싶지만 놀랍도록 잘 된다.
                sel.ShapeRange[2].Height = sel.ShapeRange[1].Height;
            }
        }
        #region
        private void BtnTL_Click(object sender, RibbonControlEventArgs e) // 앵커 버튼 누를 때마다, 함수 호출해서 값 설정
        {
            intAnchor = btnAnchor_Clicked(0);
        }

        private void BtnTC_Click(object sender, RibbonControlEventArgs e)
        {
            intAnchor = btnAnchor_Clicked(1);
        }

        private void BtnTR_Click(object sender, RibbonControlEventArgs e)
        {
            intAnchor = btnAnchor_Clicked(2);
        }

        private void BtnML_Click(object sender, RibbonControlEventArgs e)
        {
            intAnchor = btnAnchor_Clicked(3);
        }

        private void BtnMC_Click(object sender, RibbonControlEventArgs e)
        {
            intAnchor = btnAnchor_Clicked(4);
        }

        private void BtnMR_Click(object sender, RibbonControlEventArgs e)
        {
            intAnchor = btnAnchor_Clicked(5);
        }

        private void BtnBL_Click(object sender, RibbonControlEventArgs e)
        {
            intAnchor = btnAnchor_Clicked(6);
        }

        private void BtnBC_Click(object sender, RibbonControlEventArgs e)
        {
            intAnchor = btnAnchor_Clicked(7);
        }

        private void BtnBR_Click(object sender, RibbonControlEventArgs e)
        {
            intAnchor = btnAnchor_Clicked(8);
        }
        #endregion
        private void Button1_Click(object sender, RibbonControlEventArgs e) // 변수 체크용 메세지 박스 띄우는 임시 버튼
        {
            MessageBox.Show(intAnchor.ToString());
        }
    }
    //PowerPoint._Application myPPT = Globals.ThisAddIn.Application;
    //PowerPoint.Slide curSlide = myPPT.ActiveWindow.View.Slide;
}


