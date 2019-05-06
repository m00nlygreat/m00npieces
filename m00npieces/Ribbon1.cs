using Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;




namespace m00npieces
{
    public partial class Ribbon1
    {
        int intAnchorPoint;
        
        // 버튼의 누름 상태를 표시하는 열거형
        enum Stage {None, Swapped=10, SizeMatched=20, WidthMatched, Aligned=30 }
        Stage onStage = Stage.None;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            intAnchorPoint = btnAnchor_Clicked(0); // 초기 Anchor 설정
            Globals.ThisAddIn.Application.SlideSelectionChanged += SlideNoOnEdtbx; // 슬라이드 이동(포커스 이동)시 슬라이드 번호를 에디트 박스에 입력
            Globals.ThisAddIn.Application.WindowSelectionChange += UpdateObjectName;
            Globals.ThisAddIn.Application.WindowSelectionChange += offtheStageWhenYouSelectAnother;
        }

        private void offtheStageWhenYouSelectAnother(PowerPoint.Selection Sel) { GetOffTheStage(); }
        private void GetOffTheStage() // Stage 상태를 해제한다.
        {
            onStage = Stage.None;
            btnSwap.Image = btnSwap.Image = global::m00npieces.Properties.Resources.swap;
            btnSwap.Label = "교체";
        }

        private void UpdateObjectName(PowerPoint.Selection sel) // 이름을 표시한다.
        {
            try { ebxName.Text = (sel.ShapeRange.Count == 1) ? sel.ShapeRange.Name : ""; }  catch { ebxName.Text = ""; } 
        }
        private void EbxName_TextChanged(object sender, RibbonControlEventArgs e) // 이름을 바꾸면, 개체의 이름도 바꿈.
        {
            try { Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange.Name = ebxName.Text; } catch { }
        }

        private void btnSwap_Clicked(object sender, RibbonControlEventArgs e)
        {
            var sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            if (sel.ShapeRange.Count == 2) // 2개를 선택했을때만 작동.
            {
                switch (onStage)
                {
                    case Stage.Swapped: // 직전에 스왑을 실행했다면, 원본 개체를 삭제한다.
                        sel.ShapeRange[1].Delete();
                        GetOffTheStage();
                        break;
                    default:
                        ObjectSwap(sel.ShapeRange[1], sel.ShapeRange[2]);
                        onStage = Stage.Swapped;
                        btnSwap.Label = "원본 삭제";
                        btnSwap.Image = global::m00npieces.Properties.Resources.fingersnap;
                        break;
                }
            }
        }

        private void ObjectSwap(PowerPoint.Shape shape1, PowerPoint.Shape shape2)
        {
            float firstTop;
            float firstLeft;
            float secondTop;
            float secondLeft;
            GetAnchored(shape1, shape2, out firstTop, out firstLeft);
            GetAnchored(shape2, shape1, out secondTop, out secondLeft);
            shape1.Top = secondTop; // 서로의 위치를 바꾸는 부분
            shape1.Left = secondLeft;
            shape2.Top = firstTop;
            shape2.Left = firstLeft;

            int intFirstZorder = WhereInSlide(shape1); // 선택한 두 도형의 최초 순서를 변수에 담는다. 이름은 ZOrder이나 ZOrderPosition 속성과는 관련없는 커스텀 구현된 위치임.
            int intSecondZorder = WhereInSlide(shape2);
            if (intFirstZorder < intSecondZorder) // First냐, Second냐의 차이는 클릭 순서에 따라 달라진다. 암튼 뭐가 더 위에 있느냐에 따라, 앞으로 보내거나 뒤로 보내야 하는 방향이 달라짐.
            {
                for (int i = intFirstZorder; i < intSecondZorder; i++)
                {
                    shape1.ZOrder(MsoZOrderCmd.msoBringForward);
                }
                for (int i = intSecondZorder - 1; i > intFirstZorder; i--) // 두번째 도형의 순서를 옮길 때는, 한 번 덜 가야 한다. 왜냐면 첫번째 도형이 이미 두 번째 도형의 위치까지 와 있어서 서로 순서가 교체되었으므로.
                {
                    shape2.ZOrder(MsoZOrderCmd.msoSendBackward);
                }

            }
            else // 같은 내용을 방향 바꿔서 보냄.
            {
                for (int i = intFirstZorder; i > intSecondZorder; i--)
                {
                    shape1.ZOrder(MsoZOrderCmd.msoSendBackward);
                }
                for (int i = intSecondZorder + 1; i < intFirstZorder; i++)
                {
                    shape2.ZOrder(MsoZOrderCmd.msoBringForward);
                }
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
            if (onStage != Stage.None){ Globals.ThisAddIn.Application.CommandBars.ExecuteMso("Undo");}
            var sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            try
            {

                    switch (onStage) // Stage에 따라, 크기맞춤, 가로/세로 맞춤을 토글하여 실행한다.
                    {
                        default:
                            for (int i = 2; i <= sel.ShapeRange.Count; i++)
                            {
                                float secondTop;
                                float secondLeft;
                                MatchSizeGetAnchored(sel.ShapeRange[1], sel.ShapeRange[i], out secondTop, out secondLeft);
                                MessageBox.Show(secondLeft.ToString() + " " + secondTop.ToString() + " " + sel.ShapeRange[i].Left.ToString() + " " + sel.ShapeRange[i].Top.ToString());
                                sel.ShapeRange[i].Top = secondTop; // 크기변경은 먼저 위치를 옮기고 해야된다.
                                sel.ShapeRange[i].Left = secondLeft;
                                sel.ShapeRange[i].Width = sel.ShapeRange[1].Width; // 이걸로 잘 될까 싶지만 놀랍도록 잘 된다.
                                sel.ShapeRange[i].Height = sel.ShapeRange[1].Height;
                                onStage = Stage.SizeMatched;
                            }
                            break;
                        case Stage.SizeMatched:
                            for (int i = 2; i <= sel.ShapeRange.Count; i++)
                            {
                                float secondTop;
                                float secondLeft;
                                //Globals.ThisAddIn.Application.CommandBars.ExecuteMso("Undo");
                                MatchSizeGetAnchored(sel.ShapeRange[1], sel.ShapeRange[i], out secondTop, out secondLeft);
                                //MessageBox.Show(secondLeft.ToString() + " " + secondTop.ToString() + " " + sel.ShapeRange[i].Left.ToString() + " " + sel.ShapeRange[i].Top.ToString());
                                sel.ShapeRange[i].Left = secondLeft;
                                sel.ShapeRange[i].Width = sel.ShapeRange[1].Width;
                                onStage = Stage.WidthMatched;
                            }
                            break;
                        case Stage.WidthMatched:
                            for (int i = 2; i <= sel.ShapeRange.Count; i++)
                            {
                                float secondTop;
                                float secondLeft;
                                //Globals.ThisAddIn.Application.CommandBars.ExecuteMso("Undo");
                                MatchSizeGetAnchored(sel.ShapeRange[1], sel.ShapeRange[i], out secondTop, out secondLeft);
                                sel.ShapeRange[i].Top = secondTop;
                                sel.ShapeRange[i].Height = sel.ShapeRange[1].Height;
                                onStage = Stage.None;
                            }
                            break;
                }
            } catch { }
        }
        public void GetAnchored(PowerPoint.Shape first, PowerPoint.Shape second, out float top, out float left) // 도형 2개를 근거로, 위치 변경 후 2번째 도형의 바뀌어야 하는 Top, Left값 산출
        {
            switch (intAnchorPoint)
            {
                case 1:top = first.Top;left = first.Left;break;
                case 2:top = first.Top;left = first.Left + first.Width / 2 - second.Width / 2;break;
                case 3:top = first.Top;left = first.Left + first.Width - second.Width;break;
                case 4:top = first.Top + first.Height / 2 - second.Height / 2;left = first.Left;break;
                case 5:top = first.Top + first.Height / 2 - second.Height / 2;left = first.Left + first.Width / 2 - second.Width / 2;break;
                case 6:top = first.Top + first.Height / 2 - second.Height / 2;left = first.Left + first.Width - second.Width;break;
                case 7:top = first.Top + first.Height - second.Height;left = first.Left;break;
                case 8:top = first.Top + first.Height - second.Height;left = first.Left + first.Width / 2 - second.Width / 2;break;
                case 9:top = first.Top + first.Height - second.Height;left = first.Left + first.Width - second.Width;break;
                default:top = first.Top;left = first.Left;break;
            }
        }
        public void MatchSizeGetAnchored(PowerPoint.Shape first, PowerPoint.Shape second, out float top, out float left) // 도형 2개를 근거로, 크기 변경 후 앵커 설정에 의해 바뀌어야 하는 2번째 도형의 위치값 산출
        {
            switch (intAnchorPoint)
            {
                default:top = second.Top;left = second.Left;break;
                case 1:top = second.Top;left = second.Left;break;
                case 2:top = second.Top;left = second.Left - (first.Width - second.Width) / 2;break;
                case 3:top = second.Top;left = second.Left - (first.Width - second.Width);break;
                case 4:top = second.Top - (first.Height - second.Height) / 2;left = second.Left;break;
                case 5:top = second.Top - (first.Height - second.Height) / 2;left = second.Left - (first.Width - second.Width) / 2;break;
                case 6:top = second.Top - (first.Height - second.Height) / 2;left = second.Left - (first.Width - second.Width);break;
                case 7:top = second.Top - (first.Height - second.Height);left = second.Left;break;
                case 8:top = second.Top - (first.Height - second.Height);left = second.Left - (first.Width - second.Width) / 2;break;
                case 9:top = second.Top - (first.Height - second.Height);left = second.Left - (first.Width - second.Width);break;
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

        private void btnFontAntiAlias_Clicked(object sender, RibbonControlEventArgs e) // 글씨를 예쁘게
        {
            var sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            if (sel.ShapeRange.HasTextFrame == MsoTriState.msoTrue)
            {
                try
                {
                    sel.TextRange2.Font.Line.Visible = MsoTriState.msoTrue;
                    sel.TextRange2.Font.Line.Transparency = 1;
                    Globals.ThisAddIn.Application.StartNewUndoEntry();
                }

                catch
                {

                }
            }
        }

        private void EdtGoToSlide_changed(object sender, RibbonControlEventArgs e) // 슬라이드 번호 입력시, 해당 슬라이드로 이동
        {
            try
            {
                int slideNo = Convert.ToInt32(edtGoToSlide.Text);
                Globals.ThisAddIn.Application.ActivePresentation.Slides[slideNo].Select(); 
            }
            catch // 에러시 그냥 원래 슬라이드 번호 다시 입력
            {
                edtGoToSlide.Text = Globals.ThisAddIn.Application.ActiveWindow.View.Slide.SlideIndex.ToString();
            }
        }
        private void SlideNoOnEdtbx(PowerPoint.SlideRange SldRange) // 슬라이드 이동시, 현재 슬라이드 번호를 에디트 박스에 넣어줌 ㅎ.ㅎ
        {
            try
            {
                edtGoToSlide.Text = SldRange.SlideNumber.ToString();
            }
            catch // 슬라이드 여러 개 선택하면 에러남
            {

            }
        }

        private void btnAdjoinHorizontal_Clicked(object sender, RibbonControlEventArgs e) // 가로로 붙이기
        {
            var sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            List<PowerPoint.Shape> shapesInSlide = new List<PowerPoint.Shape>();
            for (int i = 1; i <= sel.ShapeRange.Count; i++)
            {
                shapesInSlide.Add(sel.ShapeRange[i]);
            }
            List<PowerPoint.Shape> leftsortedShape = shapesInSlide.OrderBy(o => o.Left).ToList();
            for (int i = 1; i <= leftsortedShape.Count - 1; i++)
            {
                leftsortedShape[i].Left = leftsortedShape[i - 1].Left + leftsortedShape[i - 1].Width;
            }
            //for (int i = 1; i <= leftsortedShape.Count; i++)
            //{
            //    leftsortedShape[i-1].TextFrame.TextRange.Text = i.ToString();
            //}

        }

        private void BtnAdjoinVertical_Click(object sender, RibbonControlEventArgs e) // 세로로 붙이기
        {
            var sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            List<PowerPoint.Shape> shapesInSlide = new List<PowerPoint.Shape>();
            for (int i = 1; i <= sel.ShapeRange.Count; i++)
            {
                shapesInSlide.Add(sel.ShapeRange[i]);
            }
            List<PowerPoint.Shape> topsortedShape = shapesInSlide.OrderBy(o => o.Top).ToList();
            for (int i = 1; i <= topsortedShape.Count - 1; i++)
            {
                topsortedShape[i].Top = topsortedShape[i - 1].Top + topsortedShape[i - 1].Height;
            }
        }
        private void Btn_Expand_Click(object sender, RibbonControlEventArgs e) // 늘이기 
        {
            var sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            List<PowerPoint.Shape> shapesInSlide = new List<PowerPoint.Shape>();
            for (int i = 1; i <= sel.ShapeRange.Count; i++)
            {
                shapesInSlide.Add(sel.ShapeRange[i]);
            }
            switch (intAnchorPoint)
            {
                case 2:
                    List<PowerPoint.Shape> topsortedShape = shapesInSlide.OrderBy(o => o.Top).ToList();
                    for (int i = 1; i <= topsortedShape.Count - 1; i++)
                    {
                        float dif = topsortedShape[0].Top - topsortedShape[i].Top;
                        topsortedShape[i].IncrementTop(dif);
                        topsortedShape[i].Height = topsortedShape[i].Height - dif;
                    }
                    break;
                case 4:
                    List<PowerPoint.Shape> leftsortedShape = shapesInSlide.OrderBy(o => o.Left).ToList();
                    for (int i = 1; i <= leftsortedShape.Count - 1; i++)
                    {
                        float dif = leftsortedShape[0].Left - leftsortedShape[i].Left;
                        leftsortedShape[i].IncrementLeft(dif);
                        leftsortedShape[i].Width = leftsortedShape[i].Width - dif;
                    }
                    break;
                case 6:
                    List<PowerPoint.Shape> rightsortedShape = shapesInSlide.OrderByDescending(o => Right(o)).ToList();
                    for (int i = 1; i <= rightsortedShape.Count - 1; i++)
                    {
                        float dif = Right(rightsortedShape[0]) - Right(rightsortedShape[i]);
                        rightsortedShape[i].Width = rightsortedShape[i].Width + dif;
                    }
                    break;
                case 8:
                    List<PowerPoint.Shape> bottomsortedShape = shapesInSlide.OrderByDescending(o => Bottom(o)).ToList();
                    for (int i = 1; i <= bottomsortedShape.Count - 1; i++)
                    {
                        float dif = Bottom(bottomsortedShape[0]) - Bottom(bottomsortedShape[i]);
                        bottomsortedShape[i].Height = bottomsortedShape[i].Height + dif;
                    }
                    break;
                default:
                    break;
            } // 앵커 위치에 따라 늘이는 방향 달라짐.
        } 
        public float Right(PowerPoint.Shape o) { return o.Left + o.Width ; }
        public float Bottom(PowerPoint.Shape o) { return o.Top + o.Height; }
        // 모으기 구현
        private void BtnGather_Click(object sender, RibbonControlEventArgs e) // 모으기
        {
            var sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            for (int i = 2; i <= sel.ShapeRange.Count; i++)
            {
                float difTop, difLeft;
                GatherGetAnchored(sel.ShapeRange[1], sel.ShapeRange[i], out difTop, out difLeft);
                sel.ShapeRange[i].IncrementTop(difTop);
                sel.ShapeRange[i].IncrementLeft(difLeft);
            }
        }
        public void GatherGetAnchored(PowerPoint.Shape first, PowerPoint.Shape second, out float top, out float left) // 앵커를 기준으로 모아줘야할 나머지 도형의 위치값 증감분을 계산
        {

            switch (intAnchorPoint)
            {
                case 1:top = first.Top - second.Top;left = first.Left - second.Left;break;
                case 2:top = first.Top - second.Top;left = first.Left - second.Left - (second.Width - first.Width) / 2;break;
                case 3:top = first.Top - second.Top;left = first.Left - second.Left - (second.Width - first.Width);break;
                case 4:top = first.Top - second.Top - (second.Height - first.Height) / 2;left = first.Left - second.Left;break;
                case 5:top = first.Top - second.Top - (second.Height - first.Height) / 2;left = first.Left - second.Left - (second.Width - first.Width) / 2;break;
                case 6:top = first.Top - second.Top - (second.Height - first.Height) / 2;left = first.Left - second.Left - (second.Width - first.Width);break;
                case 7:top = first.Top - second.Top - (second.Height - first.Height);left = first.Left - second.Left;break;
                case 8:top = first.Top - second.Top - (second.Height - first.Height);left = first.Left - second.Left - (second.Width - first.Width) / 2;break;
                case 9:top = first.Top - second.Top - (second.Height - first.Height);left = first.Left - second.Left - (second.Width - first.Width);break;
                default:top = first.Top - second.Top;left = first.Left - second.Left;break;
            }
        }

        private void BtnSync_Click(object sender, RibbonControlEventArgs e)
        {

            //try
            //{
            //    var sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            //    Show(sel.SlideRange.SlideIndex.ToString());
            //    PowerPoint.Slide curSlide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            //    //foreach (PowerPoint.Shape shp in curSlide.Shapes)
            //    //{
            //    //    if (sel.ShapeRange.Name == shp.Name)
            //    //    {
            //    //        sel.ShapeRange.Duplicate();
            //    //    }
            //    //}
            //}
            //catch { }
            //태그 사용예제..ㅠㅠ
            //var sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            //foreach (PowerPoint.Tags tag in sel.ShapeRange)
            //{ tag.}
            //MessageBox.Show();
        }
    }


    //PowerPoint._Application myPPT = Globals.ThisAddIn.Application;
    //PowerPoint.Slide curSlide = myPPT.ActiveWindow.View.Slide;
}
// 표의 텍스트 외곽선주기. 외 않데? 는지 도통 모르겄따
//try
//  {
//      var sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
//      var tab = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange.Table;
//      int col = sel.ShapeRange.Table.Columns.Count;
//      int row = sel.ShapeRange.Table.Rows.Count;
//      for (int i = 1; i <= row; i++)
//      {
//          for (int j = 1; j <= col; j++)
//          {
//              tab.Cell(i, j).Shape.TextFrame2.TextRange.Font.Bold = MsoTriState.msoTrue;
//              tab.Cell(i, j).Shape.TextFrame2.TextRange.Font.Line.Visible = MsoTriState.msoTrue;
//              tab.Cell(i, j).Shape.TextFrame2.TextRange.Font.Line.Transparency = 1;
//          }
//      }
//  }
//catch
//  {

//  }