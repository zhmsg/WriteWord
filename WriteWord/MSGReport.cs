using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace WriteWord
{
    class MSGReport
    {
        _Application wordApp;
        _Document wordDoc;
        Object Nothing;
        string FontName = "造字工房悦黑体验版常规体";
        string GreenPicPath;
        string RedPicPath;
        string YellowPicPath;
        float RiskPicWidth = 1.6f;
        float RiskPicHeight = 0.4f;
        public MSGReport()
        {
            Console.WriteLine("初始化MSGReport");
            wordApp = new ApplicationClass();
            Nothing = Missing.Value;
            wordDoc = wordApp.Documents.Add(ref Nothing, ref Nothing, ref Nothing, ref Nothing); 
            wordDoc.PageSetup.PaperSize = WdPaperSize.wdPaperA4;//设置纸张样式
            wordDoc.PageSetup.PageWidth = wordApp.CentimetersToPoints(26);
            wordDoc.PageSetup.PageHeight = wordApp.CentimetersToPoints(26);
            wordDoc.PageSetup.TopMargin = wordApp.CentimetersToPoints(0.68f);
            wordDoc.PageSetup.BottomMargin = wordApp.CentimetersToPoints(0.68f);
            wordDoc.PageSetup.LeftMargin = wordApp.CentimetersToPoints(1.21f);
            wordDoc.PageSetup.RightMargin = wordApp.CentimetersToPoints(1.21f);
            YellowPicPath = Environment.CurrentDirectory + "\\yellow.png";
            RedPicPath = Environment.CurrentDirectory + "\\red.png";
            GreenPicPath = Environment.CurrentDirectory + "\\green.png";
        }

        private void WriteSpeech(string name)
        {
            Console.WriteLine("开始Write Speech");

            #region 插入左边图片
            string SpeechPicPath = Environment.CurrentDirectory + "\\speech.jpg";
            InlineShape InShapeSpeech = wordDoc.Paragraphs.Last.Range.InlineShapes.AddPicture(SpeechPicPath); 
            Shape ShapeSpeech = InShapeSpeech.ConvertToShape();
            ShapeSpeech.WrapFormat.Type = WdWrapType.wdWrapSquare;// 四周型
            ShapeSpeech.WrapFormat.Side = WdWrapSideType.wdWrapRight; // 设置图片 文字环绕 -> 自动换行 -> 只在右侧
            #endregion

            wordDoc.Content.InsertAfter("\n");
            Range r = wordDoc.Paragraphs.Last.Range;
            r.Text = "致辞\n";
            r.Font.Size = 20f;
            r.Font.Name = FontName;
            r.Select();
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            wordApp.Selection.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceDouble;
            wordApp.Selection.ParagraphFormat.LeftIndent = wordApp.CentimetersToPoints(12);

            r = wordDoc.Paragraphs.Last.Range;
            r.Text = "尊敬的" + name + "您好！\n";
            r.Font.Size = 12f;
            r.Select();

            string ZhiCiPath = Environment.CurrentDirectory + "\\zhici.txt";
            StreamReader sr = new StreamReader(ZhiCiPath);
            r = wordDoc.Paragraphs.Last.Range;
            r.Text = sr.ReadToEnd();
            sr.Close();
            r.Font.Size = 10f;
            r.Select();
            wordApp.Selection.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
            wordApp.Selection.ParagraphFormat.LineSpacing = 20f;

            #region 设置页眉
            wordApp.ActiveWindow.View.Type = WdViewType.wdOutlineView;//视图样式
            wordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekPrimaryHeader;//进入页眉设置，其中页眉边距在页面设置中已完成

            wordApp.ActiveWindow.ActivePane.Selection.ParagraphFormat.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;//去掉页眉的横线

            wordApp.ActiveWindow.ActivePane.Selection.Borders[WdBorderType.wdBorderBottom].Visible = false;
            wordApp.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekMainDocument;//退出页眉设置
            wordDoc.Sections.Last.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
            wordDoc.Sections.Last.PageSetup.HeaderDistance = wordApp.CentimetersToPoints(0.2f);//页眉位置
            wordDoc.Sections.Last.PageSetup.TopMargin = wordApp.CentimetersToPoints(0.2f);
            
            #endregion

            object bs = WdBreakType.wdSectionBreakNextPage;
            wordDoc.Paragraphs.Last.Range.InsertBreak(ref bs);
        }

        private void WriteContents(ReportInfo ri)
        {
            Console.WriteLine("开始Write Contents");

            #region 设置页面布局
            wordDoc.Sections.Last.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
            wordDoc.Sections.Last.PageSetup.TopMargin = wordApp.CentimetersToPoints(1f);
            wordDoc.Sections.Last.PageSetup.BottomMargin = wordApp.CentimetersToPoints(1f);
            wordDoc.Sections.Last.PageSetup.HeaderDistance = 20.5f; //页眉位置
            #endregion

            int LeftIndent = 6;
            Range r = wordDoc.Paragraphs.Last.Range;
            r.Text = "目录\n";
            r.Font.Size = 20f;
            r.Font.Name = FontName;
            r.Select();
            wordApp.Selection.ParagraphFormat.LineUnitBefore = 4;
            wordApp.Selection.ParagraphFormat.CharacterUnitLeftIndent = LeftIndent;
            r = wordDoc.Paragraphs.Last.Range;
            r.Text = "Contents\n";
            r.Font.Size = 20f;
            r.Select();
            wordApp.Selection.ParagraphFormat.Reset();
            wordApp.Selection.ParagraphFormat.CharacterUnitLeftIndent = LeftIndent;
            wordApp.Selection.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
            wordApp.Selection.ParagraphFormat.LineSpacing = 22f;
            
            object bc = WdBreakType.wdSectionBreakContinuous;
            wordDoc.Paragraphs.Last.Range.InsertBreak(ref bc);

            r = wordDoc.Paragraphs.Last.Range;
            int RangeStart = r.Start;
            for (int i = 0; i < ri.PartInfo.Count; i++)
            {
                r = wordDoc.Paragraphs.Last.Range;
                r.Text = ri.PartInfo[i].PartStr + "\n";
                r.Font.Size = 12f;
                r.Select();
                wordApp.Selection.ParagraphFormat.Reset();
                wordApp.Selection.ParagraphFormat.LineUnitBefore = 3;
                wordApp.Selection.ParagraphFormat.CharacterUnitLeftIndent = LeftIndent;
                wordApp.Selection.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
                wordApp.Selection.ParagraphFormat.LineSpacing = 19f;

                r = wordDoc.Paragraphs.Last.Range;
                r.Text = ri.PartInfo[i].ChTitle + "\n";
                r.Select();
                wordApp.Selection.ParagraphFormat.Reset();
                wordApp.Selection.ParagraphFormat.CharacterUnitLeftIndent = LeftIndent;
                wordApp.Selection.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
                wordApp.Selection.ParagraphFormat.LineSpacing = 19f;

                r = wordDoc.Paragraphs.Last.Range;
                r.Text = ri.PartInfo[i].EnTitle + "\n";
                r.Font.Size = 10f;
            }
            r.Start = RangeStart;
            r.Select();
            wordApp.Selection.PageSetup.TextColumns.SetCount(2);
            wordApp.Selection.PageSetup.TextColumns.Spacing = 2f;


            
            object bs = WdBreakType.wdSectionBreakNextPage;
            wordDoc.Paragraphs.Last.Range.InsertBreak(ref bs);

            r = wordDoc.Paragraphs.Last.Range;
            r.Select();
            wordApp.Selection.ParagraphFormat.Reset();
            wordApp.Selection.PageSetup.TextColumns.SetCount(1);
        }

        private void WritePartTitle(ReportPartInfo PI, bool InsertTitle = true)
        {
            Console.WriteLine("开始Write Part Title " + PI.ChTitle);
            Range r;
            if (InsertTitle == true)
            {
                r = wordDoc.Paragraphs.Last.Range;
                r.Text = PI.ChTitle + "\n";
                r.Font.Size = 28.5f;
                r.Font.Bold = 0;
                r.Font.Name = FontName;
                r.Select();

                wordApp.Selection.ParagraphFormat.Reset();
                wordApp.Selection.PageSetup.TextColumns.SetCount(1);
                wordApp.Selection.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1;
                wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                wordApp.Selection.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
                wordApp.Selection.ParagraphFormat.LineSpacing = 43.4f;

                r = wordDoc.Paragraphs.Last.Range;
                r.Text = PI.EnTitle + "\n";
                r.Font.Size = 11.5f;
                r.Select();
                wordApp.Selection.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevelBodyText;
                wordApp.Selection.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
                wordApp.Selection.ParagraphFormat.LineSpacing = 26.1f;
            }
            #region 设置页面布局
            wordDoc.Sections.Last.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
            wordDoc.Sections.Last.PageSetup.TopMargin = wordApp.CentimetersToPoints(1f);
            wordDoc.Sections.Last.PageSetup.BottomMargin = wordApp.CentimetersToPoints(1f);
            wordDoc.Sections.Last.PageSetup.HeaderDistance = 20.5f;//页眉位置
            #endregion

            #region 设置页眉
            wordApp.ActiveWindow.View.Type = WdViewType.wdOutlineView;//视图样式
            wordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekPrimaryHeader;//进入页眉设置，其中页眉边距在页面设置中已完成
            try
            {
                while (true)
                    wordApp.ActiveWindow.ActivePane.View.NextHeaderFooter();
            }
            catch (Exception ex)
            {
            }
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            wordApp.ActiveWindow.ActivePane.Selection.ParagraphFormat.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;//去掉页眉的横线

            wordApp.ActiveWindow.ActivePane.Selection.Borders[WdBorderType.wdBorderBottom].Visible = false;
            wordApp.Selection.WholeStory();
            wordApp.Selection.Range.Text = "";
            r = wordApp.Selection.Paragraphs.Last.Range;
            r.Text = PI.PartStr + PI.ChTitle + "  " + PI.EnTitle + "\n";
            r.Font.Size = 7.68f;
            r.Font.Name = "造字工房悦黑体验版常规体";
            r.Start = r.Start + PI.PartStr.Length + PI.ChTitle.Length + 2;
            r.Font.Name = "Agency FB";
            r.Start = r.End;

            string HeaderPicPath = Environment.CurrentDirectory + "\\header.png";
            r.InlineShapes.AddPicture(HeaderPicPath);
            wordApp.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekMainDocument;//退出页眉设置
            #endregion

        }

        private void WriteHealthSurvey(ReportSurvey rs)
        {
            Console.WriteLine("开始写体检人健康调查情况汇总");

            Range r = wordDoc.Paragraphs.Last.Range;
            r.Text = "体检人健康调查情况汇总\n";
            r.Font.Size = 20f;
            r.Select();
            wordApp.Selection.ParagraphFormat.Reset();
            wordApp.Selection.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel2;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            float LeftIndent = 5;
            WriteThirdTitle("个人信息", true, LeftIndent);
            WriteNormalParagraph(rs.PersonalInfo, LeftIndent);
            WriteThirdTitle("疾病家族史", true, LeftIndent);
            WriteThirdTitle("生活习惯", true, LeftIndent);
            WriteThirdTitle("疾病既往史", true, LeftIndent);

        }

        private string GetRiskPicPath(string RiskLevel, bool IsDisease = true)
        {
            string RiskPicPath = "";
            if (IsDisease == true)
            {
                if (RiskLevel == "2")
                    RiskPicPath = RedPicPath;
                else if (RiskLevel == "1")
                    RiskPicPath = YellowPicPath;
                else if (RiskLevel == "0")
                    RiskPicPath = GreenPicPath;
                else
                    RiskPicPath = GreenPicPath;
            }
            else
            {
                if (RiskLevel == "0")
                    RiskPicPath = RedPicPath;
                else if (RiskLevel == "1")
                    RiskPicPath = GreenPicPath;
                else
                    RiskPicPath = GreenPicPath;
            }
            return RiskPicPath;
        }

        private void WriteOverviewContent(string Title, List<ReportTestItem> ListRTI, bool IsDisease = true)
        {
            Console.WriteLine("开始写概况中" + Title + "目录");
            Range r = wordDoc.Paragraphs.Last.Range;
            r.Text = "\n";
            r.Font.Size = 5;
            r = wordDoc.Paragraphs.Last.Range;
            r.Text = Title + "\n";
            r.Font.Size = 13f;
            r.Select();
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            wordApp.Selection.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceDouble;
            object bc = WdBreakType.wdSectionBreakContinuous;
            wordDoc.Paragraphs.Last.Range.InsertBreak(ref bc);
            int RangeStart = r.End;
            bool SplitAll = false;
            int ColumnCount = 2;
            float RightPos = 0;
            float RightIndent = 0;
            int KeepWithNext = 1;
            for (int i = 0; i < ListRTI.Count; i++)
            {
                if (i % ColumnCount == 0)
                {
                    RightPos = -16f;
                    RightIndent = 5f;
                    KeepWithNext = 0;
                }
                else
                {
                    RightPos = -6.2f;
                    RightIndent = 7f;
                    KeepWithNext = -1;
                }
                r = wordDoc.Paragraphs.Last.Range;
                r.Text = ListRTI[i].disease_name + " ";
                r.Font.Size = 9.6f;
                r.Select();
                wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                wordApp.Selection.ParagraphFormat.RightIndent = wordApp.CentimetersToPoints(RightIndent);
                wordApp.Selection.ParagraphFormat.KeepWithNext = KeepWithNext;
                r.Start = r.End;
                string DiseaseRiskPicPath = GetRiskPicPath(ListRTI[i].risk_level, IsDisease);
                InlineShape InShapeRisk = r.InlineShapes.AddPicture(DiseaseRiskPicPath);
                InShapeRisk.Height = wordApp.CentimetersToPoints(RiskPicHeight);
                InShapeRisk.Width = wordApp.CentimetersToPoints(RiskPicWidth);
                
                

                Shape ShapeRisk = InShapeRisk.ConvertToShape();
                ShapeRisk.WrapFormat.Type = WdWrapType.wdWrapSquare;// 紧密型
                ShapeRisk.WrapFormat.Side = WdWrapSideType.wdWrapLeft; // 设置图片 文字环绕 -> 自动换行 -> 只在左侧
                ShapeRisk.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionLine;
                ShapeRisk.Top = wordApp.CentimetersToPoints(0.4f);
                ShapeRisk.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionRightMarginArea;
                ShapeRisk.Left = wordApp.CentimetersToPoints(RightPos);
                wordDoc.Content.InsertAfter("\n");
                SplitAll = false;
                if (i % ColumnCount == 0)
                    RangeStart = r.Start;
                else if (i % ColumnCount == ColumnCount - 1)
                {
                    r.Start = RangeStart;
                    r.Select();
                    wordApp.Selection.PageSetup.TextColumns.SetCount(ColumnCount);
                    wordDoc.Paragraphs.Last.Range.InsertBreak(ref bc);
                    SplitAll = true;
                }
            }
            if (SplitAll == false)
            {
                r.Start = RangeStart;
                r.Select();
                wordApp.Selection.PageSetup.TextColumns.SetCount(ColumnCount);
                wordDoc.Paragraphs.Last.Range.InsertBreak(ref bc);
                SplitAll = true;
            }
            r = wordDoc.Paragraphs.Last.Range;
            r.Select();
            wordApp.Selection.ParagraphFormat.Reset();
            wordApp.Selection.PageSetup.TextColumns.SetCount(1);
        }

        private void WriteOverviewContent(string Title, List<ReportTestMedicine> ListRTM, bool IsDisease = true)
        {
            Console.WriteLine("开始写概况中" + Title + "目录");
            Range r = wordDoc.Paragraphs.Last.Range;
            r.Text = "\n";
            r.Font.Size = 5;
            r = wordDoc.Paragraphs.Last.Range;
            r.Text = Title + "\n";
            r.Font.Size = 13f;
            r.Select();
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            wordApp.Selection.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceDouble;
            object bc = WdBreakType.wdSectionBreakContinuous;
            wordDoc.Paragraphs.Last.Range.InsertBreak(ref bc);
            int RangeStart = r.End;
            bool SplitAll = false;
            int ColumnCount = 2;
            float RightPos = 0;
            float RightIndent = 0;
            int KeepWithNext = 1;
            for (int i = 0; i < ListRTM.Count; i++)
            {
                if (i % ColumnCount == 0)
                {
                    RightPos = -16f;
                    RightIndent = 5f;
                    KeepWithNext = 0;
                }
                else
                {
                    RightPos = -6.2f;
                    RightIndent = 7f;
                    KeepWithNext = -1;
                }
                r = wordDoc.Paragraphs.Last.Range;
                r.Text = ListRTM[i].disease_name + " ";
                r.Font.Size = 9.6f;
                r.Select();
                wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                wordApp.Selection.ParagraphFormat.RightIndent = wordApp.CentimetersToPoints(RightIndent);
                wordApp.Selection.ParagraphFormat.KeepWithNext = KeepWithNext;
                r.Start = r.End;
                string DiseaseRiskPicPath = GetRiskPicPath(ListRTM[i].risk_level, IsDisease);
                InlineShape InShapeRisk = r.InlineShapes.AddPicture(DiseaseRiskPicPath);
                InShapeRisk.Height = wordApp.CentimetersToPoints(RiskPicHeight);
                InShapeRisk.Width = wordApp.CentimetersToPoints(RiskPicWidth);



                Shape ShapeRisk = InShapeRisk.ConvertToShape();
                ShapeRisk.WrapFormat.Type = WdWrapType.wdWrapSquare;// 紧密型
                ShapeRisk.WrapFormat.Side = WdWrapSideType.wdWrapLeft; // 设置图片 文字环绕 -> 自动换行 -> 只在左侧
                ShapeRisk.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionLine;
                ShapeRisk.Top = wordApp.CentimetersToPoints(0.4f);
                ShapeRisk.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionRightMarginArea;
                ShapeRisk.Left = wordApp.CentimetersToPoints(RightPos);
                wordDoc.Content.InsertAfter("\n");
                SplitAll = false;
                if (i % ColumnCount == 0)
                    RangeStart = r.Start;
                else if (i % ColumnCount == ColumnCount - 1)
                {
                    r.Start = RangeStart;
                    r.Select();
                    wordApp.Selection.PageSetup.TextColumns.SetCount(ColumnCount);
                    wordDoc.Paragraphs.Last.Range.InsertBreak(ref bc);
                    SplitAll = true;
                }
            }
            if (SplitAll == false)
            {
                r.Start = RangeStart;
                r.Select();
                wordApp.Selection.PageSetup.TextColumns.SetCount(ColumnCount);
                wordDoc.Paragraphs.Last.Range.InsertBreak(ref bc);
                SplitAll = true;
            }
            r = wordDoc.Paragraphs.Last.Range;
            r.Select();
            wordApp.Selection.ParagraphFormat.Reset();
            wordApp.Selection.PageSetup.TextColumns.SetCount(1);
        }

        private void WriteGeneticTestOverview(ReportInfo ri)
        {
            WritePartTitle(ri.PartInfo[2]);
            WriteOverviewContent(ri.PartInfo[3].ChTitle, ri.tomour_list);
            WriteOverviewContent(ri.PartInfo[4].ChTitle, ri.chronic_disease_list);
            WriteOverviewContent(ri.PartInfo[5].ChTitle, ri.medicine_list, false);
            WriteOverviewContent(ri.PartInfo[6].ChTitle, ri.nutrition_list, false);
            object bs = WdBreakType.wdSectionBreakNextPage;
            wordDoc.Paragraphs.Last.Range.InsertBreak(ref bs);
        }

        private void WriteSecondTitle(string DiseaseName, string RiskLevel, bool IsDisease = true)
        {
            Range r = wordDoc.Paragraphs.Last.Range;
            r.Text = DiseaseName + "   ";
            r.Font.Size = 13.44f;
            r.Select();
            wordApp.Selection.ParagraphFormat.Reset();
            wordApp.Selection.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel2;
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            wordApp.Selection.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceDouble;
            r.Start = r.End;
            string DiseaseRiskPicPath = GetRiskPicPath(RiskLevel, IsDisease);
            InlineShape InShapeRisk = r.InlineShapes.AddPicture(DiseaseRiskPicPath);
            InShapeRisk.Height = wordApp.CentimetersToPoints(RiskPicHeight);
            InShapeRisk.Width = wordApp.CentimetersToPoints(RiskPicWidth);

            Shape ShapeRisk = InShapeRisk.ConvertToShape();
            ShapeRisk.WrapFormat.Type = WdWrapType.wdWrapTight;// 紧密型
            ShapeRisk.WrapFormat.Side = WdWrapSideType.wdWrapLeft; // 设置图片 文字环绕 -> 自动换行 -> 只在左侧
            ShapeRisk.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionLine;
            ShapeRisk.Top = wordApp.CentimetersToPoints(0.4f);
            wordDoc.Content.InsertAfter("\n");
        }

        private void WriteThirdTitle(string TitleName, bool InsertP=true, float LeftIndent = 0)
        {
            Range r;
            r = wordDoc.Paragraphs.Last.Range;
            r.Text = TitleName + "\n";
            r.Font.Size = 11.5f;
            r.Select();
            wordApp.Selection.ParagraphFormat.Reset();
            wordApp.Selection.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel3;
            wordApp.Selection.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpace1pt5;
            if (InsertP == true)
                wordApp.Selection.ParagraphFormat.LineUnitBefore = 1.5f;
            if(LeftIndent > 0)
                wordApp.Selection.ParagraphFormat.CharacterUnitLeftIndent = LeftIndent;
            r = wordDoc.Paragraphs.Last.Range;
            r.Select();
            wordApp.Selection.ParagraphFormat.Reset();
        }

        private Table CreateTable(int RowsCount, string[]T_Title)
        {
            Table t = wordDoc.Tables.Add(wordDoc.Paragraphs.Last.Range, RowsCount, T_Title.Length, ref Nothing, ref Nothing);//RowsCount行6列的表
            t.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;
            t.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
            t.Select();
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;//水平居中
            wordApp.Selection.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
            wordApp.Selection.ParagraphFormat.LineSpacing = 22f;
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter; // 垂直居中
            wordApp.Selection.Font.Size = 9.6f;
            t.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
            t.Borders.OutsideColor = (WdColor)(0x1 * 199 + 0x100 * 199 + 0x10000 * 199);
            Range r;
            for (int i = 1; i <= T_Title.Length; i++)
            {
                r = t.Cell(1, i).Range;
                r.Text = T_Title[i - 1];
            }
            t.Rows[1].Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
            t.Rows[1].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
            t.Rows[1].Borders[WdBorderType.wdBorderBottom].Color = (WdColor)(0x1 * 199 + 0x100 * 199 + 0x10000 * 199);
            return t;
        }

        private string CalcFreq2(string Freq)
        {
            Freq = Freq.Trim();
            if (Freq != "")
            {
                Freq += "000000";

                if (Freq[0] == '1')
                    Freq = "100.0%";
                else
                    Freq = Freq.Substring(2, 2) + "." + Freq.Substring(4, 2) + "%";
                Freq = Freq.TrimStart('0');
                if (Freq[0] == '.')
                    Freq = "0" + Freq;
            }
            
            return Freq;
        }

        public string CalcFreq(string GeneType, string Freq)
        {
            Freq = Freq.Trim();
            if (Freq != "")
            {
                float FFreq = float.Parse(Freq);
                string []GS = GeneType.Split('/');
                if (GS[0] == GS[1])
                {
                    FFreq = FFreq * FFreq * 100;
                }
                else
                {
                    FFreq = (1 - FFreq) * FFreq * 200;
                }
                Freq = String.Format("{0:F}%", FFreq);
            }

            return Freq;
        }

        private void WriteNormalParagraph(string Content, float LeftIndent = 0)
        {
            Range r = wordDoc.Paragraphs.Last.Range;
            r.Text = Content.Replace("<br>", "\n").Trim('\n') + "\n";
            r.Font.Size = 9.6f;
            r.Select();
            wordApp.Selection.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
            wordApp.Selection.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevelBodyText;
            wordApp.Selection.ParagraphFormat.LineSpacing = 19.2f;
            if (LeftIndent > 0)
                wordApp.Selection.ParagraphFormat.CharacterUnitLeftIndent = LeftIndent;
        }

        private void WriteCancerRiskInfo(ReportTestItem rti)
        {
            Console.WriteLine("开始Write Cancer Info " + rti.disease_name);

            #region 插入 疾病名称
            WriteSecondTitle(rti.disease_name, rti.risk_level);
            #endregion

            #region 插入 疾病介绍
            WriteNormalParagraph(rti.introduction);
            #endregion

            #region 插入 疑似易感性突变列表
            WriteThirdTitle("疑似易感性突变列表");

            int SampleVariantCount = rti.patient_sample_variant.Count;
            if (SampleVariantCount > 10)
                SampleVariantCount = 10;
            string[] title = { "序号", "位点编号", "基因名称", "基因型", "风险等级", "人群频率" };
            Table t = CreateTable(SampleVariantCount + 1, title);
            float[] ColWid = { 1.5f, 5f, 7f, 2f, 3f, 2.3f };
            Range r;
            for (int i = 1; i <= 6; i++)
            {
                t.Columns[i].Width = wordApp.CentimetersToPoints(ColWid[i - 1]);
                r = t.Cell(1, i).Range;
                r.Text = title[i - 1];
                r.Font.Size = 9.6f;
            }
            string VariantRiskLev = "";
            string VariantRiskPicPath = "";
            string Freq = "";
            string VariantRead = "";
            for (int i = 0; i < SampleVariantCount; i++)
            {
                r = t.Cell(i + 2, 1).Range;
                r.Text = (i + 1) + "";

                r = t.Cell(i + 2, 2).Range;
                r.Text = rti.patient_sample_variant[i][1];

                r = t.Cell(i + 2, 3).Range;
                r.Text = rti.patient_sample_variant[i][0];

                r = t.Cell(i + 2, 4).Range;
                r.Text = rti.patient_sample_variant[i][2];

                VariantRiskLev = rti.patient_sample_variant[i][4];

                if (VariantRiskLev == "2")
                {
                    VariantRiskPicPath = RedPicPath;
                    if (rti.patient_sample_variant[i][3] != "" && rti.patient_sample_variant[i][3] != "无")
                        VariantRead += rti.patient_sample_variant[i][3] + "\n";
                }
                else if (VariantRiskLev == "1")
                {
                    VariantRiskPicPath = YellowPicPath;
                    if (rti.patient_sample_variant[i][3] != "" && rti.patient_sample_variant[i][3] != "无")
                        VariantRead += rti.patient_sample_variant[i][3] + "\n";
                }
                else if (VariantRiskLev == "0")
                    VariantRiskPicPath = GreenPicPath;
                else
                    Console.WriteLine(VariantRiskLev);
                r = t.Cell(i + 2, 5).Range;
                InlineShape InShape = r.InlineShapes.AddPicture(VariantRiskPicPath);
                InShape.Height = wordApp.CentimetersToPoints(RiskPicHeight);
                InShape.Width = wordApp.CentimetersToPoints(RiskPicWidth);

                Freq = CalcFreq(rti.patient_sample_variant[i][2], rti.patient_sample_variant[i][5]);
                r = t.Cell(i + 2, 6).Range;
                r.Text = Freq;
            }
            #endregion

            #region 插入 基因检测结果解读
            WriteThirdTitle("基因检测结果解读");

            if (VariantRead == "")
                VariantRead = "暂无";
            WriteNormalParagraph(VariantRead);
            #endregion

            #region 插入 预防和干预
            WriteThirdTitle("预防和干预");

            WriteNormalParagraph(rti.prevention);
            #endregion

            #region 插入 早期发现
            WriteThirdTitle("早期发现");

            WriteNormalParagraph(rti.early_detection);
            #endregion

            #region 插入 患者用药提示
            WriteThirdTitle("患者用药提示");

            int MedTipCount = rti.medication_tips.Count;
            if (MedTipCount <= 0)
                WriteNormalParagraph("暂无建议");
            else
            {
                string[] Title_Tip = { "药物", "风险等级" };
                Table T_Tip = CreateTable(MedTipCount + 1, Title_Tip);
                string TipRiskPicPath = "";
                for (int i = 0; i < MedTipCount; i++)
                {
                    r = T_Tip.Cell(i + 2, 1).Range;
                    r.Text = rti.medication_tips[i][0];

                    string TipLevel = rti.medication_tips[i][1];
                    if (TipLevel == "0")
                    {
                        TipRiskPicPath = RedPicPath;
                    }
                    else if (TipLevel == "1")
                        TipRiskPicPath = GreenPicPath;
                    else
                    {
                        Console.WriteLine(rti.disease_name + " " + rti.medication_tips[i][0] + " [" + VariantRiskLev + "]");
                        continue;
                    }
                    r = T_Tip.Cell(i + 2, 2).Range;
                    InlineShape InShape = r.InlineShapes.AddPicture(TipRiskPicPath);
                    InShape.Height = wordApp.CentimetersToPoints(RiskPicHeight);
                    InShape.Width = wordApp.CentimetersToPoints(RiskPicWidth);
                }
            }
            
            #endregion

            #region 插入 参考文献
            WriteThirdTitle("参考文献");
            
            WriteNormalParagraph(rti.reference);
            #endregion

            r = wordDoc.Paragraphs.Last.Range;
            object bs = WdBreakType.wdPageBreak;
            r.InsertBreak(ref bs);
        }

        private void WriteChronicDiseaseRiskInfo(ReportTestItem rti)
        {
            Console.WriteLine("开始Write Chronic Disease Info " + rti.disease_name);

            #region 插入 疾病名称
            WriteSecondTitle(rti.disease_name, rti.risk_level);
            #endregion

            #region 插入 疾病介绍
            WriteNormalParagraph(rti.introduction);
            #endregion

            #region 插入 疑似易感性突变列表
            WriteThirdTitle("疑似易感性突变列表");

            int SampleVariantCount = rti.patient_sample_variant.Count;
            if (SampleVariantCount > 10)
                SampleVariantCount = 10;
            string[] title = { "序号", "位点编号", "基因名称", "基因型", "风险等级", "人群频率" };
            Table t = CreateTable(SampleVariantCount + 1, title);
            float[] ColWid = { 1.5f, 5f, 7f, 2f, 3f, 2.3f };
            Range r;
            for (int i = 1; i <= 6; i++)
            {
                t.Columns[i].Width = wordApp.CentimetersToPoints(ColWid[i - 1]);
                r = t.Cell(1, i).Range;
                r.Text = title[i - 1];
                r.Font.Size = 9.6f;
            }
            string VariantRiskLev = "";
            string VariantRiskPicPath = "";
            string Freq = "";
            string VariantRead = "";
            for (int i = 0; i < SampleVariantCount; i++)
            {
                r = t.Cell(i + 2, 1).Range;
                r.Text = (i + 1) + "";

                r = t.Cell(i + 2, 2).Range;
                r.Text = rti.patient_sample_variant[i][1];

                r = t.Cell(i + 2, 3).Range;
                r.Text = rti.patient_sample_variant[i][0];

                r = t.Cell(i + 2, 4).Range;
                r.Text = rti.patient_sample_variant[i][2];

                VariantRiskLev = rti.patient_sample_variant[i][4];

                if (VariantRiskLev == "2")
                {
                    VariantRiskPicPath = RedPicPath;
                    if (rti.patient_sample_variant[i][3] != "" && rti.patient_sample_variant[i][3] != "无")
                        VariantRead += rti.patient_sample_variant[i][3] + "\n";
                }
                else if (VariantRiskLev == "1")
                {
                    VariantRiskPicPath = YellowPicPath;
                    if (rti.patient_sample_variant[i][3] != "" && rti.patient_sample_variant[i][3] != "无")
                        VariantRead += rti.patient_sample_variant[i][3] + "\n";
                }
                else if (VariantRiskLev == "0")
                    VariantRiskPicPath = GreenPicPath;
                else
                    Console.WriteLine(VariantRiskLev);
                r = t.Cell(i + 2, 5).Range;
                InlineShape InShape = r.InlineShapes.AddPicture(VariantRiskPicPath);
                InShape.Height = wordApp.CentimetersToPoints(RiskPicHeight);
                InShape.Width = wordApp.CentimetersToPoints(RiskPicWidth);

                r = t.Cell(i + 2, 6).Range;
                Freq = CalcFreq(rti.patient_sample_variant[i][2], rti.patient_sample_variant[i][5]);
                r.Text = Freq;
            }
            #endregion

            #region 插入 基因检测结果解读
            WriteThirdTitle("基因检测结果解读");

            if (VariantRead == "")
                VariantRead = "暂无";
            WriteNormalParagraph(VariantRead);
            #endregion

            #region 插入 早期发现
            WriteThirdTitle("早期发现");

            WriteNormalParagraph(rti.early_detection);
            #endregion

            #region 插入 患者用药提示
            WriteThirdTitle("患者用药提示");
            bool HasMT = false;
            int MedTipCount = rti.medication_tips.Count;
            if (MedTipCount <= 0)
                WriteNormalParagraph("暂无建议");
            else
            {
                string[] Title_Tip = { "药物", "风险等级" };
                Table T_Tip = CreateTable(MedTipCount + 1, Title_Tip);
                
                string TipRiskPicPath = "";
                for (int i = 0; i < MedTipCount; i++)
                {
                    r = T_Tip.Cell(i + 2, 1).Range;
                    r.Text = rti.medication_tips[i][0];

                    string TipLevel = rti.medication_tips[i][1];
                    if (TipLevel == "0")
                    {
                        TipRiskPicPath = RedPicPath;
                    }
                    else if (TipLevel == "1")
                        TipRiskPicPath = GreenPicPath;
                    else
                    {
                        Console.WriteLine(rti.disease_name + " " + rti.medication_tips[i][0] + " [" + VariantRiskLev + "]");
                        continue;
                    }
                    r = T_Tip.Cell(i + 2, 2).Range;
                    InlineShape InShape = r.InlineShapes.AddPicture(TipRiskPicPath);
                    InShape.Height = wordApp.CentimetersToPoints(RiskPicHeight);
                    InShape.Width = wordApp.CentimetersToPoints(RiskPicWidth);
                }
                HasMT = true;
            }

            #endregion

            r = wordDoc.Paragraphs.Last.Range;
            r.InsertBreak();
            if (HasMT == true)
            {
                r.Start = r.Start - 2;
                r.End = r.Start + 1;
                r.Select();
                wordApp.Selection.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
                wordApp.Selection.ParagraphFormat.LineSpacing = 1f;
            }

        }

        private void WriteCurativeSide(ReportTestMedicine rtm)
        {
            Console.WriteLine("开始Write Curative Side Info " + rtm.disease_name);

            #region 插入 疾病名称
            WriteSecondTitle(rtm.disease_name, rtm.risk_level, false);
            #endregion

            #region 插入 与disease_name有效性相关的基因检测结果
            WriteThirdTitle("与" + rtm.disease_name + "有效性相关的基因检测结果", false);

            Range r;
            string GeneRiskLev = "";
            string GeneRiskPicPath = "";
            string Freq = "";
            string GeneRead = "";
            string[] title = { "基因", "位点", "基因型", "风险等级", "人群频率" };
            int ValidCount = rtm.patient_sample_variant[0].Count;
            if (ValidCount <= 0)
                WriteNormalParagraph("暂无数据");
            else
            {
                if (ValidCount > 5)
                    ValidCount = 5;
                Table T_Valid = CreateTable(ValidCount + 1, title);
                
                for (int i = 0; i < ValidCount; i++)
                {
                    r = T_Valid.Cell(i + 2, 1).Range;
                    r.Text = rtm.patient_sample_variant[0][i][0];

                    r = T_Valid.Cell(i + 2, 2).Range;
                    r.Text = rtm.patient_sample_variant[0][i][1];

                    r = T_Valid.Cell(i + 2, 3).Range;
                    r.Text = rtm.patient_sample_variant[0][i][2];

                    GeneRiskLev = rtm.patient_sample_variant[0][i][4];
                    if (GeneRiskLev == "0")
                    {
                        GeneRiskPicPath = RedPicPath;
                        if (rtm.patient_sample_variant[0][i][3] != "" && rtm.patient_sample_variant[0][i][3] != "无")
                            GeneRead += rtm.patient_sample_variant[0][i][3] + "\n";
                    }
                    else if (GeneRiskLev == "1")
                        GeneRiskPicPath = GreenPicPath;
                    else
                        return;
                    r = T_Valid.Cell(i + 2, 4).Range;
                    InlineShape InShape = r.InlineShapes.AddPicture(GeneRiskPicPath);
                    InShape.Height = wordApp.CentimetersToPoints(RiskPicHeight);
                    InShape.Width = wordApp.CentimetersToPoints(RiskPicWidth);

                    r = T_Valid.Cell(i + 2, 5).Range;
                    Freq = CalcFreq(rtm.patient_sample_variant[0][i][2], rtm.patient_sample_variant[0][i][5]);
                    r.Text = Freq;
                }
            }
            #endregion

            #region 插入 与disease_name毒性相关的基因检测结果
            WriteThirdTitle("与" + rtm.disease_name +"毒性相关的基因检测结果");

            int PoisonCount = rtm.patient_sample_variant[1].Count;
            if (PoisonCount == 0)
                WriteNormalParagraph("暂无数据");
            else
            {
                if (PoisonCount > 5)
                    PoisonCount = 5;
                Table T_Poison = CreateTable(PoisonCount + 1, title);
                List<List<string>> PoisonList = rtm.patient_sample_variant[1];
                for (int i = 0; i < PoisonCount; i++)
                {
                    r = T_Poison.Cell(i + 2, 1).Range;
                    r.Text = PoisonList[i][0];

                    r = T_Poison.Cell(i + 2, 2).Range;
                    r.Text = PoisonList[i][1];

                    r = T_Poison.Cell(i + 2, 3).Range;
                    r.Text = PoisonList[i][2];

                    GeneRiskLev = PoisonList[i][4];
                    string OneGeneRead = PoisonList[i][3];
                    if (GeneRiskLev == "0")
                    {
                        GeneRiskPicPath = RedPicPath;
                        if (OneGeneRead != "" && OneGeneRead != "无")
                            GeneRead += OneGeneRead + "\n";
                    }
                    else if (GeneRiskLev == "1")
                        GeneRiskPicPath = GreenPicPath;
                    else
                        return;
                    r = T_Poison.Cell(i + 2, 4).Range;
                    InlineShape InShape = r.InlineShapes.AddPicture(GeneRiskPicPath);
                    InShape.Height = wordApp.CentimetersToPoints(RiskPicHeight);
                    InShape.Width = wordApp.CentimetersToPoints(RiskPicWidth);

                    r = T_Poison.Cell(i + 2, 5).Range;
                    Freq = CalcFreq(PoisonList[i][2], PoisonList[i][5]);
                    r.Text = Freq;
                }
            }
            #endregion

            #region 插入 受检人基因型解读
            WriteThirdTitle("受检人基因型解读");

            if (GeneRead == "")
                GeneRead = "暂无";
            WriteNormalParagraph(GeneRead);
            #endregion

            r = wordDoc.Paragraphs.Last.Range;
            r.InsertBreak();

        }

        private void WriteNutritionInfo(ReportTestItem rti)
        {
            Console.WriteLine("开始Write Nutrition Info " + rti.disease_name);

            #region 插入 疾病名称
            WriteSecondTitle(rti.disease_name, rti.risk_level, false);
            #endregion

            #region 插入 疾病介绍
            WriteNormalParagraph(rti.introduction);
            #endregion

            #region 插入 受检人基因检测结果
            WriteThirdTitle("受检人基因检测结果");

            Range r;
            string[] title = { "基因", "位点", "基因型", "风险等级", "人群频率" };
            int VariantCount = rti.patient_sample_variant.Count;
            if (VariantCount > 5)
                VariantCount = 5;
            string GeneRiskLev = "";
            string GeneRiskPicPath = "";
            string Freq = "";
            string GeneRead = "";
            if (VariantCount == 0)
                WriteNormalParagraph("暂无数据");
            else
            {
                Table T_Poison = CreateTable(VariantCount + 1, title);
                List<List<string>> VariantList = rti.patient_sample_variant;
                for (int i = 0; i < VariantCount; i++)
                {
                    r = T_Poison.Cell(i + 2, 1).Range;
                    r.Text = VariantList[i][0];

                    r = T_Poison.Cell(i + 2, 2).Range;
                    r.Text = VariantList[i][1];

                    r = T_Poison.Cell(i + 2, 3).Range;
                    r.Text = VariantList[i][2];

                    GeneRiskLev = VariantList[i][4];
                    string OneGeneRead = VariantList[i][3];
                    if (GeneRiskLev == "0")
                    {
                        GeneRiskPicPath = RedPicPath;
                        if (OneGeneRead != "" && OneGeneRead != "无")
                            GeneRead += OneGeneRead + "\n";
                    }
                    else if (GeneRiskLev == "1")
                        GeneRiskPicPath = GreenPicPath;
                    else
                        return;
                    r = T_Poison.Cell(i + 2, 4).Range;
                    InlineShape InShape = r.InlineShapes.AddPicture(GeneRiskPicPath);
                    InShape.Height = wordApp.CentimetersToPoints(RiskPicHeight);
                    InShape.Width = wordApp.CentimetersToPoints(RiskPicWidth);

                    r = T_Poison.Cell(i + 2, 5).Range;
                    Freq = CalcFreq(VariantList[i][2], VariantList[i][5]);
                    r.Text = Freq;
                }
            }
            #endregion


            #region 插入 受检人基因型解读
            WriteThirdTitle("受检人基因型解读");

            if (GeneRead == "")
                GeneRead = "暂无";
            WriteNormalParagraph(GeneRead);
            #endregion

            r = wordDoc.Paragraphs.Last.Range;
            r.InsertBreak();

        }


        private void ReplacePMToM()
        {
            Console.WriteLine("开始将所有换行分页符替换成分页符");
            wordDoc.Content.Select();
            wordApp.Selection.Find.ClearFormatting();
            wordApp.Selection.Find.Text = "^p^m";
            wordApp.Selection.Find.Wrap = WdFindWrap.wdFindContinue;
            wordApp.Selection.Find.MatchWholeWord = false;
            wordApp.Selection.Find.Format = false;
            wordApp.Selection.Find.MatchCase = false;
            wordApp.Selection.Find.Forward = true;
            wordApp.Selection.Find.Replacement.ClearFormatting();
            wordApp.Selection.Find.Replacement.Text = "^m";
            wordApp.Selection.Find.MatchWholeWord = false;
            wordApp.Selection.Find.MatchWildcards = false;
            wordApp.Selection.Find.MatchSoundsLike = false;
            wordApp.Selection.Find.MatchByte = false;
            object ReplacAll =  WdReplace.wdReplaceAll;
            wordApp.Selection.Find.Execute(ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref ReplacAll);
        }

        private void SetPageBreak()
        {
            Console.WriteLine("开始将所有分页符的行间距设成1磅");
            wordDoc.Content.Select();
            wordApp.Selection.Find.ClearFormatting();
            wordApp.Selection.Find.Text = "^m";
            wordApp.Selection.Find.Wrap = WdFindWrap.wdFindContinue;
            wordApp.Selection.Find.MatchWholeWord = false;
            wordApp.Selection.Find.Format = false;
            wordApp.Selection.Find.MatchCase = false;
            wordApp.Selection.Find.Forward = true;
            wordApp.Selection.Find.Replacement.Text = "";
            wordApp.Selection.Find.MatchWholeWord = false;
            wordApp.Selection.Find.MatchWildcards = false;
            wordApp.Selection.Find.MatchSoundsLike = false;
            wordApp.Selection.Find.MatchByte = false;
            wordApp.Selection.Find.Execute();
            while (wordApp.Selection.Find.Found == true)
            {
                wordApp.Selection.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
                wordApp.Selection.ParagraphFormat.LineSpacing = 12f;
                wordApp.Selection.Find.Execute();
            }
        }

        private void SetEnFont()
        {
            wordDoc.Content.Select();
            wordApp.Selection.Find.ClearFormatting();
            wordApp.Selection.Find.Text = "^$";
            wordApp.Selection.Find.Wrap = WdFindWrap.wdFindContinue;
            wordApp.Selection.Find.MatchWholeWord = false;
            wordApp.Selection.Find.Format = false;
            wordApp.Selection.Find.MatchCase = false;
            wordApp.Selection.Find.Forward = true;
            wordApp.Selection.Find.Replacement.Text = "";
            wordApp.Selection.Find.MatchWholeWord = false;
            wordApp.Selection.Find.MatchWildcards = false;
            wordApp.Selection.Find.MatchSoundsLike = false;
            wordApp.Selection.Find.MatchByte = false;
            wordApp.Selection.Find.Execute();
            while (wordApp.Selection.Find.Found == true)
            {
                wordApp.Selection.Font.Name = "Agency FB";
                wordApp.Selection.Find.Execute();
            }
        }

        public void WriteReport(ReportInfo ri)
        {
            Console.WriteLine("开始写" + ri.survey_result.name + "的报告");
            //WriteSpeech(ri.survey_result.name);

            //WriteContents(ri);

            //WritePartTitle(ri.PartInfo[0], false);
            //WriteHealthSurvey(ri.survey_result);

            object bs = WdBreakType.wdSectionBreakNextPage;
            //wordDoc.Paragraphs.Last.Range.InsertBreak(ref bs);

            //WriteGeneticTestOverview(ri);

            WritePartTitle(ri.PartInfo[3]);
            for (int i = 0; i < ri.tomour_list.Count; i++)
            {
                if (ri.tomour_list[i].risk_level == "2" || ri.tomour_list[i].risk_level == "1")
                {
                    ReportTestItem rti = ri.tomour_list[i];
                    WriteCancerRiskInfo(rti);
                    //break;
                }
            }
            wordDoc.Paragraphs.Last.Range.InsertBreak(ref bs);
            WritePartTitle(ri.PartInfo[4]);
            for (int i = 0; i < ri.chronic_disease_list.Count; i++)
            {
                if (ri.chronic_disease_list[i].risk_level == "2" || ri.chronic_disease_list[i].risk_level == "1")
                {
                    ReportTestItem rti = ri.chronic_disease_list[i];
                    WriteChronicDiseaseRiskInfo(rti);
                    //break;
                }
            }
            wordDoc.Paragraphs.Last.Range.InsertBreak(ref bs);
            WritePartTitle(ri.PartInfo[5]);
            for (int i = 0; i < ri.medicine_list.Count; i++)
            {
                ReportTestMedicine rtm = ri.medicine_list[i];
                WriteCurativeSide(rtm);
                //break;
            }
            wordDoc.Paragraphs.Last.Range.InsertBreak(ref bs);
            WritePartTitle(ri.PartInfo[6]);
            for (int i = 0; i < ri.nutrition_list.Count; i++)
            {
                ReportTestItem rti = ri.nutrition_list[i];
                WriteNutritionInfo(rti);
                //break;
            }
            wordDoc.Paragraphs.Last.Range.InsertBreak(ref bs);
            WritePartTitle(ri.PartInfo[7]);
            for (int i = 0; i < ri.tomour_list.Count; i++)
            {
                if (ri.tomour_list[i].risk_level == "0")
                {
                    ReportTestItem rti = ri.tomour_list[i];
                    WriteCancerRiskInfo(rti);
                    //break;
                }
            }
            for (int i = 0; i < ri.chronic_disease_list.Count; i++)
            {
                if (ri.chronic_disease_list[i].risk_level == "0")
                {
                    ReportTestItem rti = ri.chronic_disease_list[i];
                    WriteChronicDiseaseRiskInfo(rti);
                    //break;
                }
            }
            ReplacePMToM();
        }

        public void SaveASDocx(object WordPath)
        {
            Console.WriteLine("开始保存为DOCX");
            object WordFormat = WdSaveFormat.wdFormatDocumentDefault;
            wordDoc.SaveAs(ref WordPath, ref WordFormat, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
        }

        public void SaveASHTML(object HTMLPath)
        {
            Console.WriteLine("开始保存为HTML");
            object HTMLFormat = WdSaveFormat.wdFormatHTML;
            wordDoc.SaveAs(ref HTMLPath, ref HTMLFormat, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
        }

        public void Close()
        {
            wordDoc.Close(ref Nothing, ref Nothing, ref Nothing);
            wordApp.Quit(ref Nothing, ref Nothing, ref Nothing);
        }
    }
}
