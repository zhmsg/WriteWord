using MSWord = Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using Microsoft.Office.Interop.Word;

namespace WriteWord
{
    class MSGWord
    {
        public void CreateReport()
        {
            object path;//文件路径
            string strContent;//文件内容
            _Application wordApp;//Word应用程序变量
            MSWord._Document wordDoc;//Word文档变量
            Console.WriteLine(Environment.CurrentDirectory);
            path = Environment.CurrentDirectory + "\\myWord.docx"; //保存为Word2003文档
            wordApp = new MSWord.ApplicationClass();//初始化
            if (File.Exists((string)path))
            {
                File.Delete((string)path);
            }
            Object Nothing = Missing.Value;
            wordDoc = wordApp.Documents.Add(ref Nothing, ref Nothing, ref Nothing, ref Nothing);

            #region 页面设置
            //
            wordDoc.PageSetup.PaperSize = Microsoft.Office.Interop.Word.WdPaperSize.wdPaperA4;//设置纸张样式
            
            wordDoc.PageSetup.Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientPortrait;//排列方式为垂直方向
            //wordDoc.PageSetup.TopMargin = 57.0f;
            //wordDoc.PageSetup.BottomMargin = 57.0f;
            //wordDoc.PageSetup.LeftMargin = 57.0f;
            //wordDoc.PageSetup.RightMargin = 57.0f;
            //wordDoc.PageSetup.HeaderDistance = 30.0f;//页眉位置
            #endregion

            #region 设置页眉
            wordApp.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdOutlineView;//视图样式
            wordApp.ActiveWindow.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekPrimaryHeader;//进入页眉设置，其中页眉边距在页面设置中已完成
            wordApp.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight;
            wordApp.Selection.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphRight;
            wordApp.ActiveWindow.ActivePane.Selection.InsertAfter("洛阳晶云信息科技有限公司");
            //去掉页眉的横线
            wordApp.ActiveWindow.ActivePane.Selection.ParagraphFormat.Borders[MSWord.WdBorderType.wdBorderBottom].LineStyle = MSWord.WdLineStyle.wdLineStyleNone;
            wordApp.ActiveWindow.ActivePane.Selection.Borders[MSWord.WdBorderType.wdBorderBottom].Visible = false;
            wordApp.ActiveWindow.ActivePane.View.SeekView = MSWord.WdSeekView.wdSeekMainDocument;//退出页眉设置
            #endregion

            #region 设置标头
            string filename = Environment.CurrentDirectory + "\\title_icon.png";
            //定义要向文档中插入图片的位置
            object range = wordApp.Selection.Range;
            //定义该图片是否为外部链接
            object linkToFile = false;//默认
            //定义插入的图片是否随word一起保存
            object saveWithDocument = true;
            //向word中写入图片
            object left = 82;
            object top = 65;
            Shape TitlePic = wordDoc.Shapes.AddPicture(filename, ref linkToFile, ref saveWithDocument, ref left, ref top, ref Nothing, ref Nothing, ref Nothing);
            
            TitlePic.Height = wordApp.CentimetersToPoints(2.84f);
            TitlePic.Width = wordApp.CentimetersToPoints(3.22f);
            
            //TitlePic.ConvertToShape().WrapFormat.Type = MSWord.WdWrapType.wdWrapFront;// 浮于文字上方
            TitlePic.Select(ref Nothing);
            wordDoc.Content.InsertAfter("\n");
            
            //object unite = WdUnits.wdStory;
            //wordApp.Selection.EndKey(ref unite, ref Nothing);
            //wordApp.Selection.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphLeft;//左对齐显示图片


            //object unite2 = WdUnits.wdStory;
            //wordApp.Selection.EndKey(ref unite2, ref Nothing);
            //wordApp.Selection.MoveDown(ref Nothing, ref Nothing, ref Nothing); //移动焦点
            //wordDoc.Paragraphs.Last.Range.Select();
            object unitline = WdUnits.wdLine;
            wordApp.Selection.MoveUp(ref unitline, ref Nothing, ref Nothing); //移动焦点
            wordApp.Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 0.5f; //首行缩进0.5字符
            wordApp.Selection.ParagraphFormat.CharacterUnitLeftIndent = 9.07f;  //左缩进9.07字符
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
            strContent = "中国人民解放军总医院病理科";
            Range r = wordApp.Selection.Range;
            r.Text = strContent;
            r.Font.Name = "黑体";
            r.Font.Size = 22;
            r.Font.Bold = 1;//Bold=0为不加粗
            

            //移动焦点并换行
            wordApp.Selection.MoveDown(ref Nothing, ref Nothing, ref Nothing); //移动焦点

            
            wordApp.Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 2.98f; //首行缩进2.98字符
            wordApp.Selection.ParagraphFormat.CharacterUnitLeftIndent = 9.07f;  //左缩进9.07字符
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
            strContent = "分子病理检测报告\n";
            r = wordDoc.Paragraphs.Last.Range;
            r.Text = strContent;
            r.Font.Name = "黑体";
            r.Font.Size = 22;
            r.Font.Bold = 1;  //Bold=0为不加粗
            #endregion

            #region 设置基本信息
            wordApp.Selection.MoveDown(ref Nothing, ref Nothing, ref Nothing);//移动焦点
            wordApp.Selection.ParagraphFormat.Reset();
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            strContent = "\n";
            r = wordDoc.Paragraphs.Last.Range;
            r.Text = strContent;
            r.Font.Name = "黑体";
            r.Font.Size = 12;
            r.Font.Bold = 0;
            

            wordApp.Selection.MoveDown(ref Nothing, ref Nothing, ref Nothing);//移动焦点
            wordApp.Selection.ParagraphFormat.Reset();
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            
            wordApp.Selection.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
            wordApp.Selection.ParagraphFormat.LineSpacing = 17;
            strContent = "姓 名:       性 别:         年 龄:       病理号：\n";
            r = wordDoc.Paragraphs.Last.Range;
            r.Text = strContent;
            r.Font.Name = "宋体";
            //r.Font.Size = 12;
            //r.Font.Bold = 0;

            wordApp.Selection.MoveDown(ref Nothing, ref Nothing, ref Nothing);//移动焦点
            //wordApp.Selection.ParagraphFormat.Reset();
            //wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            //wordApp.Selection.ParagraphFormat.LineSpacing = 17;
            strContent = "出生地：     送检部位：     标本类型：石蜡样本           门诊号：\n";
            r = wordDoc.Paragraphs.Last.Range;
            r.Text = strContent;
            //r.Font.Name = "宋体";
            //r.Font.Size = 12;
            //r.Font.Bold = 0;

            wordApp.Selection.MoveDown(ref Nothing, ref Nothing, ref Nothing);//移动焦点
            //wordApp.Selection.ParagraphFormat.Reset();
            //wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            //wordApp.Selection.ParagraphFormat.LineSpacing = 17;
            strContent = "检测方法：NGS   申请医生:      申请日期：       住院号：\n";
            r = wordDoc.Paragraphs.Last.Range;
            r.Text = strContent;
            //r.Font.Name = "宋体";
            //r.Font.Size = 12;
            //r.Font.Bold = 0;

            wordApp.Selection.MoveDown(ref Nothing, ref Nothing, ref Nothing);//移动焦点
            Microsoft.Office.Interop.Word.Shape ShapeLine = wordDoc.Shapes.AddLine(90, 210, 360, 210, ref Nothing);
            
            ShapeLine.Width = wordApp.CentimetersToPoints(14.92f);
            ShapeLine.Line.ForeColor.RGB = 0;  // 线条颜色 黑色
            ShapeLine.Line.Weight = 1.5f;  // 线条粗细1.5磅
            wordDoc.Content.InsertAfter("\n");

            wordApp.Selection.MoveDown(ref Nothing, ref Nothing, ref Nothing);//移动焦点
            //wordApp.Selection.ParagraphFormat.Reset();
            //wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            wordApp.Selection.ParagraphFormat.Space1();
            strContent = "主要设备：ION TORRENT PGM 测序仪\n";
            r = wordDoc.Paragraphs.Last.Range;
            r.Text = strContent;
            r.Font.Size = 15;
            r.Start = r.Start + 5;
            r.End = r.End - 4;
            r.Font.Name = "Times New Roman";
            r.Font.Size = 14;

            wordApp.Selection.MoveDown(ref Nothing, ref Nothing, ref Nothing);//移动焦点
            //wordApp.Selection.ParagraphFormat.Reset();
            //wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            //wordApp.Selection.ParagraphFormat.Space1();
            strContent = "检测试剂盒：Ion PGMTM Hi-QTM OT2 Kit-200、\n";
            r = wordDoc.Paragraphs.Last.Range;
            r.Text = strContent;
            r.Font.Size = 15;
            r.Start = r.Start + 6;
            r.End = r.End - 1;
            r.Font.Name = "Times New Roman";
            r.Font.Size = 14;
            r.Start = r.Start + 7;// 选择第一个TM 设为上标
            r.End = r.Start + 2;
            r.Font.Superscript = 1;
            r.Start = r.Start + 7; // 选择第二个TM 设为上标
            r.End = r.Start + 2;
            r.Font.Superscript = 1;

            wordApp.Selection.MoveDown(ref Nothing, ref Nothing, ref Nothing);//移动焦点
            //wordApp.Selection.ParagraphFormat.Reset();
            //wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            wordApp.Selection.ParagraphFormat.CharacterUnitFirstLineIndent = 6.5f; //首行缩进2.98字符
            wordApp.Selection.ParagraphFormat.Space15();
            strContent = "Ion PGMTM Hi-QTM Sequencing Kit\n";
            r = wordDoc.Paragraphs.Last.Range;
            r.Font.Name = "Times New Roman";
            r.Text = strContent;
            r.Font.Size = 15;
            r.Font.Size = 14;
            r.Start = r.Start + 7;// 选择第一个TM 设为上标
            r.End = r.Start + 2;
            r.Font.Superscript = 1;
            r.Start = r.Start + 7; // 选择第二个TM 设为上标
            r.End = r.Start + 2;
            r.Font.Superscript = 1;
            #endregion

            #region 第一个表格 检测结果
            wordApp.Selection.MoveDown(ref Nothing, ref Nothing, ref Nothing);//移动焦点
            wordApp.Selection.ParagraphFormat.Reset();
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
            wordApp.Selection.ParagraphFormat.Space15();
            strContent = "检测结果：\n";
            r = wordDoc.Paragraphs.Last.Range;
            r.Font.Name = "宋体";
            r.Text = strContent;
            r.Font.Size = 11.5f;
            r.Font.Bold = 1;

            wordApp.Selection.MoveDown(ref Nothing, ref Nothing, ref Nothing);//移动焦点
            //wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;//水平居中
            Table t = wordDoc.Tables.Add(wordApp.Selection.Range, 16, 6, ref Nothing, ref Nothing);//16行6列的表
            t.Rows.Alignment = MSWord.WdRowAlignment.wdAlignRowCenter;
            t.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            t.Select();
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;//水平居中
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter; // 垂直居中
            t.Columns[2].Select();
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;//左对齐居中
            float [] ColWid = {1.97f, 3.56f, 1.84f, 3.28f, 4.13f, 2.54f};
            string[] title = { "基因名称", "外显子", "突变类型", "碱基变化", "氨基酸变化", "COSMIC ID"};
            for (int i = 1; i <= 6; i++)
            {
                t.Columns[i].Width = wordApp.CentimetersToPoints(ColWid[i - 1]);
                r = t.Cell(1, i).Range;
                r.Text = title[i - 1];
                r.Font.Size = 9;
                t.Cell(1, i).Select();
                wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;//水平居中
            }
            t.Rows[1].Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
            t.Rows[1].Borders[WdBorderType.wdBorderTop].LineWidth = WdLineWidth.wdLineWidth225pt;
            t.Rows[1].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
            t.Rows[1].Borders[WdBorderType.wdBorderBottom].LineWidth = WdLineWidth.wdLineWidth150pt;
            string[] GeneName = { "EGFR", "KRAS", "BRAF", "PIK3CA", "PDGFRA", "NRAS", "KIT", "DDR2", "EBRR2", "FLT3", "NPM1", "SMO", "DNMT3A", "ABL1", "TSC1"};
            string[] wai = {"18、19、20、21", "2、3", "11、15", "9、20", "12、14、18", "2、3", "9、11、13、14、17", "18", "20", "14、15、20", "11", "8", "15、16、17、18、19、20、21、22、23", "4、5、6", "15"};
            for (int i = 0; i < GeneName.Length; i++)
            {
                r = t.Cell(i + 2, 1).Range;
                r.Text = GeneName[i];
                r.Font.Size = 10.5f;
                r.Font.Name = "Times New Roman";
                r.Bold = 0;

                r = t.Cell(i + 2, 2).Range;
                r.Text = wai[i];
                r.Font.Size = 10.5f;
                r.Font.Name = "Times New Roman";
                r.Bold = 0;

                r = t.Cell(i + 2, 4).Range;
                r.Text = "未见突变";
                r.Font.Size = 10.5f;
                r.Bold = 0;
            }
            t.Rows[t.Rows.Count].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
            object movecount = t.Rows.Count;
            #endregion

            object bs = WdBreakType.wdSectionBreakNextPage;
            wordDoc.Paragraphs.Last.Range.InsertBreak(ref bs);
            Console.WriteLine(wordApp.Selection.MoveDown(ref Nothing, ref Nothing, ref Nothing));//移动焦点


            #region 第二个表格 ALK、ROS1、RET基因热点区域融合检测结果
            wordApp.Selection.ParagraphFormat.Reset();
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
            wordApp.Selection.ParagraphFormat.LineSpacing = 17;
            strContent = "ALK、ROS1、RET基因热点区域融合检测结果：\n";
            r = wordDoc.Paragraphs.Last.Range;
            r.Select();
            r.Font.Name = "宋体";
            r.Text = strContent;
            r.Font.Size = 10.5f;
            r.Font.Bold = 0;
            r.End = r.Start + 3;
            r.Font.Name = "Times New Roman";
            r.Start = r.End + 1;
            r.End = r.Start + 4;
            r.Font.Name = "Times New Roman";
            r.Start = r.End + 1;
            r.End = r.Start + 3;
            r.Font.Name = "Times New Roman";

            Console.WriteLine(wordApp.Selection.MoveDown(ref Nothing, ref Nothing, ref Nothing));//移动焦点

            
            wordApp.Selection.ParagraphFormat.Reset();
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;//水平居中
            Table t2 = wordDoc.Tables.Add(wordDoc.Paragraphs.Last.Range, 4, 5, ref Nothing, ref Nothing);//4行5列的表
            t2.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
            
            t2.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            t2.Select();
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;//水平居中
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter; // 垂直居中
            t2.Columns[2].Select();
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;//左对齐居中
            float[] ColPro = { 13.6f, 44.6f, 12, 16, 13.5f};
            string[] HotTitle = { "基因名称", "检测范围", "结果类型", "具体位点", "COSMIC ID"};
            for (int i = 1; i <= 5; i++)
            {
                //t2.Columns[i].SetWidth(ColPro[i - 1], WdRulerStyle.wdAdjustProportional);
                t2.Columns[i].Width = ColPro[i - 1];
                r = t2.Cell(1, i).Range;
                r.Text = HotTitle[i - 1];
                r.Font.Size = 9;
                r.Font.Bold = 1;
                t2.Cell(1, i).Select();
                wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;//水平居中
            }
            t2.Rows[1].Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
            t2.Rows[1].Borders[WdBorderType.wdBorderTop].LineWidth = WdLineWidth.wdLineWidth225pt;
            t2.Rows[1].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
            t2.Rows[1].Borders[WdBorderType.wdBorderBottom].LineWidth = WdLineWidth.wdLineWidth150pt;
            string[] GeneName2 = { "EML4-ALK", "ROS1", "KIF5B-RET" };
            string[] TestR = { "E13;A20、E20;A20、E6a;A20、E6b;A20、E14;ins11del49 A20、E2;A20、E2;ins117A20、E13;ins69A20、E14;del12A20", "SLC34A2 e4-ROS1 e32、SLC34A2 e4-ROS1 e34、SLC34A2 e13-ROS1 e32、SLC34A2 e13-ROS1 e34、CD74 e6-ROS1 e32、CD74 e6-ROS1 e34、SDC4 e2-ROS1 e32、SDC4 e2-ROS1 e34、SDC4 e4-ROS1 e32、SDC4 e4-ROS1 e34、EZR e10-ROS1 e34、TPM3 e8-ROS1 e35、LRIG3 e16-ROS1 e35、FIG e8-ROS1 e35、FIG e4-ROS1 e36", "K15;R12、K15;R11、K16;R12、K22;R12、K23;R12、K24;R8、K24;R11" };
            for (int i = 0; i < GeneName2.Length; i++)
            {
                r = t2.Cell(i + 2, 1).Range;
                r.Text = GeneName2[i];
                r.Font.Size = 10.5f;
                r.Font.Name = "Times New Roman";
                r.Bold = 0;

                r = t2.Cell(i + 2, 2).Range;
                r.Text = TestR[i];
                r.Font.Size = 10.5f;
                r.Font.Name = "Times New Roman";
                r.Bold = 0;

            }
            t2.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);
            t2.Rows[t2.Rows.Count].Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
            object t2rows = t2.Rows.Count;
            wordApp.Selection.MoveDown(ref unitline, ref t2rows, ref Nothing);//移动焦点
            r = t2.Range;
            r.Start = r.End;
            r.InsertBreak(ref Nothing);
            #endregion

            Console.WriteLine(wordApp.Selection.MoveDown(ref Nothing, ref Nothing, ref Nothing));//移动焦点

            #region 第三个表格
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            wordApp.Selection.ParagraphFormat.Space1();
            strContent = "阳性突变\n检测患者相关基因，对测序顺序进行深入的分析，确定突变基因、突变位点以及临床意义的详细信息如下：";
            r = wordApp.Selection.Range;
            r.Font.Name = "宋体";
            r.Text = strContent;
            r.Font.Size = 10.5f;
            r.Font.Bold = 1;
            r.Select();
            

            wordApp.Selection.MoveDown(ref Nothing, ref Nothing, ref Nothing);//移动焦点
            Table t3 = wordDoc.Tables.Add(wordDoc.Paragraphs.Last.Range, 3, 12, ref Nothing, ref Nothing);//3行12列的表
            t3.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            t3.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
            t3.Select();
            wordApp.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;//水平居中
            wordApp.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter; // 垂直居中
            string[] DetailTitle = { "基因", "外显子", "碱基突变", "氨基酸突变", "突变比例", "Cosmic ID", "肿瘤", "标签", "等级", "说明来源", "CFDA批准的药物", "FDA批准的药物"};
            for (int i = 1; i <= DetailTitle.Length; i++)
            {
                r = t3.Cell(1, i).Range;
                r.Text = DetailTitle[i - 1];
                r.Font.Size = 10.5f;
                r.Font.Bold = 1;
            }

            t3.Rows[1].Shading.BackgroundPatternColor = (Microsoft.Office.Interop.Word.WdColor)(0x1 * 124 + 0x100 * 187 + 0x10000 * 225);
            #endregion

            wordApp.Selection.MoveDown(ref Nothing, ref Nothing, ref Nothing);//移动焦点

            wordDoc.Sections.Last.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;

            #region 设置页眉
            wordApp.ActiveWindow.View.Type = WdViewType.wdOutlineView;//视图样式
            wordApp.ActiveWindow.View.SeekView = WdSeekView.wdSeekPrimaryHeader;//进入页眉设置，其中页眉边距在页面设置中已完成
            
            wordApp.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight;
            wordApp.Selection.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphRight;
            wordApp.Selection.Paragraphs.Last.Range.Text = "洛阳晶云信息科技有限公司2";
            
            //去掉页眉的横线
            wordApp.ActiveWindow.ActivePane.Selection.ParagraphFormat.Borders[MSWord.WdBorderType.wdBorderBottom].LineStyle = MSWord.WdLineStyle.wdLineStyleNone;
            wordApp.ActiveWindow.ActivePane.Selection.Borders[MSWord.WdBorderType.wdBorderBottom].Visible = false;
            wordApp.ActiveWindow.ActivePane.View.SeekView = MSWord.WdSeekView.wdSeekMainDocument;//退出页眉设置
            #endregion
            
            #region 设置页码
            MSWord.PageNumbers pns = wordApp.Selection.Sections[1].Headers[MSWord.WdHeaderFooterIndex.wdHeaderFooterEvenPages].PageNumbers;//获取当前页的号码
            pns.NumberStyle = WdPageNumberStyle.wdPageNumberStyleNumberInDash;
            pns.HeadingLevelForChapter = 0;
            pns.IncludeChapterNumber = false;
            pns.RestartNumberingAtSection = false;
            pns.StartingNumber = 0;
            object pagenmbetal = WdPageNumberAlignment.wdAlignPageNumberCenter;//将号码设置在中间
            object first = true;
            wordApp.Selection.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterEvenPages].PageNumbers.Add(ref pagenmbetal, ref first);
            #endregion

            #region 保存为DOCX和PDF
            object format = MSWord.WdSaveFormat.wdFormatDocumentDefault;
            object pdfformat = MSWord.WdSaveFormat.wdFormatPDF;
            //object htmlformat = WdSaveFormat.wdFormatHTML;
            object pdfpath = Environment.CurrentDirectory + "\\myPDF.pdf";
            //object htmlpath = Environment.CurrentDirectory + "\\myHTML.html"; 
            wordDoc.SaveAs(ref path, ref format, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing); 
            wordDoc.SaveAs(ref pdfpath, ref pdfformat, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
            //wordDoc.SaveAs(ref htmlpath, ref htmlformat, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
            wordDoc.Close(ref Nothing, ref Nothing, ref Nothing);
            wordApp.Quit(ref Nothing, ref Nothing, ref Nothing);
            #endregion
        }
        
    }
}
