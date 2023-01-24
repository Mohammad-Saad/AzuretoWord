using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using AngleSharp;
using AngleSharp.Dom;
using Microsoft.Office.Interop.Word;

namespace HtmlToWord
{
    public partial class AzureConverter : Form
    {
        public AzureConverter()
        {
            InitializeComponent();
        }

        private async void button1_Click(object sender, EventArgs e)
        {


            string htmlpath = txthtmlPath.Text;
            string screenshotpath = txtScreenShots.Text;


            var html = File.ReadAllText(htmlpath);
            var config = Configuration.Default;
            var context = BrowsingContext.New(config);
            var doc = await context.OpenAsync(req => req.Content(html));
            List<TestResultModal> testResultModals = new List<TestResultModal>();

            string outcome = doc.QuerySelector("span[class = 'iteration-details-title-style']").TextContent;
            string pagetitle = doc.QuerySelector("td[class = 'query-title']").TextContent.Split('/')[1];
            //.All.Where(m => m.LocalName == "td" && m.ClassName == "query-title").ToString();


            IEnumerable<IElement> blueListItemsLinq = doc.All.Where(m => m.LocalName == "div" && m.ClassList.Contains("test-step-container"));

            int count = 0;
            foreach (var item in blueListItemsLinq)
            {

                TestResultModal model = new TestResultModal();

                model.ActualResult = item.QuerySelector("td[class = 'test-step-inline-title test-step-title']").TextContent;

                if (item.QuerySelector("tr[class = 'test-step-expected-result-row']") == null)
                {
                    model.expectedResult = String.Empty;
                }
                else
                {
                    string temp = item.QuerySelector("tr[class = 'test-step-expected-result-row']").TextContent;
                    model.expectedResult = temp.Substring(15);
                }



                var screenshots = item.QuerySelectorAll("img[class = 'result-step-thumbnail']");

                if (screenshots != null)
                {

                    foreach (IElement element in screenshots)
                    {
                        string str = element.OuterHtml;
                        string[] spearator = { ".png", ".PNG", "Runs - Test Plans_files/" };

                        string[] strlist = str.Split(spearator, StringSplitOptions.RemoveEmptyEntries);
                        model.ScreenShot.Add(strlist[1].ToString());
                    }
                }
                testResultModals.Add(model);

                count++;
            }

            CreateDocument(testResultModals, screenshotpath, pagetitle, outcome);


            txthtmlPath.Clear();
            txtScreenShots.Clear();

            MessageBox.Show("Conversion Complete");
            //CreateScreenshots(testResultModals, Microsoft.Office.Interop.Word.Table screenshots);

        }

        void CreateDocument(List<TestResultModal> testResultModals, string screenshotfolder, string pageTitle, string Outcome)
        {
            try
            {
                string directory = Directory.GetCurrentDirectory();
                string logoPath = directory + @"\Pass.png";
                string titleLogo = directory + @"\logo.png";
                string failedLogo = directory + @"\failedLogo.png";

                object oMissing = System.Reflection.Missing.Value;
                object oEndOfDoc = "\\endofdoc";
                Microsoft.Office.Interop.Word._Application objWord;
                Microsoft.Office.Interop.Word._Document objDoc;
                objWord = new Microsoft.Office.Interop.Word.Application();
                objWord.Visible = true;
                objDoc = objWord.Documents.Add(ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing);

                objDoc.PageSetup.Orientation = WdOrientation.wdOrientLandscape;
                Color sysColor = Color.FromArgb(0, 39, 205, 79);
                var color = (Microsoft.Office.Interop.Word.WdColor)(sysColor.R + 0x100 * sysColor.G + 0x10000 * sysColor.B);

                Microsoft.Office.Interop.Word.Paragraph Picture;
                Picture = objDoc.Content.Paragraphs.Add(ref oMissing);
                Picture.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                if (Outcome.Contains("failed"))
                {
                    Picture.Range.InlineShapes.AddPicture(failedLogo);
                }
                else
                {
                    Picture.Range.InlineShapes.AddPicture(titleLogo);
                }
                //Picture.Range.InlineShapes.AddPicture(logoPath).Width = 25;

                Microsoft.Office.Interop.Word.Paragraph title;
                title = objDoc.Content.Paragraphs.Add(ref oMissing);
                title.Range.Text = pageTitle;
                title.Range.Font.Bold = 1;
                title.Range.Font.Size = 16;
                if (Outcome.Contains("failed"))
                {
                    title.Range.Font.Color = WdColor.wdColorRed;
                }
                else
                {
                    title.Range.Font.Color = color;
                }

                title.Range.Font.Name = "Times New Roman";
                objDoc.Words.Last.InsertBreak(Microsoft.Office.Interop.Word.WdBreakType.wdLineBreak);
                Microsoft.Office.Interop.Word.Table screenshottableTable;
                Microsoft.Office.Interop.Word.Range wrdRng2 = objDoc.Bookmarks.get_Item(ref oEndOfDoc).Range; // objDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                screenshottableTable = objDoc.Tables.Add(wrdRng2, 5, 1, ref oMissing, ref oMissing);
                screenshottableTable.set_Style("Grid Table 2 - Accent 3");
                screenshottableTable.Rows.SetHeight(40, WdRowHeightRule.wdRowHeightAtLeast);
                screenshottableTable.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                screenshottableTable.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                screenshottableTable.Borders.InsideColor = WdColor.wdColorGray20;
                screenshottableTable.Borders.OutsideColor = WdColor.wdColorGray20;
                screenshottableTable.Range.Font.Color = WdColor.wdColorBlack;

                screenshottableTable.Range.Font.Size = 11;
                screenshottableTable.Range.Font.Bold = 1;
                screenshottableTable.Cell(1, 1).Range.Text = "Test Case Detail: ";
                screenshottableTable.Cell(2, 1).Range.Text = "Compliance:";
                screenshottableTable.Cell(3, 1).Range.Text = "Build Number: ";
                screenshottableTable.Cell(4, 1).Range.Text = "Execution Start Time: ";
                screenshottableTable.Cell(4, 1).Split(1, 3);
                screenshottableTable.Cell(4, 2).Range.Text = "Execution End Time: ";
                screenshottableTable.Cell(4, 3).Range.Text = "Duration: ";
                screenshottableTable.Cell(5, 1).Range.Text = "Browser: ";


                objDoc.Words.Last.InsertBreak(Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak);

                Microsoft.Office.Interop.Word.Table objTable;
                Microsoft.Office.Interop.Word.Range wrdRng = objDoc.Bookmarks.get_Item(ref oEndOfDoc).Range; // objDoc.Range(objDoc.Content.Start, ref oMissing);
                objTable = objDoc.Tables.Add(wrdRng, testResultModals.Count, 3, ref oMissing, ref oMissing);
                objTable.Range.ParagraphFormat.SpaceAfter = 7;
                objTable.set_Style("Grid Table 2 - Accent 3");
                objTable.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                objTable.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                objTable.Borders.InsideColor = WdColor.wdColorGray20;
                objTable.Borders.OutsideColor = WdColor.wdColorGray20;
                objTable.Range.Font.Size = 11;
                objTable.Range.Font.Color = WdColor.wdColorBlack;


                string strText = string.Empty;
                string screenshotdate = string.Empty;
                string[] TestDataSeperator = { "Test Data:" };
                string[] ComplianceSeperator = { "Compliance:", "COMPLIANCE:" };


                // tblHeader.Cell(1, 2).Range.InlineShapes.AddPicture(logoPath, Type.Missing, Type.Missing, Type.Missing);

                for (int Row = 1; Row <= testResultModals.Count; Row++)
                {
                    for (int Cell = 1; Cell <= 3; Cell++)
                    {

                        if (Cell == 1)
                        {
                            objTable.Columns[Cell].Width = 50f;

                            objTable.Cell(Row, Cell).Range.Text = "\n" + "   " + Row.ToString();
                            objTable.Cell(Row, Cell).Range.InlineShapes.AddPicture(logoPath);

                        }

                        else if (Cell == 2)
                        {

                            string Actual = testResultModals[Row - 1].ActualResult;

                            string[] spiltstring = Actual.Split(TestDataSeperator, StringSplitOptions.RemoveEmptyEntries);


                            // Actual Steps
                            strText = "Steps: ";
                            foreach (var item in spiltstring)
                            {
                                strText += item + "\n";
                            }


                            // Expected Steps:
                            string expected = testResultModals[Row - 1].expectedResult;

                            if (!string.IsNullOrEmpty(expected))
                            {
                                string[] expectedSplit = expected.Split(ComplianceSeperator, StringSplitOptions.RemoveEmptyEntries);

                                strText += "Expected: " + expectedSplit[0] + "\n";

                                if (expectedSplit.Length > 1)
                                {
                                    strText += "\n" + "Compliance: " + expectedSplit[1] + "\n";
                                }
                            }

                            if (testResultModals[Row - 1].ScreenShot.Count > 0)
                            {
                                strText += "Date Time: " + testResultModals[Row - 1].ScreenShot.Last().Replace("Screenshot-", "");
                            }

                            objTable.Columns[Cell].Width = 200f;
                            objTable.Cell(Row, Cell).Range.Text = strText;
                            strText = String.Empty;
                            objTable.Cell(Row, Cell).Range.Bold = 0;

                        }


                        else if (Cell == 3)
                        {
                            foreach (var item in testResultModals[Row - 1].ScreenShot)
                            {
                                screenshotdate += System.Environment.NewLine + item;
                                string screenshotpath = screenshotfolder + "\\" + item + ".png";
                                objTable.Columns[Cell].Width = 400f;
                                objTable.Cell(Row, Cell).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                                objTable.Cell(Row, Cell).Range.Bold = 1;
                                objTable.Cell(Row, Cell).Range.InlineShapes.AddPicture(screenshotpath);
                                objTable.Cell(Row, Cell).Range.InsertBefore(item);
                            }

                            screenshotdate = String.Empty;
                        }
                    }
                }
                //screenshottableTable.Range.InlineShapes.AddPicture
                string temp = string.Empty;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }




    public class TestResultModal
    {
        public string ActualResult { get; set; }
        public string expectedResult { get; set; }
        public List<String> ScreenShot { get; set; }
        public TestResultModal()
        {
            ScreenShot = new List<String>();
        }
    }
}
