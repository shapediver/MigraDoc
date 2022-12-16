using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using System.Reflection;
using System.IO;
using Microsoft.SqlServer.Server;
using System.Text;

namespace MigraDoc.Tests
{
    [TestClass]
    public class TemplateTest
    {
        static string TestDirectory;

        static string TestText = "What a loooong test text.";
        static string TestTextLong;

        /// <summary>
        /// Test setup for complete class, will be called once for all tests contained herein
        /// Change signature to "async static Task" in case of async tests
        /// </summary>
        /// <param name="context"></param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
            TestDirectory = Path.GetDirectoryName(context.TestRunDirectory);

            var sb = new StringBuilder();
            for (int i = 0; i<10; i++)
            {
                sb.AppendLine(TestText);
            }
            TestTextLong = sb.ToString();
        }

        /// <summary>
        /// Test setup per test, will be called once for each test
        /// Change signature to "async Task" in case of async tests
        /// </summary>
        [TestInitialize]
        public void TestInitialize()
        {
            // TODO
        }

        /// <summary>
        /// Test method
        /// Change signature to "async Task" in case of async tests
        /// </summary>
        [TestMethod]
        public void Test_00_Page_10_by_10()
        {
            // page height and width
            int pageHeight = 10;
            int pageWidth = 10;

            // new document and page
            var pdfDoc = new PdfDocument();
            var pdfPage = pdfDoc.AddPage();

            pdfDoc.Options.CompressContentStreams = true;
            pdfDoc.Options.EnableCcittCompressionForBilevelImages = true;
            pdfDoc.Options.FlateEncodeMode = PdfFlateEncodeMode.BestCompression;
            pdfDoc.Options.NoCompression = false;

            // page width and height
            if (pageHeight < pageWidth)
                pdfPage.Orientation = PdfSharp.PageOrientation.Landscape;
            else
                pdfPage.Orientation = PdfSharp.PageOrientation.Portrait;

            pdfPage.Height = pageHeight;
            pdfPage.Width = pageWidth;

            // creator info
            var assembly = Assembly.GetExecutingAssembly();
            pdfDoc.Info.Creator = assembly.ToString();

            // draw stuff onto the pdf
            using (var g = XGraphics.FromPdfPage(pdfPage))
            {
                // remember the transformation state before drawing
                var state = g.Save();

                // rectangle height and width
                int rectHeight = 10;
                int rectWidth = 10;

                var doc = new MigraDoc.DocumentObjectModel.Document();
                var sec = doc.AddSection();
                sec.PageSetup.PageHeight = rectHeight;
                sec.PageSetup.PageWidth = rectWidth;
                sec.PageSetup.LeftMargin = 0;
                sec.PageSetup.RightMargin = 0;
                sec.PageSetup.TopMargin = 0;
                sec.PageSetup.BottomMargin = 0;
                sec.PageSetup.HeaderDistance = 0;
                sec.PageSetup.FooterDistance = 0;

                var tab = sec.AddTable();
                tab.Borders.Color = MigraDoc.DocumentObjectModel.Colors.Transparent;
                tab.Borders.Width = 0;

                tab.LeftPadding = 0;
                tab.RightPadding = 0;
                tab.TopPadding = 0;
                tab.BottomPadding = 0;

                tab.Format.Alignment = MigraDoc.DocumentObjectModel.ParagraphAlignment.Left;

                tab.Rows.LeftIndent = 0;
           
                var col = tab.AddColumn(rectWidth);
                var row = tab.AddRow();
                row.Height = rectHeight;

                var cell = row.Cells[0];
                cell.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Top;

                var para = cell.AddParagraph();

                para.Format.Font.Color = new MigraDoc.DocumentObjectModel.Color(0, 0, 0);
                para.Format.Font.Name = "Times New";
                para.Format.Font.Size = 1;

                var tFormat = new MigraDoc.DocumentObjectModel.TextFormat();

                para.AddFormattedText(TestTextLong, tFormat);

                var docRenderer = new MigraDoc.Rendering.DocumentRenderer(doc);
                docRenderer.PrepareDocument();
                docRenderer.RenderPage(g, 1, MigraDoc.Rendering.PageRenderOptions.RenderContent);

                // restore state
                g.Restore(state);
            }

            // save the pdf
            pdfDoc.Save(Path.Combine(TestDirectory, "test_00.pdf"));

        }

        [TestMethod]
        public void Test_01_Page_10_by_10()
        {
            // page height and width
            int pageHeight = 10;
            int pageWidth = 10;

            // new document and page
            var pdfDoc = new PdfDocument();
            var pdfPage = pdfDoc.AddPage();

            pdfDoc.Options.CompressContentStreams = true;
            pdfDoc.Options.EnableCcittCompressionForBilevelImages = true;
            pdfDoc.Options.FlateEncodeMode = PdfFlateEncodeMode.BestCompression;
            pdfDoc.Options.NoCompression = false;

            // page width and height
            if (pageHeight < pageWidth)
                pdfPage.Orientation = PdfSharp.PageOrientation.Landscape;
            else
                pdfPage.Orientation = PdfSharp.PageOrientation.Portrait;

            pdfPage.Height = pageHeight;
            pdfPage.Width = pageWidth;

            // creator info
            var assembly = Assembly.GetExecutingAssembly();
            pdfDoc.Info.Creator = assembly.ToString();

            // draw stuff onto the pdf
            using (var g = XGraphics.FromPdfPage(pdfPage))
            {
                // remember the transformation state before drawing
                var state = g.Save();

                // rectangle height and width
                int rectHeight = 10;
                int rectWidth = 10;

                var doc = new MigraDoc.DocumentObjectModel.Document();
                var sec = doc.AddSection();
                sec.PageSetup.PageHeight = rectHeight;
                sec.PageSetup.PageWidth = rectWidth;
                sec.PageSetup.LeftMargin = 0;
                sec.PageSetup.RightMargin = 0;
                sec.PageSetup.TopMargin = 0;
                sec.PageSetup.BottomMargin = 0;
                sec.PageSetup.HeaderDistance = 0;
                sec.PageSetup.FooterDistance = 0;

                var tab = sec.AddTable();

                tab.Borders.Color = MigraDoc.DocumentObjectModel.Colors.Transparent;
                tab.Borders.Width = 0;

                tab.LeftPadding = 0;
                tab.RightPadding = 0;
                tab.TopPadding = 0;
                tab.BottomPadding = 0;

                tab.Format.Alignment = MigraDoc.DocumentObjectModel.ParagraphAlignment.Left;
                tab.Format.Font.Color = new MigraDoc.DocumentObjectModel.Color(0, 0, 0);
                tab.Format.Font.Name = "Times New";
                tab.Format.Font.Size = 1;

                tab.Rows.Height = rectHeight;
                tab.Rows.HeightRule = DocumentObjectModel.Tables.RowHeightRule.Exactly;
                tab.Rows.LeftIndent = 0;
                tab.Rows.VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Top; ;

                var col = tab.AddColumn(rectWidth);
                var row = tab.AddRow();
                var cell = row.Cells[0];

                var para = cell.AddParagraph();
                var tFormat = new MigraDoc.DocumentObjectModel.TextFormat();
                para.AddFormattedText(TestTextLong, tFormat);

                var docRenderer = new MigraDoc.Rendering.DocumentRenderer(doc);
                docRenderer.PrepareDocument();
                docRenderer.RenderPage(g, 1, MigraDoc.Rendering.PageRenderOptions.RenderContent);

                // restore state
                g.Restore(state);
            }

            // save the pdf
            pdfDoc.Save(Path.Combine(TestDirectory, "test_01.pdf"));

        }

        [TestMethod]
        public void Test_02_Page_210_by_297()
        {
            // page height and width
            int pageHeight = 297;
            int pageWidth = 210;

            // new document and page
            var pdfDoc = new PdfDocument();
            var pdfPage = pdfDoc.AddPage();

            pdfDoc.Options.CompressContentStreams = true;
            pdfDoc.Options.EnableCcittCompressionForBilevelImages = true;
            pdfDoc.Options.FlateEncodeMode = PdfFlateEncodeMode.BestCompression;
            pdfDoc.Options.NoCompression = false;

            // page width and height
            if (pageHeight < pageWidth)
                pdfPage.Orientation = PdfSharp.PageOrientation.Landscape;
            else
                pdfPage.Orientation = PdfSharp.PageOrientation.Portrait;

            pdfPage.Height = pageHeight;
            pdfPage.Width = pageWidth;

            // creator info
            var assembly = Assembly.GetExecutingAssembly();
            pdfDoc.Info.Creator = assembly.ToString();

            // draw stuff onto the pdf
            using (var g = XGraphics.FromPdfPage(pdfPage))
            {
                // remember the transformation state before drawing
                var state = g.Save();

                // rectangle height and width
                int rectHeight = 297;
                int rectWidth = 210;

                var doc = new MigraDoc.DocumentObjectModel.Document();
                var sec = doc.AddSection();
                sec.PageSetup.PageHeight = rectHeight;
                sec.PageSetup.PageWidth = rectWidth;
                sec.PageSetup.LeftMargin = 0;
                sec.PageSetup.RightMargin = 0;
                sec.PageSetup.TopMargin = 0;
                sec.PageSetup.BottomMargin = 0;
                sec.PageSetup.HeaderDistance = 0;
                sec.PageSetup.FooterDistance = 0;

                var tab = sec.AddTable();
                tab.Borders.Color = MigraDoc.DocumentObjectModel.Colors.Transparent;
                tab.Borders.Width = 0;
                tab.Rows.LeftIndent = 0;

                var col = tab.AddColumn(rectWidth);
                var row = tab.AddRow();
                row.Height = rectHeight;
                var para = row.Cells[0].AddParagraph();

                row.Format.Alignment = MigraDoc.DocumentObjectModel.ParagraphAlignment.Left;
                para.Format.Alignment = MigraDoc.DocumentObjectModel.ParagraphAlignment.Left;

                row.Cells[0].VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment.Top;

                para.Format.Font.Name = "Times New";
                para.Format.Font.Size = 10;

                para.Format.Font.Color = new MigraDoc.DocumentObjectModel.Color(0, 0, 0);

                var tFormat = new MigraDoc.DocumentObjectModel.TextFormat();

                para.AddFormattedText("What a loooong test text", tFormat);

                var docRenderer = new MigraDoc.Rendering.DocumentRenderer(doc);
                docRenderer.PrepareDocument();
                docRenderer.RenderPage(g, 1, MigraDoc.Rendering.PageRenderOptions.RenderContent);

                // restore state
                g.Restore(state);
            }

            // save the pdf
            pdfDoc.Save(Path.Combine(TestDirectory, "test_02.pdf"));

        }

        /// <summary>
        /// Test cleanup per test, will be called once for each test
        /// Change signature to "async Task" in case of async tests
        /// </summary>
        [TestCleanup]
        public void TestCleanup()
        {
            // TODO 
        }

        /// <summary>
        /// Test cleanup for complete class, will be called once for all tests contained herein
        /// Change signature to "async static Task" in case of async tests
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            // TODO 
        }
    }
}