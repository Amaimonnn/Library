using System;
using System.IO;
using PdfSharp;
using MigraDoc;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.Rendering;
using PdfSharp.Pdf;
using Color = MigraDoc.DocumentObjectModel.Color;

namespace ConsoleApplication1
{
    internal class Program
    {
        public static void Main(string[] args)
        {
            Document document = new Document();
            Section section = document.AddSection();
            section.PageSetup.PageFormat = PageFormat.A4;// формат листа а4
            section.PageSetup.Orientation = Orientation.Portrait;
            section.PageSetup.BottomMargin = 10;//нижний отступ
            section.PageSetup.TopMargin = 10;//верхний отступ
            
            Paragraph paragraph = new Paragraph();
            section.Add(paragraph);
            paragraph.Format.Font.Name = "Arial";
            paragraph.Format.Font.Size = "22";
            paragraph.Format.Alignment = ParagraphAlignment.Center;
            paragraph.Format.Font.Color = Color.FromRgb(100, 100, 100);
            paragraph.Format.SpaceAfter = "3cm";
            paragraph.Format.SpaceBefore = "3cm";
            Text text = new Text("");
            paragraph.AddText("Мои мечты стремятся вдаль,");
            paragraph.AddFormattedText(TextFormat.Bold);
            paragraph.AddLineBreak();
            paragraph.AddText("Где слышны вопли и рыданья,");
            paragraph.AddFormattedText(TextFormat.Italic);// форматированный текст
            paragraph.AddLineBreak();
            paragraph.AddText("Чужую разделить печаль");
            paragraph.AddLineBreak();
            paragraph.AddText("И муки тяжкого страданья");
            paragraph.AddLineBreak();
            paragraph.AddLineBreak();
            paragraph.AddText("Я там могу найти себе");
            paragraph.AddLineBreak();
            paragraph.AddText("Отраду в жизни, упоенье,");
            paragraph.AddLineBreak();
            paragraph.AddText("И там, наперекор судьбе,");
            paragraph.AddLineBreak();
            paragraph.AddText("Искать я буду вдохновенья.");
            Section imageSection = document.AddSection();
            Image image = new Image("C:/Users/Валерия/Downloads/Есенин.jpg");
            image.Width = 100;
            image.Height = 100;
            section.Add(image);
            var pdfRenderer = new PdfDocumentRenderer(true);
            pdfRenderer.Document = document;
            pdfRenderer.RenderDocument();
            var filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "123.pdf");
            pdfRenderer.PdfDocument.Save(filePath);// сохраняем 
        }
        
    }
}