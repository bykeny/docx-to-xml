using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Documents;
using System.Xml.Linq;

namespace DocxToXmlWpf
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effects = DragDropEffects.Copy;
        }

        private void Window_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (var file in files)
                {
                    if (Path.GetExtension(file).ToLower() == ".docx")
                    {
                        ConvertDocxToXml(file);
                    }
                    else
                    {
                        MessageBox.Show("Only .docx files are supported.", "Unsupported File", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
            }
        }

        private void ConvertDocxToXml(string docxPath)
        {
            string xmlPath = Path.ChangeExtension(docxPath, ".xml");

            try
            {
                XDocument xmlDoc = new XDocument(new XElement("Constitution"));

                XElement currentChapter = null;
                XElement currentArticle = null;
                int paragraphId = 1;
                int articleId = 1;

                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(docxPath, false))
                {
                    Body? body = wordDoc.MainDocumentPart?.Document.Body; // Fix CS8602 by using null-conditional operator

                    if (body == null)
                    {
                        MessageBox.Show("The document body is null or invalid.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    foreach (var para in body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>()) // Fix CS0104 by fully qualifying Paragraph
                    {
                        string text = para.InnerText.Trim();
                        if (string.IsNullOrWhiteSpace(text)) continue;

                        if (Regex.IsMatch(text, @"^\s*(Chapter|Fəsil)\s+\w+", RegexOptions.IgnoreCase))
                        {
                            currentChapter = new XElement("Chapter",
                                new XAttribute("title", text));
                            xmlDoc.Root.Add(currentChapter);
                            currentArticle = null;
                        }
                        else if (Regex.IsMatch(text, @"^\s*(Maddə)\s+\d+|^M\d+", RegexOptions.IgnoreCase))
                        {
                            currentArticle = new XElement("Madde",
                                new XAttribute("id", articleId++),
                                new XCData(text));
                            if (currentChapter != null)
                                currentChapter.Add(currentArticle);
                            else
                                xmlDoc.Root.Add(currentArticle);
                        }
                        else
                        {
                            var paraElem = new XElement("Paragraph",
                                new XAttribute("id", paragraphId++),
                                new XCData(text));
                            if (currentArticle != null)
                                currentArticle.Add(paraElem);
                            else if (currentChapter != null)
                                currentChapter.Add(paraElem);
                            else
                                xmlDoc.Root.Add(paraElem);
                        }
                    }
                }

                xmlDoc.Save(xmlPath);
                MessageBox.Show($"Structured XML saved at:\n{xmlPath}", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error:\n{ex.Message}", "Conversion Failed", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
