using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Xml.Linq;
using DocxToXmlWpf.Models;

namespace DocxToXmlWpf
{
    public partial class MainWindow : Window
    {
        private string selectedFilePath;

        public ObservableCollection<MappingRule> MappingRules { get; set; } = new();

        public MainWindow()
        {
            InitializeComponent();
            MappingGrid.ItemsSource = MappingRules;

            AllowDrop = true;
            DragEnter += MainWindow_DragEnter;
            Drop += MainWindow_Drop;
        }

        private void MainWindow_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effects = DragDropEffects.Copy;
        }

        private void MainWindow_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0 && Path.GetExtension(files[0]) == ".docx")
                {
                    selectedFilePath = files[0];
                    MessageBox.Show($"File selected: {Path.GetFileName(selectedFilePath)}");
                }
                else
                {
                    MessageBox.Show("Please drop a valid .docx file.");
                }
            }
        }

        private void ConvertButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(selectedFilePath))
            {
                MessageBox.Show("Please drag and drop a .docx file first.");
                return;
            }

            try
            {
                string outputPath = Path.Combine(Path.GetDirectoryName(selectedFilePath),
                    Path.GetFileNameWithoutExtension(selectedFilePath) + "_converted.xml");

                ConvertDocxToXml(selectedFilePath, outputPath);
                MessageBox.Show($"Conversion complete! XML saved to:\n{outputPath}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during conversion:\n{ex.Message}");
            }
        }

        private void ConvertDocxToXml(string docxPath, string outputPath)
        {
            using var wordDoc = WordprocessingDocument.Open(docxPath, false);
            var paragraphs = wordDoc.MainDocumentPart.Document.Body.Elements<Paragraph>();

            var xmlDoc = new XDocument(new XElement("Document"));
            XElement currentChapter = null;
            XElement currentArticle = null;

            foreach (var paragraph in paragraphs)
            {
                string text = paragraph.InnerText.Trim();
                if (string.IsNullOrEmpty(text)) continue;

                bool matched = false;

                foreach (var rule in MappingRules)
                {
                    if (!string.IsNullOrWhiteSpace(rule.Prefix) &&
                        !string.IsNullOrWhiteSpace(rule.Tag) &&
                        text.StartsWith(rule.Prefix, StringComparison.OrdinalIgnoreCase))
                    {
                        var customElement = new XElement(rule.Tag, new XAttribute("title", text));

                        if (rule.Tag.Equals("Chapter", StringComparison.OrdinalIgnoreCase))
                        {
                            currentChapter = customElement;
                            xmlDoc.Root.Add(currentChapter);
                            currentArticle = null;
                        }
                        else if (rule.Tag.Equals("Article", StringComparison.OrdinalIgnoreCase))
                        {
                            currentArticle = customElement;
                            if (currentChapter != null)
                                currentChapter.Add(currentArticle);
                            else
                                xmlDoc.Root.Add(currentArticle);
                        }
                        else
                        {
                            if (currentArticle != null)
                                currentArticle.Add(customElement);
                            else if (currentChapter != null)
                                currentChapter.Add(customElement);
                            else
                                xmlDoc.Root.Add(customElement);
                        }

                        matched = true;
                        break;
                    }
                }

                if (!matched)
                {
                    var p = new XElement("Paragraph", text);
                    if (currentArticle != null)
                        currentArticle.Add(p);
                    else if (currentChapter != null)
                        currentChapter.Add(p);
                    else
                        xmlDoc.Root.Add(p);
                }
            }

            xmlDoc.Save(outputPath);
        }
    }
}
