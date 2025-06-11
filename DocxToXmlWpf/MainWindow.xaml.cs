using DocxToXmlWpf.Models;
using Microsoft.Win32;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Xml.Linq;
using Xceed.Words.NET;

namespace DocxToXmlWpf
{
    public partial class MainWindow : Window
    {
        public ObservableCollection<MappingRule> Mappings { get; set; } = new();
        private string droppedFilePath;
        private bool fileAlreadyDropped = false;

        public MainWindow()
        {
            InitializeComponent();
            MappingGrid.ItemsSource = Mappings;
        }

        private void MainWindow_DragEnter(object sender, DragEventArgs e)
        {
            if (!fileAlreadyDropped && e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }

        private void MainWindow_Drop(object sender, DragEventArgs e)
        {
            if (fileAlreadyDropped)
                return;

            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0 && Path.GetExtension(files[0]) == ".docx")
                {
                    droppedFilePath = files[0];
                    DropStatusText.Text = $"File dropped: {Path.GetFileName(droppedFilePath)}";
                    fileAlreadyDropped = true;
                }
                else
                {
                    MessageBox.Show("Please drop a .docx file.");
                }
            }
        }

        private void ConvertButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(droppedFilePath) || !File.Exists(droppedFilePath))
            {
                MessageBox.Show("Please drag and drop a .docx file before converting.", "Missing File", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var xml = ConvertDocxToXml(droppedFilePath);
            ShowXmlPreview(xml);
        }

        private XDocument ConvertDocxToXml(string path)
        {
            var doc = DocX.Load(path);
            var root = new XElement("Document");

            foreach (var paragraph in doc.Paragraphs)
            {
                string text = paragraph.Text.Trim();

                if (string.IsNullOrWhiteSpace(text))
                    continue;

                XElement element = null;

                foreach (var mapping in Mappings)
                {
                    if (!string.IsNullOrEmpty(mapping.Prefix) && text.StartsWith(mapping.Prefix))
                    {
                        element = new XElement(mapping.Tag, text);
                        break;
                    }
                }

                if (element == null)
                {
                    element = new XElement("Paragraph", text);
                }

                root.Add(element);
            }

            return new XDocument(root);
        }

        private void ShowXmlPreview(XDocument xmlDoc)
        {
            var previewWindow = new Window
            {
                Title = "XML Preview",
                Width = 600,
                Height = 500,
                Owner = this,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Background = Brushes.White
            };

            var textBox = new TextBox
            {
                Text = xmlDoc.ToString(),
                FontFamily = new FontFamily("Consolas"),
                FontSize = 14,
                IsReadOnly = true,
                TextWrapping = TextWrapping.Wrap,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                Margin = new Thickness(10)
            };

            var saveButton = new Button
            {
                Content = "Save XML...",
                Width = 120,
                Height = 35,
                Margin = new Thickness(10),
                HorizontalAlignment = HorizontalAlignment.Right
            };
            saveButton.Click += (s, e) =>
            {
                SaveFileDialog dialog = new()
                {
                    Filter = "XML files (*.xml)|*.xml",
                    Title = "Save XML File",
                    FileName = "converted.xml"
                };
                if (dialog.ShowDialog() == true)
                {
                    xmlDoc.Save(dialog.FileName);
                    MessageBox.Show("File saved successfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                    previewWindow.Close();
                }
            };

            var panel = new DockPanel();
            DockPanel.SetDock(saveButton, Dock.Bottom);
            panel.Children.Add(saveButton);
            panel.Children.Add(textBox);

            previewWindow.Content = panel;
            previewWindow.ShowDialog();
        }
    }
}
