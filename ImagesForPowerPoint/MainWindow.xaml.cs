using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Syncfusion.Presentation;
using System.IO;

namespace ImagesForPowerPoint
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        

        public MainWindow()
        {
            InitializeComponent();

          
        }



        private void SubmitButton_Click(object sender, RoutedEventArgs e)
        {
            string mylist = (" ");
            List<string> boldTexts = new List<string>();
            foreach (Paragraph p in rchTextbox.Document.Blocks)
            {
                foreach (var inline in p.Inlines)
                {
                    if (inline.FontWeight == FontWeights.Bold)
                    {
                        var textRange = new TextRange(inline.ContentStart, inline.ContentEnd);
                        boldTexts.Add(textRange.Text);
                        //MessageBox.Show(textRange.Text);
                        mylist += (" " + textRange.Text);
                        

                        


                    }
            myWeb.Source = new Uri("https://www.google.com/search?tbm=isch&q=" + titleWord.Text + " " + mylist);


                }

            }
            
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            //Create a new instance of PowerPoint Presentation file
            IPresentation pptxDoc = Presentation.Create();

            //Add a new slide to file and apply background color
            ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.TitleOnly);

            //Add title content to the slide by accessing the title placeholder of the TitleOnly layout-slide
            IShape titleShape = slide.Shapes[0] as IShape;
            titleShape.TextBody.AddParagraph(titleWord.Text).HorizontalAlignment = HorizontalAlignmentType.Center;

            //Add description content to the slide by adding a new TextBox
            IShape descriptionShape = slide.AddTextBox(53.22, 141.73, 874.19, 77.70);
            descriptionShape.TextBody.Text = bodyPPT.Text;
            //Gets a picture as stream.
            //Stream pictureStream = File.Open("C:/Users/dell/Downloads/download.png", FileMode.Open);

            //Adds the picture to a slide by specifying its size and position.
            //slide.Shapes.AddPicture(pictureStream, 499.79, 238.59, 364.54, 192.16);

            //Save the PowerPoint Presentation 
            //pptxDoc.Save("Sample.pptx");

            //Close the PowerPoint presentation
            //pptxDoc.Close();
        }
    }

        
}


       
    

    
