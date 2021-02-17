using Syncfusion.Presentation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media.Imaging;

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
                        mylist += (textRange.Text + " ");




                    }
                    myWeb.Source = new Uri("https://www.google.com/search?tbm=isch&q=" + titleWord.Text + mylist);
                    string urls = (" ");
                    //create web client
                    WebClient googleImages = new WebClient();
                    //This regex searches for image urls in the html from google
                    Regex googleRegex = new Regex(@"src=""https://[^""]*""", RegexOptions.Compiled | RegexOptions.IgnoreCase);
                    //get google html for image search
                    string html = googleImages.DownloadString("https://www.google.com/search?tbm=isch&q=" + titleWord.Text + mylist);
                    MatchCollection googleMatches = googleRegex.Matches(html);

                    //MessageBox.Show("" + googleMatches.Count);
                    foreach (Match m in googleMatches)
                    {
                        //add the match to the arraylist
                        urls += (m.Value + " ");
                    }
                    int iterator;
                    string[] imgInfo = urls.Split(' ');
                    
                    foreach (string info in imgInfo) //(int x = 0; x < urls.Count; x++)
                    {
                        iterator = 1;
                        //Console.WriteLine(x);
                        //Console.WriteLine(x.Substring(5, x.Length - 6));

                        //create the image and add it to the listbox
                        Image googleImage = new Image();
                        googleImage.Name = "image" + iterator;
                        //googleImage.Source = new BitmapImage(new Uri(info.Substring(0, info.Length)));
                        BitmapImage bi3 = new BitmapImage();
                        bi3.BeginInit();
                        bi3.UriSource = new Uri(info, UriKind.Relative);
                        bi3.EndInit();
                        googleImage.Source = bi3;
                        string saveDirectory = @"C:/ImagesForPowerPoint/ImagesForPowerPoint/Images/";
                        //googleImage.Name = Path.Combine(saveDirectory, googleImage.Name);


                        //increment iterator
                        if (iterator < 7)
                        {
                            iterator++;
                        }




                        
                        
                    }



                }

            }

        }



        private void Button_Click_1(object sender, RoutedEventArgs e)
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
            //Stream pictureStream = File.Open("C:/Users/dell/source/SEH/ref1.png", FileMode.Open);

            //Adds the picture to a slide by specifying its size and position.
            //slide.Shapes.AddPicture(pictureStream, 499.79, 238.59, 364.54, 192.16);

            //Save the PowerPoint Presentation 
            //pptxDoc.Save("Sample.pptx");

            //Close the PowerPoint presentation
            //pptxDoc.Close(); 
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
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
            //Stream pictureStream = File.Open("C:/Users/dell/source/SEH/ref2.png", FileMode.Open);

            //Adds the picture to a slide by specifying its size and position.
            //slide.Shapes.AddPicture(pictureStream, 499.79, 238.59, 364.54, 192.16);

            //Save the PowerPoint Presentation 
            //pptxDoc.Save("Sample.pptx");

            //Close the PowerPoint presentation
            //pptxDoc.Close(); 
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
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
            //Stream pictureStream = File.Open("C:/Users/dell/source/SEH/ref3.png", FileMode.Open);

            //Adds the picture to a slide by specifying its size and position.
            //slide.Shapes.AddPicture(pictureStream, 499.79, 238.59, 364.54, 192.16);

            //Save the PowerPoint Presentation 
            //pptxDoc.Save("Sample.pptx");

            //Close the PowerPoint presentation
            //pptxDoc.Close(); 
        }

        private void Button_Click_4(object sender, RoutedEventArgs e)
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
            //Stream pictureStream = File.Open("C:/Users/dell/source/SEH/ref4.png", FileMode.Open);

            //Adds the picture to a slide by specifying its size and position.
            //slide.Shapes.AddPicture(pictureStream, 499.79, 238.59, 364.54, 192.16);

            //Save the PowerPoint Presentation 
            //pptxDoc.Save("Sample.pptx");

            //Close the PowerPoint presentation
            //pptxDoc.Close(); 
        }

        private void Button_Click_5(object sender, RoutedEventArgs e)
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
            //Stream pictureStream = File.Open("C:/Users/dell/source/SEH/ref5.png", FileMode.Open);

            //Adds the picture to a slide by specifying its size and position.
            //slide.Shapes.AddPicture(pictureStream, 499.79, 238.59, 364.54, 192.16);

            //Save the PowerPoint Presentation 
            //pptxDoc.Save("Sample.pptx");

            //Close the PowerPoint presentation
            //pptxDoc.Close(); 
        }

        private void Button_Click_6(object sender, RoutedEventArgs e)
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
            //Stream pictureStream = File.Open("C:/Users/dell/source/SEH/ref6.png", FileMode.Open);

            //Adds the picture to a slide by specifying its size and position.
            //slide.Shapes.AddPicture(pictureStream, 499.79, 238.59, 364.54, 192.16);

            //Save the PowerPoint Presentation 
            //pptxDoc.Save("Sample.pptx");

            //Close the PowerPoint presentation
            //pptxDoc.Close(); 
        }
    }


}


