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
using System.Linq;
using System.Collections;


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
            List<string> uriResult = new List<string>();
            int iterator;
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
                    Regex googleRegex = new Regex(@"""https://[^""]*""", RegexOptions.Compiled | RegexOptions.IgnoreCase);

                    //get google html for image search
                    string html = googleImages.DownloadString("https://www.google.com/search?tbm=isch&q=" + titleWord.Text + mylist);
                    MatchCollection googleMatches = googleRegex.Matches(html);

                    //MessageBox.Show("" + googleMatches.Count);
                    foreach (Match m in googleMatches)
                    {
                        //add the match to the arraylist
                        urls += (m.Value + " ");
                    }

                    string[] imgInfo = urls.Split(' ');

                    foreach (string info in imgInfo.Where((info => !string.IsNullOrEmpty(info)))) //(int x = 0; x < urls.Count; x++)
                    {
                        //iterator = 1;
                        //Console.WriteLine(info);
                        //Console.WriteLine(info.Substring(0, info.Length));

                        //create the image and add it to the listbox
                        foreach (string s in info.Split(' '))
                        {
                            if (!String.IsNullOrEmpty(s))
                                uriResult.Add(s);
                        }
                    }
                    iterator = 1;
                    string bmp = uriResult.ElementAt(iterator);
                    Image googleImage = new Image();
                    googleImage.Name = "image";
                    ////this.RegisterName(googleImage.Name, googleImage);
                    googleImage.Source = new BitmapImage(new Uri(bmp.Substring(1, bmp.Length-5)));

                    //bi3 = new System.Drawing.Image()
                    //bi3.BeginInit();
                    //bi3.Source = new BitmapImage(new Uri(bmp.Substring(1, bmp.Length - 5)));
                    //bi3.EndInit();
                    //googleImage = BitmapImage (bi3);
                    //string saveDirectory = @"C:/ImagesForPowerPoint/ImagesForPowerPoint/Images/";
                    //googleImage = Path.Combine(saveDirectory, googleImage.Name + iterator + ".png");
                    //bi3.Save("C:/ImagesForPowerPoint/ImagesForPowerPoint/Images/" + googleImage.Name + iterator, System.Drawing.ImageFormat.png);
                    if (iterator < 7)
                    {
                        iterator++;
                    };


                    string bnmp1 = uriResult.ElementAt(1);
            button1.Content = new Image

            {
                Source = new BitmapImage(new Uri(bnmp1.Substring(1, bnmp1.Length - 5)))

                    //Source = new BitmapImage(new Uri(info, UriKind.Relative));


            };

                    //iterator++;
                    string bnmp2 = uriResult.ElementAt(2);        
            button2.Content = new Image
            {
                Source = new BitmapImage(new Uri(bnmp2.Substring(1,bnmp2.Length-5)))
            };

                    //iterator++;
                    string bnmp3 = uriResult.ElementAt(3);
            button3.Content = new Image
            {
                Source = new BitmapImage(new Uri(bnmp3.Substring(1, bnmp3.Length - 5)))
            };

                    //iterator++;
                    string bnmp4 = uriResult.ElementAt(4);
            button4.Content = new Image
            {
                Source = new BitmapImage(new Uri(bnmp4.Substring(1, bnmp4.Length - 5)))
            };        //iterator++;


                    string bnmp5 = uriResult.ElementAt(5);
            button5.Content = new Image
            {
                Source = new BitmapImage(new Uri(bnmp5.Substring(1, bnmp5.Length - 5)))
            };

                    //iterator++;
                    string bnmp6 = uriResult.ElementAt(6);
            button6.Content = new Image
            {
                Source = new BitmapImage(new Uri(bnmp6.Substring(1, bnmp6.Length - 5)))
            };








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
        
            Stream pictureStream = File.Open("C:/ImagesForPowerPoint/ImagesForPowerPoint/Images/image1.png", FileMode.Open);
            

            //Adds the picture to a slide by specifying its size and position.
            slide.Shapes.AddPicture(pictureStream, 499.79, 238.59, 364.54, 192.16);
            pictureStream.Close();

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
            Stream pictureStream = File.Open("C:/ImagesForPowerPoint/ImagesForPowerPoint/Images/image2.png", FileMode.Open);

            //Adds the picture to a slide by specifying its size and position.
            slide.Shapes.AddPicture(pictureStream, 499.79, 238.59, 364.54, 192.16);
            pictureStream.Close();

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
            Stream pictureStream = File.Open("C:/ImagesForPowerPoint/ImagesForPowerPoint/Images/image3.png", FileMode.Open);

            //Adds the picture to a slide by specifying its size and position.
            slide.Shapes.AddPicture(pictureStream, 499.79, 238.59, 364.54, 192.16);
            pictureStream.Close();


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
            Stream pictureStream = File.Open("C:/ImagesForPowerPoint/ImagesForPowerPoint/Images/image4.png", FileMode.Open);
            
            //Adds the picture to a slide by specifying its size and position.
            slide.Shapes.AddPicture(pictureStream, 499.79, 238.59, 364.54, 192.16);
            pictureStream.Close();

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
            Stream pictureStream = File.Open("C:/ImagesForPowerPoint/ImagesForPowerPoint/Images/image5.png", FileMode.Open);

            //Adds the picture to a slide by specifying its size and position.
            slide.Shapes.AddPicture(pictureStream, 499.79, 238.59, 364.54, 192.16);
            pictureStream.Close();

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
            Stream pictureStream = File.Open("C:/ImagesForPowerPoint/ImagesForPowerPoint/Images/image6.png", FileMode.Open);

            //Adds the picture to a slide by specifying its size and position.
            slide.Shapes.AddPicture(pictureStream, 499.79, 238.59, 364.54, 192.16);
            pictureStream.Close();
            //Save the PowerPoint Presentation 
            //pptxDoc.Save("Sample.pptx");

            //Close the PowerPoint presentation
            //pptxDoc.Close(); 
        }

        
    }


}


