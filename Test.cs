using System;
using System.Collections.Generic;
using System.IO;
using System.Drawing;
using SautinSoft.Document;
using System.Collections;

namespace wordtoindd
{
    class Program
    {
        static void Main(string[] args)
        {
/*            string response = "{\"response\":\"message\",\"errorCode\":\"eCode\"}";
            string filePath = args[0];
            
            var dirName = Path.GetDirectoryName(filePath);
            var fileName = Path.GetFileName(filePath);
            string outputFile = "response_" + fileName + ".json";
            // Create application instance
            var saveFile = Path.Combine(dirName +"//"+ Path.GetFileNameWithoutExtension(fileName) + ".indd");
            Console.WriteLine("Arguments " + args.ToString() );*/
            //INPUT
            double pgWidth = 5.6;
            double pgHeight = 11;
            String InputFile = @"D:\worddocument\Test\Input\BRA1982398_Proof - Darlington _ Stockton Times - 29.01.21 - PLA.docx";
            String OutputFile = @"D:\worddocument\Test\output\BRA1982398_Proof - Darlington _ Stockton Times - 29.01.21 - PLA.docx";


            try
            {
                Console.WriteLine("starting application");
                Type t = Type.GetTypeFromProgID("InDesign.Application");
                InDesign.Application application = (InDesign.Application)Activator.CreateInstance(t);
                Console.WriteLine("attempting to create document");
                application.ViewPreferences.HorizontalMeasurementUnits = InDesign.idMeasurementUnits.idCentimeters;
                application.ViewPreferences.VerticalMeasurementUnits = InDesign.idMeasurementUnits.idCentimeters;
                
                
                //add doc and set preferences
                InDesign.Document doc = application.Documents.Add(false);
                doc.TextPreferences.SmartTextReflow = true;
                doc.Save(OutputFile, false, "bot generated", false);
                
                // chge defolt doc size
                doc.DocumentPreferences.PageWidth = pgWidth;
                doc.DocumentPreferences.PageHeight = pgHeight;

                
                InDesign.Window window = (InDesign.Window)doc.Windows.Add();
                Console.WriteLine("pages " + doc.Pages.Count);

                // add page to active document
                InDesign.Page page = (InDesign.Page)doc.Pages.FirstItem();
                page.MarginPreferences.Left = "0 cm";
                page.MarginPreferences.Right = "0 cm";
                page.MarginPreferences.Bottom = "0 cm";
                page.MarginPreferences.Top = "0 cm";

                InDesign.Layer layer = (InDesign.Layer)doc.Layers[1];
                layer.Name = "bot";

                // "D:\worddocument\Test\Input\Legal Notice Advts\BDF8544478\in\BRA1982398_1\BRA1982398_1\BRA1982398_Proof - Darlington _ Stockton Times - 29.01.21 - PLA.docx"
                //page.Place(@"D:\worddocument\Test\Input\BarnumsHumbug.docx", page ,layer, false, true);
                //InDesign.idLeading leading = (InDesign.idLeading)1.5;

                // Text Frame
                page.TextFrames.Add();
                InDesign.TextFrame textFrame = (InDesign.TextFrame)page.TextFrames.FirstItem();
                textFrame.GeometricBounds = new string[4] { "0.15 cm ", "0.15 cm", $"{pgHeight - 0.15} cm ", $"{pgWidth - 0.15} cm" };
                textFrame.Place(InputFile, false);
                application.DoScript(@"D:\worddocument\scripts\FindChangeByList.jsx", InDesign.idScriptLanguage.idJavascript);

                // IEnumerator paragraphs = textFrame.Paragraphs.GetEnumerator();
                for (int j = 1; j <= textFrame.Paragraphs.Count; j++)
                {
                    Console.WriteLine("paragraph style change");
                    
                    InDesign.Paragraph paragraph = (InDesign.Paragraph) textFrame.Paragraphs[j];
                    Console.WriteLine("content " + paragraph.FontStyle);
                    for(int i =1; i <= paragraph.Characters.Count; i++)
                    {
                        InDesign.Character characters = (InDesign.Character)paragraph.Characters[i];
                        if (characters.FontStyle == "Bold")
                        {
                            characters.FontStyle = "Bold";
                            characters.AppliedFont = "Sans";
                        }
                        else
                        {
                            characters.FontStyle = "Regular";
                            characters.AppliedFont = "Sans";
                        }
                    }
                    
                    paragraph.Justification = InDesign.idJustification.idLeftAlign;
                    paragraph.PointSize = 7;
                    paragraph.Leading = 7;
                    paragraph.SpaceAfter = 0.12;
                    paragraph.Hyphenation = false;                    
                }

                //Grep Changes
                application.FindTextPreferences = application.ChangeTextPreferences = InDesign.idNothingEnum.idNothing;
                InDesign.FindTextPreference findGrep = (InDesign.FindTextPreference) application.FindTextPreferences;
                findGrep.FindWhat = "([\\S]+[.][\\S]+)";
                InDesign.ChangeGrepPreference ChangeGrep = (InDesign.ChangeGrepPreference)application.ChangeGrepPreferences;
                ChangeGrep.AppliedCharacterStyle = InDesign.idNothingEnum.idNothing;
                ChangeGrep.ChangeTo = "$1";

                application.DoScript(@"D:\worddocument\scripts\FindChangeByList.jsx", InDesign.idScriptLanguage.idJavascript);
                // Logo
                Console.WriteLine("lago");
                var img =Image.FromFile(@"D:\worddocument\Test\Input\aotm_logo.png");
                InDesign.TextFrame frames = (InDesign.TextFrame)page.TextFrames.FirstItem();
                double height = img.Height * 0.0104166667;
                double width = img.Width * 0.0104166667;
                Console.WriteLine("height and width " + img.Height + " , " + img.Width);

                // set x,y and width, height for box
                //frames.GeometricBounds = new string[4] { "0.85 in", "2.9 in",$"{height + 1.4} in  " , $"{width + 3.95} in " };
                //var tfp = frames.TextFramePreferences;
                //tfp.AutoSizingReferencePoint = InDesign.idAutoSizingReferenceEnum.idTopLeftPoint;
                //tfp.AutoSizingType = InDesign.idAutoSizingTypeEnum.idHeightAndWidthProportionally;
                //frames.Fit(InDesign.idFitOptions.idProportionally);
                //frames.Place(@"D:\worddocument\Test\Input\aotm_logo.png", false);
                Console.WriteLine("created document"+ doc.FullName);

                // save
                //"D:\worddocument\Test\output\Legal Notice Advts\BDF8544478\in\BRA1982398_1\BRA1982398_1\BRA1982398_Proof - Darlington _ Stockton Times - 29.01.21 - PLA.docx"
                doc.Save(OutputFile, false, "bot generated", false);
              //  doc.Close();
                
            }
            catch (Exception e)
            {
                Console.WriteLine("error occured"+ e.Message);
            }
        }

    }
  
}
