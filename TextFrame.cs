using InDesign;
using System;

namespace wordtoindd
{
    class Program
    {
        static void Main(string[] args)
        {

            // Create application instance
            try
            {
                Console.WriteLine("starting application");
                Type t = Type.GetTypeFromProgID("InDesign.Application");
                InDesign.Application application = (InDesign.Application)Activator.CreateInstance(t);

                Console.WriteLine("attempting to create document");
                Document doc = application.Documents.Add(false);
               
                doc.TextPreferences.SmartTextReflow = true;
                Window window = (Window)doc.Windows.Add();
                Console.WriteLine("pages " + doc.Pages.Count);
                
                InDesign.Page page = (InDesign.Page)doc.Pages[1];

                InDesign.Layer layer = (InDesign.Layer)doc.Layers[1];
                page.TextFrames.Add();
                Console.WriteLine("page frames " + page.TextFrames.Count);
                
                Console.WriteLine("created document");
                InDesign.TextFrame frames = (InDesign.TextFrame)page.TextFrames.FirstItem();
                
                // set x,y and width, height for box
                frames.GeometricBounds = new string[4] { "4p", "4p", "62p", "47p" };
                frames.Place("file path and name to save", false);

                doc.Save("file path and name to save", false, "Comments here", false);
                doc.Close();
                
            }
            catch (Exception e)
            {
                Console.WriteLine("error occured"+ e.Message);
            }
        }
    }
}
