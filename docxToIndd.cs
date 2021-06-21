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
		// instance of application (InDesign)
                Type t = Type.GetTypeFromProgID("InDesign.Application");
                InDesign.Application application = (InDesign.Application)Activator.CreateInstance(t);

		// Create a document 
                Document doc = application.Documents.Add(false);
		// Create window
                Window window = (Window)doc.Windows.Add();
                Console.WriteLine("pages " + doc.Pages.Count);

		// get page object
                InDesign.Page page = (InDesign.Page)doc.Pages[1];

		// get Layers info
                InDesign.Layer layer = (InDesign.Layer)doc.Layers[1];

		// get text from externl source 
                page.Place("Folder\\File_Name.docx", page ,layer, false, true);
                
		// renaming and saving document
                Console.WriteLine("created document");
                doc.Save("Folder\\File_Name.indd", false, "Ypor Comments here", false);
		doc.close();

            }
            catch (Exception e)
            {
                Console.WriteLine("error occured"+ e.Message);
            }
        }
    }
}
