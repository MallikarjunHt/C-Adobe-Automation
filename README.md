# Csharp-Adobe-Automation
Automating Adobe InDesign using C# programing

docTOIndd ➡️ import **docx** file and extract text and place in **InDesign** and sving file.

word manuplation => https://www.c-sharpcorner.com/forums/how-to-get-current-document-styles-from-word-document
https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.paragraphformat?view=word-pia


InDesign.Texts texts= (InDesign.Texts)frames.Texts;
                InDesign.Text text = (InDesign.Text)texts.FirstItem();
                
 for (int i = 1; i <= application.Fonts.Count; i++)
                {
                    InDesign.Font font = (InDesign.Font)application.Fonts[i];
                    if(font.FullName.Equals("Acumin Variable Concept SemiCondensed Black Italic"))
                    {
                        text.AppliedFont = font;
                        break;
                    }
                } 
