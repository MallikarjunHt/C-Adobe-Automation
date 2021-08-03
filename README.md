# Csharp-Adobe-Automation
Automating Adobe InDesign using C# programing

docTOIndd ➡️ import **docx** file and extract text and place in **InDesign** and sving file.

word manuplation => https://www.c-sharpcorner.com/forums/how-to-get-current-document-styles-from-word-document
https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.paragraphformat?view=word-pia

```javaScript
InDesign.Texts texts= (InDesign.Texts)frames.Texts;
InDesign.Text text = (InDesign.Text)texts.FirstItem();
```                
 ```javaScript               
 for (int i = 1; i <= application.Fonts.Count; i++)
                {
                    InDesign.Font font = (InDesign.Font)application.Fonts[i];
                    if(font.FullName.Equals("Acumin Variable Concept SemiCondensed Black Italic"))
                    {
                        text.AppliedFont = font;
                        break;
                    }
                } 
```

```javaScript
if(doc.Hyperlinks.Count > 0)
                {
                    for (int i = doc.Hyperlinks.Count; i >= 0; i--)
                    {
                        InDesign.Hyperlink hyperlink = (InDesign.Hyperlink)doc.Hyperlinks[i];
                        hyperlink.Delete();
                    }
                }
```
http://blog.gilbertconsulting.com/2007/10/use-grep-to-find-url.html

`[\w-]+(?:\.[\w-]+)*@(?:[\w-]+\.)+[a-zA-Z]{2,7}` finds Email
Both Email and URL `[\S]+[.][\S]+`
