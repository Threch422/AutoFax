using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System;
using System.IO;

namespace AutoFax
{
    public class WordHandler
    {
        protected string filePath;

        public WordHandler()
        {
            this.filePath = string.Empty;
        }

        public WordHandler(string filePath)
        {
            this.filePath = filePath;
        }

        // Returning the FaxNumber and RecipientName
        protected internal void GenerateDocument(List<Dictionary<string, string>> replaceList)
        {
            var wordApp = new Application();

            var templateDoc = wordApp.Documents.Open(filePath, false, true);

            foreach (var replaceFile in replaceList)
            {
                var fileName = replaceFile["{#VenueName#}"];

                foreach (var replaceWordPair in replaceFile)
                {
                    templateDoc.Content.Find.Execute(
                        FindText: replaceWordPair.Key,
                            MatchCase: true,
                            MatchWholeWord: true,
                            MatchWildcards: false,
                            MatchSoundsLike: false,
                            MatchAllWordForms: false,
                            Forward: true,
                            Wrap: false,
                            Format: false,
                            ReplaceWith: replaceWordPair.Value,
                            Replace: WdReplace.wdReplaceAll
                            );
                }
                templateDoc.SaveAs2($"{Directory.GetCurrentDirectory()}/wordResult/{fileName}.doc");

                Console.WriteLine($"{fileName}.doc is generated.");
            }

            templateDoc.Close();

            wordApp.Quit();
        }
    }
}
