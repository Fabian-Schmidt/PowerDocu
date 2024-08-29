using PowerDocu.AppDocumenter;
using PowerDocu.Common;
using PowerDocu.FlowDocumenter;
using System;
using System.IO;

namespace PowerDocu.SolutionDocumenter
{
    public static class SolutionDocumentationGenerator
    {
        public static void GenerateDocumentation(string filePath, string fileFormat, bool documentDefaultChangesOnly, bool documentDefaults, bool documentSampleData, string flowActionSortOrder, string wordTemplate = null, string outputPath = null)
        {
            if (File.Exists(filePath))
            {
                var startDocGeneration = DateTime.Now;
                var flows = FlowDocumentationGenerator.GenerateDocumentation(
                    filePath,
                    fileFormat,
                    flowActionSortOrder,
                    wordTemplate,
                    outputPath
                );
                var apps = AppDocumentationGenerator.GenerateDocumentation(
                    filePath,
                    fileFormat,
                    documentDefaultChangesOnly,
                    documentDefaults,
                    documentSampleData,
                    wordTemplate,
                    outputPath
                );
                var solutionParser = new SolutionParser(filePath);
                if (solutionParser.solution != null)
                {
                    var path = outputPath == null ?
                        Path.Combine(Path.GetDirectoryName(filePath), "Solution " + CharsetHelper.GetSafeName(Path.GetFileNameWithoutExtension(filePath))) :
                        Path.Combine(outputPath, CharsetHelper.GetSafeName(Path.GetFileNameWithoutExtension(filePath)));

#if DEBUG
                    path = outputPath;
#endif

                    var solutionContent = new SolutionDocumentationContent(solutionParser.solution, apps, flows, path);
                    if (fileFormat.Equals(OutputFormatHelper.Word) || fileFormat.Equals(OutputFormatHelper.All))
                    {
                        //create the Word document
                        NotificationHelper.SendNotification("Creating Solution documentation");
                        if (String.IsNullOrEmpty(wordTemplate) || !File.Exists(wordTemplate))
                        {
                            var wordzip = new SolutionWordDocBuilder(solutionContent, null);
                        }
                        else
                        {
                            var wordzip = new SolutionWordDocBuilder(solutionContent, wordTemplate);
                        }
                    }
                    if (fileFormat.Equals(OutputFormatHelper.Markdown) || fileFormat.Equals(OutputFormatHelper.All))
                    {
                        var mdDoc = new SolutionMarkdownBuilder(solutionContent);
                    }
                    var endDocGeneration = DateTime.Now;
                    NotificationHelper.SendNotification("SolutionDocumenter: Created documentation for " + filePath + ". Total solution documentation completed in " + (endDocGeneration - startDocGeneration).TotalSeconds + " seconds.");
                }
            }
            else
            {
                NotificationHelper.SendNotification("File not found: " + filePath);
            }
        }
    }
}