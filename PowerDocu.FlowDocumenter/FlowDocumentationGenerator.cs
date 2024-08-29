using PowerDocu.Common;
using System;
using System.Collections.Generic;
using System.IO;

namespace PowerDocu.FlowDocumenter
{
    public static class FlowDocumentationGenerator
    {
        public static List<FlowEntity> GenerateDocumentation(string filePath, string fileFormat, string flowActionSortOrder, string wordTemplate = null, string outputPath = null)
        {
            if (File.Exists(filePath))
            {
                var path = outputPath == null
                    ? Path.GetDirectoryName(filePath)
                    : Path.Combine(outputPath, Path.GetFileNameWithoutExtension(filePath));
#if DEBUG
                path = outputPath;
#endif
                var startDocGeneration = DateTime.Now;
                var flowParserFromZip = new FlowParser(filePath);
                if (outputPath == null && flowParserFromZip.packageType == FlowParser.PackageType.SolutionPackage)
                {
                    path = Path.Combine(path, "Solution " + CharsetHelper.GetSafeName(Path.GetFileNameWithoutExtension(filePath)));
                }
                var flows = flowParserFromZip.getFlows();
                foreach (var flow in flows)
                {
                    var gbzip = new GraphBuilder(flow, path);
                    gbzip.buildTopLevelGraph();
                    gbzip.buildDetailedGraph();
                    var sortOrder = flowActionSortOrder switch
                    {
                        "By order of appearance" => FlowActionSortOrder.SortByOrder,
                        "By name" => FlowActionSortOrder.SortByName,
                        _ => FlowActionSortOrder.SortByName
                    };
                    var content = new FlowDocumentationContent(flow, path, sortOrder);
                    if (fileFormat.Equals(OutputFormatHelper.Word) || fileFormat.Equals(OutputFormatHelper.All))
                    {
                        NotificationHelper.SendNotification("Creating Word documentation");
                        if (String.IsNullOrEmpty(wordTemplate) || !File.Exists(wordTemplate))
                        {
                            var wordzip = new FlowWordDocBuilder(content, null);
                        }
                        else
                        {
                            var wordzip = new FlowWordDocBuilder(content, wordTemplate);
                        }
                    }
                    if (fileFormat.Equals(OutputFormatHelper.Markdown) || fileFormat.Equals(OutputFormatHelper.All))
                    {
                        NotificationHelper.SendNotification("Creating Markdown documentation");
                        var markdownFile = new FlowMarkdownBuilder(content);
                    }
                }
                var endDocGeneration = DateTime.Now;
                NotificationHelper.SendNotification("FlowDocumenter: Created documentation for " + filePath + ". A total of " + flowParserFromZip.getFlows().Count + " files were processed in " + (endDocGeneration - startDocGeneration).TotalSeconds + " seconds.");
                return flows;
            }
            else
            {
                NotificationHelper.SendNotification("File not found: " + filePath);
            }
            return null;
        }
    }
}