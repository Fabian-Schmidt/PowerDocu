using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PowerDocu.Common;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;

namespace PowerDocu.AppDocumenter
{
    class AppWordDocBuilder : WordDocBuilder
    {
        private readonly AppDocumentationContent content;
        private readonly bool DetailedDocumentation = false;
        private readonly bool documentChangedDefaultsOnly;
        private readonly bool showDefaults;
        private readonly bool documentSampleData;

        public AppWordDocBuilder(AppDocumentationContent contentDocumentation, string template, bool documentChangedDefaultsOnly = false, bool showDefaults = true, bool documentSampleData = false)
        {
            content = contentDocumentation;
            this.documentChangedDefaultsOnly = documentChangedDefaultsOnly;
            this.showDefaults = showDefaults;
            this.documentSampleData = documentSampleData;
            Directory.CreateDirectory(content.folderPath);
            do
            {
                var filename = InitializeWordDocument(content.folderPath + content.filename + (DetailedDocumentation ? " detailed" : ""), template);
                using var wordDocument = WordprocessingDocument.Open(filename, true);
                mainPart = wordDocument.MainDocumentPart;
                body = mainPart.Document.Body;
                PrepareDocument(!String.IsNullOrEmpty(template));
                addAppProperties();
                addAppVariablesInfo();
                addAppDataSources();
                addAppResources();
                addAppControlsOverview(wordDocument);
                if (DetailedDocumentation)
                {
                    addDetailedAppControls();
                }

                DetailedDocumentation = !DetailedDocumentation;
            } while (DetailedDocumentation);
            NotificationHelper.SendNotification("Created Word documentation for " + contentDocumentation.Name);
        }

        private void addAppProperties()
        {
            var para = body.AppendChild(new Paragraph());
            var run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.appProperties.header));
            ApplyStyleToParagraph("Heading1", para);
            body.AppendChild(new Paragraph(new Run()));
            var table = CreateTable();
            table.Append(CreateRow(new Text("App Name"), new Text(content.Name)));
            //if there is a custom logo we add it to the documentation as well. Icon based logos currently not supported
            if (!String.IsNullOrEmpty(content.appProperties.appLogo))
            {
                if (content.ResourceStreams.TryGetValue(content.appProperties.appLogo, out var resourceStream))
                {
                    Drawing icon;
                    var imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
                    int imageWidth, imageHeight;
                    using (var image = Image.FromStream(resourceStream, false, false))
                    {
                        imageWidth = image.Width;
                        imageHeight = image.Height;
                    }
                    resourceStream.Position = 0;
                    imagePart.FeedData(resourceStream);
                    var usedWidth = (imageWidth > 400) ? 400 : imageWidth;
                    icon = InsertImage(mainPart.GetIdOfPart(imagePart), usedWidth, usedWidth * imageHeight / imageWidth);
                    var tr = CreateRow(new Text("App Logo"), icon);
                    if (!String.IsNullOrEmpty(content.appProperties.appBackgroundColour))
                    {
                        var tc = (TableCell)tr.LastChild;
                        var shading = new Shading()
                        {
                            Color = "auto",
                            Fill = ColourHelper.ParseColor(content.appProperties.appBackgroundColour),
                            Val = ShadingPatternValues.Clear
                        };
                        tc.TableCellProperties.Append(shading);
                    }
                    table.Append(tr);
                }
            }
            if (!CommandLineHelper.NoTimestamp)
            {
                table.Append(CreateRow(new Text(content.appProperties.headerDocumentationGenerated), new Text(DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToShortTimeString())));
            }

            var statisticsTable = CreateTable();
            foreach (var stats in content.appProperties.statisticsTable)
            {
                statisticsTable.Append(CreateRow(new Text(stats.Key), new Text(stats.Value)));
            }
            table.Append(CreateRow(new Text(content.appProperties.headerAppStatistics), statisticsTable));
            body.Append(table);
            body.AppendChild(new Paragraph(new Run(new Break())));
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.appProperties.headerAppProperties));
            ApplyStyleToParagraph("Heading1", para);
            body.AppendChild(new Paragraph(new Run()));
            table = CreateTable();
            foreach (var property in content.appProperties.appProperties)
            {
                if (!content.appProperties.propertiesToSkip.Contains(property.expressionOperator) && (content.appProperties.OverviewProperties.Contains(property.expressionOperator) || DetailedDocumentation))
                {
                    AddExpressionTable(property, table, 1, false, true);
                }
            }
            body.Append(table);
            body.AppendChild(new Paragraph(new Run(new Break())));
            addAppControlsTable(content.appControls.controls.First<ControlEntity>(o => o.Type == "appinfo"));
            if (DetailedDocumentation)
            {
                para = body.AppendChild(new Paragraph());
                run = para.AppendChild(new Run());
                run.AppendChild(new Text(content.appProperties.headerAppPreviewFlags));
                ApplyStyleToParagraph("Heading1", para);
                body.AppendChild(new Paragraph(new Run()));
                table = CreateTable();
                var appPreviewsFlagProperty = content.appProperties.appPreviewsFlagProperty;
                if (appPreviewsFlagProperty != null)
                {
                    foreach (Expression flagProp in appPreviewsFlagProperty.expressionOperands)
                    {
                        AddExpressionTable(flagProp, table, 1, false, true);
                    }
                }
                body.Append(table);
                body.AppendChild(new Paragraph(new Run(new Break())));
            }
        }

        private void addAppVariablesInfo()
        {
            var para = body.AppendChild(new Paragraph());
            var run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.appVariablesInfo.header));
            ApplyStyleToParagraph("Heading1", para);
            body.AppendChild(new Paragraph(new Run(new Text(content.appVariablesInfo.infoText))));
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.appVariablesInfo.headerGlobalVariables));
            ApplyStyleToParagraph("Heading2", para);
            var table = CreateTable();
            table.Append(CreateHeaderRow(new Text("Variable Name"), new Text("Used In")));
            foreach (var var in content.appVariablesInfo.globalVariables)
            {
                var varReferenceTable = CreateTable();
                content.appVariablesInfo.variableCollectionControlReferences.TryGetValue(var, out var references);
                if (references != null)
                {
                    varReferenceTable.Append(CreateHeaderRow(new Text("Control"), new Text("Property")));
                    foreach (var reference in references.OrderBy(o => o.Control.Name).ThenBy(o => o.RuleProperty))
                    {
                        varReferenceTable.Append(CreateRow(new Text(reference.Control.Name + " (" + reference.Control.Screen()?.Name + ")"), new Text(reference.RuleProperty)));
                    }
                }
                table.Append(CreateRow(new Text(var), varReferenceTable));
            }
            body.Append(table);
            body.AppendChild(new Paragraph(new Run(new Break())));
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.appVariablesInfo.headerContextVariables));
            ApplyStyleToParagraph("Heading2", para);
            table = CreateTable();
            table.Append(CreateHeaderRow(new Text("Variable Name"), new Text("Used In")));
            foreach (var var in content.appVariablesInfo.contextVariables)
            {
                var varReferenceTable = CreateTable();
                content.appVariablesInfo.variableCollectionControlReferences.TryGetValue(var, out var references);
                if (references != null)
                {
                    varReferenceTable.Append(CreateHeaderRow(new Text("Control"), new Text("Property")));
                    foreach (var reference in references.OrderBy(o => o.Control.Name).ThenBy(o => o.RuleProperty))
                    {
                        varReferenceTable.Append(CreateRow(new Text(reference.Control.Name + " (" + reference.Control.Screen()?.Name + ")"), new Text(reference.RuleProperty)));
                    }
                }
                table.Append(CreateRow(new Text(var), varReferenceTable));
            }
            body.Append(table);
            body.AppendChild(new Paragraph(new Run(new Break())));
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.appVariablesInfo.headerCollections));
            ApplyStyleToParagraph("Heading2", para);
            table = CreateTable();
            table.Append(CreateHeaderRow(new Text("Collection Name"), new Text("Used In")));
            foreach (var coll in content.appVariablesInfo.collections)
            {
                var collReferenceTable = CreateTable();
                content.appVariablesInfo.variableCollectionControlReferences.TryGetValue(coll, out var references);
                if (references != null)
                {
                    collReferenceTable.Append(CreateHeaderRow(new Text("Control"), new Text("Property")));
                    foreach (var reference in references.OrderBy(o => o.Control.Name).ThenBy(o => o.RuleProperty))
                    {
                        collReferenceTable.Append(CreateRow(new Text(reference.Control.Name), new Text(reference.RuleProperty)));
                    }
                }
                table.Append(CreateRow(new Text(coll), collReferenceTable));
            }
            body.Append(table);
            body.AppendChild(new Paragraph(new Run(new Break())));
        }

        private void addAppControlsOverview(WordprocessingDocument wordDoc)
        {
            var para = body.AppendChild(new Paragraph());
            var run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.appControls.headerOverview));
            ApplyStyleToParagraph("Heading1", para);
            body.AppendChild(new Paragraph(new Run(new Text(content.appControls.infoTextScreens))));
            body.AppendChild(new Paragraph(new Run(new Text(content.appControls.infoTextControls))));
            foreach (var control in content.appControls.controls.Where(o => o.Type != "appinfo"))
            {
                para = body.AppendChild(new Paragraph());
                run = para.AppendChild(new Run());
                if (DetailedDocumentation)
                {
                    run.AppendChild(new Hyperlink(new Text("Screen: " + control.Name))
                    {
                        Anchor = CreateMD5Hash(control.Name),
                        DocLocation = ""
                    });
                }
                else
                {
                    run.AppendChild(new Text("Screen: " + control.Name));
                }
                ApplyStyleToParagraph("Heading2", para);
                body.AppendChild(CreateControlTable(control));
            }
            body.AppendChild(new Paragraph(new Run(new Break())));
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.appControls.headerScreenNavigation));
            ApplyStyleToParagraph("Heading2", para);
            body.AppendChild(new Paragraph(new Run(new Text(content.appControls.infoTextScreenNavigation))));
            var imagePart = wordDoc.MainDocumentPart.AddImagePart(ImagePartType.Png);
            int imageWidth, imageHeight;
            using (var stream = new FileStream(content.folderPath + content.appControls.imageScreenNavigation + ".png", FileMode.Open))
            {
                using (var image = Image.FromStream(stream, false, false))
                {
                    imageWidth = image.Width;
                    imageHeight = image.Height;
                }
                stream.Position = 0;
                imagePart.FeedData(stream);
            }
            var svgPart = wordDoc.MainDocumentPart.AddNewPart<ImagePart>("image/svg+xml", "rId" + new Random().Next(100000, 999999));
            using (var stream = new FileStream(content.folderPath + content.appControls.imageScreenNavigation + ".svg", FileMode.Open))
            {
                svgPart.FeedData(stream);
            }
            body.AppendChild(new Paragraph(new Run(
                InsertSvgImage(wordDoc.MainDocumentPart.GetIdOfPart(svgPart), wordDoc.MainDocumentPart.GetIdOfPart(imagePart), imageWidth, imageHeight)
            )));
            body.AppendChild(new Paragraph(new Run(new Break())));
        }

        private Table CreateControlTable(ControlEntity control)
        {
            return CreateControlTable(control, BorderValues.Single);
        }

        private Table CreateControlTable(ControlEntity control, BorderValues borderType)
        {
            var table = CreateTable();
            table.GetFirstChild<TableProperties>().TableBorders = new TableBorders(
                    SetDefaultTableBorderStyle(new TopBorder(), borderType),
                    SetDefaultTableBorderStyle(new LeftBorder(), borderType),
                    SetDefaultTableBorderStyle(new BottomBorder(), borderType),
                    SetDefaultTableBorderStyle(new RightBorder(), borderType),
                    SetDefaultTableBorderStyle(new InsideHorizontalBorder(), BorderValues.None),
                    SetDefaultTableBorderStyle(new InsideVerticalBorder(), BorderValues.None)
                );
            var controlType = control.Type;
            OpenXmlElement controlElement;
            if (DetailedDocumentation)
            {
                controlElement = new Hyperlink(new Run(new Text(control.Name + " [" + controlType + "]")))
                {
                    Anchor = CreateMD5Hash(control.Name),
                    DocLocation = ""
                };
            }
            else
            {
                controlElement = new Text(control.Name + " [" + controlType + "]");
            }
            table.Append(CreateRow(InsertSvgImage(mainPart, AppControlIcons.GetControlIcon(controlType), 32, 32), controlElement));
            foreach (var child in control.Children.OrderBy(o => o.Name).ToList())
            {
                table.Append(CreateRow(new Text(""), CreateControlTable(child, BorderValues.None)));
            }
            return table;
        }

        private void addDetailedAppControls()
        {
            var para = body.AppendChild(new Paragraph());
            var run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.appControls.headerDetails));
            ApplyStyleToParagraph("Heading1", para);
            foreach (var screen in content.appControls.controls.Where(o => o.Type == "screen").OrderBy(o => o.Name).ToList())
            {
                para = body.AppendChild(new Paragraph());
                run = para.AppendChild(new Run());
                run.AppendChild(new Text(screen.Name));
                var bookmarkID = new Random().Next(100000, 999999).ToString();
                var start = new BookmarkStart() { Name = CreateMD5Hash(screen.Name), Id = bookmarkID };
                var end = new BookmarkEnd() { Id = bookmarkID };
                para.Append(start, end);
                ApplyStyleToParagraph("Heading2", para);
                body.AppendChild(new Paragraph(new Run()));
                addAppControlsTable(screen);
                foreach (var control in content
                    .appControls
                    .allControls
                    .Where(o => o.Type != "appinfo" && o.Type != "screen" && screen.Equals(o.Screen()))
                    .OrderBy(o => o.Name)
                    .ToList()
                    )
                {
                    para = body.AppendChild(new Paragraph());
                    run = para.AppendChild(new Run());
                    run.AppendChild(new Text(control.Name));
                    bookmarkID = new Random().Next(100000, 999999).ToString();
                    start = new BookmarkStart() { Name = CreateMD5Hash(control.Name), Id = bookmarkID };
                    end = new BookmarkEnd() { Id = bookmarkID };
                    para.Append(start, end);
                    ApplyStyleToParagraph("Heading3", para);
                    body.AppendChild(new Paragraph(new Run()));
                    addAppControlsTable(control);
                }
            }
            body.AppendChild(new Paragraph(new Run(new Break())));
        }

        private void addAppControlsTable(ControlEntity control)
        {
            var defaultEntity = DefaultChangeHelper.GetEntityDefaults(control.Type);
            var table = CreateTable();
            var typeTable = CreateTable(BorderValues.None);
            typeTable.Append(CreateRow(InsertSvgImage(mainPart, AppControlIcons.GetControlIcon(control.Type), 16, 16), new Text(control.Type)));
            table.Append(CreateRow(new Text("Type"), typeTable));
            var category = "";
            foreach (var rule in control.Rules.OrderBy(o => o.Category).ThenBy(o => o.Property).ToList())
            {
                var defaultValue = defaultEntity?.Rules.Find(r => r.Property == rule.Property)?.InvariantScript;
                if (String.IsNullOrEmpty(defaultValue))
                {
                    defaultValue = DefaultChangeHelper.DefaultValueIfUnknown;
                }

                if (!documentChangedDefaultsOnly || (defaultValue != rule.InvariantScript))
                {
                    if (!content.ColourProperties.Contains(rule.Property))
                    {
                        if (rule.Category != category)
                        {
                            category = rule.Category;
                            table.Append(CreateMergedRow(new Text(category), 2, WordDocBuilder.cellHeaderBackground));
                        }
                        if (rule.InvariantScript.StartsWith("RGBA("))
                        {
                            table.Append(CreateColorTable(rule, defaultValue));
                        }
                        else
                        {
                            table.Append(CreateRowForControlProperty(rule, defaultValue));
                        }
                    }
                }
            }
            var colourPropertiesHeaderAdded = false;
            foreach (var property in content.ColourProperties)
            {
                var rule = control.Rules.Find(o => o.Property == property);
                if (rule != null)
                {
                    var defaultValue = defaultEntity?.Rules.Find(r => r.Property == rule.Property)?.InvariantScript;
                    if (String.IsNullOrEmpty(defaultValue))
                    {
                        defaultValue = DefaultChangeHelper.DefaultValueIfUnknown;
                    }

                    if (!documentChangedDefaultsOnly || defaultValue != rule.InvariantScript)
                    {
                        //we only need to add this once, and only if we add content
                        if (!colourPropertiesHeaderAdded)
                        {
                            table.Append(CreateMergedRow(new Text("Color Properties"), 2, WordDocBuilder.cellHeaderBackground));
                            colourPropertiesHeaderAdded = true;
                        }
                        if (rule.InvariantScript.StartsWith("RGBA("))
                        {
                            table.Append(CreateColorTable(rule, defaultValue));
                        }
                        else
                        {
                            table.Append(CreateRowForControlProperty(rule, defaultValue));
                        }
                    }
                }
            }
            if (control.Children.Count > 0 || control.Parent != null)
            {
                table.Append(CreateMergedRow(new Text("Child & Parent Controls"), 2, WordDocBuilder.cellHeaderBackground));
                var childtable = CreateTable(BorderValues.None);
                foreach (var childControl in control.Children)
                {
                    childtable.Append(CreateRow(new Text(childControl.Name)));
                }
                table.Append(CreateRow(new Text("Child Controls"), childtable));
                if (control.Parent != null)
                {
                    table.Append(CreateRow(new Text("Parent Control"), new Text(control.Parent.Name)));
                }
            }
            //todo isLocked property could be documented
            /* //Other properties are likely not needed for documentation, still keeping this code in case we want to show them at some point
            table.Append(CreateMergedRow(new Text("Properties"), 2, WordDocBuilder.cellHeaderBackground));
            foreach (Expression expression in control.Properties)
            {
                AddExpressionTable(expression, table);
            }*/
            body.Append(table);
            body.AppendChild(new Paragraph(new Run(new Break())));
        }

        private TableRow CreateRowForControlProperty(Rule rule, string defaultValue)
        {
            OpenXmlElement value = new Text(rule.InvariantScript);
            if (showDefaults && defaultValue != rule.InvariantScript && !content.appControls.controlPropertiesToSkip.Contains(rule.Property))
            {
                value = CreateTable(BorderValues.None);
                value.Append(CreateChangedDefaultColourRow(CreateRunWithLinebreaks(rule.InvariantScript), new Text(defaultValue)));
            }
            return CreateRow(new Text(rule.Property), value);
        }

        private TableRow CreateChangedDefaultColourRow(OpenXmlElement firstColumnElement, OpenXmlElement secondColumnElement)
        {
            var fiftyPercentWidth = new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "2500" };
            var tr = CreateRow(firstColumnElement, secondColumnElement);
            //update the cell with the current value
            var tc = (TableCell)tr.FirstChild;
            var shading = new Shading()
            {
                Color = "auto",
                Fill = "ccffcc",
                Val = ShadingPatternValues.Clear
            };
            tc.TableCellProperties.Append(shading);
            tc.TableCellProperties.TableCellWidth = (TableCellWidth)fiftyPercentWidth.Clone();
            //update the cell with the default value
            tc = (TableCell)tr.LastChild;
            shading = new Shading()
            {
                Color = "auto",
                Fill = "ffcccc",
                Val = ShadingPatternValues.Clear
            };
            tc.TableCellProperties.Append(shading);
            tc.TableCellProperties.TableCellWidth = (TableCellWidth)fiftyPercentWidth.Clone();
            return tr;
        }

        private TableRow CreateColorTable(Rule rule, string defaultValue)
        {
            var colorTable = CreateTable(BorderValues.None);
            colorTable.Append(CreateRow(new Text(rule.InvariantScript)));
            var colour = ColourHelper.ParseColor(rule.InvariantScript[..(rule.InvariantScript.IndexOf(')') + 1)]);
            if (!String.IsNullOrEmpty(colour))
            {
                colorTable.Append(CreateMergedRow(new Text(""), 1, colour));
            }
            if (showDefaults && defaultValue != rule.InvariantScript && !content.appControls.controlPropertiesToSkip.Contains(rule.Property))
            {
                var defaultTable = CreateTable(BorderValues.None);
                defaultTable.Append(CreateRow(new Text(defaultValue)));
                var defaultColour = ColourHelper.ParseColor(defaultValue);
                if (!String.IsNullOrEmpty(defaultColour))
                {
                    defaultTable.Append(CreateMergedRow(new Text(""), 1, defaultColour));
                }
                var changesTable = CreateTable(BorderValues.None);
                changesTable.Append(CreateChangedDefaultColourRow(colorTable, defaultTable));
                return CreateRow(new Text(rule.Property), changesTable);
            }
            return CreateRow(new Text(rule.Property), colorTable);
        }

        private void addAppDataSources()
        {
            var para = body.AppendChild(new Paragraph());
            var run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.appDataSources.header));
            ApplyStyleToParagraph("Heading1", para);
            body.AppendChild(new Paragraph(new Run(new Text(content.appDataSources.infoText))));
            foreach (var datasource in content.appDataSources.dataSources)
            {
                if (!datasource.isSampleDataSource() || documentSampleData)
                {
                    para = body.AppendChild(new Paragraph());
                    run = para.AppendChild(new Run());
                    run.AppendChild(new Text(datasource.Name));
                    ApplyStyleToParagraph("Heading2", para);
                    body.AppendChild(new Paragraph(new Run()));
                    var table = CreateTable();
                    table.Append(CreateRow(new Text("Name"), new Text(datasource.Name)));
                    table.Append(CreateRow(new Text("Type"), new Text(datasource.Type)));
                    if (DetailedDocumentation)
                    {
                        table.Append(CreateMergedRow(new Text("DataSource Properties"), 2, WordDocBuilder.cellHeaderBackground));
                        foreach (var expression in datasource.Properties.OrderBy(o => o.expressionOperator))
                        {
                            AddExpressionTable(expression, table);
                        }
                    }
                    body.Append(table);
                    body.AppendChild(new Paragraph(new Run(new Break())));
                }
            }
            body.AppendChild(new Paragraph(new Run(new Break())));
        }

        private void addAppResources()
        {
            var para = body.AppendChild(new Paragraph());
            var run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.appResources.header));
            ApplyStyleToParagraph("Heading1", para);
            body.AppendChild(new Paragraph(new Run(new Text(content.appResources.infoText))));
            foreach (var resource in content.appResources.resources)
            {
                if (!resource.isSampleResource())
                {
                    para = body.AppendChild(new Paragraph());
                    run = para.AppendChild(new Run());
                    run.AppendChild(new Text(resource.Name));
                    ApplyStyleToParagraph("Heading2", para);
                    body.AppendChild(new Paragraph(new Run()));
                    var table = CreateTable();
                    table.Append(CreateRow(new Text("Name"), new Text(resource.Name)));
                    table.Append(CreateRow(new Text("Content"), new Text(resource.Content)));
                    table.Append(CreateRow(new Text("Resource Kind"), new Text(resource.ResourceKind)));
                    if (resource.ResourceKind == "LocalFile" && content.ResourceStreams.TryGetValue(resource.Name, out var resourceStream))
                    {
                        try
                        {
                            Drawing icon = null;
                            var fileName = resource.Properties.First(o => o.expressionOperator == "FileName");
                            if (fileName.expressionOperands[0].ToString().EndsWith("svg", StringComparison.OrdinalIgnoreCase))
                            {
                                var svg = Encoding.Default.GetString(resourceStream.ToArray());
                                icon = InsertSvgImage(mainPart, svg, 400, 400);
                            }
                            else
                            {
                                var imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
                                int imageWidth, imageHeight;
                                using (var image = Image.FromStream(resourceStream, false, false))
                                {
                                    imageWidth = image.Width;
                                    imageHeight = image.Height;
                                }
                                resourceStream.Position = 0;
                                imagePart.FeedData(resourceStream);
                                var usedWidth = (imageWidth > 400) ? 400 : imageWidth;
                                icon = InsertImage(mainPart.GetIdOfPart(imagePart), usedWidth, usedWidth * imageHeight / imageWidth);
                            }
                            table.Append(CreateRow(new Text("Resource Preview"), icon));
                        }
                        catch (Exception e)
                        {
                            table.Append(CreateRow(new Text("Resource Preview"), new Text("Resource Preview is not available, media file is invalid.")));
                        }
                    }
                    if (DetailedDocumentation)
                    {
                        table.Append(CreateMergedRow(new Text("Resource Properties"), 2, WordDocBuilder.cellHeaderBackground));
                        foreach (var expression in resource.Properties.OrderBy(o => o.expressionOperator))
                        {
                            AddExpressionTable(expression, table);
                        }
                    }
                    body.Append(table);
                    body.AppendChild(new Paragraph(new Run(new Break())));
                }
            }
            body.AppendChild(new Paragraph(new Run(new Break())));
        }
    }
}