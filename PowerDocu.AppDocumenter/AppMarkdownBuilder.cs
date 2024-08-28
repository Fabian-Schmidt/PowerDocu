using Grynwald.MarkdownGenerator;
using PowerDocu.Common;
using Svg;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;

namespace PowerDocu.AppDocumenter
{
    class AppMarkdownBuilder : MarkdownBuilder
    {
        private readonly AppDocumentationContent content;
        private readonly string mainDocumentFileName, appDetailsFileName, variablesDocumentFileName, dataSourcesFileName, resourcesFileName, controlsFileName;
        private readonly MdDocument mainDocument, appDetailsDocument, variablesDocument, dataSourcesDocument, resourcesDocument, controlsDocument;
        private readonly Dictionary<string, MdDocument> screenDocuments = new Dictionary<string, MdDocument>();
        private readonly Dictionary<string, MdDocument> datasourcesDocuments = new Dictionary<string, MdDocument>();
        private readonly DocumentSet<MdDocument> set;
        private MdTable metadataTable;
        private readonly int appLogoWidth = 250;
        private readonly bool documentChangedDefaultsOnly;
        private readonly bool showDefaults;
        private readonly bool documentSampleData;

        public AppMarkdownBuilder(AppDocumentationContent contentdocumentation, bool documentChangedDefaultsOnly = false, bool showDefaults = true, bool documentSampleData = false)
        {
            content = contentdocumentation;
            this.documentChangedDefaultsOnly = documentChangedDefaultsOnly;
            this.showDefaults = showDefaults;
            this.documentSampleData = documentSampleData;
            Directory.CreateDirectory(content.folderPath);
            mainDocumentFileName = ("index " + content.filename + ".md").Replace(" ", "-");
            appDetailsFileName = ("appdetails " + content.filename + ".md").Replace(" ", "-");
            variablesDocumentFileName = ("variables " + content.filename + ".md").Replace(" ", "-");
            dataSourcesFileName = ("datasources " + content.filename + ".md").Replace(" ", "-");
            resourcesFileName = ("resources " + content.filename + ".md").Replace(" ", "-");
            controlsFileName = ("controls " + content.filename + ".md").Replace(" ", "-");
            set = new DocumentSet<MdDocument>();
            mainDocument = set.CreateMdDocument(mainDocumentFileName);
            appDetailsDocument = set.CreateMdDocument(appDetailsFileName);
            variablesDocument = set.CreateMdDocument(variablesDocumentFileName);
            dataSourcesDocument = set.CreateMdDocument(dataSourcesFileName);
            //a dedicated document for each datasource
            foreach (var datasource in content.appDataSources.dataSources.OrderBy(o => o.Name).ToList())
            {
                datasourcesDocuments.Add(datasource.Name, set.CreateMdDocument(("datasource " + datasource.Name + " " + content.filename + ".md").Replace(" ", "-")));
            }
            resourcesDocument = set.CreateMdDocument(resourcesFileName);
            controlsDocument = set.CreateMdDocument(controlsFileName);
            //a dedicated document for each screen
            foreach (var screen in contentdocumentation.appControls.controls.Where(o => o.Type == "screen").OrderBy(o => o.Name).ToList())
            {
                screenDocuments.Add(screen.Name, set.CreateMdDocument(("screen " + screen.Name + " " + content.filename + ".md").Replace(" ", "-")));
            }
            addAppMetadata();
            addAppOverview();
            addAppDetails();
            addAppVariablesInfo();
            addAppDataSources();
            addAppResources();
            addAppControlsOverview();
            addDetailedAppControls();
            set.Save(content.folderPath);
            NotificationHelper.SendNotification("Created Markdown documentation for " + content.Name);
        }

        private void addAppOverview()
        {
            var tableRows = new List<MdTableRow>();
            mainDocument.Root.Add(new MdHeading(content.appProperties.headerAppStatistics, 2));
            foreach (var kvp2 in content.appProperties.statisticsTable)
            {
                tableRows.Add(new MdTableRow(kvp2.Key, kvp2.Value));
            }
            mainDocument.Root.Add(new MdTable(new MdTableRow("Component Type", "Count"), tableRows));
        }

        private MdBulletList getNavigationLinks(bool topLevel = true)
        {
            var navItems = new MdListItem[] {
                new MdListItem(new MdLinkSpan("Overview", topLevel ? mainDocumentFileName : "../" + mainDocumentFileName)),
                new MdListItem(new MdLinkSpan("App Details", topLevel ? appDetailsFileName : "../" + appDetailsFileName)),
                new MdListItem(new MdLinkSpan("Variables", topLevel ? variablesDocumentFileName : "../" + variablesDocumentFileName)),
                new MdListItem(new MdLinkSpan("DataSources", topLevel ? dataSourcesFileName : "../" + dataSourcesFileName)),
                new MdListItem(new MdLinkSpan("Resources", topLevel ? resourcesFileName : "../" + resourcesFileName)),
                new MdListItem(new MdLinkSpan("Controls", topLevel ? controlsFileName : "../" + controlsFileName)),
                };
            return new MdBulletList(navItems);
        }

        private void addAppMetadata()
        {
            var tableRows = new List<MdTableRow>
            {
                new MdTableRow("App Name", content.Name)
            };
            if (!String.IsNullOrEmpty(content.appProperties.appLogo))
            {
                if (content.ResourceStreams.TryGetValue(content.appProperties.appLogo, out var resourceStream))
                {
                    Directory.CreateDirectory(content.folderPath + "resources");
                    Bitmap appLogo;
                    if (!String.IsNullOrEmpty(content.appProperties.appBackgroundColour))
                    {
                        var c = ColorTranslator.FromHtml(ColourHelper.ParseColor(content.appProperties.appBackgroundColour));
                        var bmp = new Bitmap(resourceStream);
                        appLogo = new Bitmap(bmp.Width, bmp.Height);
                        var rect = new Rectangle(Point.Empty, bmp.Size);
                        using (var G = Graphics.FromImage(appLogo))
                        {
                            G.Clear(c);
                            G.DrawImageUnscaledAndClipped(bmp, rect);
                        }
                        appLogo.Save(content.folderPath + @"resources\applogo.png");
                    }
                    else
                    {
                        using Stream streamToWriteTo = File.Open(content.folderPath + @"resources\applogo.png", FileMode.Create);
                        resourceStream.CopyTo(streamToWriteTo);
                        resourceStream.Position = 0;
                        appLogo = new Bitmap(resourceStream);
                    }
                    resourceStream.Position = 0;
                    if (appLogo.Width > appLogoWidth)
                    {
                        var resized = new Bitmap(appLogo, new Size(appLogoWidth, appLogoWidth * appLogo.Height / appLogo.Width));
                        resized.Save(content.folderPath + @"resources\applogoSmall.png");
                        tableRows.Add(new MdTableRow("App Logo", new MdImageSpan("App Logo", "resources/applogoSmall.png")));
                    }
                    else
                    {
                        tableRows.Add(new MdTableRow("App Logo", new MdImageSpan("App Logo", "resources/applogo.png")));
                    }
                }
            }
            if (!CommandLineHelper.NoTimestamp)
            {
                tableRows.Add(new MdTableRow(content.appProperties.headerDocumentationGenerated, DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToShortTimeString()));
            }

            metadataTable = new MdTable(new MdTableRow(new List<string>() { "Property", "Value" }), tableRows);
            // prepare the common sections for all documents
            foreach (var doc in set.Documents)
            {
                doc.Root.Add(new MdHeading(content.appProperties.header, 1));
                doc.Root.Add(metadataTable);
                doc.Root.Add(getNavigationLinks());
            }
        }

        private void addAppDetails()
        {
            var tableRows = new List<MdTableRow>();
            appDetailsDocument.Root.Add(new MdHeading(content.appProperties.headerAppProperties, 2));
            foreach (var property in content.appProperties.appProperties)
            {
                if (!content.appProperties.propertiesToSkip.Contains(property.expressionOperator))
                {
                    tableRows.Add(new MdTableRow(property.expressionOperator, property.expressionOperands[0].ToString()));
                }
            }
            if (tableRows.Count > 0)
            {
                appDetailsDocument.Root.Add(new MdTable(new MdTableRow("App Property", "Value"), tableRows));
            }
            appDetailsDocument.Root.Add(new MdHeading(content.appProperties.headerAppPreviewFlags, 2));
            tableRows = new List<MdTableRow>();
            if (content.appProperties.appPreviewsFlagProperty != null)
            {
                foreach (Expression flagProp in content.appProperties.appPreviewsFlagProperty.expressionOperands)
                {
                    tableRows.Add(new MdTableRow(flagProp.expressionOperator, flagProp.expressionOperands[0].ToString()));
                }
                if (tableRows.Count > 0)
                {
                    appDetailsDocument.Root.Add(new MdTable(new MdTableRow("Preview Flag", "Value"), tableRows));
                }
            }
        }

        private void addAppVariablesInfo()
        {
            variablesDocument.Root.Add(new MdHeading(content.appVariablesInfo.header, 2));
            variablesDocument.Root.Add(new MdParagraph(new MdTextSpan(content.appVariablesInfo.infoText)));
            variablesDocument.Root.Add(new MdHeading(content.appVariablesInfo.headerGlobalVariables, 3));
            foreach (var var in content.appVariablesInfo.globalVariables)
            {
                variablesDocument.Root.Add(new MdHeading(var, 4));
                content.appVariablesInfo.variableCollectionControlReferences.TryGetValue(var, out var references);
                if (references != null)
                {
                    variablesDocument.Root.Add(new MdParagraph(new MdTextSpan("Variable used in:")));
                    var tableRows = new List<MdTableRow>();
                    foreach (var reference in references.OrderBy(o => o.Control.Name).ThenBy(o => o.RuleProperty))
                    {
                        //link to the screen instead of the control directly for the moment, as the directly generated anchor link (#" + control.Name.ToLower()) doesn't work the same way in DevOps and GitHub
                        tableRows.Add(new MdTableRow(new MdLinkSpan(reference.Control.Name + " (" + reference.Control.Screen()?.Name + ")",
                                                            ("screen " + reference.Control.Screen()?.Name + " " + content.filename + ".md").Replace(" ", "-")),
                                                    reference.RuleProperty));
                    }
                    variablesDocument.Root.Add(new MdTable(new MdTableRow("Control", "Property"), tableRows));
                }
            }
            variablesDocument.Root.Add(new MdHeading(content.appVariablesInfo.headerContextVariables, 3));
            foreach (var var in content.appVariablesInfo.contextVariables)
            {
                variablesDocument.Root.Add(new MdHeading(var, 4));
                content.appVariablesInfo.variableCollectionControlReferences.TryGetValue(var, out var references);
                if (references != null)
                {
                    variablesDocument.Root.Add(new MdParagraph(new MdTextSpan("Variable used in:")));
                    var tableRows = new List<MdTableRow>();
                    foreach (var reference in references.OrderBy(o => o.Control.Name).ThenBy(o => o.RuleProperty))
                    {
                        tableRows.Add(new MdTableRow(reference.Control.Name + " (" + reference.Control.Screen()?.Name + ")", reference.RuleProperty));
                    }
                    variablesDocument.Root.Add(new MdTable(new MdTableRow("Control", "Property"), tableRows));
                }
            }
            variablesDocument.Root.Add(new MdHeading(content.appVariablesInfo.headerCollections, 3));
            foreach (var var in content.appVariablesInfo.collections)
            {
                variablesDocument.Root.Add(new MdHeading(var, 4));
                content.appVariablesInfo.variableCollectionControlReferences.TryGetValue(var, out var references);
                if (references != null)
                {
                    variablesDocument.Root.Add(new MdParagraph(new MdTextSpan("Variable used in:")));
                    var tableRows = new List<MdTableRow>();
                    foreach (var reference in references.OrderBy(o => o.Control.Name).ThenBy(o => o.RuleProperty))
                    {
                        tableRows.Add(new MdTableRow(reference.Control.Name + " (" + reference.Control.Screen()?.Name + ")", reference.RuleProperty));
                    }
                    variablesDocument.Root.Add(new MdTable(new MdTableRow("Control", "Property"), tableRows));
                }
            }
        }

        private void addAppControlsOverview()
        {
            controlsDocument.Root.Add(new MdHeading(content.appControls.headerOverview, 2));
            controlsDocument.Root.Add(new MdParagraph(new MdTextSpan(content.appControls.infoTextScreens)));
            controlsDocument.Root.Add(new MdParagraph(new MdTextSpan(content.appControls.infoTextControls)));
            foreach (var control in content.appControls.controls.Where(o => o.Type != "appinfo"))
            {
                controlsDocument.Root.Add(new MdHeading(new MdLinkSpan("Screen: " + control.Name, ("screen " + control.Name + " " + content.filename + ".md").Replace(" ", "-")), 3));
                controlsDocument.Root.Add(CreateControlList(control));
            }
            controlsDocument.Root.Add(new MdHeading(content.appControls.headerScreenNavigation, 2));
            controlsDocument.Root.Add(new MdParagraph(new MdTextSpan(content.appControls.infoTextScreenNavigation)));
            controlsDocument.Root.Add(new MdParagraph(new MdImageSpan(content.appControls.headerScreenNavigation, content.appControls.imageScreenNavigation + ".svg")));
        }

        private MdBulletList CreateControlList(ControlEntity control)
        {
            var svgDocument = SvgDocument.FromSvg<SvgDocument>(AppControlIcons.GetControlIcon(control.Type));
            //generating the PNG from the SVG with a width of 16px because some SVGs are huge and downscaled, thus can't be shown directly
            using (var bitmap = svgDocument.Draw(16, 0))
            {
                bitmap?.Save(content.folderPath + @"resources\" + control.Type + ".png");
            }
            //link to the screen instead of the control directly for the moment, as the directly generated anchor link (#" + control.Name.ToLower()) doesn't work the same way in DevOps and GitHub
            var list = new MdBulletList(){
                                     new MdListItem(new MdLinkSpan(
                                            new MdCompositeSpan(
                                                new MdImageSpan(control.Type, "resources/"+control.Type+".png"),
                                                new MdTextSpan(" "+control.Name))
                                        ,("screen " + control.Screen().Name + " " + content.filename + ".md").Replace(" ", "-")))};

            foreach (var child in control.Children.OrderBy(o => o.Name).ToList())
            {
                list.Add(new MdListItem(CreateControlList(child)));
            }
            return list;
        }

        private void addDetailedAppControls()
        {
            foreach (var screen in content.appControls.controls.Where(o => o.Type == "screen").OrderBy(o => o.Name).ToList())
            {
                screenDocuments.TryGetValue(screen.Name, out var screenDoc);
                screenDoc.Root.Add(new MdHeading(screen.Name, 2));
                addAppControlsTable(screen, screenDoc);
                foreach (var control in content.appControls.allControls.Where(o => o.Type != "appinfo" && o.Type != "screen" && screen.Equals(o.Screen())).OrderBy(o => o.Name).ToList())
                {
                    screenDoc.Root.Add(new MdHeading(control.Name, 2));
                    addAppControlsTable(control, screenDoc);
                }
            }
        }

        private void addAppControlsTable(ControlEntity control, MdDocument screenDoc)
        {
            var defaultEntity = DefaultChangeHelper.GetEntityDefaults(control.Type);
            var tableRows = new List<MdTableRow>();
            var svgDocument = SvgDocument.FromSvg<SvgDocument>(AppControlIcons.GetControlIcon(control.Type));
            //generating the PNG from the SVG with a width of 16px because some SVGs are huge and downscaled, thus can't be shown directly
            using (var bitmap = svgDocument.Draw(16, 0))
            {
                bitmap?.Save(content.folderPath + @"resources\" + control.Type + ".png");
            }
            tableRows.Add(new MdTableRow(new MdImageSpan(control.Type, "resources/" + control.Type + ".png"), new MdTextSpan("Type: " + control.Type)));

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
                            if (tableRows.Count > 0)
                            {
                                screenDoc.Root.Add(new MdTable(new MdTableRow("Property", "Value"), tableRows));
                                tableRows = new List<MdTableRow>();
                            }
                            category = rule.Category;
                            screenDoc.Root.Add(new MdHeading(category, 3));
                        }
                        if (rule.InvariantScript.StartsWith("RGBA("))
                        {
                            tableRows.Add(CreateColorTable(rule, defaultValue));
                        }
                        else
                        {
                            tableRows.Add(CreateRowForControlProperty(rule, defaultValue));
                        }
                    }
                }
            }
            if (tableRows.Count > 0)
            {
                screenDoc.Root.Add(new MdTable(new MdTableRow("Property", "Value"), tableRows));
            }
            //Colour properties
            tableRows = new List<MdTableRow>();
            screenDoc.Root.Add(new MdHeading("Color Properties", 3));
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
                        if (rule.InvariantScript.StartsWith("RGBA("))
                        {
                            tableRows.Add(CreateColorTable(rule, defaultValue));
                        }
                        else
                        {
                            tableRows.Add(new MdTableRow(rule.Property, rule.InvariantScript));
                        }
                    }
                }
            }
            if (tableRows.Count > 0)
            {
                screenDoc.Root.Add(new MdTable(new MdTableRow("Property", "Value"), tableRows));
            }

            tableRows = new List<MdTableRow>();
            screenDoc.Root.Add(new MdHeading("Child & Parent Controls", 3));

            foreach (var childControl in control.Children)
            {
                tableRows.Add(new MdTableRow("Child Control", childControl.Name));
            }
            if (control.Parent != null)
            {
                tableRows.Add(new MdTableRow("Parent Control", control.Parent.Name));
            }
            if (tableRows.Count > 0)
            {
                screenDoc.Root.Add(new MdTable(new MdTableRow("Property", "Value"), tableRows));
            }
        }

        private string CreateChangedDefaultColourRow(string firstColumnElement, string secondColumnElement)
        {
            var tr = new StringBuilder("<tr>");
            tr.Append("<td style=\"width:50%; background-color:#ccffcc; color:black;\">")
                .Append(firstColumnElement)
                .Append("</td><td style=\"width:50%; background-color:#ffcccc; color:black;\">")
                .Append(secondColumnElement)
                .Append("</td></tr>");
            return tr.ToString();
        }

        private MdTableRow CreateColorTable(Rule rule, string defaultValue)
        {
            var colourTable = new StringBuilder("<table border=\"0\">");
            colourTable.Append("<tr><td>").Append(rule.InvariantScript).Append("</td></tr>");
            var colour = ColourHelper.ParseColor(rule.InvariantScript[..(rule.InvariantScript.IndexOf(')') + 1)]);
            if (!String.IsNullOrEmpty(colour))
            {
                colourTable.Append("<tr><td style=\"background-color:").Append(colour).Append("\"></td></tr>");
            }
            colourTable.Append("</table>");
            if (showDefaults && defaultValue != rule.InvariantScript && !content.appControls.controlPropertiesToSkip.Contains(rule.Property))
            {
                var defaultTable = new StringBuilder("<table border=\"0\">");
                defaultTable.Append("<tr><td>").Append(defaultValue).Append("</td></tr>");
                var defaultColour = ColourHelper.ParseColor(defaultValue);
                if (!String.IsNullOrEmpty(defaultColour))
                {
                    defaultTable.Append("<tr><td style=\"background-color:").Append(defaultColour).Append("\"></td></tr>");
                }
                defaultTable.Append("</table>");
                var changesTable = new StringBuilder("<table border=\"0\">");
                changesTable.Append(CreateChangedDefaultColourRow(colourTable.ToString(), defaultTable.ToString()));
                return new MdTableRow(rule.Property, new MdRawMarkdownSpan(changesTable.Append("</table>").ToString()));
            }
            return new MdTableRow(rule.Property, new MdRawMarkdownSpan(colourTable.ToString()));
        }

        private MdTableRow CreateRowForControlProperty(Rule rule, string defaultValue)
        {
            if (showDefaults && defaultValue != rule.InvariantScript && !content.appControls.controlPropertiesToSkip.Contains(rule.Property))
            {
                var table = new StringBuilder("<table border=\"0\">");
                table.Append("<tr><td style=\"background-color:#ccffcc; width:50%;\">")
                     .Append(rule.InvariantScript)
                     .Append("<td style=\"background-color:#ffcccc; width:50%;\">").Append(defaultValue).Append("</td></tr></table>");
                return new MdTableRow(rule.Property, new MdRawMarkdownSpan(table.ToString()));
            }
            return new MdTableRow(rule.Property, rule.InvariantScript);
        }

        private void addAppDataSources()
        {
            dataSourcesDocument.Root.Add(new MdHeading(content.appDataSources.header, 2));
            dataSourcesDocument.Root.Add(new MdParagraph(new MdTextSpan(content.appDataSources.infoText)));

            foreach (var datasource in content.appDataSources.dataSources)
            {
                if (!datasource.isSampleDataSource() || documentSampleData)
                {
                    dataSourcesDocument.Root.Add(new MdHeading(new MdLinkSpan(datasource.Name, ("datasource " + datasource.Name + " " + content.filename + ".md").Replace(" ", "-")), 3));
                    datasourcesDocuments.TryGetValue(datasource.Name, out var dataSourceDocument);
                    dataSourceDocument.Root.Add(new MdHeading(datasource.Name, 3));
                    var tableRows = new List<MdTableRow>
                {
                    new MdTableRow("Name", datasource.Name),
                    new MdTableRow("Type", datasource.Type)
                };
                    dataSourceDocument.Root.Add(new MdTable(new MdTableRow("Property", "Value"), tableRows));
                    tableRows = new List<MdTableRow>();
                    dataSourceDocument.Root.Add(new MdHeading("DataSource Properties", 4));
                    foreach (var expression in datasource.Properties.OrderBy(o => o.expressionOperator))
                    {
                        if (expression.expressionOperator != "TableDefinition")
                        {
                            if (expression.expressionOperands.Count > 1)
                            {
                                tableRows.Add(new MdTableRow(expression.expressionOperator, new MdRawMarkdownSpan(AddExpressionDetails(new List<Expression> { expression }))));
                            }
                            else
                            {
                                tableRows.Add(new MdTableRow(expression.expressionOperator, (expression.expressionOperands.Count > 0) ? expression.expressionOperands[0].ToString() : ""));
                            }
                        }
                        else
                        {
                            //todo document the table definition ? probably make it configurable
                            // removed it for the moment as it resulted in very large files with likely little value
                        }
                    }
                    dataSourceDocument.Root.Add(new MdTable(new MdTableRow("Property", "Value"), tableRows));
                }
            }
        }

        private void addAppResources()
        {
            Directory.CreateDirectory(content.folderPath + "resources");
            resourcesDocument.Root.Add(new MdHeading(content.appResources.header, 2));
            resourcesDocument.Root.Add(new MdParagraph(new MdTextSpan(content.appResources.infoText)));
            foreach (var resource in content.appResources.resources)
            {
                if (!resource.isSampleResource())
                {
                    resourcesDocument.Root.Add(new MdHeading(resource.Name, 3));
                    var tableRows = new List<MdTableRow>
                {
                    new MdTableRow("Name", resource.Name),
                    new MdTableRow("Content", resource.Content),
                    new MdTableRow("Resource Kind", resource.ResourceKind)
                };
                    if (resource.ResourceKind == "LocalFile")
                    {
                        if (content.ResourceStreams.TryGetValue(resource.Name, out var resourceStream))
                        {
                            var fileName = resource.Properties.First(o => o.expressionOperator == "FileName");
                            using Stream streamToWriteTo = File.Open(content.folderPath + @"resources\" + fileName.expressionOperands[0].ToString(), FileMode.Create);

                            resourceStream.Position = 0;
                            resourceStream.CopyTo(streamToWriteTo);
                            var imageWidth = 400;
                            if (!fileName.expressionOperands[0].ToString().EndsWith("svg", StringComparison.OrdinalIgnoreCase))
                            {
                                using var image = Image.FromStream(resourceStream, false, false);
                                imageWidth = (image.Width > 400) ? 400 : image.Width;
                            }
                            //todo consider showing a resized copy of the image if it is wider than 400px
                            tableRows.Add(new MdTableRow("Resource Preview", new MdImageSpan(resource.Name, "resources/" + fileName.expressionOperands[0].ToString())));
                        }
                    }
                    foreach (var expression in resource.Properties.OrderBy(o => o.expressionOperator))
                    {
                        tableRows.Add(new MdTableRow(expression.expressionOperator, expression.expressionOperands?[0].ToString()));
                    }

                    resourcesDocument.Root.Add(new MdTable(new MdTableRow("Property", "Value"), tableRows));
                }
            }
        }
    }
}
