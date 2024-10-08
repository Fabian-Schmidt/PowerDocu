using Grynwald.MarkdownGenerator;
using PowerDocu.Common;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace PowerDocu.FlowDocumenter
{
    class FlowMarkdownBuilder : MarkdownBuilder
    {
        private readonly string mainDocumentFileName, connectionsDocumentFileName, variablesDocumentFileName, triggerActionsFileName;
        private readonly MdDocument mainDocument, connectionsDocument, variablesDocument, triggerActionsDocument;
        private readonly FlowDocumentationContent content;
        private readonly DocumentSet<MdDocument> set;
        private MdTable metadataTable;

        public FlowMarkdownBuilder(FlowDocumentationContent contentdocumentation)
        {
            content = contentdocumentation;
            Directory.CreateDirectory(content.folderPath);
            mainDocumentFileName = ("index" + ".md").Replace(" ", "-");
            connectionsDocumentFileName = ("connections" + ".md").Replace(" ", "-");
            variablesDocumentFileName = ("variables" + ".md").Replace(" ", "-");
            triggerActionsFileName = ("triggersactions" + ".md").Replace(" ", " - ");
            set = new DocumentSet<MdDocument>();
            mainDocument = set.CreateMdDocument(mainDocumentFileName);
            connectionsDocument = set.CreateMdDocument(connectionsDocumentFileName);
            variablesDocument = set.CreateMdDocument(variablesDocumentFileName);
            triggerActionsDocument = set.CreateMdDocument(triggerActionsFileName);

            //add all the relevant content
            addFlowMetadata();
            addFlowOverview();
            addConnectionReferenceInfo();
            addTriggerInfo();
            addVariablesInfo();
            addActionInfo();
            addFlowDetails();
            set.Save(content.folderPath);
            NotificationHelper.SendNotification("Created Markdown documentation for " + content.metadata.Name);
        }

        private MdBulletList getNavigationLinks(bool topLevel = true)
        {
            var navItems = new MdListItem[] {
                new MdListItem(new MdLinkSpan("Overview", topLevel ? mainDocumentFileName : "../" + mainDocumentFileName)),
                new MdListItem(new MdLinkSpan("Connection References",topLevel ? connectionsDocumentFileName : "../" + connectionsDocumentFileName)),
                new MdListItem(new MdLinkSpan("Variables", topLevel ? variablesDocumentFileName : "../" + variablesDocumentFileName)),
                new MdListItem(new MdLinkSpan("Triggers & Actions", topLevel ? triggerActionsFileName : "../" + triggerActionsFileName))
                };
            return new MdBulletList(navItems);
        }

        private void addFlowMetadata()
        {
            var tableRows = new List<MdTableRow>();
            foreach (var kvp in content.metadata.metadataTable)
            {
                tableRows.Add(new MdTableRow(kvp.Key, kvp.Value));
            }
            metadataTable = new MdTable(new MdTableRow(new List<string>() { "Flow Name", content.metadata.Name }), tableRows);
            // prepare the common sections for all documents
            foreach (var doc in set.Documents)
            {
                doc.Root.Add(new MdHeading(content.metadata.header, 1));
                doc.Root.Add(metadataTable);
                doc.Root.Add(getNavigationLinks());
            }
        }

        private void addFlowOverview()
        {
            mainDocument.Root.Add(new MdHeading(content.overview.header, 2));
            mainDocument.Root.Add(new MdParagraph(new MdTextSpan(content.overview.infoText)));
            mainDocument.Root.Add(new MdParagraph(new MdImageSpan("Flow Overview Diagram", content.overview.svgFile)));
        }

        private void addConnectionReferenceInfo()
        {
            connectionsDocument.Root.Add(new MdHeading(content.connectionReferences.header, 2));
            connectionsDocument.Root.Add(new MdParagraph(new MdTextSpan(content.connectionReferences.infoText)));
            foreach (var kvp in content.connectionReferences.connectionTable)
            {
                var connectorUniqueName = kvp.Key;
                var connectorIcon = ConnectorHelper.getConnectorIcon(connectorUniqueName);
                connectionsDocument.Root.Add(new MdHeading((connectorIcon != null) ? connectorIcon.Name : connectorUniqueName, 3));

                var tableRows = new List<MdTableRow>();
                foreach (var kvp2 in kvp.Value)
                {
                    tableRows.Add(new MdTableRow(kvp2.Key, kvp2.Value));
                }
                var table = new MdTable(new MdTableRow(new MdTextSpan("Connector"), getConnectorNameAndIcon(connectorUniqueName, "https://docs.microsoft.com/connectors/" + connectorUniqueName)), tableRows);
                connectionsDocument.Root.Add(table);
            }
        }

        private MdLinkSpan getConnectorNameAndIcon(string connectorUniqueName, string url, bool fromSubfolder = false)
        {
            var connectorIcon = ConnectorHelper.getConnectorIcon(connectorUniqueName);
            if (ConnectorHelper.getConnectorIconFile(connectorUniqueName) != "")
            {
                return new MdLinkSpan(new MdCompositeSpan(
                                            new MdImageSpan(connectorUniqueName, (fromSubfolder ? "../" : "") + connectorUniqueName + "32.png"),
                                            new MdTextSpan(" " + ((connectorIcon != null) ? connectorIcon.Name : connectorUniqueName))
                                        ), url);
            }
            else
            {
                return new MdLinkSpan((connectorIcon != null) ? connectorIcon.Name : connectorUniqueName, url);
            }
        }

        private void addVariablesInfo()
        {
            variablesDocument.Root.Add(new MdHeading(content.variables.header, 2));
            foreach (var kvp in content.variables.variablesTable)
            {
                variablesDocument.Root.Add(new MdHeading(kvp.Key, 3));

                var tableRows = new List<MdTableRow>();

                foreach (var kvp2 in kvp.Value)
                {
                    if (!kvp2.Key.Equals("Initial Value"))
                    {
                        tableRows.Add(new MdTableRow(kvp2.Key, new MdCodeSpan(kvp2.Value)));
                    }
                }
                var table = new MdTable(new MdTableRow(new List<string>() { "Property", "Value" }), tableRows);
                variablesDocument.Root.Add(table);
                if (kvp.Value.ContainsKey("Initial Value"))
                {
                    tableRows = new List<MdTableRow>();
                    content.variables.initialValTable.TryGetValue(kvp.Key, out var initialValues);
                    foreach (var initialVal in initialValues)
                    {
                        tableRows.Add(new MdTableRow(new MdCodeSpan(initialVal.Key), new MdCodeSpan(initialVal.Value)));
                    }
                    if (tableRows.Count > 0)
                    {
                        table = new MdTable(new MdTableRow(new List<string>() { "Variable Property", "Initial Value" }), tableRows);
                        variablesDocument.Root.Add(table);
                    }
                }
                content.variables.referencesTable.TryGetValue(kvp.Key, out var references);
                if (references?.Count > 0)
                {
                    tableRows = new List<MdTableRow>();
                    foreach (var action in references.OrderBy(o => o.Name).ToList())
                    {
                        tableRows.Add(new MdTableRow(new MdLinkSpan(action.Name, "actions/" + getLinkFromAction(action.Name))));
                    }
                    table = new MdTable(new MdTableRow(new List<string>() { "Variable Used In" }), tableRows);
                    variablesDocument.Root.Add(table);
                }
            }
        }

        private void addTriggerInfo()
        {
            var triggerDocFileName = ("actions/" + content.trigger.header + ".md").Replace(" ", "-");
            var triggerDoc = set.CreateMdDocument(triggerDocFileName);
            triggerDoc.Root.Add(new MdHeading(content.metadata.header, 1));
            triggerDoc.Root.Add(metadataTable);
            triggerDoc.Root.Add(getNavigationLinks(false));
            triggerDoc.Root.Add(new MdHeading(content.trigger.header, 2));
            triggerActionsDocument.Root.Add(new MdHeading(content.trigger.header, 2));
            triggerActionsDocument.Root.Add(new MdBulletList(new MdListItem(new MdLinkSpan(content.trigger.header, triggerDocFileName))));

            var tableRows = new List<MdTableRow>();
            foreach (var kvp in content.trigger.triggerTable)
            {
                if (kvp.Value == "mergedrow")
                {
                    tableRows.Add(new MdTableRow(new MdCompositeSpan(kvp.Key)));
                }
                else
                {
                    tableRows.Add(new MdTableRow(kvp.Key, new MdCodeSpan(kvp.Value)));
                }
            }
            var table = new MdTable(new MdTableRow(new List<string>() { "Property", "Value" }), tableRows);
            triggerDoc.Root.Add(table);
            if (content.trigger.inputs?.Count > 0)
            {
                triggerDoc.Root.Add(new MdHeading(content.trigger.inputsHeader, 3));
                triggerDoc.Root.Add(new MdParagraph(new MdRawMarkdownSpan(AddExpressionDetails(content.trigger.inputs))));
            }
            if (content.trigger.triggerProperties?.Count > 0)
            {
                triggerDoc.Root.Add(new MdHeading("Other Trigger Properties", 3));
                triggerDoc.Root.Add(new MdParagraph(new MdRawMarkdownSpan(AddExpressionDetails(content.trigger.triggerProperties))));
            }
        }

        private void addActionInfo()
        {
            var triggerActionsLinks = new List<MdListItem>();
            var actionNodesList = content.actions.actionNodesList;
            triggerActionsDocument.Root.Add(new MdHeading(content.actions.header, 2));
            triggerActionsDocument.Root.Add(new MdParagraph(new MdTextSpan(content.actions.infoText)));

            foreach (var action in actionNodesList)
            {
                var actionDocFileName = ("actions/" + CharsetHelper.GetSafeName(action.Name) + ".md").Replace(" ", "-");
                triggerActionsLinks.Add(new MdListItem(new MdLinkSpan(action.Name, actionDocFileName)));
                var actionsDoc = set.CreateMdDocument(actionDocFileName);
                actionsDoc.Root.Add(new MdHeading(content.metadata.header, 1));
                actionsDoc.Root.Add(metadataTable);
                actionsDoc.Root.Add(getNavigationLinks(false));

                actionsDoc.Root.Add(new MdHeading(action.Name, 2));
                var tableRows = new List<MdTableRow>
                {
                    new MdTableRow("Name", action.Name),
                    new MdTableRow("Type", action.Type)
                };
                if (!String.IsNullOrEmpty(action.Description))
                {
                    tableRows.Add(new MdTableRow("Description / Note", action.Description));
                }
                if (!String.IsNullOrEmpty(action.Connection))
                {
                    tableRows.Add(new MdTableRow("Connection",
                                                getConnectorNameAndIcon(action.Connection, "https://docs.microsoft.com/connectors/" + action.Connection, true)));
                }

                //TODO provide more details, such as information about subaction, subsequent actions, switch actions, ...
                if (action.actionExpression != null || !String.IsNullOrEmpty(action.Expression))
                {
                    tableRows.Add(new MdTableRow("Expression", new MdCodeSpan((action.actionExpression != null) ? AddExpressionTable(action.actionExpression).ToString() : action.Expression)));
                }
                var table = new MdTable(new MdTableRow(new List<string>() { "Property", "Value" }), tableRows);
                actionsDoc.Root.Add(table);
                if (action.actionInputs.Count > 0 || !String.IsNullOrEmpty(action.Inputs))
                {
                    tableRows = new List<MdTableRow>();
                    actionsDoc.Root.Add(new MdHeading("Inputs", 3));
                    if (action.actionInputs.Count > 0)
                    {
                        foreach (var actionInput in action.actionInputs)
                        {
                            var operandsCell = new StringBuilder();
                            if (actionInput.expressionOperands.Count > 1)
                            {
                                var operandsTable = new StringBuilder("<table>");
                                foreach (var actionInputOperand in actionInput.expressionOperands)
                                {
                                    if (actionInputOperand.GetType() == typeof(Expression))
                                    {
                                        operandsTable.Append(AddExpressionTable((Expression)actionInputOperand, false));
                                    }
                                    else
                                    {
                                        operandsTable.Append("<tr><td>").Append(actionInputOperand.ToString()).Append("</td></tr>");
                                    }
                                }
                                operandsCell.Append(operandsTable.Append("</table>"));
                            }
                            else
                            {
                                if (actionInput.expressionOperands.Count == 0)
                                {
                                    operandsCell.Append("<tr><td></td></tr>");
                                }
                                else
                                {
                                    if (actionInput.expressionOperands[0]?.GetType() == typeof(Expression))
                                    {
                                        operandsCell.Append(AddExpressionTable((Expression)actionInput.expressionOperands[0]));
                                    }
                                    else if (actionInput.expressionOperands[0]?.GetType() == typeof(List<object>))
                                    {
                                        operandsCell.Append("<table>");
                                        foreach (var obj in (List<object>)actionInput.expressionOperands[0])
                                        {
                                            if (obj.GetType().Equals(typeof(Expression)))
                                            {
                                                operandsCell.Append(AddExpressionTable((Expression)obj, false));
                                            }
                                            else if (obj.GetType().Equals(typeof(List<object>)))
                                            {
                                                foreach (var o in (List<object>)obj)
                                                {
                                                    operandsCell.Append(AddExpressionTable((Expression)o, false));
                                                }
                                            }
                                            else
                                            {
                                                var s = "";
                                            }
                                        }
                                        operandsCell.Append("</table>");
                                    }
                                    else
                                    {
                                        operandsCell.Append(actionInput.expressionOperands[0]?.ToString());
                                    }
                                }
                            }
                            tableRows.Add(new MdTableRow(actionInput.expressionOperator, new MdCodeSpan(operandsCell.ToString())));
                        }
                    }
                    if (!String.IsNullOrEmpty(action.Inputs))
                    {
                        tableRows.Add(new MdTableRow("Value", new MdCodeSpan(action.Inputs)));
                    }
                    table = new MdTable(new MdTableRow(new List<string>() { "Property", "Value" }), tableRows);
                    actionsDoc.Root.Add(table);
                }

                if (action.Subactions.Count > 0 || action.Elseactions.Count > 0)
                {
                    if (action.Subactions.Count > 0)
                    {
                        tableRows = new List<MdTableRow>();
                        actionsDoc.Root.Add(new MdHeading(action.Type == "Switch" ? "Switch Actions" : "Subactions", 3));
                        if (action.Type == "Switch")
                        {
                            foreach (var subaction in action.Subactions)
                            {
                                if (action.switchRelationship.TryGetValue(subaction, out var switchValue))
                                {
                                    tableRows.Add(new MdTableRow(switchValue, new MdLinkSpan(subaction.Name, getLinkFromAction(subaction.Name))));
                                }
                            }
                            table = new MdTable(new MdTableRow(new List<string>() { "Case Values", "Action" }), tableRows);
                            actionsDoc.Root.Add(table);
                        }
                        else
                        {
                            foreach (var subaction in action.Subactions)
                            {
                                //adding a link to the subaction's section in the documentation
                                tableRows.Add(new MdTableRow(new MdLinkSpan(subaction.Name, getLinkFromAction(subaction.Name))));
                            }
                            table = new MdTable(new MdTableRow(new List<string>() { "Action" }), tableRows);
                            actionsDoc.Root.Add(table);
                        }
                    }
                    if (action.Elseactions.Count > 0)
                    {
                        tableRows = new List<MdTableRow>();
                        actionsDoc.Root.Add(new MdHeading("Elseactions", 3));
                        foreach (var elseaction in action.Elseactions)
                        {
                            //adding a link to the elseaction's section
                            tableRows.Add(new MdTableRow(new MdLinkSpan(elseaction.Name, getLinkFromAction(elseaction.Name))));
                        }
                        table = new MdTable(new MdTableRow(new List<string>() { "Elseactions" }), tableRows);
                        actionsDoc.Root.Add(table);
                    }
                }
                if (action.Neighbours.Count > 0)
                {
                    actionsDoc.Root.Add(new MdHeading("Next Action(s) Conditions", 3));
                    tableRows = new List<MdTableRow>();
                    foreach (var nextAction in action.Neighbours)
                    {
                        var raConditions = action.nodeRunAfterConditions[nextAction];
                        tableRows.Add(new MdTableRow(new MdLinkSpan(nextAction.Name + " [" + string.Join(", ", raConditions) + "]", getLinkFromAction(nextAction.Name))));
                    }
                    table = new MdTable(new MdTableRow(new List<string>() { "Next Action" }), tableRows);
                    actionsDoc.Root.Add(table);
                }
            }
            triggerActionsDocument.Root.Add(new MdBulletList(triggerActionsLinks));
        }

        private void addFlowDetails()
        {
            mainDocument.Root.Add(new MdHeading(content.details.header, 2));
            mainDocument.Root.Add(new MdParagraph(new MdTextSpan(content.details.infoText)));
            mainDocument.Root.Add(new MdParagraph(new MdImageSpan(content.details.header, content.details.imageFileName + ".svg")));
        }

        private string getLinkFromAction(string name)
        {
            return (CharsetHelper.GetSafeName(name) + ".md").Replace(" ", "-");
        }
    }
}