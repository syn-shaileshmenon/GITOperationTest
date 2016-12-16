// <copyright file="DocumentManager.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>

namespace Mkl.WebTeam.DocumentGenerator
{
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.Drawing;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Web;
    using Aspose.Words;
    using Aspose.Words.Drawing;
    using Aspose.Words.Fields;
    using Aspose.Words.Markup;
    using Aspose.Words.Tables;
    using DecisionModel.Models.Policy;
    using Mkl.WebTeam.DocumentGenerator.Helpers;
    using Mkl.WebTeam.SubmissionShared.Enumerations;
    using Mkl.WebTeam.WebCore2;

    /// <summary>
    /// This is an example of a class to utilize the Aspose library for WORD Doc interactions
    /// </summary>
    /// <seealso cref="Mkl.WebTeam.DocumentGenerator.IDocumentManager" />
    public class DocumentManager : IDocumentManager
    {
        /// <summary>
        /// The transaction identifier
        /// </summary>
        private string transactionId = string.Empty;

        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentManager" /> class.
        /// </summary>
        /// <param name="docPath">The document path.</param>
        /// <param name="id">Transaction Id</param>
        /// <exception cref="System.ArgumentException">Specified document does not exist. Please verify path is for a valid document.;docPath</exception>
        public DocumentManager(string docPath, string id)
        {
            this.transactionId = id;
            if (File.Exists(docPath))
            {
                this.OriginalFileName = Path.GetFileNameWithoutExtension(docPath);
                Stream readStream = File.OpenRead(docPath);
                this.WordDoc = new Aspose.Words.Document(readStream);
                readStream.Close();
            }
            else
            {
                LogEvent.Log.Error("Transaction: " + id + ": Specified document " + docPath + " does not exist. Please verify path is for a valid document.");
                LogEvent.TraceLog.Error("Transaction: " + id + ": Specified document " + docPath + " does not exist.");
                throw new ArgumentException("Transaction: " + id + ": Specified document " + docPath + " does not exist. Please verify path is for a valid document.");
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentManager" /> class.
        /// Initializes a new instance of the DocumentManager class with a Steam object passed
        /// </summary>
        /// <param name="docStream">Stream containing the document to load</param>
        /// <param name="id">Transaction id</param>
        public DocumentManager(Stream docStream, string id)
        {
            this.transactionId = id;
            if (docStream != null)
            {
                docStream.Seek(0, SeekOrigin.Begin);
                this.WordDoc = new Aspose.Words.Document(docStream);
                docStream.Close();
            }
            else
            {
                LogEvent.Log.Error("Transaction: " + id + ": Passed document stream is invalid or is not readable");
                LogEvent.TraceLog.Error("Transaction: " + id + ": Passed document stream is invalid or is not readable");
                throw new ArgumentException("Transaction: " + id + ": Passed document stream is invalid or is not readable");
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentManager" /> class. merging all docs in the docPath array and adding them to.
        /// memory stream to work with
        /// </summary>
        /// <param name="docPath">Array of document paths.</param>
        /// <param name="id">Transaction Id</param>
        /// <exception cref="System.ArgumentException">Specified document does not exist. Please verify path is for a valid document.;docPath</exception>
        public DocumentManager(string[] docPath, string id)
        {
            this.transactionId = id;
            this.OriginalFileName = Path.GetFileNameWithoutExtension(docPath[0]);
            Stream readStream = File.OpenRead(docPath[0]);
            this.WordDoc = new Aspose.Words.Document(readStream);
            readStream.Close();

            // start with the 2nd doc appending them to the first
            for (int i = 1; i < docPath.Length; i++)
            {
                if (File.Exists(docPath[i]))
                {
                    Aspose.Words.Document srcDoc = new Aspose.Words.Document(docPath[i]);
                    srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.NewPage;
                    srcDoc.FirstSection.PageSetup.RestartPageNumbering = false;
                    this.WordDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
                }
                else
                {
                    LogEvent.Log.Error("Transaction: " + id + ": Specified document " + docPath[i] + " does not exist. Please verify path is for a valid document.");
                    LogEvent.TraceLog.Error("Transaction: " + id + ": Specified document " + docPath[i] + " does not exist.");
                    throw new ArgumentException("Transaction: " + id + ": Specified document " + docPath[i] + " does not exist. Please verify path is for a valid document.");
                }
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentManager" /> class.
        /// </summary>
        /// <param name="doc">The document.</param>
        public DocumentManager(Aspose.Words.Document doc)
        {
            this.WordDoc = doc;
        }

        /// <summary>
        /// File format types
        /// </summary>
        public enum FileFormat
        {
            /// <summary>
            /// The docx
            /// </summary>
            DOCX,

            /// <summary>
            /// The document
            /// </summary>
            DOC,

            /// <summary>
            /// The PDF
            /// </summary>
            PDF,

            /// <summary>
            /// The PNG
            /// </summary>
            PNG,

            /// <summary>
            /// The TIFF
            /// </summary>
            TIFF,

            /// <summary>
            /// The BMP
            /// </summary>
            BMP,

            /// <summary>
            /// The JPEG
            /// </summary>
            JPEG,

            /// <summary>
            /// The none
            /// </summary>
            None
        }

        /// <summary>
        /// Gets the document.
        /// </summary>
        /// <value>The document.</value>
        public Aspose.Words.Document Document
        {
            get { return this.WordDoc; }
        }

        /// <summary>
        /// Gets or sets the name of the original file.
        /// </summary>
        /// <value>The name of the original file.</value>
        public string OriginalFileName { get; set; }

        /// <summary>
        /// Gets or sets the save name with ext.
        /// </summary>
        /// <value>The save name with ext.</value>
        public string SaveNameWithExt { get; set; }

        /// <summary>
        /// Gets or sets the quantity order.
        /// </summary>
        /// <value>The quantity order.</value>
        public int QuantityOrder { get; set; }

        /// <summary>
        /// Gets or sets the form edition date.
        /// </summary>
        /// <value>The edition date.</value>
        public DateTime EditionDate { get; set; }

        /// <summary>
        /// Gets or sets the normalized form number.
        /// </summary>
        /// <value>The normalized form number.</value>
        public string FormNormalizedNumber { get; set; }

        /// <summary>
        /// Gets or sets the word document.
        /// </summary>
        /// <value>The word document.</value>
        protected Aspose.Words.Document WordDoc { get; set; }

        /// <summary>
        /// Appends the specified dest.
        /// run this to create the document that needs to be run through the population process????
        /// </summary>
        /// <param name="dest">The dest.</param>
        /// <param name="source">The source.</param>
        public void Append(string dest, string source)
        {
            var dstDoc = new Aspose.Words.Document(dest);
            var srcDoc = new Aspose.Words.Document(source);
            srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.NewPage;
            srcDoc.FirstSection.PageSetup.RestartPageNumbering = false;
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
        }

        /// <summary>
        /// Appends a document to the one being managed and optionally inserts a page break before the new document.
        /// </summary>
        /// <param name="document">The document.</param>
        /// <param name="insertPageBreakBefore">if set to <c>true</c> [insert page break before].</param>
        /// <param name="restartNumbering">if set to <c>true</c> [restart numbering].</param>
        public void AppendDocument(Aspose.Words.Document document, bool insertPageBreakBefore = true, bool restartNumbering = true)
        {
            var builder = new DocumentBuilder(this.WordDoc);
            builder.MoveToDocumentEnd();

            if (insertPageBreakBefore)
            {
                builder.InsertBreak(BreakType.PageBreak);
                ////builder.InsertBreak(BreakType.SectionBreakNewPage);
            }

            builder.Document.AppendDocument(document, ImportFormatMode.KeepSourceFormatting);
        }

        /// <summary>
        /// Updates the total page number footer field.
        /// </summary>
        public void UpdateTotalPageNumberFooterField()
        {
            var builder = new DocumentBuilder(this.WordDoc);
            builder.Document.UpdatePageLayout();
            var numberOfPages = this.WordDoc.PageCount;
            foreach (Section section in this.WordDoc)
            {
                HeaderFooter footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
                if (footer != null)
                {
                    this.ReplaceHeaderFooterFieldWithResult(builder, footer, FieldType.FieldPage);
                    this.ReplaceHeaderFooterFieldWithResult(builder, footer, FieldType.FieldNumPages);
                }

                footer = section.HeadersFooters[HeaderFooterType.FooterFirst];
                if (footer != null)
                {
                    this.ReplaceHeaderFooterFieldWithResult(builder, footer, FieldType.FieldPage);
                    this.ReplaceHeaderFooterFieldWithResult(builder, footer, FieldType.FieldNumPages);

                    if (numberOfPages < 2 && section.HeadersFooters[HeaderFooterType.FooterEven] != null)
                    {
                        section.HeadersFooters[HeaderFooterType.FooterEven].RemoveAllChildren();
                        foreach (Node node in footer.ChildNodes)
                        {
                            section.HeadersFooters[HeaderFooterType.FooterEven].AppendChild(node.Clone(true));
                        }
                    }
                }

                if (numberOfPages >= 2)
                {
                    footer = section.HeadersFooters[HeaderFooterType.FooterEven];
                    if (footer != null)
                    {
                        this.ReplaceHeaderFooterFieldWithResult(builder, footer, FieldType.FieldPage);
                        this.ReplaceHeaderFooterFieldWithResult(builder, footer, FieldType.FieldNumPages);
                    }
                }
            }
        }

        /// <summary>
        /// Combine all source documents into destination document.
        /// </summary>
        /// <param name="dest">Destination file</param>
        /// <param name="source">Source files</param>
        public void Combine(string dest, string[] source)
        {
            Aspose.Words.Document dstDoc = new Aspose.Words.Document(dest);

            foreach (var sourceDocument in source)
            {
                Aspose.Words.Document srcDoc = new Aspose.Words.Document(sourceDocument);
                srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.NewPage;
                srcDoc.FirstSection.PageSetup.RestartPageNumbering = false;
                dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            }
        }

        /// <summary>
        /// Get a list of StructuredDocumentControls using the Tag property as an ID.
        /// </summary>
        /// <returns>List of controls</returns>
        public List<StructuredDocumentTag> GetControls()
        {
            return this.WordDoc.GetChildNodes(NodeType.StructuredDocumentTag, true).Cast<StructuredDocumentTag>().ToList();
        }

        /// <summary>
        /// Get the bookmarks from a word  template
        /// </summary>
        /// <returns>BookmarkCollection</returns>
        public BookmarkCollection GetControlsBookMarks()
        {
            return this.WordDoc.Range.Bookmarks;
        }

        /// <summary>
        /// Get the bookmark control from a word  template using bookmark name
        /// </summary>
        /// <param name="bookmarkName">The bookmark name.</param>
        /// <returns>Bookmark</returns>
        public Bookmark GetBookMarkByName(string bookmarkName)
        {
            var bookmark = this.WordDoc.Range.Bookmarks[bookmarkName];
            if (bookmark == null)
            {
                foreach (Bookmark mark in this.WordDoc.Range.Bookmarks)
                {
                    if (string.Compare(mark.Name, bookmarkName, StringComparison.InvariantCultureIgnoreCase) == 0)
                    {
                        return mark;
                    }
                }
            }

            return bookmark;
        }

        /// <summary>
        /// Inserts the date.
        /// </summary>
        /// <param name="date">The date.</param>
        /// <param name="control">The control.</param>
        public void InsertDate(DateTime date, StructuredDocumentTag control)
        {
            control.DateDisplayFormat = "MM/dd/yyyy";
            control.DateStorageFormat = SdtDateStorageFormat.Date;
            control.FullDate = date;

            control.RemoveAllChildren();
            Run run = new Run(this.WordDoc);
            run.Text = date.ToString(control.DateDisplayFormat);
            control.AppendChild(run);
        }

        /// <summary>
        /// Insert a hyperlink into Rich Text
        /// </summary>
        /// <param name="link">The link.</param>
        /// <param name="text">The text.</param>
        /// <param name="controlId">The control identifier.</param>
        /// <exception cref="ArgumentException">Value must not be empty.;value
        /// or The given control must exist.;control
        /// or Value must not be empty.;value
        /// or A controlId is required.;controlId</exception>
        public void InsertLink(string link, string text, string controlId)
        {
            if (link == null)
            {
                throw new ArgumentException("Value must not be empty.", "value");
            }

            if (controlId == null)
            {
                throw new ArgumentException("The given control must exist.", "control");
            }

            if (text == null)
            {
                throw new ArgumentException("Value must not be empty.", "value");
            }

            if (string.IsNullOrWhiteSpace(controlId))
            {
                throw new ArgumentException("A controlId is required.", "controlId");
            }

            DocumentBuilder builder = new DocumentBuilder(this.WordDoc);
            foreach (var control in this.GetControlByTag(controlId))
            {
                builder.MoveTo(control);
                builder.InsertHyperlink(text, link, false);
            }
        }

        /// <summary>
        /// Inserts the picture.
        /// </summary>
        /// <param name="imagePath">The image path.</param>
        /// <param name="control">The control.</param>
        public void InsertPicture(string imagePath, StructuredDocumentTag control)
        {
            Shape drawing = control.GetChild(NodeType.Shape, 0, true) as Shape;
            if (drawing.HasImage)
            {
                drawing.AlternativeText = imagePath;
                drawing.ImageData.SetImage(imagePath);
            }
        }

        /// <summary>
        /// Insert a table from a DataTable at the location specified by the controlId.
        /// </summary>
        /// <param name="dataTable">The data table.</param>
        /// <param name="controlId">ID of the control to use as a location for the table</param>
        /// <param name="useHeading">Use the headings from the DataTable. Default is true.</param>
        /// <exception cref="ArgumentException">DataTable must have rows.;dataTable
        /// or
        /// A controlId is required.;controlId</exception>
        public void InsertTable(DataTable dataTable, string controlId, bool useHeading = true)
        {
            if (dataTable == null || dataTable.Rows.Count <= 0)
            {
                throw new ArgumentException("DataTable must have rows.", "dataTable");
            }

            if (string.IsNullOrWhiteSpace(controlId))
            {
                throw new ArgumentException("A controlId is required.", "controlId");
            }

            foreach (var control in this.GetControlByTag(controlId))
            {
                this.InsertTable(dataTable, control, useHeading);
            }
        }

        /// <summary>
        /// Inserts the table.
        /// </summary>
        /// <param name="dataTable">The data table.</param>
        /// <param name="control">The control.</param>
        /// <param name="useHeading">if set to <c>true</c> [use heading].</param>
        /// <exception cref="Exception">Specified control must be a Paragraph or Inline control, or must come after a Paragraph or Inline control.</exception>
        public void InsertTable(DataTable dataTable, StructuredDocumentTag control, bool useHeading = true)
        {
            DocumentBuilder builder = new DocumentBuilder(this.WordDoc);
            if (!(control.PreviousSibling is Paragraph))
            {
                throw new Exception("Specified control must be a Paragraph or Inline control, or must come after a Paragraph or Inline control.");
            }

            builder.MoveTo(control.PreviousSibling);
            builder.InsertBreak(BreakType.ParagraphBreak);
            builder.StartBookmark(control.Tag);
            builder.EndBookmark(control.Tag);
            Table table = this.BuildTableFromDataTable(dataTable, builder, useHeading);
        }

        /// <summary>
        /// replace bookmark text
        /// </summary>
        /// <param name="placeHolder">placeHolder</param>
        /// <param name="replaceText">replaceText</param>
        public void InsertFieldText(string placeHolder, string replaceText)
        {
            FormFieldCollection documentFormFields = this.WordDoc.Range.FormFields;
            FormField docField = documentFormFields[placeHolder];
            if (docField != null && docField.Type == FieldType.FieldFormTextInput)
            {
                DocumentBuilder builder = new DocumentBuilder(this.WordDoc);

                builder.MoveToBookmark(placeHolder);

                builder.Write(replaceText);
                docField.Result = string.Empty;
            }
            else if (docField != null && docField.Type == FieldType.FieldFormDropDown)
            {
                DocumentBuilder builder = new DocumentBuilder(this.WordDoc);

                builder.MoveToBookmark(placeHolder);
                docField.Result = replaceText;
            }
        }

        /// <summary>
        /// Remove bookmark
        /// </summary>
        /// <param name="bookmark">Bookmark</param>
        public void RemoveBookmark(Bookmark bookmark)
        {
            Node currentNode = bookmark.BookmarkStart;

            List<KeyValuePair<string, Node>> childNodes = new List<KeyValuePair<string, Node>>();
            Node nextNode = currentNode.NextPreOrder(this.WordDoc);
            while (nextNode != bookmark.BookmarkEnd)
            {
                if (nextNode.NodeType == NodeType.FormField)
                {
                    var kvp = new KeyValuePair<string, Node>(((FormField)nextNode).Type.ToString(), nextNode);
                    childNodes.Add(kvp);
                }
                else
                {
                    var kvp = new KeyValuePair<string, Node>(nextNode.NodeType.ToString(), nextNode);
                    childNodes.Add(kvp);
                }

                nextNode = nextNode.NextPreOrder(this.WordDoc);
            }

            if (childNodes.Count == 0)
            {
                bookmark.Remove();
            }
            else
            {
                if (childNodes.Contains(new KeyValuePair<string, Node>(FieldType.FieldFormCheckBox.ToString(), null), new KeyComparer()))
                {
                    // We are currently ignoring check boxes so leave them be for now.
                }
                else if (childNodes.Contains(new KeyValuePair<string, Node>(FieldType.FieldFormDropDown.ToString(), null), new KeyComparer()))
                {
                    foreach (var keyValuePair in childNodes.Where(kvp => kvp.Key == FieldType.FieldFormDropDown.ToString()))
                    {
                        var dropDown = keyValuePair.Value as FormField;
                        if (dropDown != null)
                        {
                            dropDown.Remove();
                        }
                    }
                }
                else if (childNodes.Contains(new KeyValuePair<string, Node>(FieldType.FieldFormTextInput.ToString(), null), new KeyComparer()))
                {
                    foreach (var keyValuePair in childNodes.Where(kvp => kvp.Key == FieldType.FieldFormTextInput.ToString()))
                    {
                        var input = keyValuePair.Value as FormField;
                        if (input != null)
                        {
                            input.Remove();
                        }
                    }
                }
                else
                {
                    if (childNodes.Contains(new KeyValuePair<string, Node>(NodeType.Run.ToString(), null), new KeyComparer()))
                    {
                        var lastNode = childNodes.Last(kvp => kvp.Key == NodeType.Run.ToString());
                        ((Run)lastNode.Value).Remove();
                    }
                }
            }
        }

        /// <summary>
        /// Replaces the node value
        /// </summary>
        /// <param name="bookmark">bookmark</param>
        /// <param name="replaceText">replaceText</param>
        public void ReplaceNodeValue(Bookmark bookmark, string replaceText)
        {
            Node currentNode = bookmark.BookmarkStart;

            List<KeyValuePair<string, Node>> childNodes = new List<KeyValuePair<string, Node>>();
            Node nextNode = currentNode.NextPreOrder(this.WordDoc);
            while (nextNode != null &&
                    bookmark.BookmarkEnd.Name != (nextNode as BookmarkEnd)?.Name)
            {
                if (nextNode.NodeType == NodeType.FormField)
                {
                    var kvp = new KeyValuePair<string, Node>(((FormField)nextNode).Type.ToString(), nextNode);
                    childNodes.Add(kvp);
                }
                else
                {
                    var kvp = new KeyValuePair<string, Node>(nextNode.NodeType.ToString(), nextNode);
                    childNodes.Add(kvp);
                }

                nextNode = nextNode.NextPreOrder(this.WordDoc);
            }

            if (childNodes.Count == 0)
            {
                bookmark.Text = replaceText;
            }
            else
            {
                if (childNodes.Contains(new KeyValuePair<string, Node>(FieldType.FieldFormCheckBox.ToString(), null), new KeyComparer()))
                {
                    foreach (var keyValuePair in childNodes.Where(kvp => kvp.Key == FieldType.FieldFormCheckBox.ToString()))
                    {
                        var checkBox = keyValuePair.Value as FormField;
                        if (checkBox != null && !string.IsNullOrWhiteSpace(replaceText))
                        {
                            checkBox.Checked = true;
                        }
                    }
                }
                else if (childNodes.Contains(new KeyValuePair<string, Node>(FieldType.FieldFormDropDown.ToString(), null), new KeyComparer()))
                {
                    foreach (var keyValuePair in childNodes.Where(kvp => kvp.Key == FieldType.FieldFormDropDown.ToString()))
                    {
                        var dropDown = keyValuePair.Value as FormField;
                        if (dropDown != null)
                        {
                            dropDown.DropDownItems.Add(replaceText);
                            dropDown.DropDownSelectedIndex = dropDown.DropDownItems.Count - 1;
                        }
                    }
                }
                else if (childNodes.Contains(new KeyValuePair<string, Node>(FieldType.FieldFormTextInput.ToString(), null), new KeyComparer()))
                {
                    foreach (var keyValuePair in childNodes.Where(kvp => kvp.Key == FieldType.FieldFormTextInput.ToString()))
                    {
                        var input = keyValuePair.Value as FormField;
                        if (input != null)
                        {
                            input.Result = replaceText;
                        }
                    }
                }
                else
                {
                    if (childNodes.Contains(new KeyValuePair<string, Node>(NodeType.Run.ToString(), null), new KeyComparer()))
                    {
                        var lastNode = childNodes.Last(kvp => kvp.Key == NodeType.Run.ToString());
                        ((Run)lastNode.Value).Text = replaceText;
                    }
                }
            }
        }

        /// <summary>
        /// Insert string into Plain Text or markup into Rich Text
        /// </summary>
        /// <param name="value">String of text or Rich Text markup to insert</param>
        /// <param name="control">Plain Text or Rich Text control</param>
        public void InsertText(string value, StructuredDocumentTag control)
        {
            if (control.Level == MarkupLevel.Block)
            {
                DocumentBuilder builder = new DocumentBuilder(this.WordDoc);
                if (control.PreviousSibling is Paragraph)
                {
                    builder.MoveTo(control.PreviousSibling);
                }

                builder.StartBookmark(control.Tag);
                builder.EndBookmark(control.Tag);
                builder.Font.Bold = control.ContentsFont.Bold;
                builder.Font.Name = control.ContentsFont.Name;
                builder.Font.Size = control.ContentsFont.Size;

                // grab the paragraph from the control
                NodeCollection childNodes = control.GetChildNodes(NodeType.Paragraph, false);
                int count = (childNodes.Count < 3) ? childNodes.Count - 1 : 2;
                if (count > -1)
                {
                    Paragraph q = (Paragraph)childNodes[count];

                    // create a new paragraph and passover spacing and tabs to it
                    builder.ParagraphFormat.SpaceAfter = q.ParagraphFormat.SpaceAfter;
                    builder.ParagraphFormat.Alignment = q.ParagraphFormat.Alignment;
                    TabStop[] tabs = q.GetEffectiveTabStops();
                    foreach (TabStop tab in tabs)
                    {
                        builder.ParagraphFormat.TabStops.Add(tab);
                    }
                }

                builder.InsertBreak(BreakType.ParagraphBreak);
                builder.Write(value);
                control.Remove();
            }
            else
            {
                control.RemoveAllChildren();
                Run run = new Run(this.WordDoc, value);
                run.Font.Bold = control.ContentsFont.Bold;
                run.Font.Name = control.ContentsFont.Name;
                run.Font.Size = control.ContentsFont.Size;
                control.AppendChild(run);
            }
        }

        /// <summary>
        /// Inset a value into a Content Control within the Word Document by Tag. This will replace any value already in the Content Control.
        /// </summary>
        /// <param name="value">The value to insert into the control corresponding to the control type expected.
        /// Eg: String value for Text control. DateTime for Date Picker control, Image path for Picture Content control.</param>
        /// <param name="controlId">Tag value of Control used as ID. If more than one match is found, value will be inserted into all.</param>
        public void InsertValue(object value, string controlId)
        {
            if (value == null)
            {
                throw new ArgumentException("Value must not be empty.", "value");
            }

            if (string.IsNullOrWhiteSpace(controlId))
            {
                throw new ArgumentException("A controlId is required.", "controlId");
            }

            foreach (var control in this.GetControlByTag(controlId))
            {
                this.InsertValue(value, control);
            }

            foreach (var control in this.GetControls())
            {
                string finval = control.Tag.ToString();
            }
        }

        /// <summary>
        /// Inset a value into a Content Control within the Word Document by Tag. This will replace any value already in the Content Control.
        /// </summary>
        /// <param name="value">The value to insert into the control corresponding to the control type expected.
        /// Eg: String value for Text control. DateTime for Date Picker control, Image path for Picture Content control.</param>
        /// <param name="control">The Content Control as StructuredDocumentTag to insert into.</param>
        /// <param name="suffix">The suffix.</param>
        /// <param name="text">The text.</param>
        /// <exception cref="System.ArgumentException">Value must not be empty.;value or The given control must exist.;control</exception>
        public void InsertValue(object value, StructuredDocumentTag control, string suffix = "", string text = "")
        {
            if (value == null)
            {
                string error = " Value is coming in as null for control [" + control.Tag + "] in transaction [" + this.transactionId + "].";
                LogEvent.Log.Error(error);
                LogEvent.TraceLog.Error(error);
                value = string.Empty;
            }

            if (control == null)
            {
                string error = " The passed control is null in transaction [" + this.transactionId + "].";
                LogEvent.Log.Error(error);
                LogEvent.TraceLog.Error(error);
                return;
            }

            switch (control.SdtType)
            {
                case SdtType.RichText:
                case SdtType.PlainText:

                    if (text == string.Empty)
                    {
                        text = control.GetText();
                    }

                    float val = 0;
                    if (text.Length > 0)
                    {
                        if (text == " ")
                        {
                            text = string.Empty;
                        }

                        if (text.Length > 1 && text.Substring(text.Length - 1) == "\r")
                        {
                            text = text.Substring(0, text.Length - 1);
                        }

                        if (text.Length > 1 && text.Substring(text.Length - 1) == "\v")
                        {
                            suffix = "\v";
                            text = text.Substring(0, text.Length - 1);
                        }

                        if (text.EndsWith("%"))
                        {
                            suffix = "%" + suffix;
                            text = text.Substring(0, text.Length - 1);
                        }

                        if (text.IndexOf("0.00") != -1)
                        {
                            if (float.TryParse((string)value, out val))
                            {
                                value = text.Replace("0.00", Convert.ToDouble(value).ToString("N2"));
                            }
                        }
                        else if (text.IndexOf("00") != -1)
                        {
                            if (float.TryParse((string)value, out val))
                            {
                                value = text.Replace("00", Convert.ToDouble(value).ToString("N0"));
                            }
                            else if (string.IsNullOrWhiteSpace(value.ToString()))
                            {
                                value = text.Replace("00", "0");
                            }
                        }
                        else if (text.IndexOf("_Date:") != -1)
                        {
                            DateTime valDate = default(DateTime);
                            if (DateTime.TryParse(value as string, out valDate))
                            {
                                string dateFormat = text.Replace("_Date:", string.Empty);
                                value = valDate.ToString(dateFormat);
                            }
                        }
                        else
                        {
                            if (text == " ")
                            {
                                text = string.Empty;
                            }

                            value = text + value.ToString();
                        }

                        value = value + suffix;
                    }

                    this.InsertText(value as string, control);
                    break;
                case SdtType.Date:
                    DateTime date = default(DateTime);
                    if (DateTime.TryParse(value as string, out date))
                    {
                        this.InsertDate(date, control);
                    }

                    break;
                case SdtType.ComboBox:
                case SdtType.DropDownList:
                    this.SelectDropDown(value as string, control);
                    break;
                case SdtType.Picture:
                    if (!string.IsNullOrWhiteSpace(value as string) && File.Exists(value as string))
                    {
                        this.InsertPicture(value as string, control);
                    }

                    break;
                case SdtType.Checkbox:
                    bool isChecked = false;
                    var checkedVal = value.ToString().ToLower();
                    if (checkedVal == "on" || checkedVal == "true")
                    {
                        isChecked = true;
                    }

                    this.ToggleCheckBox(isChecked, control);
                    break;
            }
        }

        /// <summary>
        /// Determines whether [is block control] [the specified control].
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns>Boolean</returns>
        public bool IsBlockControl(StructuredDocumentTag control)
        {
            return control.Level == MarkupLevel.Block;
        }

        /// <summary>
        /// Find and remove all controls that match a given controlId
        /// </summary>
        /// <param name="controlId">Tag value of Control used as ID. If more than one match is found, all will be removed.</param>
        /// <exception cref="ArgumentException">A controlId is required.;controlId</exception>
        public void Remove(string controlId)
        {
            if (string.IsNullOrWhiteSpace(controlId))
            {
                throw new ArgumentException("A controlId is required.", "controlId");
            }

            foreach (var control in this.GetControlByTag(controlId))
            {
                control.Remove();
            }
        }

        /// <summary>
        /// Save the current document to the file system at the path specified in a given file format. Default format is DOCX.
        /// </summary>
        /// <param name="directory">The directory.</param>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="fileFormat">Format to save the document as. Default is: DocX</param>
        public void Save(string directory, string fileName, FileFormat fileFormat = FileFormat.DOCX)
        {
            Directory.CreateDirectory(directory);

            string savePath = Path.Combine(directory, fileName);
            this.WordDoc.Save(savePath, this.GetSaveFormat(fileFormat));
            this.SetSaveName(fileFormat);
        }

        /// <summary>
        /// Saves the specified stream.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="fileFormat">The file format.</param>
        public void Save(Stream stream, FileFormat fileFormat = FileFormat.DOCX)
        {
            this.WordDoc.Save(stream, this.GetSaveFormat(fileFormat));
            this.SetSaveName(fileFormat);
        }

        /// <summary>
        /// Saves the specified response.
        /// </summary>
        /// <param name="response">The response.</param>
        /// <param name="fileName">Name of the file.</param>
        public void Save(HttpResponse response, string fileName)
        {
            // extension for format is determined by file ext - this ext should be controlled for validation
            this.WordDoc.Save(response, fileName, ContentDisposition.Attachment, null);
            this.SetSaveName(this.GetFormatFromStr(Path.GetExtension(fileName)));
        }

        /// <summary>
        /// Generates Tiff from wordDoc and converts it to a PNG.
        /// </summary>
        /// <param name="docStream">The document stream.</param>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="saveToFile">if set to <c>true</c> [save to file].</param>
        /// <returns>the image in byte[] format</returns>
        public object SaveToPng(MemoryStream docStream, string fileName, bool saveToFile = false)
        {
            docStream.Position = 0;

            MemoryStream ms;
            Bitmap myBMP;
            Image myImage;
            int height = 0;
            int width = 0;
            System.Collections.ArrayList myImages = new System.Collections.ArrayList();

            // take image from stream
            myImage = Image.FromStream(docStream);
            Guid myGuid = myImage.FrameDimensionsList[0];
            System.Drawing.Imaging.FrameDimension myDimension = new System.Drawing.Imaging.FrameDimension(myGuid);
            int myPageCount = myImage.GetFrameCount(myDimension);

            // cycle through the tiff pages and save them into the bitmap file array list
            for (int i = 0; i < myPageCount; i++)
            {
                ms = new MemoryStream();
                myImage.SelectActiveFrame(myDimension, i);
                myImage.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp);
                myBMP = new Bitmap(ms);
                height += myBMP.Height;
                width = myBMP.Width > width ? myBMP.Width : width;

                myImages.Add(myBMP);
                ms.Close();
            }

            docStream.Close();
            System.Drawing.Bitmap finalImage = new System.Drawing.Bitmap((int)(width * .85), (int)(height * .85));

            // cycle through the bitmap file array and attach each to the finalImage image
            using (System.Drawing.Graphics g = System.Drawing.Graphics.FromImage(finalImage))
            {
                // set background color
                g.Clear(System.Drawing.Color.White);

                // go through each image and draw it on the final image
                int offset = 0;
                foreach (System.Drawing.Bitmap image in myImages)
                {
                    g.DrawImage(image, new System.Drawing.Rectangle(0, offset, (int)(image.Width * .85), (int)(image.Height * .85)));
                    offset += (int)(image.Height * .85);
                }
            }

            if (saveToFile)
            {
                // save image as a png file
                string path = System.Configuration.ConfigurationManager.AppSettings["tempPath"].ToString();
                string savePath = System.Web.Hosting.HostingEnvironment.MapPath(Path.Combine(path, fileName + ".png"));
                finalImage.Save(savePath, System.Drawing.Imaging.ImageFormat.Png);
            }

            ms = new MemoryStream();
            finalImage.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
            var myfile = ms.ToArray();
            return myfile;
        }

        /// <summary>
        /// Set selected value in a Drop Down List or Combo Box
        /// </summary>
        /// <param name="selectedValue">Value in list to select. Set to NULL to clear selection</param>
        /// <param name="control">A Drop Down List or Combo Box</param>
        public void SelectDropDown(string selectedValue, StructuredDocumentTag control)
        {
            if (string.IsNullOrWhiteSpace(selectedValue))
            {
                control.ListItems.SelectedValue = null;
            }
            else
            {
                SdtListItem listItem = control.ListItems.Cast<SdtListItem>().FirstOrDefault(li => li.Value == selectedValue);
                control.ListItems.SelectedValue = listItem;
            }
        }

        /// <summary>
        /// Strips the HTML from a string.
        /// </summary>
        /// <param name="text">The text.</param>
        /// <returns>string</returns>
        public string StripHTML(string text)
        {
            text = System.Text.RegularExpressions.Regex.Replace(text, @"&nbsp;", " ");
            return System.Text.RegularExpressions.Regex.Replace(text, @"<[^>]*?>|\r\n", string.Empty);
        }

        /// <summary>
        /// Toggles the CheckBox.
        /// </summary>
        /// <param name="isChecked">if set to <c>true</c> [is checked].</param>
        /// <param name="control">The control.</param>
        public void ToggleCheckBox(bool isChecked, StructuredDocumentTag control)
        {
            DocumentBuilder builder = new DocumentBuilder(this.WordDoc);
            builder.MoveTo(control);
            builder.InsertCheckBox("addReview", isChecked, 0);
        }

        /// <summary>
        /// Inserts a watermark with the specified text into the generated document
        /// </summary>
        /// <param name="waterMarkText">Watermark text that appears in the background of the document</param>
        public void InsertWaterMark(string waterMarkText = "SPECIMEN")
        {
            Shape waterMark = new Shape(this.WordDoc, ShapeType.TextPlainText)
            {
                DistanceBottom = -5,
                Width = 400,
                Height = 80,
                Rotation = -40,
                StrokeColor = Color.LightGray,
                FillColor = Color.Gray,
                BehindText = true,
                RelativeHorizontalPosition = RelativeHorizontalPosition.Page,
                RelativeVerticalPosition = RelativeVerticalPosition.Page,
                WrapType = WrapType.None,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
            };
            waterMark.TextPath.Text = waterMarkText;
            waterMark.TextPath.TextPathAlignment = TextPathAlignment.WordJustify;

            // Insert the watermark into all headers of each document section
            foreach (Section section in this.WordDoc.Sections)
            {
                this.InsertWatermarkIntoHeader(waterMark, section, HeaderFooterType.HeaderPrimary);
                this.InsertWatermarkIntoHeader(waterMark, section, HeaderFooterType.HeaderFirst);
                this.InsertWatermarkIntoHeader(waterMark, section, HeaderFooterType.HeaderEven);
            }
        }

        /// <summary>
        /// Updates the repeater links.
        /// </summary>
        public void UpdateRepeaterLinks()
        {
            this.WordDoc.UpdateFields();
        }

        /// <summary>
        /// Removes from the document a table containing bookmark with the specified name
        /// </summary>
        /// <param name="bookmarkName">Bookmark name</param>
        /// <param name="removeParagraphAfterTable">If paragraph below the table should be removed</param>
        public void RemoveTableWithBookmark(string bookmarkName, bool removeParagraphAfterTable)
        {
            Table table = this.GetTableContainingBookmark(bookmarkName);
            Node para = table.NextSibling;
            if (table != null)
            {
                table.Remove();

                if (para != null && typeof(Paragraph) == para.GetType() && removeParagraphAfterTable)
                {
                    para.Remove();
                }
            }
        }

        /// <summary>
        /// Replaces the question bookmark value.
        /// </summary>
        /// <param name="bookmarkName">Name of the bookmark.</param>
        /// <param name="question">The question.</param>
        public void ReplaceQuestionBookmarkValue(string bookmarkName, Question question)
        {
            this.ReplaceQuestionBookmarkValue(bookmarkName, question, null, null);
        }

        /// <summary>
        /// Replaces the question bookmark value.
        /// </summary>
        /// <param name="bookmarkName">Name of the bookmark.</param>
        /// <param name="question">The question.</param>
        /// <param name="formatFunction">Format function applied to the text value</param>
        /// <param name="formatMethodInfo">The format method information.</param>
        public void ReplaceQuestionBookmarkValue(string bookmarkName, Question question, Func<string, string> formatFunction, MethodInfo formatMethodInfo = null)
        {
            // Get the bookmark from the document
            var bookmark = this.GetBookMarkByName(bookmarkName);
            if (bookmark == null)
            {
                // If there is no bookmark with that name, continue
                return;
            }

            this.ReplaceQuestionBookmarkValue(bookmark, question, formatFunction, formatMethodInfo);
        }

        /// <summary>
        /// Replaces the question bookmark value.
        /// </summary>
        /// <param name="bookmark">The bookmark.</param>
        /// <param name="question">The question.</param>
        public void ReplaceQuestionBookmarkValue(Bookmark bookmark, Question question)
        {
            this.ReplaceQuestionBookmarkValue(bookmark, question, null, null);
        }

        /// <summary>
        /// Replaces the question bookmark value.
        /// </summary>
        /// <param name="bookmark">The bookmark.</param>
        /// <param name="question">The question.</param>
        /// <param name="formatFunction">Format function applied to the text value</param>
        /// <param name="formatMethodInfo">Format function applied to the text value as a MethodInfo</param>
        public void ReplaceQuestionBookmarkValue(Bookmark bookmark, Question question, Func<string, string> formatFunction, MethodInfo formatMethodInfo = null)
        {
            if (bookmark == null)
            {
                return;
            }

            var nodeValue = this.GetQuestionValue(bookmark.Name, question);
            if (formatFunction != null)
            {
                nodeValue = formatFunction(nodeValue);
            }

            // Handle bookmarks configured as TextBox type Q&A
            if (!string.IsNullOrWhiteSpace(nodeValue))
            {
                this.ReplaceNodeValue(bookmark, nodeValue);
            }
        }

        /// <summary>
        /// Gets the question value.
        /// </summary>
        /// <param name="bookmarkName">Name of the bookmark.</param>
        /// <param name="question">The question.</param>
        /// <returns>System.String.</returns>
        public string GetQuestionValue(string bookmarkName, Question question)
        {
            var bookmarkNameSeparator = new char[] { ',' };

            Answer selectedAnswer;
            switch (question.DisplayFormat)
            {
                case DisplayFormat.CheckBox:
                    // Set the checkbox to checked if the answer code
                    selectedAnswer = question.Answers.FirstOrDefault(a => a.MergeFieldName.Split(bookmarkNameSeparator).Contains(bookmarkName, StringComparer.InvariantCultureIgnoreCase) &&
                                                                            a.IsSelected);
                    if (selectedAnswer != null)
                    {
                        return "CHECKED";
                    }

                    break;

                case DisplayFormat.RadioButton:
                case DisplayFormat.DropDown:
                    // If the questions merge field name is defined and it matches the current bookmark, we can
                    // fill in the answer, otherwise you get the value for the bookmark from the selected answer for the question
                    if (!string.IsNullOrWhiteSpace(question.MergeFieldName) &&
                        question.MergeFieldName.Split(bookmarkNameSeparator).Contains(bookmarkName, StringComparer.InvariantCultureIgnoreCase))
                    {
                        return question.Answers.FirstOrDefault(a => a.Value.EqualsIgnoreCase(question.AnswerValue))?.Verbiage ?? string.Empty;
                    }

                    selectedAnswer = question.Answers.FirstOrDefault(a => a.Value.EqualsIgnoreCase(question.AnswerValue));
                    if (selectedAnswer != null)
                    {
                        var blankMergeField = selectedAnswer.MergeFieldName;
                        var verbiage = selectedAnswer.Verbiage;
                        if (blankMergeField != null &&
                            blankMergeField.Split(bookmarkNameSeparator).Contains(bookmarkName, StringComparer.InvariantCultureIgnoreCase))
                        {
                            return verbiage;
                        }
                    }

                    break;

                case DisplayFormat.StaticText:
                    return question.Verbiage;

                default:
                    return question.AnswerValue;
            }

            return string.Empty;
        }

        /// <summary>
        /// Checks the string for nulls.
        /// </summary>
        /// <param name="obj">The object.</param>
        /// <param name="str">The string.</param>
        /// <returns>String</returns>
        protected string CheckStringForNulls(object obj, string str = "")
        {
            return obj?.ToString() ?? str;
        }

        /// <summary>
        /// Returns a .NET Image object from the specified byte array.
        /// </summary>
        /// <param name="imageBytes">The image bytes.</param>
        /// <returns>Image</returns>
        private static Image GetImageFromByteArray(byte[] imageBytes)
        {
            // Some drivers can pick up some junk data to the start of binary storage fields.
            // This means we cannot directly read the bytes into an image, we first need
            // to skip past until we find the start of the image.
            string imageString = System.Text.Encoding.ASCII.GetString(imageBytes);
            int index = imageString.IndexOf("BM");
            return Image.FromStream(new MemoryStream(imageBytes, index, imageBytes.Length - index));
        }

        /// <summary>
        /// Returns a table containing the specified bookmark.
        /// </summary>
        /// <param name="bookmarkName">Bookmark name</param>
        /// <returns>Table</returns>
        private Table GetTableContainingBookmark(string bookmarkName)
        {
            var bookmark = this.GetBookMarkByName(bookmarkName);
            Table table = (Table)bookmark.BookmarkStart.GetAncestor(NodeType.Table);
            return table;
        }

        /// <summary>
        /// Inserts the watermark into the specific section Header
        /// </summary>
        /// <param name="waterMark">The water mark.</param>
        /// <param name="section">Section to set the Header and footer for the Watermark</param>
        /// <param name="headerType">HeaderType for the watermark section</param>
        private void InsertWatermarkIntoHeader(Shape waterMark, Section section, HeaderFooterType headerType)
        {
            bool headerExists = false;
            if (headerType == HeaderFooterType.HeaderPrimary)
            {
                foreach (HeaderFooter item in section.HeadersFooters)
                {
                    if (item.IsHeader)
                    {
                        headerExists = true;
                        break;
                    }
                }
            }

            if (!headerExists)
            {
                section.PageSetup.HeaderDistance = 0;
            }

            HeaderFooter header = section.HeadersFooters[headerType];

            if (header == null)
            {
                // There is no header in the document template for the current section, create one
                header = new HeaderFooter(section.Document, headerType);
                section.HeadersFooters.Add(header);
            }

            Paragraph headerParagraph = null;
            var headerParagraphs = header.GetChildNodes(NodeType.Paragraph, true);
            if (headerParagraphs != null ||
                headerParagraphs.Count == 0)
            {
                headerParagraph = headerParagraphs[0] as Paragraph;
            }

            if (headerParagraph == null)
            {
                headerParagraph = new Paragraph(this.WordDoc);
                header.AppendChild(headerParagraph);
            }

            headerParagraph.AppendChild(waterMark.Clone(true));

            ////// Insert a clone of the watermark into the header
            ////header.AppendChild(watermarkParaph.Clone(true));
        }

        /// <summary>
        /// Builds the table from data table.
        /// </summary>
        /// <param name="dataTable">The data table.</param>
        /// <param name="builder">The builder.</param>
        /// <param name="importHeadings">if set to <c>true</c> [import headings].</param>
        /// <returns>Table</returns>
        private Table BuildTableFromDataTable(DataTable dataTable, DocumentBuilder builder, bool importHeadings)
        {
            Table table = builder.StartTable();
            if (importHeadings)
            {
                // store original values
                bool bold = builder.Font.Bold;
                ParagraphAlignment parAlign = builder.ParagraphFormat.Alignment;

                // override for heading
                builder.Font.Bold = true;
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

                // loop through columns and set headings
                foreach (DataColumn col in dataTable.Columns)
                {
                    builder.InsertCell();
                    builder.Write(col.ColumnName);
                }

                builder.EndRow();

                // reset to original vals
                builder.Font.Bold = bold;
                builder.ParagraphFormat.Alignment = parAlign;
            }

            foreach (DataRow row in dataTable.Rows)
            {
                foreach (var item in row.ItemArray)
                {
                    builder.InsertCell();
                    switch (item.GetType().Name)
                    {
                        case "Byte[]":
                            // Assume a byte array is an image. Other data types can be added here. Might not even need to account for images here
                            // but keep this in check just in case
                            builder.InsertImage(GetImageFromByteArray((byte[])item), 50, 50);
                            break;
                        default:
                            // By default any other item will be inserted as text.
                            builder.Write(item.ToString());
                            break;
                    }
                }

                builder.EndRow();
            }

            builder.EndTable();
            return table;
        }

        /// <summary>
        /// Get a list of StructuredDocumentControls using the Tag property as an ID.
        /// </summary>
        /// <param name="controlId">ID of control as used in the control's Tag property</param>
        /// <returns>List of structured document Tags</returns>
        private List<StructuredDocumentTag> GetControlByTag(string controlId)
        {
            return this.WordDoc.GetChildNodes(NodeType.StructuredDocumentTag, true).Cast<StructuredDocumentTag>().Where(sdt => sdt.Tag == controlId).ToList();
        }

        /// <summary>
        /// Gets the format from string.
        /// </summary>
        /// <param name="extension">The extension.</param>
        /// <returns>File Format</returns>
        private FileFormat GetFormatFromStr(string extension)
        {
            if (extension.StartsWith("."))
            {
                extension = extension.Substring(1);
            }

            FileFormat format;
            if (Enum.TryParse<FileFormat>(extension, true, out format))
            {
                return format;
            }

            return FileFormat.None;
        }

        /// <summary>
        /// Gets the save format.
        /// </summary>
        /// <param name="format">The format.</param>
        /// <returns>Save Format</returns>
        private Aspose.Words.SaveFormat GetSaveFormat(FileFormat format)
        {
            SaveFormat asposeFormat;
            if (!Enum.TryParse<SaveFormat>(format.ToString(), true, out asposeFormat))
            {
                asposeFormat = SaveFormat.Docx;
            }

            return asposeFormat;
        }

        /// <summary>
        /// Sets the name of the save.
        /// </summary>
        /// <param name="fileFormat">The file format.</param>
        private void SetSaveName(FileFormat fileFormat)
        {
            this.SaveNameWithExt = this.OriginalFileName + "." + fileFormat.ToString().ToLower();
        }

        /// <summary>
        /// Replaces the header/footer field with result.
        /// </summary>
        /// <param name="builder">The builder.</param>
        /// <param name="headerFooter">The header footer.</param>
        /// <param name="fieldType">Type of the field.</param>
        /// <param name="replacementText">The replacement text.</param>
        private void ReplaceHeaderFooterFieldWithResult(DocumentBuilder builder, HeaderFooter headerFooter, FieldType fieldType, string replacementText = null)
        {
            if (builder == null)
            {
                builder = new DocumentBuilder(this.WordDoc);
            }

            var field = this.GetFieldByFieldType(headerFooter.Range.Fields, fieldType);
            if (field != null)
            {
                var value = replacementText ?? field.Result;
                builder.MoveToField(field, false);
                builder.Write(value);
                field.Remove();
            }
        }

        /// <summary>
        /// Gets the field by the field type.
        /// </summary>
        /// <param name="fields">The fields.</param>
        /// <param name="fieldType">Type of the field.</param>
        /// <returns>Field.</returns>
        private Field GetFieldByFieldType(FieldCollection fields, FieldType fieldType)
        {
            return fields?.Cast<Field>().FirstOrDefault(field => field.Type == fieldType);
        }
    }
}