// ***********************************************************************
// Assembly         : CorrespondenceServices.Tests
// Author           : rsteelea
// Created          : 08-16-2016
//
// Last Modified By : rsteelea
// Last Modified On : 08-16-2016
// ***********************************************************************
// <copyright file="PolicyDocumentManagerStub.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************

namespace CorrespondenceServices.Tests.Stubs
{
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.Diagnostics.CodeAnalysis;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Web;
    using Aspose.Words;
    using Aspose.Words.Markup;
    using DecisionModel.Interfaces;
    using DecisionModel.Models.Policy;
    using Mkl.WebTeam.DocumentGenerator;
    using Document = DecisionModel.Models.Policy.Document;

    /// <summary>
    /// Class PolicyDocumentManagerStub.
    /// </summary>
    /// <seealso cref="Mkl.WebTeam.DocumentGenerator.IPolicyDocumentManager" />
    [ExcludeFromCodeCoverage]
    public class PolicyDocumentManagerStub : IPolicyDocumentManager
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PolicyDocumentManagerStub"/> class.
        /// </summary>
        public PolicyDocumentManagerStub()
        {
            this.Document = new Aspose.Words.Document();
            this.Builder = new DocumentBuilder(this.Document);
        }

        /// <summary>
        /// Gets the document.
        /// </summary>
        /// <value>The document.</value>
        public Aspose.Words.Document Document { get; }

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
        /// Gets or sets the normalized form number.
        /// </summary>
        /// <value>The normalized form number.</value>
        public string FormNormalizedNumber { get; set; }

        /// <summary>
        /// Gets or sets the form edition date.
        /// </summary>
        /// <value>The form edition date.</value>
        public DateTime EditionDate { get; set; }

        /// <summary>
        /// Gets or sets the builder.
        /// </summary>
        /// <value>The builder.</value>
        private DocumentBuilder Builder { get; set; }

        /// <summary>
        /// Builds the unordered list.
        /// </summary>
        /// <param name="list">The list.</param>
        /// <param name="control">The control.</param>
        /// <param name="isBind">if set to <c>true</c> [is bind].</param>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public void BuildUnorderedList(IEnumerable<Subjectivity> list, StructuredDocumentTag control, bool isBind = false)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Inserts the buildings table.
        /// </summary>
        /// <param name="buildings">The buildings.</param>
        /// <param name="control">The control.</param>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public void InsertBuildingsTable(IOrderedEnumerable<CfRiskUnit> buildings, StructuredDocumentTag control)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Inserts the clauses.
        /// </summary>
        /// <param name="clauses">The clauses.</param>
        /// <param name="control">The control.</param>
        /// <param name="documents">The documents.</param>
        /// <param name="useHeading">if set to <c>true</c> [use heading].</param>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public void InsertClauses(IEnumerable<Clause> clauses, StructuredDocumentTag control, IEnumerable<Document> documents, bool useHeading = true)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Inserts the coverage options.
        /// </summary>
        /// <param name="clauses">The clauses.</param>
        /// <param name="control">The control.</param>
        /// <param name="documents">The documents.</param>
        /// <param name="optCov">The optional coverages.</param>
        /// <param name="useHeading">if set to <c>true</c> [use heading].</param>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public void InsertCoverageOptions(IEnumerable<Clause> clauses, StructuredDocumentTag control, IEnumerable<Document> documents, IEnumerable<OptionalCoverage> optCov, bool useHeading = true)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Inserts the coverages table.
        /// </summary>
        /// <param name="coverages">The coverages.</param>
        /// <param name="control">The control.</param>
        /// <param name="glOccurLimit">The gl occur limit.</param>
        /// <param name="glDeductible">The gl deductible.</param>
        /// <param name="showLimit">if set to <c>true</c> [show limit].</param>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public void InsertCoveragesTable(IEnumerable<OptionalCoverage> coverages, StructuredDocumentTag control, int glOccurLimit, int glDeductible, bool showLimit = true)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Inserts the exposures table.
        /// </summary>
        /// <param name="risks">The risks.</param>
        /// <param name="control">The control.</param>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public void InsertExposuresTable(IEnumerable<IBaseGlRiskUnit> risks, StructuredDocumentTag control)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Inserts the im items table.
        /// </summary>
        /// <param name="imRisk">The im risk.</param>
        /// <param name="control">The control.</param>
        /// <param name="useHeading">if set to <c>true</c> [use heading].</param>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public void InsertImItemsTable(ImRiskUnit imRisk, StructuredDocumentTag control, bool useHeading = true)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Inserts the xs layers table.
        /// </summary>
        /// <param name="layers">The layers.</param>
        /// <param name="control">The control.</param>
        /// <param name="useHeading">if set to <c>true</c> [use heading].</param>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public void InsertLayerTable(List<XsLayer> layers, StructuredDocumentTag control, bool useHeading = true)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Inserts the link list table.
        /// </summary>
        /// <param name="docs">The docs.</param>
        /// <param name="control">The control.</param>
        /// <param name="useHeading">if set to <c>true</c> [use heading].</param>
        /// <param name="isBindLetter">Indicates if the letter is a bind letter</param>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public void InsertLinkListTable(List<Document> docs, StructuredDocumentTag control, bool useHeading = true, bool isBindLetter = false)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Inserts the name amount table.
        /// </summary>
        /// <param name="fees">The fees.</param>
        /// <param name="control">The control.</param>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public void InsertNameAmountTable(IOrderedEnumerable<TaxFee> fees, StructuredDocumentTag control)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Inserts the excess underlying coverage table.
        /// </summary>
        /// <param name="limits">The limits.</param>
        /// <param name="coverage">The coverage.</param>
        /// <param name="control">The control.</param>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public void InsertUnderlyingCoverageTable(List<Limit> limits, XsUnderlyingCoverage coverage, StructuredDocumentTag control)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Insert XS auto liability vehicleType.
        /// </summary>
        /// <param name="vehicleTypes">The vehicle types.</param>
        /// <param name="control">The control.</param>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public void InsertXSAutoLiabilityVehicleType(IEnumerable<XsAutoLiabilityVehicleType> vehicleTypes, StructuredDocumentTag control)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Inserts the exposures table.
        /// </summary>
        /// <param name="warranties">The warranties.</param>
        /// <param name="control">The control.</param>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public void InsertWarrantiesTable(IEnumerable<Warranty> warranties, StructuredDocumentTag control)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Inserts the policy form table
        /// </summary>
        /// <param name="documents">Collection of documents</param>
        /// <param name="control">Book mark</param>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public void InsertPolicyFormTable(Dictionary<dynamic, Dictionary<string, DecisionModel.Models.Policy.Document>> documents, Bookmark control)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Builds the mailing address
        /// </summary>
        /// <param name="streetAddress">streetAddress</param>
        /// <param name="control">Document Control</param>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public void InsertMailingAddress(Address streetAddress, StructuredDocumentTag control)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Appends the specified dest.
        /// run this to create the document that needs to be run through the population process????
        /// </summary>
        /// <param name="dest">The dest.</param>
        /// <param name="source">The source.</param>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public void Append(string dest, string source)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Appends a document to the one being managed and optionally inserts a page break before the new document.
        /// </summary>
        /// <param name="document">The document.</param>
        /// <param name="insertPageBreakBefore">if set to <c>true</c> [insert page break before].</param>
        /// <param name="restartNumbering">if set to <c>true</c> [restart numbering].</param>
        public void AppendDocument(Aspose.Words.Document document, bool insertPageBreakBefore = true, bool restartNumbering = true)
        {
            this.Builder.MoveToDocumentEnd();

            if (insertPageBreakBefore)
            {
                this.Builder.InsertBreak(BreakType.PageBreak);
            }

            this.Builder.Document.AppendDocument(document, ImportFormatMode.KeepSourceFormatting);
        }

        /// <summary>
        /// Updates the total page number footer field.
        /// </summary>
        public void UpdateTotalPageNumberFooterField()
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Combine all source documents into destination document.
        /// </summary>
        /// <param name="dest">Destination file</param>
        /// <param name="source">Source files</param>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public void Combine(string dest, string[] source)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Get a list of StructuredDocumentControls using the Tag property as an ID.
        /// </summary>
        /// <returns>List of controls</returns>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public List<StructuredDocumentTag> GetControls()
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Get the bookmarks from a word  template
        /// </summary>
        /// <returns>BookmarkCollection</returns>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public BookmarkCollection GetControlsBookMarks()
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Get the bookmark control from a word  template using bookmark name
        /// </summary>
        /// <param name="bookmarkName">The bookmark name.</param>
        /// <returns>Bookmark</returns>
        public Bookmark GetBookMarkByName(string bookmarkName)
        {
            this.Builder.StartBookmark("testBookmark");
            this.Builder.Writeln("mybookmark");
            this.Builder.EndBookmark("testBookmark");
            return this.Document.Range.Bookmarks["testBookmark"];
        }

        /// <summary>
        /// Inserts the date.
        /// </summary>
        /// <param name="date">The date.</param>
        /// <param name="control">The control.</param>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public void InsertDate(DateTime date, StructuredDocumentTag control)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Insert a hyperlink into Rich Text
        /// </summary>
        /// <param name="link">The link.</param>
        /// <param name="text">The text.</param>
        /// <param name="controlId">The control identifier.</param>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public void InsertLink(string link, string text, string controlId)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Inserts the picture.
        /// </summary>
        /// <param name="imagePath">The image path.</param>
        /// <param name="control">The control.</param>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public void InsertPicture(string imagePath, StructuredDocumentTag control)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Insert a table from a DataTable at the location specified by the controlId.
        /// </summary>
        /// <param name="dataTable">The data table.</param>
        /// <param name="controlId">ID of the control to use as a location for the table</param>
        /// <param name="useHeading">Use the headings from the DataTable. Default is true.</param>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public void InsertTable(DataTable dataTable, string controlId, bool useHeading = true)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Inserts the table.
        /// </summary>
        /// <param name="dataTable">The data table.</param>
        /// <param name="control">The control.</param>
        /// <param name="useHeading">if set to <c>true</c> [use heading].</param>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public void InsertTable(DataTable dataTable, StructuredDocumentTag control, bool useHeading = true)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// replace bookmark text
        /// </summary>
        /// <param name="placeHolder">placeHolder</param>
        /// <param name="replaceText">replaceText</param>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public void InsertFieldText(string placeHolder, string replaceText)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Remove bookmark
        /// </summary>
        /// <param name="bookmark">Bookmark</param>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public void RemoveBookmark(Bookmark bookmark)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Replaces the node value
        /// </summary>
        /// <param name="bookmark">bookmark</param>
        /// <param name="replaceText">replaceText</param>
        public void ReplaceNodeValue(Bookmark bookmark, string replaceText)
        {
            this.Document.Range.Bookmarks[0].Text = replaceText;
        }

        /// <summary>
        /// Insert string into Plain Text or markup into Rich Text
        /// </summary>
        /// <param name="value">String of text or Rich Text markup to insert</param>
        /// <param name="control">Plain Text or Rich Text control</param>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public void InsertText(string value, StructuredDocumentTag control)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Inset a value into a Content Control within the Word Document by Tag. This will replace any value already in the Content Control.
        /// </summary>
        /// <param name="value">The value to insert into the control corresponding to the control type expected.
        /// Eg: String value for Text control. DateTime for Date Picker control, Image path for Picture Content control.</param>
        /// <param name="controlId">Tag value of Control used as ID. If more than one match is found, value will be inserted into all.</param>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public void InsertValue(object value, string controlId)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Inset a value into a Content Control within the Word Document by Tag. This will replace any value already in the Content Control.
        /// </summary>
        /// <param name="value">The value to insert into the control corresponding to the control type expected.
        /// Eg: String value for Text control. DateTime for Date Picker control, Image path for Picture Content control.</param>
        /// <param name="control">The Content Control as StructuredDocumentTag to insert into.</param>
        /// <param name="suffix">The suffix.</param>
        /// <param name="text">The text.</param>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public void InsertValue(object value, StructuredDocumentTag control, string suffix = "", string text = "")
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Determines whether [is block control] [the specified control].
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns>Boolean</returns>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public bool IsBlockControl(StructuredDocumentTag control)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Find and remove all controls that match a given controlId
        /// </summary>
        /// <param name="controlId">Tag value of Control used as ID. If more than one match is found, all will be removed.</param>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public void Remove(string controlId)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Save the current document to the file system at the path specified in a given file format. Default format is DOCX.
        /// </summary>
        /// <param name="directory">The directory.</param>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="fileFormat">Format to save the document as. Default is: DocX</param>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public void Save(string directory, string fileName, DocumentManager.FileFormat fileFormat = DocumentManager.FileFormat.DOCX)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Saves the specified stream.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="fileFormat">The file format.</param>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public void Save(Stream stream, DocumentManager.FileFormat fileFormat = DocumentManager.FileFormat.DOCX)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Saves the specified response.
        /// </summary>
        /// <param name="response">The response.</param>
        /// <param name="fileName">Name of the file.</param>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public void Save(HttpResponse response, string fileName)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Generates Tiff from wordDoc and converts it to a PNG.
        /// </summary>
        /// <param name="docStream">The document stream.</param>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="saveToFile">if set to <c>true</c> [save to file].</param>
        /// <returns>the image in byte[] format</returns>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public object SaveToPng(MemoryStream docStream, string fileName, bool saveToFile = false)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Set selected value in a Drop Down List or Combo Box
        /// </summary>
        /// <param name="selectedValue">Value in list to select. Set to NULL to clear selection</param>
        /// <param name="control">A Drop Down List or Combo Box</param>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public void SelectDropDown(string selectedValue, StructuredDocumentTag control)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Strips the HTML from a string.
        /// </summary>
        /// <param name="text">The text.</param>
        /// <returns>string</returns>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public string StripHTML(string text)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Toggles the CheckBox.
        /// </summary>
        /// <param name="isChecked">if set to <c>true</c> [is checked].</param>
        /// <param name="control">The control.</param>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public void ToggleCheckBox(bool isChecked, StructuredDocumentTag control)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Inserts a watermark with the specified text into the generated document
        /// </summary>
        /// <param name="waterMarkText">Watermark text that appears in the background of the document</param>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public void InsertWaterMark(string waterMarkText = "SPECIMEN")
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Updates the repeater links.
        /// </summary>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public void UpdateRepeaterLinks()
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Removes from the document a table containing bookmark with the specified name
        /// </summary>
        /// <param name="bookmarkName">Bookmark name</param>
        /// <param name="removeParagraphAfterTable">If paragraph below the table should be removed</param>
        /// <exception cref="NotImplementedException">The feature has not been implemented</exception>
        public void RemoveTableWithBookmark(string bookmarkName, bool removeParagraphAfterTable)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Replaces the question bookmark value.
        /// </summary>
        /// <param name="bookmarkName">Name of the bookmark.</param>
        /// <param name="question">The question.</param>
        public void ReplaceQuestionBookmarkValue(string bookmarkName, Question question)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Replaces the question bookmark value.
        /// </summary>
        /// <param name="bookmark">The bookmark.</param>
        /// <param name="question">The question.</param>
        public void ReplaceQuestionBookmarkValue(Bookmark bookmark, Question question)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Replaces the question bookmark value.
        /// </summary>
        /// <param name="bookmarkName">Name of the bookmark.</param>
        /// <param name="question">The question.</param>
        /// <param name="formatFunction">Format function applied to the text value</param>
        /// <param name="formatMethodInfo">Format function applied to the text value as a MethodInfo</param>
        public void ReplaceQuestionBookmarkValue(string bookmarkName, Question question, Func<string, string> formatFunction, MethodInfo formatMethodInfo = null)
        {
            throw new NotImplementedException();
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
            throw new NotImplementedException();
        }

        /// <summary>
        /// Gets the question value.
        /// </summary>
        /// <param name="bookmarkName">Name of the bookmark.</param>
        /// <param name="question">The question.</param>
        /// <returns>System.String.</returns>
        public string GetQuestionValue(string bookmarkName, Question question)
        {
            throw new NotImplementedException();
        }
    }
}
