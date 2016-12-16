// <copyright file="IPolicyDocumentManager.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>

namespace Mkl.WebTeam.DocumentGenerator
{
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Web;
    using Aspose.Words;
    using Aspose.Words.Markup;
    using DecisionModel.Interfaces;
    using DecisionModel.Models.Policy;

    /// <summary>
    /// Interface IPolicyDocumentManager
    /// </summary>
    public interface IPolicyDocumentManager
    {
        /// <summary>
        /// Gets the document.
        /// </summary>
        /// <value>
        /// The document.
        /// </value>
        Aspose.Words.Document Document { get; }

        /// <summary>
        /// Gets or sets the name of the original file.
        /// </summary>
        /// <value>
        /// The name of the original file.
        /// </value>
        string OriginalFileName { get; set; }

        /// <summary>
        /// Gets or sets the save name with ext.
        /// </summary>
        /// <value>
        /// The save name with ext.
        /// </value>
        string SaveNameWithExt { get; set; }

        /// <summary>
        /// Gets or sets the quantity order.
        /// </summary>
        /// <value>The quantity order.</value>
        int QuantityOrder { get; set; }

        /// <summary>
        /// Gets or sets the form edition date.
        /// </summary>
        /// <value>The edition date.</value>
        DateTime EditionDate { get; set; }

        /// <summary>
        /// Gets or sets the normalized form number.
        /// </summary>
        /// <value>The normalized form number.</value>
        string FormNormalizedNumber { get; set; }

        /// <summary>
        /// Builds the unordered list.
        /// </summary>
        /// <param name="list">The list.</param>
        /// <param name="control">The control.</param>
        /// <param name="isBind">if set to <c>true</c> [is bind].</param>
        /// <exception cref="System.Exception">Specified control must be a Paragraph or Inline control, or must come after a Paragraph or Inline control.</exception>
        void BuildUnorderedList(IEnumerable<Subjectivity> list, StructuredDocumentTag control, bool isBind = false);

        /// <summary>
        /// Inserts the buildings table.
        /// </summary>
        /// <param name="buildings">The buildings.</param>
        /// <param name="control">The control.</param>
        /// <exception cref="Exception">Specified control must be a Paragraph or Inline control, or must come after a Paragraph or Inline control.</exception>
        void InsertBuildingsTable(IOrderedEnumerable<CfRiskUnit> buildings, StructuredDocumentTag control);

        /// <summary>
        /// Inserts the clauses.
        /// </summary>
        /// <param name="clauses">The clauses.</param>
        /// <param name="control">The control.</param>
        /// <param name="documents">The documents.</param>
        /// <param name="useHeading">if set to <c>true</c> [use heading].</param>
        /// <exception cref="Exception">Specified control must be a Paragraph or Inline control, or must come after a Paragraph or Inline control.</exception>
        void InsertClauses(IEnumerable<Clause> clauses, StructuredDocumentTag control, IEnumerable<DecisionModel.Models.Policy.Document> documents, bool useHeading = true);

        /// <summary>
        /// Inserts the coverage options.
        /// </summary>
        /// <param name="clauses">The clauses.</param>
        /// <param name="control">The control.</param>
        /// <param name="documents">The documents.</param>
        /// <param name="optCov">The optional coverages.</param>
        /// <param name="useHeading">if set to <c>true</c> [use heading].</param>
        /// <exception cref="Exception">Specified control must be a Paragraph or Inline control, or must come after a Paragraph or Inline control.</exception>
        void InsertCoverageOptions(IEnumerable<Clause> clauses, StructuredDocumentTag control, IEnumerable<DecisionModel.Models.Policy.Document> documents, IEnumerable<OptionalCoverage> optCov, bool useHeading = true);

        /// <summary>
        /// Inserts the coverages table.
        /// </summary>
        /// <param name="coverages">The coverages.</param>
        /// <param name="control">The control.</param>
        /// <param name="glOccurLimit">The gl occur limit.</param>
        /// <param name="glDeductible">The gl deductible.</param>
        /// <param name="showLimit">if set to <c>true</c> [show limit].</param>
        /// <exception cref="System.Exception">Specified control must be a Paragraph or Inline control, or must come after a Paragraph or Inline control.</exception>
        void InsertCoveragesTable(IEnumerable<OptionalCoverage> coverages, StructuredDocumentTag control, int glOccurLimit, int glDeductible, bool showLimit = true);

        /// <summary>
        /// Inserts the exposures table.
        /// </summary>
        /// <param name="risks">The risks.</param>
        /// <param name="control">The control.</param>
        /// <exception cref="Exception">Specified control must be a Paragraph or Inline control, or must come after a Paragraph or Inline control.</exception>
        void InsertExposuresTable(IEnumerable<IBaseGlRiskUnit> risks, StructuredDocumentTag control);

        /// <summary>
        /// Inserts the im items table.
        /// </summary>
        /// <param name="imRisk">The im risk.</param>
        /// <param name="control">The control.</param>
        /// <param name="useHeading">if set to <c>true</c> [use heading].</param>
        /// <exception cref="Exception">Specified control must be a Paragraph or Inline control, or must come after a Paragraph or Inline control.</exception>
        void InsertImItemsTable(ImRiskUnit imRisk, StructuredDocumentTag control, bool useHeading = true);

        /// <summary>
        /// Inserts the xs layers table.
        /// </summary>
        /// <param name="layers">The layers.</param>
        /// <param name="control">The control.</param>
        /// <param name="useHeading">if set to <c>true</c> [use heading].</param>
        /// <exception cref="System.Exception">Specified control must be a Paragraph or Inline control, or must come after a Paragraph or Inline control.</exception>
        void InsertLayerTable(List<XsLayer> layers, StructuredDocumentTag control, bool useHeading = true);

        /// <summary>
        /// Inserts the link list table.
        /// </summary>
        /// <param name="docs">The docs.</param>
        /// <param name="control">The control.</param>
        /// <param name="useHeading">if set to <c>true</c> [use heading].</param>
        /// <param name="isBindLetter">Indicates if the letter is a bind letter</param>
        /// <exception cref="Exception">Specified control must be a Paragraph or Inline control, or must come after a Paragraph or Inline control.</exception>
        void InsertLinkListTable(List<DecisionModel.Models.Policy.Document> docs, StructuredDocumentTag control, bool useHeading = true, bool isBindLetter = false);

        /// <summary>
        /// Inserts the name amount table.
        /// </summary>
        /// <param name="fees">The fees.</param>
        /// <param name="control">The control.</param>
        /// <exception cref="Exception">Specified control must be a Paragraph or Inline control, or must come after a Paragraph or Inline control.</exception>
        void InsertNameAmountTable(IOrderedEnumerable<TaxFee> fees, StructuredDocumentTag control);

        /// <summary>
        /// Inserts the excess underlying coverage table.
        /// </summary>
        /// <param name="limits">The limits.</param>
        /// <param name="coverage">The coverage.</param>
        /// <param name="control">The control.</param>
        /// <exception cref="System.Exception">Specified control must be a Paragraph or Inline control, or must come after a Paragraph or Inline control.</exception>
        void InsertUnderlyingCoverageTable(List<Limit> limits, XsUnderlyingCoverage coverage, StructuredDocumentTag control);

        /// <summary>
        /// Insert XS auto liability vehicleType.
        /// </summary>
        /// <param name="vehicleTypes">The vehicle types.</param>
        /// <param name="control">The control.</param>
        void InsertXSAutoLiabilityVehicleType(IEnumerable<XsAutoLiabilityVehicleType> vehicleTypes, StructuredDocumentTag control);

        /// <summary>
        /// Inserts the exposures table.
        /// </summary>
        /// <param name="warranties">The warranties.</param>
        /// <param name="control">The control.</param>
        /// <exception cref="System.Exception">Specified control must be a Paragraph or Inline control, or must come after a Paragraph or Inline control.</exception>
        void InsertWarrantiesTable(IEnumerable<Warranty> warranties, StructuredDocumentTag control);

        /// <summary>
        /// Inserts the policy form table
        /// </summary>
        /// <param name="documents">Collection of documents</param>
        /// <param name="control">Book mark</param>
        void InsertPolicyFormTable(Dictionary<dynamic, Dictionary<string, DecisionModel.Models.Policy.Document>> documents, Aspose.Words.Bookmark control);

        /// <summary>
        /// Builds the mailing address
        /// </summary>
        /// <param name="streetAddress">streetAddress</param>
        /// <param name="control">Document Control</param>
        void InsertMailingAddress(Address streetAddress, StructuredDocumentTag control);

        /// <summary>
        /// Appends the specified dest.
        /// run this to create the document that needs to be run through the population process????
        /// </summary>
        /// <param name="dest">The dest.</param>
        /// <param name="source">The source.</param>
        void Append(string dest, string source);

        /// <summary>
        /// Appends a document to the one being managed and optionally inserts a page break before the new document.
        /// </summary>
        /// <param name="document">The document.</param>
        /// <param name="insertPageBreakBefore">if set to <c>true</c> [insert page break before].</param>
        /// <param name="restartNumbering">if set to <c>true</c> [restart numbering].</param>
        void AppendDocument(Aspose.Words.Document document, bool insertPageBreakBefore = true, bool restartNumbering = true);

        /// <summary>
        /// Updates the total page number footer field.
        /// </summary>
        void UpdateTotalPageNumberFooterField();

        /// <summary>
        /// Combine all source documents into destination document.
        /// </summary>
        /// <param name="dest">Destination file</param>
        /// <param name="source">Source files</param>
        void Combine(string dest, string[] source);

        /// <summary>
        /// Get a list of StructuredDocumentControls using the Tag property as an ID.
        /// </summary>
        /// <returns>
        /// List of controls
        /// </returns>
        List<StructuredDocumentTag> GetControls();

        /// <summary>
        /// Get the bookmarks from a word  template
        /// </summary>
        /// <returns>BookmarkCollection</returns>
        BookmarkCollection GetControlsBookMarks();

        /// <summary>
        /// Get the bookmark control from a word  template using bookmark name
        /// </summary>
        /// <param name="bookmarkName">The bookmark name.</param>
        /// <returns>Bookmark</returns>
        Bookmark GetBookMarkByName(string bookmarkName);

        /// <summary>
        /// Inserts the date.
        /// </summary>
        /// <param name="date">The date.</param>
        /// <param name="control">The control.</param>
        void InsertDate(DateTime date, StructuredDocumentTag control);

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
        void InsertLink(string link, string text, string controlId);

        /// <summary>
        /// Inserts the picture.
        /// </summary>
        /// <param name="imagePath">The image path.</param>
        /// <param name="control">The control.</param>
        void InsertPicture(string imagePath, StructuredDocumentTag control);

        /// <summary>
        /// Insert a table from a DataTable at the location specified by the controlId.
        /// </summary>
        /// <param name="dataTable">The data table.</param>
        /// <param name="controlId">ID of the control to use as a location for the table</param>
        /// <param name="useHeading">Use the headings from the DataTable. Default is true.</param>
        /// <exception cref="ArgumentException">
        /// DataTable must have rows.;dataTable
        /// or
        /// A controlId is required.;controlId
        /// </exception>
        void InsertTable(DataTable dataTable, string controlId, bool useHeading = true);

        /// <summary>
        /// Inserts the table.
        /// </summary>
        /// <param name="dataTable">The data table.</param>
        /// <param name="control">The control.</param>
        /// <param name="useHeading">if set to <c>true</c> [use heading].</param>
        /// <exception cref="Exception">Specified control must be a Paragraph or Inline control, or must come after a Paragraph or Inline control.</exception>
        void InsertTable(DataTable dataTable, StructuredDocumentTag control, bool useHeading = true);

        /// <summary>
        /// replace bookmark text
        /// </summary>
        /// <param name="placeHolder">placeHolder</param>
        /// <param name="replaceText">replaceText</param>
        void InsertFieldText(string placeHolder, string replaceText);

        /// <summary>
        /// Remove bookmark
        /// </summary>
        /// <param name="bookmark">Bookmark</param>
        void RemoveBookmark(Bookmark bookmark);

        /// <summary>
        /// Replaces the node value
        /// </summary>
        /// <param name="bookmark">bookmark</param>
        /// <param name="replaceText">replaceText</param>
        void ReplaceNodeValue(Bookmark bookmark, string replaceText);

        /// <summary>
        /// Insert string into Plain Text or markup into Rich Text
        /// </summary>
        /// <param name="value">String of text or Rich Text markup to insert</param>
        /// <param name="control">Plain Text or Rich Text control</param>
        void InsertText(string value, StructuredDocumentTag control);

        /// <summary>
        /// Inset a value into a Content Control within the Word Document by Tag. This will replace any value already in the Content Control.
        /// </summary>
        /// <param name="value">
        /// The value to insert into the control corresponding to the control type expected.
        /// Eg: String value for Text control. DateTime for Date Picker control, Image path for Picture Content control.
        /// </param>
        /// <param name="controlId">Tag value of Control used as ID. If more than one match is found, value will be inserted into all.</param>
        void InsertValue(object value, string controlId);

        /// <summary>
        /// Inset a value into a Content Control within the Word Document by Tag. This will replace any value already in the Content Control.
        /// </summary>
        /// <param name="value">The value to insert into the control corresponding to the control type expected.
        /// Eg: String value for Text control. DateTime for Date Picker control, Image path for Picture Content control.</param>
        /// <param name="control">The Content Control as StructuredDocumentTag to insert into.</param>
        /// <param name="suffix">The suffix.</param>
        /// <param name="text">The text.</param>
        /// <exception cref="System.ArgumentException">Value must not be empty.;value or The given control must exist.;control</exception>
        void InsertValue(object value, StructuredDocumentTag control, string suffix = "", string text = "");

        /// <summary>
        /// Determines whether [is block control] [the specified control].
        /// </summary>
        /// <param name="control">The control.</param>
        /// <returns>Boolean</returns>
        bool IsBlockControl(StructuredDocumentTag control);

        /// <summary>
        /// Find and remove all controls that match a given controlId
        /// </summary>
        /// <param name="controlId">Tag value of Control used as ID. If more than one match is found, all will be removed.</param>
        /// <exception cref="ArgumentException">A controlId is required.;controlId</exception>
        void Remove(string controlId);

        /// <summary>
        /// Save the current document to the file system at the path specified in a given file format. Default format is DOCX.
        /// </summary>
        /// <param name="directory">The directory.</param>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="fileFormat">Format to save the document as. Default is: DocX</param>
        void Save(string directory, string fileName, DocumentManager.FileFormat fileFormat = DocumentManager.FileFormat.DOCX);

        /// <summary>
        /// Saves the specified stream.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="fileFormat">The file format.</param>
        void Save(Stream stream, DocumentManager.FileFormat fileFormat = DocumentManager.FileFormat.DOCX);

        /// <summary>
        /// Saves the specified response.
        /// </summary>
        /// <param name="response">The response.</param>
        /// <param name="fileName">Name of the file.</param>
        void Save(HttpResponse response, string fileName);

        /// <summary>
        /// Generates Tiff from wordDoc and converts it to a PNG.
        /// </summary>
        /// <param name="docStream">The document stream.</param>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="saveToFile">if set to <c>true</c> [save to file].</param>
        /// <returns>
        /// the image in byte[] format
        /// </returns>
        object SaveToPng(MemoryStream docStream, string fileName, bool saveToFile = false);

        /// <summary>
        /// Set selected value in a Drop Down List or Combo Box
        /// </summary>
        /// <param name="selectedValue">Value in list to select. Set to NULL to clear selection</param>
        /// <param name="control">A Drop Down List or Combo Box</param>
        void SelectDropDown(string selectedValue, StructuredDocumentTag control);

        /// <summary>
        /// Strips the HTML from a string.
        /// </summary>
        /// <param name="text">The text.</param>
        /// <returns>string</returns>
        string StripHTML(string text);

        /// <summary>
        /// Toggles the CheckBox.
        /// </summary>
        /// <param name="isChecked">if set to <c>true</c> [is checked].</param>
        /// <param name="control">The control.</param>
        void ToggleCheckBox(bool isChecked, StructuredDocumentTag control);

        /// <summary>
        /// Inserts a watermark with the specified text into the generated document
        /// </summary>
        /// <param name="waterMarkText">Watermark text that appears in the background of the document</param>
        void InsertWaterMark(string waterMarkText = "SPECIMEN");

        /// <summary>
        /// Updates the repeater links.
        /// </summary>
        void UpdateRepeaterLinks();

        /// <summary>
        /// Removes from the document a table containing bookmark with the specified name
        /// </summary>
        /// <param name="bookmarkName">Bookmark name</param>
        /// <param name="removeParagraphAfterTable">If paragraph below the table should be removed</param>
        void RemoveTableWithBookmark(string bookmarkName, bool removeParagraphAfterTable);

        /// <summary>
        /// Replaces the question bookmark value.
        /// </summary>
        /// <param name="bookmarkName">Name of the bookmark.</param>
        /// <param name="question">The question.</param>
        void ReplaceQuestionBookmarkValue(string bookmarkName, Question question);

        /// <summary>
        /// Replaces the question bookmark value.
        /// </summary>
        /// <param name="bookmark">The bookmark.</param>
        /// <param name="question">The question.</param>
        void ReplaceQuestionBookmarkValue(Bookmark bookmark, Question question);

        /// <summary>
        /// Replaces the question bookmark value.
        /// </summary>
        /// <param name="bookmarkName">Name of the bookmark.</param>
        /// <param name="question">The question.</param>
        /// <param name="formatFunction">Format function applied to the text value</param>
        /// <param name="formatMethodInfo">Format function applied to the text value as a MethodInfo</param>
        void ReplaceQuestionBookmarkValue(string bookmarkName, Question question, Func<string, string> formatFunction, MethodInfo formatMethodInfo = null);

        /// <summary>
        /// Replaces the question bookmark value.
        /// </summary>
        /// <param name="bookmark">The bookmark.</param>
        /// <param name="question">The question.</param>
        /// <param name="formatFunction">Format function applied to the text value</param>
        /// <param name="formatMethodInfo">Format function applied to the text value as a MethodInfo</param>
        void ReplaceQuestionBookmarkValue(Bookmark bookmark, Question question, Func<string, string> formatFunction, MethodInfo formatMethodInfo = null);

        /// <summary>
        /// Gets the question value.
        /// </summary>
        /// <param name="bookmarkName">Name of the bookmark.</param>
        /// <param name="question">The question.</param>
        /// <returns>System.String.</returns>
        string GetQuestionValue(string bookmarkName, Question question);
    }
}