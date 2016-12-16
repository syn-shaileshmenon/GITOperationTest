// ***********************************************************************
// Assembly         : CorrespondenceServices.Tests
// Author           : rsteelea
// Created          : 08-16-2016
//
// Last Modified By : rsteelea
// Last Modified On : 08-16-2016
// ***********************************************************************
// <copyright file="DocumentManagerTests.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************

namespace CorrespondenceServices.Tests
{
    using System.Diagnostics.CodeAnalysis;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Class DocumentManagerTests.
    /// </summary>
    [TestClass]
    [ExcludeFromCodeCoverage]
    public class DocumentManagerTests
    {
        ////#region ReplaceQuestionBookmarkValue

        /////// <summary>
        /////// Determines whether this instance [can replace question CheckBox bookmark selected value].
        /////// </summary>
        ////[TestMethod]
        ////public void CanReplaceQuestionCheckBoxBookmarkSelectedValue()
        ////{
        ////    var bookmarkName = "MergeField_1";
        ////    var bookmark = IssuanceManagerTests.MockBookmark(bookmarkName);
        ////    var policyDocumentMock = new Mock<IPolicyDocumentManager>();
        ////    var expectedResult = string.Empty;

        ////    policyDocumentMock.Setup(s => s.GetBookMarkByName(It.IsAny<string>())).Returns(bookmark);
        ////    policyDocumentMock.Setup(s => s.ReplaceNodeValue(bookmark, It.IsAny<string>()))
        ////        .Callback((Bookmark b, string s) => expectedResult = s);

        ////    var answerOther = IssuanceManagerTests.CreateAnswer("MergeField_2", "Other");
        ////    var answerSomething = IssuanceManagerTests.CreateAnswer(bookmarkName, "Something", true);

        ////    var question = IssuanceManagerTests.CreateQuestion(
        ////        DisplayFormat.CheckBox,
        ////        "Test1",
        ////        "This is a checkbox question",
        ////        mergeFieldName: bookmarkName,
        ////        groupingNumber: 2,
        ////        answers: new List<Answer>(new[] { answerOther, answerSomething }));

        ////    var po = new PrivateType(typeof(IssuanceManager));
        ////    var args = new object[] { policyDocumentMock.Object, bookmarkName, question };
        ////    var replacementQuestion = po.InvokeStatic("ReplaceQuestionBookmarkValue", args);

        ////    Assert.AreEqual("CHECKED", expectedResult);
        ////}

        /////// <summary>
        /////// Tests that the not replace question CheckBox bookmark when not selected value.
        /////// </summary>
        ////[TestMethod]
        ////public void DoesNotReplaceQuestionCheckBoxBookmarkWhenNotSelectedValue()
        ////{
        ////    var bookmarkName = "MergeField_1";
        ////    var documentMock = new Mock<Aspose.Words.Document>();
        ////    var bookmarkStartMock = new Mock<BookmarkStart>(MockBehavior.Loose, documentMock.Object, bookmarkName);
        ////    var bookmarkEndMock = new Mock<BookmarkEnd>(MockBehavior.Loose, documentMock.Object, bookmarkName);
        ////    var bookmarkMock = new PrivateObject(typeof(Bookmark), bookmarkStartMock.Object, bookmarkEndMock.Object);
        ////    var bookmark = (Bookmark)bookmarkMock.Target;
        ////    var policyDocumentMock = new Mock<IPolicyDocumentManager>();
        ////    var expectedResult = string.Empty;

        ////    var manager = new DocumentManager(documentMock.Object);
        ////    policyDocumentMock.Setup(s => s.GetBookMarkByName(It.IsAny<string>())).Returns(bookmark);
        ////    policyDocumentMock.Setup(s => s.ReplaceNodeValue(bookmark, It.IsAny<string>()))
        ////        .Callback((Bookmark b, string s) => expectedResult = s);

        ////    var answerOtherAdditional = IssuanceManagerTests.CreateAnswer("MergeField_2", "Other");
        ////    var answerSomethingAdditional = IssuanceManagerTests.CreateAnswer(bookmarkName, "Something");

        ////    var question = IssuanceManagerTests.CreateQuestion(
        ////        DisplayFormat.CheckBox,
        ////        "Test1",
        ////        "This is a checkbox question",
        ////        mergeFieldName: bookmarkName,
        ////        groupingNumber: 2,
        ////        answers: new List<Answer>(new[] { answerSomethingAdditional, answerOtherAdditional }));

        ////    manager.ReplaceQuestionBookmarkValue(bookmarkName, question);

        ////    var po = new PrivateType(typeof(IssuanceManager));
        ////    var args = new object[] { policyDocumentMock.Object, bookmarkName, question };
        ////    var replacementQuestion = po.InvokeStatic("ReplaceQuestionBookmarkValue", args);

        ////    Assert.AreEqual(string.Empty, expectedResult);
        ////}

        /////// <summary>
        /////// Determines whether this instance [can replace question RadioButton bookmark selected value on question merge field name].
        /////// </summary>
        ////[TestMethod]
        ////public void CanReplaceQuestionRadioButtonBookmarkSelectedValueOnQuestionMergeFieldName()
        ////{
        ////    var bookmarkName = "MergeField_1";
        ////    var documentMock = new Mock<Aspose.Words.Document>();
        ////    var bookmarkStartMock = new Mock<BookmarkStart>(MockBehavior.Loose, documentMock.Object, bookmarkName);
        ////    var bookmarkEndMock = new Mock<BookmarkEnd>(MockBehavior.Loose, documentMock.Object, bookmarkName);
        ////    var bookmarkMock = new PrivateObject(typeof(Bookmark), bookmarkStartMock.Object, bookmarkEndMock.Object);
        ////    var bookmark = (Bookmark)bookmarkMock.Target;
        ////    var policyDocumentMock = new Mock<IPolicyDocumentManager>();
        ////    var expectedResult = string.Empty;

        ////    policyDocumentMock.Setup(s => s.GetBookMarkByName(It.IsAny<string>())).Returns(bookmark);
        ////    policyDocumentMock.Setup(s => s.ReplaceNodeValue(bookmark, It.IsAny<string>()))
        ////        .Callback((Bookmark b, string s) => expectedResult = s);

        ////    var answerOther = IssuanceManagerTests.CreateAnswer(null, "Other");
        ////    var answerSomething = IssuanceManagerTests.CreateAnswer(null, "Something", true);

        ////    var question = IssuanceManagerTests.CreateQuestion(
        ////        DisplayFormat.RadioButton,
        ////        "Test1",
        ////        "This is a radio button question",
        ////        mergeFieldName: bookmarkName,
        ////        groupingNumber: 2,
        ////        answerValue: "Something",
        ////        answers: new List<Answer>(new[] { answerSomething, answerOther }));

        ////    var po = new PrivateType(typeof(IssuanceManager));
        ////    var args = new object[] { policyDocumentMock.Object, bookmarkName, question };
        ////    var replacementQuestion = po.InvokeStatic("ReplaceQuestionBookmarkValue", args);

        ////    Assert.AreEqual("Something", expectedResult);
        ////}

        /////// <summary>
        /////// Determines whether this instance [can replace question RadioButton bookmark selected value on answer merge field name].
        /////// </summary>
        ////[TestMethod]
        ////public void CanReplaceQuestionRadioButtonBookmarkSelectedValueOnAnswerMergeFieldName()
        ////{
        ////    var bookmarkName = "MergeField_1";
        ////    var documentMock = new Mock<Aspose.Words.Document>();
        ////    var bookmarkStartMock = new Mock<BookmarkStart>(MockBehavior.Loose, documentMock.Object, bookmarkName);
        ////    var bookmarkEndMock = new Mock<BookmarkEnd>(MockBehavior.Loose, documentMock.Object, bookmarkName);
        ////    var bookmarkMock = new PrivateObject(typeof(Bookmark), bookmarkStartMock.Object, bookmarkEndMock.Object);
        ////    var bookmark = (Bookmark)bookmarkMock.Target;
        ////    var policyDocumentMock = new Mock<IPolicyDocumentManager>();
        ////    var expectedResult = string.Empty;

        ////    policyDocumentMock.Setup(s => s.GetBookMarkByName(It.IsAny<string>())).Returns(bookmark);
        ////    policyDocumentMock.Setup(s => s.ReplaceNodeValue(bookmark, It.IsAny<string>()))
        ////        .Callback((Bookmark b, string s) => expectedResult = s);

        ////    var answerOther = IssuanceManagerTests.CreateAnswer("MergeField_2", "Other");
        ////    var answerSomething = IssuanceManagerTests.CreateAnswer(bookmarkName, "Something", true);

        ////    var question = IssuanceManagerTests.CreateQuestion(
        ////        DisplayFormat.RadioButton,
        ////        "Test1",
        ////        "This is a radio button question",
        ////        groupingNumber: 2,
        ////        answerValue: "Something",
        ////        answers: new List<Answer>(new[] { answerSomething, answerOther }));

        ////    var po = new PrivateType(typeof(IssuanceManager));
        ////    var args = new object[] { policyDocumentMock.Object, bookmarkName, question };
        ////    var replacementQuestion = po.InvokeStatic("ReplaceQuestionBookmarkValue", args);

        ////    Assert.AreEqual("Something", expectedResult);
        ////}

        /////// <summary>
        /////// Tests that the not replace question RadioButton bookmark when not selected value.
        /////// </summary>
        ////[TestMethod]
        ////public void DoesNotReplaceQuestionRadioButtonBookmarkWhenNotSelectedValue()
        ////{
        ////    var bookmarkName = "MergeField_1";
        ////    var documentMock = new Mock<Aspose.Words.Document>();
        ////    var bookmarkStartMock = new Mock<BookmarkStart>(MockBehavior.Loose, documentMock.Object, bookmarkName);
        ////    var bookmarkEndMock = new Mock<BookmarkEnd>(MockBehavior.Loose, documentMock.Object, bookmarkName);
        ////    var bookmarkMock = new PrivateObject(typeof(Bookmark), bookmarkStartMock.Object, bookmarkEndMock.Object);
        ////    var bookmark = (Bookmark)bookmarkMock.Target;
        ////    var policyDocumentMock = new Mock<IPolicyDocumentManager>();
        ////    var expectedResult = string.Empty;

        ////    policyDocumentMock.Setup(s => s.GetBookMarkByName(It.IsAny<string>())).Returns(bookmark);
        ////    policyDocumentMock.Setup(s => s.ReplaceNodeValue(bookmark, It.IsAny<string>()))
        ////        .Callback((Bookmark b, string s) => expectedResult = s);

        ////    var answerOther = IssuanceManagerTests.CreateAnswer(null, "Other");
        ////    var answerSomething = IssuanceManagerTests.CreateAnswer(null, "Something");

        ////    var question = IssuanceManagerTests.CreateQuestion(
        ////        DisplayFormat.RadioButton,
        ////        "Test1",
        ////        "This is a radio button question",
        ////        mergeFieldName: bookmarkName,
        ////        groupingNumber: 2,
        ////        answers: new List<Answer>(new[] { answerSomething, answerOther }));

        ////    var po = new PrivateType(typeof(IssuanceManager));
        ////    var args = new object[] { policyDocumentMock.Object, bookmarkName, question };
        ////    var replacementQuestion = po.InvokeStatic("ReplaceQuestionBookmarkValue", args);

        ////    Assert.AreEqual(string.Empty, expectedResult);
        ////}

        /////// <summary>
        /////// Tests that the name of the not replace question RadioButton bookmark when not selected value on answer merge field.
        /////// </summary>
        ////[TestMethod]
        ////public void DoesNotReplaceQuestionRadioButtonBookmarkWhenNotSelectedValueOnAnswerMergeFieldName()
        ////{
        ////    var bookmarkName = "MergeField_1";
        ////    var documentMock = new Mock<Aspose.Words.Document>();
        ////    var bookmarkStartMock = new Mock<BookmarkStart>(MockBehavior.Loose, documentMock.Object, bookmarkName);
        ////    var bookmarkEndMock = new Mock<BookmarkEnd>(MockBehavior.Loose, documentMock.Object, bookmarkName);
        ////    var bookmarkMock = new PrivateObject(typeof(Bookmark), bookmarkStartMock.Object, bookmarkEndMock.Object);
        ////    var bookmark = (Bookmark)bookmarkMock.Target;
        ////    var policyDocumentMock = new Mock<IPolicyDocumentManager>();
        ////    var expectedResult = string.Empty;

        ////    policyDocumentMock.Setup(s => s.GetBookMarkByName(It.IsAny<string>())).Returns(bookmark);
        ////    policyDocumentMock.Setup(s => s.ReplaceNodeValue(bookmark, It.IsAny<string>()))
        ////        .Callback((Bookmark b, string s) => expectedResult = s);

        ////    var answerOther = IssuanceManagerTests.CreateAnswer("MergeField_2", "Other", true);
        ////    var answerSomething = IssuanceManagerTests.CreateAnswer(bookmarkName, "Something");

        ////    var question = IssuanceManagerTests.CreateQuestion(
        ////        DisplayFormat.RadioButton,
        ////        "Test1",
        ////        "This is a radio button question",
        ////        groupingNumber: 2,
        ////        answerValue: "Other",
        ////        answers: new List<Answer>(new[] { answerSomething, answerOther }));

        ////    var po = new PrivateType(typeof(IssuanceManager));
        ////    var args = new object[] { policyDocumentMock.Object, bookmarkName, question };
        ////    var replacementQuestion = po.InvokeStatic("ReplaceQuestionBookmarkValue", args);

        ////    Assert.AreEqual(string.Empty, expectedResult);
        ////}

        /////// <summary>
        /////// Determines whether this instance [can replace question drop down bookmark selected value on question merge field name].
        /////// </summary>
        ////[TestMethod]
        ////public void CanReplaceQuestionDropDownBookmarkSelectedValueOnQuestionMergeFieldName()
        ////{
        ////    var bookmarkName = "MergeField_1";
        ////    var documentMock = new Mock<Aspose.Words.Document>();
        ////    var bookmarkStartMock = new Mock<BookmarkStart>(MockBehavior.Loose, documentMock.Object, bookmarkName);
        ////    var bookmarkEndMock = new Mock<BookmarkEnd>(MockBehavior.Loose, documentMock.Object, bookmarkName);
        ////    var bookmarkMock = new PrivateObject(typeof(Bookmark), bookmarkStartMock.Object, bookmarkEndMock.Object);
        ////    var bookmark = (Bookmark)bookmarkMock.Target;
        ////    var policyDocumentMock = new Mock<IPolicyDocumentManager>();
        ////    var expectedResult = string.Empty;

        ////    policyDocumentMock.Setup(s => s.GetBookMarkByName(It.IsAny<string>())).Returns(bookmark);
        ////    policyDocumentMock.Setup(s => s.ReplaceNodeValue(bookmark, It.IsAny<string>()))
        ////        .Callback((Bookmark b, string s) => expectedResult = s);

        ////    var answerOther = IssuanceManagerTests.CreateAnswer(null, "Other");
        ////    var answerSomething = IssuanceManagerTests.CreateAnswer(null, "Something", true);

        ////    var question = IssuanceManagerTests.CreateQuestion(
        ////        DisplayFormat.DropDown,
        ////        "Test1",
        ////        "This is a dropdown question",
        ////        mergeFieldName: bookmarkName,
        ////        groupingNumber: 2,
        ////        answerValue: "Something",
        ////        answers: new List<Answer>(new[] { answerSomething, answerOther }));

        ////    var po = new PrivateType(typeof(IssuanceManager));
        ////    var args = new object[] { policyDocumentMock.Object, bookmarkName, question };
        ////    var replacementQuestion = po.InvokeStatic("ReplaceQuestionBookmarkValue", args);

        ////    Assert.AreEqual("Something", expectedResult);
        ////}

        /////// <summary>
        /////// Determines whether this instance [can replace question drop down bookmark selected value on answer merge field name].
        /////// </summary>
        ////[TestMethod]
        ////public void CanReplaceQuestionDropDownBookmarkSelectedValueOnAnswerMergeFieldName()
        ////{
        ////    var bookmarkName = "MergeField_1";
        ////    var documentMock = new Mock<Aspose.Words.Document>();
        ////    var bookmarkStartMock = new Mock<BookmarkStart>(MockBehavior.Loose, documentMock.Object, bookmarkName);
        ////    var bookmarkEndMock = new Mock<BookmarkEnd>(MockBehavior.Loose, documentMock.Object, bookmarkName);
        ////    var bookmarkMock = new PrivateObject(typeof(Bookmark), bookmarkStartMock.Object, bookmarkEndMock.Object);
        ////    var bookmark = (Bookmark)bookmarkMock.Target;
        ////    var policyDocumentMock = new Mock<IPolicyDocumentManager>();
        ////    var expectedResult = string.Empty;

        ////    policyDocumentMock.Setup(s => s.GetBookMarkByName(It.IsAny<string>())).Returns(bookmark);
        ////    policyDocumentMock.Setup(s => s.ReplaceNodeValue(bookmark, It.IsAny<string>()))
        ////        .Callback((Bookmark b, string s) => expectedResult = s);

        ////    var answerOther = IssuanceManagerTests.CreateAnswer("MergeField_2", "Other");
        ////    var answerSomething = IssuanceManagerTests.CreateAnswer(bookmarkName, "Something", true);

        ////    var question = IssuanceManagerTests.CreateQuestion(
        ////        DisplayFormat.DropDown,
        ////        "Test1",
        ////        "This is a dropdown question",
        ////        answerValue: "Something",
        ////        groupingNumber: 2,
        ////        answers: new List<Answer>(new[] { answerSomething, answerOther }));

        ////    var po = new PrivateType(typeof(IssuanceManager));
        ////    var args = new object[] { policyDocumentMock.Object, bookmarkName, question };
        ////    var replacementQuestion = po.InvokeStatic("ReplaceQuestionBookmarkValue", args);

        ////    Assert.AreEqual("Something", expectedResult);
        ////}

        /////// <summary>
        /////// Tests that the not replace question drop down bookmark when not selected value.
        /////// </summary>
        ////[TestMethod]
        ////public void DoesNotReplaceQuestionDropDownBookmarkWhenNotSelectedValue()
        ////{
        ////    var bookmarkName = "MergeField_1";
        ////    var documentMock = new Mock<Aspose.Words.Document>();
        ////    var bookmarkStartMock = new Mock<BookmarkStart>(MockBehavior.Loose, documentMock.Object, bookmarkName);
        ////    var bookmarkEndMock = new Mock<BookmarkEnd>(MockBehavior.Loose, documentMock.Object, bookmarkName);
        ////    var bookmarkMock = new PrivateObject(typeof(Bookmark), bookmarkStartMock.Object, bookmarkEndMock.Object);
        ////    var bookmark = (Bookmark)bookmarkMock.Target;
        ////    var policyDocumentMock = new Mock<IPolicyDocumentManager>();
        ////    var expectedResult = string.Empty;

        ////    policyDocumentMock.Setup(s => s.GetBookMarkByName(It.IsAny<string>())).Returns(bookmark);
        ////    policyDocumentMock.Setup(s => s.ReplaceNodeValue(bookmark, It.IsAny<string>()))
        ////        .Callback((Bookmark b, string s) => expectedResult = s);

        ////    var answerOther = IssuanceManagerTests.CreateAnswer(null, "Other");
        ////    var answerSomething = IssuanceManagerTests.CreateAnswer(null, "Something");

        ////    var question = IssuanceManagerTests.CreateQuestion(
        ////        DisplayFormat.DropDown,
        ////        "Test1",
        ////        "This is a dropdown question",
        ////        mergeFieldName: bookmarkName,
        ////        groupingNumber: 2,
        ////        answers: new List<Answer>(new[] { answerOther, answerSomething }));

        ////    var po = new PrivateType(typeof(IssuanceManager));
        ////    var args = new object[] { policyDocumentMock.Object, bookmarkName, question };
        ////    var replacementQuestion = po.InvokeStatic("ReplaceQuestionBookmarkValue", args);

        ////    Assert.AreEqual(string.Empty, expectedResult);
        ////}

        /////// <summary>
        /////// Tests that the name of the not replace question drop down bookmark when not selected value on answer merge field.
        /////// </summary>
        ////[TestMethod]
        ////public void DoesNotReplaceQuestionDropDownBookmarkWhenNotSelectedValueOnAnswerMergeFieldName()
        ////{
        ////    var bookmarkName = "MergeField_1";
        ////    var documentMock = new Mock<Aspose.Words.Document>();
        ////    var bookmarkStartMock = new Mock<BookmarkStart>(MockBehavior.Loose, documentMock.Object, bookmarkName);
        ////    var bookmarkEndMock = new Mock<BookmarkEnd>(MockBehavior.Loose, documentMock.Object, bookmarkName);
        ////    var bookmarkMock = new PrivateObject(typeof(Bookmark), bookmarkStartMock.Object, bookmarkEndMock.Object);
        ////    var bookmark = (Bookmark)bookmarkMock.Target;
        ////    var policyDocumentMock = new Mock<IPolicyDocumentManager>();
        ////    var expectedResult = string.Empty;

        ////    policyDocumentMock.Setup(s => s.GetBookMarkByName(It.IsAny<string>())).Returns(bookmark);
        ////    policyDocumentMock.Setup(s => s.ReplaceNodeValue(bookmark, It.IsAny<string>()))
        ////        .Callback((Bookmark b, string s) => expectedResult = s);

        ////    var answerOther = IssuanceManagerTests.CreateAnswer("MergeField_2", "Other", true);
        ////    var answerSomething = IssuanceManagerTests.CreateAnswer(bookmarkName, "Something");

        ////    var question = IssuanceManagerTests.CreateQuestion(
        ////        DisplayFormat.DropDown,
        ////        "Test1",
        ////        "This is a dropdown question",
        ////        groupingNumber: 0,
        ////        answerValue: "Other",
        ////        answers: new List<Answer>(new[] { answerSomething, answerOther }));

        ////    var po = new PrivateType(typeof(IssuanceManager));
        ////    var args = new object[] { policyDocumentMock.Object, bookmarkName, question };
        ////    var replacementQuestion = po.InvokeStatic("ReplaceQuestionBookmarkValue", args);

        ////    Assert.AreEqual(string.Empty, expectedResult);
        ////}

        /////// <summary>
        /////// Ensure the text answer from a text box replaces the bookmark
        /////// </summary>
        ////[TestMethod]
        ////public void CanReplaceQuestionTextBoxBookmark()
        ////{
        ////    var bookmarkName = "MergeField_1";
        ////    var documentMock = new Mock<Aspose.Words.Document>();
        ////    var bookmarkStartMock = new Mock<BookmarkStart>(MockBehavior.Loose, documentMock.Object, bookmarkName);
        ////    var bookmarkEndMock = new Mock<BookmarkEnd>(MockBehavior.Loose, documentMock.Object, bookmarkName);
        ////    var bookmarkMock = new PrivateObject(typeof(Bookmark), bookmarkStartMock.Object, bookmarkEndMock.Object);
        ////    var bookmark = (Bookmark)bookmarkMock.Target;
        ////    var policyDocumentMock = new Mock<IPolicyDocumentManager>();
        ////    var expectedResult = string.Empty;

        ////    policyDocumentMock.Setup(s => s.GetBookMarkByName(It.IsAny<string>())).Returns(bookmark);
        ////    policyDocumentMock.Setup(s => s.ReplaceNodeValue(bookmark, It.IsAny<string>()))
        ////        .Callback((Bookmark b, string s) => expectedResult = s);

        ////    var question = IssuanceManagerTests.CreateQuestion(
        ////        DisplayFormat.TextBox,
        ////        "Test1",
        ////        "This is a textbox question",
        ////        groupingNumber: 2,
        ////        answerValue: "A Text Answer");

        ////    var po = new PrivateType(typeof(IssuanceManager));
        ////    var args = new object[] { policyDocumentMock.Object, bookmarkName, question };
        ////    var replacementQuestion = po.InvokeStatic("ReplaceQuestionBookmarkValue", args);

        ////    Assert.AreEqual("A Text Answer", expectedResult);
        ////}

        ////#endregion
    }
}
