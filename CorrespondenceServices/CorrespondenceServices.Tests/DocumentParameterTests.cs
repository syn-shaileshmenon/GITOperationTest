// ***********************************************************************
// Assembly         : CorrespondenceServices.Tests
// Author           : rsteelea
// Created          : 11-01-2016
//
// Last Modified By : rsteelea
// Last Modified On : 11-01-2016
// ***********************************************************************
// <copyright file="DocumentParameterTests.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************
namespace CorrespondenceServices.Tests
{
    using System.Diagnostics.CodeAnalysis;
    using CorrespondenceServices.Classes.Parameters;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Mkl.WebTeam.DocumentGenerator;

    /// <summary>
    /// Class DocumentParameterTests.
    /// </summary>
    [TestClass]
    [ExcludeFromCodeCoverage]
    public class DocumentParameterTests
    {
        /// <summary>
        /// Determines whether this instance [can test document parameter initialization].
        /// </summary>
        [TestMethod]
        public void CanTestDocumentParameterInitialization()
        {
            var docParam = new DocumentParameters();
            Assert.IsNotNull(docParam);
            Assert.AreEqual(string.Empty, docParam.FileName);
            Assert.AreEqual(DocumentManager.FileFormat.DOCX, docParam.FileFormat);
            Assert.IsNotNull(docParam.NamedParameters);
            Assert.AreEqual(0, docParam.NamedParameters.Count);
        }

        /// <summary>
        /// Determines whether this instance [can test named value initialization].
        /// </summary>
        [TestMethod]
        public void CanTestNamedValueInitialization()
        {
            var named = new DocumentParameters.NamedValue();
            Assert.IsNotNull(named);
            Assert.IsNull(named.Name);
            Assert.IsNull(named.Value);
        }

        /// <summary>
        /// Determines whether this instance [can test named value get set name].
        /// </summary>
        public void CanTestNamedValueGetSetName()
        {
            const string Name = "Named";
            var named = new DocumentParameters.NamedValue();
            Assert.IsNotNull(named);
            Assert.IsNull(named.Name);

            named.Name = Name;
            Assert.AreEqual(Name, named.Name);
        }

        /// <summary>
        /// Determines whether this instance [can test named value get set value string].
        /// </summary>
        public void CanTestNamedValueGetSetValueString()
        {
            var named = new DocumentParameters.NamedValue();
            Assert.IsNotNull(named);
            Assert.IsNull(named.Value);

            named.Value = "Value";
            Assert.AreEqual("Value", named.Value);
        }

        /// <summary>
        /// Determines whether this instance [can test named value get set value int].
        /// </summary>
        public void CanTestNamedValueGetSetValueInt()
        {
            Assert.IsNotNull(new DocumentParameters.NamedValue());
            Assert.IsNull(new DocumentParameters.NamedValue().Value);
            new DocumentParameters.NamedValue().Value = 1;
            Assert.AreEqual(1, new DocumentParameters.NamedValue().Value);
        }

        /// <summary>
        /// Determines whether this instance [can test named value get set value object].
        /// </summary>
        public void CanTestNamedValueGetSetValueObject()
        {
            var value = new object();
            var named = new DocumentParameters.NamedValue();
            Assert.IsNotNull(named);
            Assert.IsNull(named.Value);

            named.Value = value;
            Assert.AreEqual(value, named.Value);
        }
    }
}
