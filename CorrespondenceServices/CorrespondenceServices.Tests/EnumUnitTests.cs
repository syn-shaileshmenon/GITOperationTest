// ***********************************************************************
// Assembly         : CorrespondenceServices.Tests
// Author           : rsteelea
// Created          : 11-02-2016
//
// Last Modified By : rsteelea
// Last Modified On : 11-02-2016
// ***********************************************************************
// <copyright file="EnumUnitTests.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************
namespace CorrespondenceServices.Tests
{
    using System;
    using System.Diagnostics.CodeAnalysis;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Mkl.WebTeam.DocumentGenerator;
    using Mkl.WebTeam.DocumentGenerator.Enumerations;

    /// <summary>
    /// Class EnumUnitTests.
    /// </summary>
    [TestClass]
    [ExcludeFromCodeCoverage]
    public class EnumUnitTests
    {
        /// <summary>
        /// Determines whether this instance [can test all enumerated values].
        /// </summary>
        [TestMethod]
        public void CanTestAllEnumeratedValues()
        {
            var formatDataTypes = Enum.GetValues(typeof(FormatDataType));
            Assert.AreEqual(6, formatDataTypes.Length);
            Assert.IsNotNull(FormatDataType.Currency.ToString());
            Assert.IsNotNull(FormatDataType.Decimal.ToString());
            Assert.IsNotNull(FormatDataType.Longdate.ToString());
            Assert.IsNotNull(FormatDataType.Number.ToString());
            Assert.IsNotNull(FormatDataType.Percent.ToString());
            Assert.IsNotNull(FormatDataType.Shortdate.ToString());

            var fileFormats = Enum.GetValues(typeof(DocumentManager.FileFormat));
            Assert.AreEqual(8, fileFormats.Length);
            Assert.IsNotNull(DocumentManager.FileFormat.None);
            Assert.IsNotNull(DocumentManager.FileFormat.DOCX);
            Assert.IsNotNull(DocumentManager.FileFormat.BMP);
            Assert.IsNotNull(DocumentManager.FileFormat.DOC);
            Assert.IsNotNull(DocumentManager.FileFormat.JPEG);
            Assert.IsNotNull(DocumentManager.FileFormat.PDF);
            Assert.IsNotNull(DocumentManager.FileFormat.PNG);
            Assert.IsNotNull(DocumentManager.FileFormat.TIFF);
        }
    }
}
