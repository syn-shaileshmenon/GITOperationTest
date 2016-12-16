// ***********************************************************************
// Assembly         : CorrespondenceServices.Tests
// Author           : rsteelea
// Created          : 11-02-2016
//
// Last Modified By : rsteelea
// Last Modified On : 11-02-2016
// ***********************************************************************
// <copyright file="DownloadParametersTests.cs" company="Markel">
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
    /// Class DownloadParametersTests.
    /// </summary>
    [TestClass]
    [ExcludeFromCodeCoverage]
    public class DownloadParametersTests
    {
        /// <summary>
        /// Determines whether this instance [can test download parameters initialize].
        /// </summary>
        [TestMethod]
        public void CanTestDownloadParametersInitialize()
        {
            var param = new DownloadParameters();
            Assert.IsNotNull(param);
            Assert.IsNull(param.TempFileName);
            Assert.AreEqual(DocumentManager.FileFormat.DOCX, param.SaveFormat);
        }

        /// <summary>
        /// Determines whether this instance [can test download parameters temporary file name get set].
        /// </summary>
        [TestMethod]
        public void CanTestDownloadParametersTempFileNameGetSet()
        {
            const string TempFileName = "TempFileName";
            var param = new DownloadParameters();
            Assert.IsNotNull(param);
            Assert.IsNull(param.TempFileName);
            Assert.AreEqual(DocumentManager.FileFormat.DOCX, param.SaveFormat);

            param.TempFileName = TempFileName;
            Assert.AreEqual(TempFileName, param.TempFileName);
            Assert.AreEqual(DocumentManager.FileFormat.DOCX, param.SaveFormat);
        }

        /// <summary>
        /// Determines whether this instance [can test download parameters save format get set].
        /// </summary>
        [TestMethod]
        public void CanTestDownloadParamatersSaveFormatGetSet()
        {
            var saveFormat = DocumentManager.FileFormat.None;
            var param = new DownloadParameters();
            Assert.IsNotNull(param);
            Assert.IsNull(param.TempFileName);
            Assert.AreEqual(DocumentManager.FileFormat.DOCX, param.SaveFormat);

            param.SaveFormat = saveFormat;
            Assert.AreEqual(saveFormat, param.SaveFormat);
            Assert.IsNull(param.TempFileName);
        }
    }
}
