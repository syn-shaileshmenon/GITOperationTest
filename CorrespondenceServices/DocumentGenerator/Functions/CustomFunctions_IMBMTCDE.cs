// ***********************************************************************
// Assembly         : Mkl.WebTeam.DocumentGenerator
// Author           : rsteelea
// Created          : 11-07-2016
//
// Last Modified By : rsteelea
// Last Modified On : 11-09-2016
// ***********************************************************************
// <copyright file="CustomFunctions_IMBMTCDE.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************

namespace Mkl.WebTeam.DocumentGenerator.Functions
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using Aspose.Words;
    using Aspose.Words.Tables;
    using DecisionModel.Models.Policy;
    using Mkl.WebTeam.DocumentGenerator.Helpers;

    /// <summary>
    /// Class CustomFunctions.
    /// </summary>
    public partial class CustomFunctions
    {
        /// <summary>
        /// Generates the IMB MTC-DE 07 03 form.
        /// </summary>
        /// <param name="policy">The policy.</param>
        /// <param name="documentManager">The document manager.</param>
        /// <param name="extraParameters">The extra parameters.</param>
        public void GenerateIMBMTCDE_0703Form(Policy policy, ref IPolicyDocumentManager documentManager, Dictionary<string, object> extraParameters)
        {
            var questionCodeList = new List<string> { "DriverName", "DriverState", "DriverLicenseNumber" };
            var blankMergeFieldList = new List<List<string>>
            {
                new List<string> { string.Empty, "3", "6", "9", "12" },
                new List<string> { "1", "4", "7", "10", "13" },
                new List<string> { "2", "5", "8", "11", "14" },
            };
            var handlerList = new List<CellCallbackType> { null, this.HandleStateSymbolsCallbackType, null };

            ProcessCustomTable(policy, documentManager, blankMergeFieldList, questionCodeList, cellHandlers: handlerList);
        }

        /// <summary>
        /// Handles the state callback.
        /// </summary>
        /// <param name="manager">The manager.</param>
        /// <param name="table">The table.</param>
        /// <param name="row">The row.</param>
        /// <param name="rowNumber">The row number.</param>
        /// <param name="columnNumber">The column number.</param>
        /// <param name="question">The question.</param>
        /// <param name="bookmark">The bookmark.</param>
        /// <param name="bookmarkName">Name of the bookmark.</param>
        /// <returns><c>true</c> if XXXX, <c>false</c> otherwise.</returns>
        public bool HandleStateSymbolsCallbackType(IPolicyDocumentManager manager, Table table, Row row, int rowNumber, int columnNumber, Question question, Bookmark bookmark, string bookmarkName)
        {
            var answer = question?.Answers?.Where(a => a.Value.EqualsIgnoreCase(question?.AnswerValue)).Select(a => a.Verbiage).FirstOrDefault();
            manager.ReplaceNodeValue(bookmark, answer);

            return false;
        }
    }
}
