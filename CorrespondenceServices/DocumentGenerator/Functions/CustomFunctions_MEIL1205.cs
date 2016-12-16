// ***********************************************************************
// Assembly         : Mkl.WebTeam.DocumentGenerator
// Author           : rsteelea
// Created          : 11-14-2016
//
// Last Modified By : rsteelea
// Last Modified On : 11-14-2016
// ***********************************************************************
// <copyright file="CustomFunctions_MEIL1205.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************

/// <summary>
/// The Functions namespace.
/// </summary>
namespace Mkl.WebTeam.DocumentGenerator.Functions
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics.CodeAnalysis;
    using System.Linq;
    using Aspose.Words;
    using Aspose.Words.Tables;
    using DecisionModel.Models.Policy;
    using Mkl.WebTeam.CommonCore.Extensions;

    /// <summary>
    /// Class CustomFunctions.
    /// </summary>
    public partial class CustomFunctions
    {
        /// <summary>
        /// Generates the meg L0263 0314 form.
        /// </summary>
        /// <param name="policy">The policy.</param>
        /// <param name="documentManager">The document manager.</param>
        /// <param name="extraParameters">The extra parameters.</param>
        [SuppressMessage("StyleCop.CSharp.SpacingRules", "SA1003:SymbolsMustBeSpacedCorrectly", Justification = "?. can split across multiple lines")]
        public void GenerateMEIL1205_0713Form(Policy policy, ref IPolicyDocumentManager documentManager, Dictionary<string, object> extraParameters)
        {
            var formNumber = documentManager.FormNormalizedNumber;
            var quantityOrder = documentManager.QuantityOrder;

            List<string> questionCodeList = new List<string> { "PremisesNo", "BuildingNo", "ProtectiveSafeguardSymbols" };
            List<List<string>> blankMergeFieldList = new List<List<string>>
            {
                new List<string> { "1", "4", "7", "10", "13" },
                new List<string> { "2", "5", "8", "11", "14" },
                new List<string> { "3", "6", "9", "12", "15" },
            };

            ProcessCustomTable(policy, documentManager, blankMergeFieldList, questionCodeList, cellHandlers: new List<CellCallbackType> { null, null, this.HandleProtectiveSafeguardSymbolsCallbackType });

            // Handle section A check boxes
            var answers = policy.RollUp
                                    ?.Documents
                                    ?.FirstOrDefault(d => d.NormalizedNumber == formNumber &&
                                                        d.QuantityOrder.GetValueOrDefault(0) == quantityOrder)
                                    ?.Questions
                                    ?.Where(q => q.Code == "ProtectiveSafeguardSymbols" &&
                                                    q.Answers.Any(a => a.IsSelected))
                                    .SelectMany(s => s.Answers.FindAll(a => a.IsSelected))
                                    .Select(s => s.Value)
                                    .Distinct()
                                    .OrderBy(o => o)
                                    .ToList();

            if (answers == null ||
                !answers.Any())
            {
                return;
            }

            foreach (var answer in answers)
            {
                Bookmark bookmark;
                switch (answer.ToLowerInvariant())
                {
                    case "a":
                        bookmark = documentManager.GetBookMarkByName("Check8");
                        documentManager.ReplaceNodeValue(bookmark, "CHECKED");
                        break;

                    case "b":
                        bookmark = documentManager.GetBookMarkByName("Check9");
                        documentManager.ReplaceNodeValue(bookmark, "CHECKED");
                        break;

                    case "c":
                        bookmark = documentManager.GetBookMarkByName("Check10");
                        documentManager.ReplaceNodeValue(bookmark, "CHECKED");
                        break;

                    case "d":
                        bookmark = documentManager.GetBookMarkByName("Check11");
                        documentManager.ReplaceNodeValue(bookmark, "CHECKED");
                        break;

                    case "e":
                        bookmark = documentManager.GetBookMarkByName("Check12");
                        documentManager.ReplaceNodeValue(bookmark, "CHECKED");
                        break;

                    case "f":
                        bookmark = documentManager.GetBookMarkByName("Check13");
                        documentManager.ReplaceNodeValue(bookmark, "CHECKED");
                        break;

                    case "g":
                        bookmark = documentManager.GetBookMarkByName("Check14");
                        documentManager.ReplaceNodeValue(bookmark, "CHECKED");
                        break;

                    case "h":
                        bookmark = documentManager.GetBookMarkByName("Check15");
                        documentManager.ReplaceNodeValue(bookmark, "CHECKED");
                        break;

                    case "i":
                        bookmark = documentManager.GetBookMarkByName("Check16");
                        documentManager.ReplaceNodeValue(bookmark, "CHECKED");
                        break;
                }
            }
        }

        /// <summary>
        /// Handles the type of the protective safeguard symbols callback.
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
        public bool HandleProtectiveSafeguardSymbolsCallbackType(IPolicyDocumentManager manager, Table table, Row row, int rowNumber, int columnNumber, Question question, Bookmark bookmark, string bookmarkName)
        {
            var answers = question?.Answers.Where(a => a.IsSelected)
                .Select(s => s.Value)
                .Distinct()
                .ToArray()
                .Join(s => s, ",");
            manager.ReplaceNodeValue(bookmark, answers);

            return false;
        }
    }
}
