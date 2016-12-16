// ***********************************************************************
// Assembly         : Mkl.WebTeam.DocumentGenerator
// Author           : smenon
// Created          : 11-12-2016
//
// Last Modified By : smenon
// Last Modified On : 11-12-2016
// ***********************************************************************
// <copyright file="CustomFunctions_IMBMTCDE.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************

namespace Mkl.WebTeam.DocumentGenerator.Functions
{
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
        public void GenerateMECP1281Form1(Policy policy, ref IPolicyDocumentManager documentManager, Dictionary<string, object> extraParameters)
        {
            List<string> blankMergeFieldList = new List<string> { string.Empty, "2", "4" };
            ////List<List<string>> blankMergeFieldList = new List<List<string>>
            ////{
            ////    new List<string> { string.Empty, "3", "6", "9", "12" },
            ////    new List<string> { "1", "4", "7", "10", "13" },
            ////    new List<string> { "2", "5", "8", "11", "14" },
            ////};

            var builder = new DocumentBuilder(documentManager.Document);

            builder.MoveToBookmark("BlankMergeField");
            var table = builder.CurrentNode.GetParentTable();
            if (table == null)
            {
                // this is an error...create the table manually?
            }

            // Fields are as follows:
            // Driver Name
            // State
            // License No
            var formNumber = documentManager.FormNormalizedNumber;
            var quantityOrder = documentManager.QuantityOrder;
            var questions = policy.RollUp.Documents.Where(d => d.NormalizedNumber == formNumber);

            for (var rowNumber = 1; rowNumber <= 10; rowNumber++)
            {
                Row workingRow;
                ////if (rowNumber <= blankMergeFieldList[0].Count)
                if (rowNumber <= blankMergeFieldList.Count)
                {
                    ////var number = blankMergeFieldList[0][rowNumber - 1];
                    var number = blankMergeFieldList[rowNumber - 1];
                    if (!string.IsNullOrWhiteSpace(number))
                    {
                        number = $"_{number}";
                    }

                    builder.MoveToBookmark($"BlankMergeField{number}");
                    workingRow = builder.CurrentNode.GetCurrentRow();
                }
                else
                {
                    workingRow = (Row)table.LastRow.Clone(true);
                    table.AppendChild(workingRow);
                }

                if (workingRow == null)
                {
                    continue;
                }

                var value = $"Premises Number {rowNumber}";  // Some questions answer
                this.ReplaceCellParagraphOrBookmark(ref documentManager, workingRow.Cells[0], value);

                value = $"Building Number {rowNumber}";
                this.ReplaceCellParagraphOrBookmark(ref documentManager, workingRow.Cells[1], value);
            }
        }
    }
}
