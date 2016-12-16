// <copyright file="CustomFunctions_MECP1288.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>

namespace Mkl.WebTeam.DocumentGenerator.Functions
{
    using System.Collections.Generic;
    using System.Linq;
    using Aspose.Words;
    using Aspose.Words.Tables;
    using DecisionModel.Models.Policy;
    using Mkl.WebTeam.DocumentGenerator.Helpers;

    /// <summary>
    /// Class CustomFunctions_MECP1288
    /// </summary>
    public partial class CustomFunctions
    {
        /// <summary>
        /// Generates the MECP 1288 form.
        /// </summary>
        /// <param name="policy">The policy.</param>
        /// <param name="documentManager">The document manager.</param>
        /// <param name="extraParameters">The extra parameters.</param>
        public void GenerateMECP_1288Form(Policy policy, ref IPolicyDocumentManager documentManager, Dictionary<string, object> extraParameters)
        {
            List<string> blankMergeFieldList = new List<string> { string.Empty, "3", "6" };

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
            ////var questions = policy.RollUp.Documents.Where(d => d.NormalizedNumber == "")

            var questions = policy.RollUp.Documents.FirstOrDefault(x => x.NormalizedNumber.Equals("MECP1288"))?.Questions;
            var questionsRowCount = policy.RollUp.Documents.FirstOrDefault(x => x.NormalizedNumber.Equals("MECP1288"))?.Questions?.Count(x => x.Code.Equals("PremisesNo"));

            var premisesNumberQuestions = questions.Where(x => x.Code.Equals("PremisesNo")).Select(x => x).OrderBy(x => x.MultipleRowGroupingNumber).ToList();
            var buildingNumberQuestions = questions.Where(x => x.Code.Equals("BuildingNo")).Select(x => x).OrderBy(x => x.MultipleRowGroupingNumber).ToList();
            var indicateApplicabilityQuestions = questions.Where(x => x.Code.Equals("IndicateApplicability")).Select(x => x).OrderBy(x => x.MultipleRowGroupingNumber).ToList();

            for (var rowNumber = 0; rowNumber < questionsRowCount; rowNumber++)
            {
                Row workingRow;
                ////if (rowNumber <= blankMergeFieldList[0].Count)
                if (rowNumber < blankMergeFieldList.Count)
                {
                    ////var number = blankMergeFieldList[0][rowNumber - 1];
                    var number = blankMergeFieldList[rowNumber];
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

                var premisesNumber = premisesNumberQuestions.Count() > rowNumber ? this.GetQuestionValue(premisesNumberQuestions[rowNumber]) : string.Empty;
                var buildingNumber = buildingNumberQuestions.Count() > rowNumber ? this.GetQuestionValue(buildingNumberQuestions[rowNumber]) : string.Empty;
                var indicateApplicability = indicateApplicabilityQuestions.Count() > rowNumber ? this.GetQuestionValue(indicateApplicabilityQuestions[rowNumber]) : string.Empty;

                var value = $"{premisesNumber}";  // Some questions answer
                this.ReplaceCellParagraphOrBookmark(ref documentManager, workingRow.Cells[0], value);

                value = $"{buildingNumber}";
                this.ReplaceCellParagraphOrBookmark(ref documentManager, workingRow.Cells[1], value);

                value = $"{indicateApplicability}";
                this.ReplaceCellParagraphOrBookmark(ref documentManager, workingRow.Cells[2], value);
            }
        }
    }
}
