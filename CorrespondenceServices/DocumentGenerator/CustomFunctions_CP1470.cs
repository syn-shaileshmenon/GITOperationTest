// <copyright file="CustomFunctions_CP1470.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>

namespace Mkl.WebTeam.DocumentGenerator.Functions
{
    using System.Collections.Generic;
    using System.Linq;
    using Aspose.Words;
    using DecisionModel.Models.Policy;

    /// <summary>
    /// Class CustomFunctions_CP1420
    /// </summary>
    public partial class CustomFunctions
    {
        /// <summary>
        /// Generates the CP 14 70 10 12 form.
        /// </summary>
        /// <param name="policy">The policy.</param>
        /// <param name="documentManager">The document manager.</param>
        /// <param name="extraParameters">The extra parameters.</param>
        public void GenerateCP1470Form(Policy policy, ref IPolicyDocumentManager documentManager, Dictionary<string, object> extraParameters)
        {
            var questionCodeList = "CausesofLossForm";

            var builder = new DocumentBuilder(documentManager.Document);

            var formNumber = documentManager.FormNormalizedNumber;
            var quantityOrder = documentManager.QuantityOrder;
            var questions = policy.RollUp?.Documents?.FirstOrDefault(d => d.NormalizedNumber == formNumber &&
                                                                          d.QuantityOrder.GetValueOrDefault(0) == quantityOrder)?.Questions?.OrderBy(o => o.MultipleRowGroupingNumber)
                .ToList();

            if (questions == null ||
                !questions.Any())
            {
                return;
            }

            var groupingNumbers = questions.Where(s => s.MultipleRowGroupingNumber > 0).Select(s => s.MultipleRowGroupingNumber).Distinct().ToList();
            int blankMergeFieldCnt = 2;

            var causeofLoss = string.Empty;
            foreach (var groupNbr in groupingNumbers)
            {
                var questionRows = questions.Where(s => s.MultipleRowGroupingNumber == groupNbr).ToList();
                var baseQuestion = questionRows?.Where(s => s.AnswerValue != null && s.Code == questionCodeList).FirstOrDefault();

                if (baseQuestion != null)
                {
                    var bookmarkName = "BlankMergeField_" + blankMergeFieldCnt;
                    documentManager.ReplaceQuestionBookmarkValue(bookmarkName, baseQuestion, null);
                    blankMergeFieldCnt++;
                }
            }
        }
    }
}
