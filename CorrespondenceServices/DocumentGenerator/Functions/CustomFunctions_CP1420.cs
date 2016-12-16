// <copyright file="CustomFunctions_CP1420.cs" company="Markel">
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
        /// Generates the CP 14 20 01 91 form.
        /// </summary>
        /// <param name="policy">The policy.</param>
        /// <param name="documentManager">The document manager.</param>
        /// <param name="extraParameters">The extra parameters.</param>
        public void GenerateCP1420_Form(Policy policy, ref IPolicyDocumentManager documentManager, Dictionary<string, object> extraParameters)
        {
            var questionCodeList = new List<string> { "PremisesNo", "BuildingNo", "PerformedBY", "PropDescLimit" };

            List<KeyValuePair<string, CP1420ORow>> answerMapping = new List<KeyValuePair<string, CP1420ORow>>()
            {
               new KeyValuePair<string, CP1420ORow>("CP_Canopy", new CP1420ORow(1, 0, 1)),
               new KeyValuePair<string, CP1420ORow>("CP_Brick", new CP1420ORow(2, 2, 3)),
               new KeyValuePair<string, CP1420ORow>("CP_Crop", new CP1420ORow(3, 4, 5)),
               new KeyValuePair<string, CP1420ORow>("CP_Swimming", new CP1420ORow(4, 6, 7)),
               new KeyValuePair<string, CP1420ORow>("CP_Waterwheel", new CP1420ORow(5, 8, 9)),
               new KeyValuePair<string, CP1420ORow>("CP_Repair", new CP1420ORow(6, 10, 11)),
               new KeyValuePair<string, CP1420ORow>("CP_Personal", new CP1420ORow(7, 12, 13)),
               new KeyValuePair<string, CP1420ORow>("CP_Content", new CP1420ORow(8, 14, 15)),
               new KeyValuePair<string, CP1420ORow>("CP_Glass", new CP1420ORow(9, 16, 17)),
               new KeyValuePair<string, CP1420ORow>("CP_Metal", new CP1420ORow(10, 18, 19)),
               new KeyValuePair<string, CP1420ORow>("CP_Ores", new CP1420ORow(11, 20, 21)),
               new KeyValuePair<string, CP1420ORow>("CP_Others", new CP1420ORow(12, 22, 23)),
               new KeyValuePair<string, CP1420ORow>("CP_OpenYard", new CP1420ORow(13, 24, 25)),
               new KeyValuePair<string, CP1420ORow>("CP_SignPrem", new CP1420ORow(14, 26, 27)),
               new KeyValuePair<string, CP1420ORow>("CP_Vending", new CP1420ORow(15, 28, 29)),
               new KeyValuePair<string, CP1420ORow>("CP_Tenants", new CP1420ORow(16, 30, 31)),
               new KeyValuePair<string, CP1420ORow>("CP_Stock", new CP1420ORow(17, 32, 33)),
               new KeyValuePair<string, CP1420ORow>("CP_BldgCook", new CP1420ORow(18, 34, 35)),
               new KeyValuePair<string, CP1420ORow>("CP_BldgRepair", new CP1420ORow(19, 36, 37)),
               new KeyValuePair<string, CP1420ORow>("CP_BldgMotor", new CP1420ORow(20, 38, 39)),
               new KeyValuePair<string, CP1420ORow>("CP_Sales", new CP1420ORow(21, 40, 41)),
               new KeyValuePair<string, CP1420ORow>("CP_Petro", new CP1420ORow(22, 42, 43))
            };

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

            foreach (var groupNbr in groupingNumbers)
            {
                var questionRows = questions.Where(s => s.MultipleRowGroupingNumber == groupNbr).ToList();
                var baseQuestion = questionRows?.Where(s => s.AnswerValue != null && s.Code == "PropDescLimit" && s.AnswerValue.Contains("CP_")).FirstOrDefault();
                var premQuestion = questionRows?.Where(s => s.Code.Contains("PremisesNo")).FirstOrDefault();
                var bldgQuestion = questionRows?.Where(s => s.Code.Contains("BuildingNo")).FirstOrDefault();

                var mappingEntity = answerMapping.Where(a => a.Key == baseQuestion?.AnswerValue).FirstOrDefault();

                if (baseQuestion != null)
                {
                    documentManager.ReplaceQuestionBookmarkValue(mappingEntity.Value.CheckField, baseQuestion, null);
                }

                if (premQuestion != null)
                {
                    documentManager.ReplaceQuestionBookmarkValue(mappingEntity.Value.PremiseNoField, premQuestion, null);
                }

                if (bldgQuestion != null)
                {
                    documentManager.ReplaceQuestionBookmarkValue(mappingEntity.Value.BldgNoField, bldgQuestion, null);
                }

                if (baseQuestion.Code == "PropDescLimit" && baseQuestion.AnswerValue == "CP_Repair")
                {
                    documentManager.ReplaceQuestionBookmarkValue("BlankMergeField_44", questionRows?.Where(s => s.Code.Contains("PerformedBY")).FirstOrDefault(), null);
                }
            }
        }

        /// <summary>
        /// Class depicting the mapping information of a row in MECP 1420
        /// </summary>
        private class CP1420ORow
        {
            /// <summary>
            /// Initializes a new instance of the <see cref="CP1420ORow"/> class.
            /// </summary>
            /// <param name="checkfield">check field</param>
            /// <param name="premField">premises number Field</param>
            /// <param name="bldgField">building Field</param>
            public CP1420ORow(int checkfield, int premField, int bldgField)
            {
                this.CheckField = "Check" + checkfield;
                if (premField == 0)
                {
                    this.PremiseNoField = "BlankMergeField";
                }
                else
                {
                    this.PremiseNoField = "BlankMergeField_" + premField;
                }

                this.BldgNoField = "BlankMergeField_" + bldgField;
            }

            /// <summary>
            /// Gets or sets the checkbox mapping detail for a limit
            /// </summary>
            public string CheckField { get; set; }

            /// <summary>
            /// Gets or sets the premises number mapping detail
            /// </summary>
            public string PremiseNoField { get; set; }

            /// <summary>
            /// Gets or sets the building number mapping detail
            /// </summary>
            public string BldgNoField { get; set; }
        }
    }
}
