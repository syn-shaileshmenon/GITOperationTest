// <copyright file="CustomFunctions_CP0320.cs" company="Markel">
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
        /// Formats the CP0320.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <param name="manager">The manager.</param>
        /// <param name="question">The question.</param>
        /// <param name="bookmark">The bookmark.</param>
        /// <param name="bookmarkName">Name of the bookmark.</param>
        /// <param name="policy">The policy.</param>
        /// <returns>System.String.</returns>
        public string Format_CP0320(string value, IPolicyDocumentManager manager, Question question, Bookmark bookmark, string bookmarkName, Policy policy)
        {
            string result = value;

            if (question.Code.EqualsIgnoreCase("CoveredCausesOfLoss"))
            {
                result = this.GetSelectedCoveredCausesOfLoss(question);
            }

            ////else if (question.Code.EqualsIgnoreCase("Deductible"))
            ////{
            ////    int deductible = 0;
            ////    if (int.TryParse(value, out deductible))
            ////    {
            ////        result = FormatHelper.FormatValue(deductible, FormatHelper.CurrencyFormat);
            ////    }
            ////}

            return result;
        }

        /// <summary>
        /// Returns the comma separated string of selected causes of loss
        /// </summary>
        /// <param name="question">Question</param>
        /// <returns>Comma separated string of selected causes of loss</returns>
        private string GetSelectedCoveredCausesOfLoss(Question question)
        {
            var selectedAnswers = new List<string>();

            foreach (var answer in question.Answers)
            {
                if (answer.IsSelected)
                {
                    switch (answer.Value.ToUpper())
                    {
                        case "COVEREDCAUSESOFLOSS":
                            selectedAnswers.Add("1");
                            break;
                        case "COVEREDCAUSESOFLOSSEXCEPTWINDSTORMORHAIL":
                            selectedAnswers.Add("2");
                            break;
                        case "COVEREDCAUSESOFLOSSEXCEPTTHEFT":
                            selectedAnswers.Add("3");
                            break;
                        case "COVEREDCAUSESOFLOSSEXCEPTWINDSTORMORHAILANDTHEFT":
                            selectedAnswers.Add("4");
                            break;
                        case "WINDSTORMORHAIL":
                            selectedAnswers.Add("5");
                            break;
                        case "THEFT":
                            selectedAnswers.Add("6");
                            break;
                    }
                }
            }

            return selectedAnswers.Any() ? string.Join(",", selectedAnswers) : string.Empty;
        }

        /// <summary>
        /// Gets the multiple document instance index
        /// </summary>
        /// <param name="extraParameters">Extra parameters for the custom function</param>
        /// <returns>Document instance index if present, null otherwise.</returns>
        private int? GetMultipleDocumentInstanceIndex(Dictionary<string, object> extraParameters)
        {
            int? index = null;

            if (extraParameters != null && extraParameters.ContainsKey("DocumentInstanceIndex"))
            {
                index = (int)extraParameters["DocumentInstanceIndex"];
            }

            return index;
        }
    }
}
