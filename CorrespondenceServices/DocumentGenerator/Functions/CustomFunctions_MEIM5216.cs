// <copyright file="CustomFunctions_MEIM5216.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>

namespace Mkl.WebTeam.DocumentGenerator.Functions
{
    using System.Collections.Generic;
    using DecisionModel.Models.Policy;

    /// <summary>
    /// The custom functions
    /// </summary>
    public partial class CustomFunctions
    {
        /// <summary>
        /// Generates the IMB MTC-DE 07 03 form.
        /// </summary>
        /// <param name="policy">The policy.</param>
        /// <param name="documentManager">The document manager.</param>
        /// <param name="extraParameters">The extra parameters.</param>
        public void GenerateMEIM5216_0712Form(Policy policy, ref IPolicyDocumentManager documentManager, Dictionary<string, object> extraParameters)
        {
            var questionCodeList = new List<string> { "CoveredPropertyDescription", "CoveredPropertyIdentificationNumber", "CoveredPropertyCoverageLimit", "CoveredPropertyPremiumPerItem" };
            var blankMergeFieldList = new List<List<string>>
            {
                new List<string> { string.Empty, "4", "8", "12", "16", "20", "24", "28", "32", "36", "40", "44", "48", "52", "56", "60", "64", "68", "72", "76", "80", "84", "88", "92" },
                new List<string> { "1", "5", "9", "13", "17", "21", "25", "29", "33", "37", "41", "45", "49", "53", "57", "61", "65", "69", "73", "77", "81", "85", "89", "93" },
                new List<string> { "2", "6", "10", "14", "18", "22", "26", "30", "34", "38", "42", "46", "50", "54", "58", "62", "66", "70", "74", "78", "82", "86", "90", "94" },
                new List<string> { "3", "7", "11", "15", "19", "23", "27", "31", "35", "39", "43", "47", "51", "55", "59", "63", "67", "71", "75", "79", "83", "87", "91", "95" }
            };

            ProcessCustomTable(policy, documentManager, blankMergeFieldList, questionCodeList);
        }
    }
}
