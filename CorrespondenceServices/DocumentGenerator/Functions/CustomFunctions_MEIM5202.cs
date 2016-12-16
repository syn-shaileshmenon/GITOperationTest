// ***********************************************************************
// Assembly         : Mkl.WebTeam.DocumentGenerator
// Author           : ssandlina
// Created          : 11-15-2016
//
// Last Modified By : ssandlina
// Last Modified On : 11-15-2016
// ***********************************************************************
// <copyright file="CustomFunctions_MEIM5202.cs" company="Markel">
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
        /// Generates the MEIM 5202 11 09 form.
        /// </summary>
        /// <param name="policy">The policy.</param>
        /// <param name="documentManager">The document manager.</param>
        /// <param name="extraParameters">The extra parameters.</param>
        public void GenerateMEIM5202_1109Form(Policy policy, ref IPolicyDocumentManager documentManager, Dictionary<string, object> extraParameters)
        {
            // Fields are as follows:
            // Covered Items / Values Description [CoveredDescription]
            // Percentage Deductible [PercentageDeductible]
            // Covered Causes of Loss Applicable to Percentage Deductible [DesignatedCoveredPercentageDeductible]
            // Dollar Deductible [DollarDeductible]
            // Covered Causes of Loss Applicable to Dollar Deductible [DesignatedCoveredDollarDeductible]
            List<string> questionCodeList = new List<string> { "CoveredDescription", "PercentageDeductible", "DesignatedCoveredPercentageDeductible", "DollarDeductible", "DesignatedCoveredDollarDeductible" };

            List<List<string>> blankMergeFieldList = new List<List<string>>
            {
                new List<string> { "1", "6", "11" },
                new List<string> { "2", "7", "12" },
                new List<string> { "3", "8", "13" },
                new List<string> { "4", "9", "14" },
                new List<string> { "5", "10", "15" },
            };

            ProcessCustomTable(policy, documentManager, blankMergeFieldList, questionCodeList);
        }
    }
}
