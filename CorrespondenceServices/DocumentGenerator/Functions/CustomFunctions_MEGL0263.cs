// ***********************************************************************
// Assembly         : Mkl.WebTeam.DocumentGenerator
// Author           : rsteelea
// Created          : 11-14-2016
//
// Last Modified By : rsteelea
// Last Modified On : 11-14-2016
// ***********************************************************************
// <copyright file="CustomFunctions_MEGL0263.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************

/// <summary>
/// The Functions namespace.
/// </summary>
namespace Mkl.WebTeam.DocumentGenerator.Functions
{
    using System.Collections.Generic;
    using DecisionModel.Models.Policy;
    using Mkl.WebTeam.DocumentGenerator.Enumerations;

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
        public void GenerateMEGL0263_0314Form(Policy policy, ref IPolicyDocumentManager documentManager, Dictionary<string, object> extraParameters)
        {
            var questionCodeList = new List<string> { "InsurerName", "PolicyNumber", "LimitsOfLiability" };
            var blankMergeFieldList = new List<List<string>>
            {
                new List<string> { string.Empty, "3", "6" },
                new List<string> { "1", "4", "7" },
                new List<string> { "2", "5", "8" },
            };

            ProcessCustomTable(policy, documentManager, blankMergeFieldList, questionCodeList, RemoveGridLines.TopBottom);
        }
    }
}
