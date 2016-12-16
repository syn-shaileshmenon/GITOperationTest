// ***********************************************************************
// Assembly         : Mkl.WebTeam.DocumentGenerator
// Author           : rsteelea
// Created          : 11-14-2016
//
// Last Modified By : rsteelea
// Last Modified On : 11-14-2016
// ***********************************************************************
// <copyright file="CustomFunctions_MDIL1005.cs" company="Markel">
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
        public void GenerateMDIL1005_0814Form(Policy policy, ref IPolicyDocumentManager documentManager, Dictionary<string, object> extraParameters)
        {
            List<string> questionCodeList = new List<string> { "NamedInsured" };
            List<List<string>> blankMergeFieldList = new List<List<string>>
            {
                new List<string> { string.Empty },
            };

            // Fill the blank merge field list
            for (var row = 1; row <= 31; row++)
            {
                blankMergeFieldList[0].Add(row.ToString());
            }

            ProcessCustomTable(policy, documentManager, blankMergeFieldList, questionCodeList);
        }
    }
}
