// <copyright file="CustomFunctions_CG0300.cs" company="Markel">
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
    /// Class CustomFunctions_CG0300
    /// </summary>
    public partial class CustomFunctions
    {
        /// <summary>
        /// Generates CG 03 00 form
        /// </summary>
        /// <param name="policy">Policy</param>
        /// <param name="documentManager">Document manager</param>
        /// <param name="extraParameters">Extra parameters</param>
        public void GenerateCG0300Form(Policy policy, ref IPolicyDocumentManager documentManager, Dictionary<string, object> extraParameters)
        {
            GlLimit limit = null;
            if (policy.SpecialEventLine != null && policy.SpecialEventLine.RiskUnits != null)
            {
                if (policy.SpecialEventLine.Limits != null)
                {
                    limit = policy.SpecialEventLine.Limits;
                }
            }

            if (policy.GlLine != null && policy.GlLine.RiskUnits != null)
            {
                if (policy.GlLine.Limits != null)
                {
                    limit = policy.GlLine.Limits;
                }
            }

            if (limit != null)
            {
                var deductible = limit.Deductible;
                string bookmarkName = string.Empty;
                if (limit.Type == "BI")
                {
                    if (limit.Basis == "PC")
                    {
                        bookmarkName = "BlankMergeField";
                    }
                    else
                    {
                        bookmarkName = "BlankMergeField_1";
                    }
                }
                else if (limit.Type == "PD")
                {
                    if (limit.Basis == "PC")
                    {
                        bookmarkName = "BlankMergeField_2";
                    }
                    else
                    {
                        bookmarkName = "BlankMergeField_3";
                    }
                }
                else if (limit.Type == "BIPD")
                {
                    if (limit.Basis == "PC")
                    {
                        bookmarkName = "BlankMergeField_4";
                    }
                    else
                    {
                        bookmarkName = "BlankMergeField_5";
                    }
                }

                this.FillBookmarkField(ref documentManager, bookmarkName, FormatHelper.FormatValue(deductible, FormatHelper.CurrencyFormat));
            }
        }
    }
}