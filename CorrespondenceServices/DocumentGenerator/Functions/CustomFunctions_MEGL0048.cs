// <copyright file="CustomFunctions_MEGL0048.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>

namespace Mkl.WebTeam.DocumentGenerator.Functions
{
    using System.Collections.Generic;
    using DecisionModel.Models.Policy;
    using Mkl.WebTeam.DocumentGenerator.Helpers;

    /// <summary>
    /// Class CustomFunctions_MEGL0048
    /// </summary>
    public partial class CustomFunctions
    {
        /// <summary>
        /// Generates the MEGL 0048 form.
        /// </summary>
        /// <param name="policy">The policy.</param>
        /// <param name="documentManager">The document manager.</param>
        /// <param name="extraParameters">The extra parameters.</param>
        public void GenerateMEGL0048Form(Policy policy, ref IPolicyDocumentManager documentManager, Dictionary<string, object> extraParameters)
        {
            this.MapDeductibles(policy, ref documentManager);
        }

        /// <summary>
        /// Maps deductible values
        /// </summary>
        /// <param name="policy">Policy object</param>
        /// <param name="documentManager">document manager</param>
        private void MapDeductibles(Policy policy, ref IPolicyDocumentManager documentManager)
        {
            GlLimit limits = policy.OCPLine != null ? policy.OCPLine.Limits :
                             policy.GlLine != null ? policy.GlLine.Limits :
                             policy.SpecialEventLine != null ? policy.SpecialEventLine.Limits :
                             null;

            if (policy.OCPLine != null || policy.SpecialEventLine != null)
            {
                switch (limits.Type.ToUpper())
                {
                    case "BI":
                        if (limits.Basis.EqualsIgnoreCase("PC"))
                        {
                            this.FillBookmarkField(ref documentManager, "BlankMergeField_1", FormatHelper.FormatValue(limits.Deductible, FormatHelper.CurrencyFormat));
                        }
                        else if (limits.Basis.EqualsIgnoreCase("PO"))
                        {
                            this.FillBookmarkField(ref documentManager, "BlankMergeField_2", FormatHelper.FormatValue(limits.Deductible, FormatHelper.CurrencyFormat));
                        }

                        break;
                    case "PD":
                        if (limits.Basis.EqualsIgnoreCase("PC"))
                        {
                            this.FillBookmarkField(ref documentManager, "BlankMergeField_3", FormatHelper.FormatValue(limits.Deductible, FormatHelper.CurrencyFormat));
                        }
                        else if (limits.Basis.EqualsIgnoreCase("PO"))
                        {
                            this.FillBookmarkField(ref documentManager, "BlankMergeField_4", FormatHelper.FormatValue(limits.Deductible, FormatHelper.CurrencyFormat));
                        }

                        break;
                    case "BIPD":
                        if (limits.Basis.EqualsIgnoreCase("PC"))
                        {
                            this.FillBookmarkField(ref documentManager, "BlankMergeField_5", FormatHelper.FormatValue(limits.Deductible, FormatHelper.CurrencyFormat));
                        }
                        else if (limits.Basis.EqualsIgnoreCase("PO"))
                        {
                            this.FillBookmarkField(ref documentManager, "BlankMergeField_6", FormatHelper.FormatValue(limits.Deductible, FormatHelper.CurrencyFormat));
                        }

                        break;
                }
            }

            if (policy.GlLine != null)
            {
                switch (limits.Type.ToUpper())
                {
                    case "BI":
                        if (limits.Basis.EqualsIgnoreCase("PC") || limits.Basis.EqualsIgnoreCase("PIPC"))
                        {
                            this.FillBookmarkField(ref documentManager, "BlankMergeField_1", FormatHelper.FormatValue(limits.Deductible, FormatHelper.CurrencyFormat));
                        }
                        else if (limits.Basis.EqualsIgnoreCase("PO"))
                        {
                            this.FillBookmarkField(ref documentManager, "BlankMergeField_2", FormatHelper.FormatValue(limits.Deductible, FormatHelper.CurrencyFormat));
                        }

                        break;
                    case "PD":
                        if (limits.Basis.EqualsIgnoreCase("PC") || limits.Basis.EqualsIgnoreCase("PIPC"))
                        {
                            this.FillBookmarkField(ref documentManager, "BlankMergeField_3", FormatHelper.FormatValue(limits.Deductible, FormatHelper.CurrencyFormat));
                        }
                        else if (limits.Basis.EqualsIgnoreCase("PO"))
                        {
                            this.FillBookmarkField(ref documentManager, "BlankMergeField_4", FormatHelper.FormatValue(limits.Deductible, FormatHelper.CurrencyFormat));
                        }

                        break;
                    case "BIPD":
                        if (limits.Basis.EqualsIgnoreCase("PC") || limits.Basis.EqualsIgnoreCase("PIPC"))
                        {
                            this.FillBookmarkField(ref documentManager, "BlankMergeField_5", FormatHelper.FormatValue(limits.Deductible, FormatHelper.CurrencyFormat));
                        }
                        else if (limits.Basis.EqualsIgnoreCase("PO"))
                        {
                            this.FillBookmarkField(ref documentManager, "BlankMergeField_6", FormatHelper.FormatValue(limits.Deductible, FormatHelper.CurrencyFormat));
                        }

                        break;
                }

                if (limits.Basis.EqualsIgnoreCase("PIPC"))
                {
                    this.FillBookmarkField(ref documentManager, "Check1", "CHECKED");
                }
            }
        }
    }
}
