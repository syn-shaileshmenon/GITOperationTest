// <copyright file="CustomFunctions.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>

namespace Mkl.WebTeam.DocumentGenerator.Functions
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Text.RegularExpressions;
    using Aspose.Words;
    using Aspose.Words.Tables;
    using DecisionModel.Models.Policy;
    using DecisionModel.Representations;
    using Mkl.WebTeam.DocumentGenerator.Entities;
    using Mkl.WebTeam.DocumentGenerator.Enumerations;
    using Mkl.WebTeam.DocumentGenerator.Helpers;
    using SubmissionShared.Enumerations;

    /// <summary>
    /// Custom functions for bookmark substitution to be used for Document Generation
    /// </summary>
    public partial class CustomFunctions
    {
        /// <summary>
        /// Blank merge field text
        /// </summary>
        private const string BlankMergeField = "BlankMergeField_";

        /// <summary>
        /// The last index for sum of taxes
        /// </summary>
        private const short SumOfTax = 48;

        /// <summary>
        /// Initializes a new instance of the <see cref="CustomFunctions" /> class.
        /// </summary>
        public CustomFunctions()
        {
        }

        /// <summary>
        /// Delegate to call for each Row of a table.
        /// </summary>
        /// <param name="manager">The manager.</param>
        /// <param name="table">The table.</param>
        /// <param name="row">The row.</param>
        /// <param name="rowNumber">The row number.</param>
        public delegate void RowCallbackType(IPolicyDocumentManager manager, Table table, Row row, int rowNumber);

        /// <summary>
        /// Delegate to call for each cell of a table
        /// </summary>
        /// <param name="manager">The manager.</param>
        /// <param name="table">The table.</param>
        /// <param name="row">The row.</param>
        /// <param name="rowNumber">The row number.</param>
        /// <param name="columnNumber">The column number.</param>
        /// <param name="question">The question.</param>
        /// <param name="bookmark">The bookmark being worked</param>
        /// <param name="bookmarkName">Name of the bookmark.</param>
        /// <returns><c>true</c> if XXXX, <c>false</c> otherwise.</returns>
        public delegate bool CellCallbackType(IPolicyDocumentManager manager, Table table, Row row, int rowNumber, int columnNumber, Question question, Bookmark bookmark, string bookmarkName);

        /// <summary>
        /// Calculates Sum of Fees and Taxes for the Policy Rollup
        /// </summary>
        /// <param name="policy">Passed policy object</param>
        /// <param name="documentManager">Passed document manager</param>
        /// <param name="extraParameters">Dictionary of extra parameters</param>
        public void CalculateTaxesFees(Policy policy, ref IPolicyDocumentManager documentManager, Dictionary<string, object> extraParameters)
        {
            ////TODO: Add processing exceptions to a List<string> and return back to the caller
            decimal? sumOfTaxesAndFees;
            string bookMarkValue = string.Empty;
            if (policy != null && policy.TaxesAndFees != null && policy.TaxesAndFees.Any())
            {
                int indexField = 1;
                foreach (TaxFee taxFee in policy.TaxesAndFees)
                {
                    this.FillBookmarkField(ref documentManager, BlankMergeField + indexField, taxFee.Name);
                    indexField++;

                    this.FillBookmarkField(ref documentManager, BlankMergeField + indexField, FormatHelper.FormatValue(taxFee.Amount, FormatHelper.CurrencyFormatWithDecimal));
                    indexField += 2;
                }

                sumOfTaxesAndFees = policy.TaxesAndFees.Sum(c => c.Amount);
                if (sumOfTaxesAndFees.HasValue)
                {
                    bookMarkValue = FormatHelper.FormatValue(sumOfTaxesAndFees, FormatHelper.CurrencyFormatWithDecimal);
                }

                this.FillBookmarkField(ref documentManager, BlankMergeField + SumOfTax, bookMarkValue);
            }
        }

        /// <summary>
        /// Get today's date in UTC
        /// </summary>
        /// <param name="policy">Passed policy object</param>
        /// <param name="documentManager">Passed document manager</param>
        /// <param name="extraParameters">Dictionary of extra parameters</param>
        /// <returns>string</returns>
        public string GetTodayDate(Policy policy, ref IPolicyDocumentManager documentManager, Dictionary<string, object> extraParameters)
        {
            string bookMarkValue = FormatHelper.FormatValue(DateTime.UtcNow, extraParameters.Count > 0 && extraParameters["Format"] != null ? extraParameters["Format"].ToString() : string.Empty);
            return bookMarkValue;
        }

        /// <summary>
        /// Get Terrorism Premium
        /// </summary>
        /// <param name="policy">Passed policy object</param>
        /// <param name="documentManager">Passed document manager</param>
        /// <param name="extraParameters">Dictionary of extra parameters</param>
        /// <returns>string</returns>
        public string GetTerrorismPremium(Policy policy, ref IPolicyDocumentManager documentManager, Dictionary<string, object> extraParameters)
        {
            var format = extraParameters.Count > 0 && extraParameters["Format"] != null ? extraParameters["Format"].ToString() : string.Empty;
            var terrorismPremium = extraParameters.Count > 0 && extraParameters["Constant"] != null ? extraParameters["Constant"].ToString() : string.Empty;

            if (policy.IsTerrorism)
            {
                terrorismPremium = FormatHelper.FormatValue(policy.PremiumTerrorism, FormatHelper.CurrencyFormat);
            }

            return terrorismPremium;
        }

        /// <summary>
        /// Get sum of taxes
        /// </summary>
        /// <param name="policy">Passed policy object</param>
        /// <param name="documentManager">Passed document manager</param>
        /// <param name="extraParameters">Dictionary of extra parameters</param>
        /// <returns>string</returns>
        public string SumOfTaxes(Policy policy, ref IPolicyDocumentManager documentManager, Dictionary<string, object> extraParameters)
        {
            decimal? sumOfTaxesAndFees = policy.IsExcludeTaxesFees == true ? 0.00m : policy.TaxesAndFees.Sum(c => c.Amount);
            string format = extraParameters["Format"] != null ? extraParameters["Format"].ToString() : string.Empty;
            if (sumOfTaxesAndFees.HasValue)
            {
                string bookMarkValue = FormatHelper.FormatValue(sumOfTaxesAndFees, format);
                return bookMarkValue;
            }
            else
            {
                sumOfTaxesAndFees = 0.00m;
                return FormatHelper.FormatValue(sumOfTaxesAndFees, format);
            }
        }

        /// <summary>
        /// Get the State transaction code used on MADUB 1000 01 15
        /// </summary>
        /// <param name="policy">Passed policy object</param>
        /// <param name="documentManager">Passed document manager</param>
        /// <param name="extraParameters">Dictionary of extra parameters</param>
        /// <returns>State Transaction Code from Policy object</returns>
        public string GetStateTransactionCode(Policy policy, ref IPolicyDocumentManager documentManager, Dictionary<string, object> extraParameters)
        {
            string stateTransactionCode = string.Empty;
            if (policy.StateTransactionCode != null && !string.IsNullOrWhiteSpace(policy.StateTransactionCode))
            {
                stateTransactionCode = policy.StateTransactionCode;
            }

            return stateTransactionCode;
        }

        /// <summary>
        /// Generates the terrorism premium for MUB TERR 2
        /// </summary>
        /// <param name="policy">Passed policy object</param>
        /// <param name="documentManager">Passed document manager</param>
        /// <param name="extraParameters">Dictionary of extra parameters</param>
        public void GenerateMUBTerr2(Policy policy, ref IPolicyDocumentManager documentManager, Dictionary<string, object> extraParameters)
        {
            var terrorismPremium = policy.CalculateTerrorismPremium(isSelectedTerrorismConsidered: false);
            var formattedTerrorismPremium = FormatHelper.FormatValue(terrorismPremium, FormatHelper.CurrencyFormat);
            this.FillBookmarkField(ref documentManager, "TerrorismPremiums", formattedTerrorismPremium);
        }

        /// <summary>
        /// Generates MDIL 1000 and 1000-CA forms for non XS
        /// </summary>
        /// <param name="policy">Passed policy object</param>
        /// <param name="documentManager">Passed document manager</param>
        /// <param name="extraParameters">Dictionary of extra parameters</param>
        public void GenerateMDIL1000Form(Policy policy, ref IPolicyDocumentManager documentManager, Dictionary<string, object> extraParameters)
        {
            const string NotCoveredText = @"Not Covered";
            const string OcpCoverageText = @"Owners and Contractors Protective Liability";
            const string TaxesFeesText = @"Taxes and Fees - See MDIL 1002";
            const string TriaText = @"Terrorism - Certified Acts";
            const string TriaExcludedText = @"Excluded";
            decimal premiumGrandTotal = default(decimal);
            string cfRollupPremium = this.GetApplicablePremium(policy, LineOfBusiness.CF);
            string glRollupPremium = this.GetApplicablePremium(policy, LineOfBusiness.GL);
            string specialEventsRollupPremium = this.GetApplicablePremium(policy, LineOfBusiness.SpecialEvent);
            string imRollupPremium = this.GetApplicablePremium(policy, LineOfBusiness.IM);
            string liquorRollupPremium = this.GetApplicablePremium(policy, LineOfBusiness.LL);
            string ocpRollupPremium = this.GetApplicablePremium(policy, LineOfBusiness.OCP);
            string policyRollupPremium = FormatHelper.FormatValue(policy.RollUp.Premium, FormatHelper.CurrencyFormatWithDecimal);
            string terrorismPremium = policy.IsTerrorism && policy.PremiumTerrorism > 0 ? FormatHelper.FormatValue(policy.PremiumTerrorism, FormatHelper.CurrencyFormatWithDecimal) : TriaExcludedText;
            premiumGrandTotal = policy.IsExcludeTaxesFees == true ? policy.RollUp.Premium : policy.TaxesAndFees.Sum(c => c.Amount).Value + policy.RollUp.Premium;
            this.FillBookmarkField(ref documentManager, "BlankMergeField_3", cfRollupPremium);

            if (policy.GlLine != null && policy.SpecialEventLine == null)
            {
                this.FillBookmarkField(ref documentManager, "BlankMergeField_4", glRollupPremium);
            }
            else
            {
                this.FillBookmarkField(ref documentManager, "BlankMergeField_4", specialEventsRollupPremium);
            }

            this.FillBookmarkField(ref documentManager, "BlankMergeField_5", imRollupPremium);
            this.FillBookmarkField(ref documentManager, "BlankMergeField_6", NotCoveredText);
            this.FillBookmarkField(ref documentManager, "BlankMergeField_7", NotCoveredText);

            this.FillBookmarkField(ref documentManager, "BlankMergeField_8", NotCoveredText);
            this.FillBookmarkField(ref documentManager, "BlankMergeField_9", liquorRollupPremium);
            this.FillBookmarkField(ref documentManager, "BlankMergeField_10", NotCoveredText);
            this.FillBookmarkField(ref documentManager, "BlankMergeField_11", TriaText);
            this.FillBookmarkField(ref documentManager, "BlankMergeField_12", terrorismPremium);

            if (policy.OCPLine != null)
            {
                this.FillBookmarkField(ref documentManager, "BlankMergeField_13", OcpCoverageText);
                this.FillBookmarkField(ref documentManager, "BlankMergeField_14", ocpRollupPremium);
            }

            this.FillBookmarkField(ref documentManager, "BlankMergeField_15", policyRollupPremium);

            if (policy.IsExcludeTaxesFees == false && policy.TaxesAndFees.Any())
            {
                this.FillBookmarkField(ref documentManager, "BlankMergeField_16", TaxesFeesText);
                this.FillBookmarkField(ref documentManager, "BlankMergeField_17", FormatHelper.FormatValue(policy.TaxesAndFees.Sum(c => c.Amount).Value, FormatHelper.CurrencyFormatWithDecimal));
            }

            this.FillBookmarkField(ref documentManager, "BlankMergeField_25", policy.StateTransactionCode != null ? policy.StateTransactionCode : string.Empty);
            this.FillBookmarkField(ref documentManager, "BlankMergeField_22", FormatHelper.FormatValue(premiumGrandTotal, FormatHelper.CurrencyFormatWithDecimal));
        }

        /// <summary>
        /// Generate MDIL 1001 form. This method is here for Excess transactions created before issuance was rolled out to all lines.
        /// </summary>
        /// <param name="policy">Passed policy object</param>
        /// <param name="documentManager">Passed document manager</param>
        /// <param name="extraParameters">Dictionary of extra parameters</param>
        public void GenerateMDIL1001Form(Policy policy, ref IPolicyDocumentManager documentManager, Dictionary<string, object> extraParameters)
        {
            Bookmark bookMark = documentManager.GetBookMarkByName("PolicyForm");

            var documents = new List<DecisionModel.Models.Policy.Document>();
            var nonZeroOrderDocuments = policy.RollUp.Documents.Where(x => x.Order != 0 && x.ReturnIsSelected() && x.Section != SubmissionShared.Enumerations.DocumentSection.Application).Select(x => x).OrderBy(x => x.Order).ThenBy(x => x.NormalizedNumber).ToList();
            var zeroOrderDocuments = policy.RollUp.Documents.Where(x => x.Order == 0 && x.ReturnIsSelected() && x.Section != SubmissionShared.Enumerations.DocumentSection.Application).Select(x => x).OrderBy(x => x.NormalizedNumber).ToList();
            documents.AddRange(nonZeroOrderDocuments);
            documents.AddRange(zeroOrderDocuments);
            var documentDictionary = new Dictionary<string, DecisionModel.Models.Policy.Document>();
            documents.ForEach(d => documentDictionary.Add(GetDocumentKey(d), d));

            var dictionary = new Dictionary<dynamic, Dictionary<string, DecisionModel.Models.Policy.Document>> { { new LobKey() { Title = "Excess Liability", Order = 1 }, documentDictionary } };
            documentManager.InsertPolicyFormTable(dictionary, bookMark);
        }

        /// <summary>
        /// Generate MADUB 1003 form
        /// </summary>
        /// <param name="policy">Passed policy object</param>
        /// <param name="documentManager">Passed document manager</param>
        /// <param name="extraParameters">Dictionary of extra parameters</param>
        public void GenerateMADUB1003Form(Policy policy, ref IPolicyDocumentManager documentManager, Dictionary<string, object> extraParameters)
        {
            ////TODO: Add processing exceptions to a List<string> and return back to the caller
            string returnValue = string.Empty;
            Dictionary<string, MappingDetail> dictionaryMADUB1003 = this.CreateMADUB1003Dictionary(policy);

            string val = string.Empty;

            List<string> liabilitiesToBeRemoved = new List<string>();

            if (!policy.XsLine.XsAutoLiabilityCoverage.IsActive)
            {
                liabilitiesToBeRemoved.Add("XsAutoLiabilityCoverage");
            }

            if (!policy.XsLine.XsEmployersLiabilityCoverage.IsActive)
            {
                liabilitiesToBeRemoved.Add("XsEmployersLiabilityCoverage");
            }

            if (!policy.XsLine.XsHiredNonOwnedCoverage.IsActive)
            {
                liabilitiesToBeRemoved.Add("XsHiredNonOwnedCoverage");
            }

            if (!policy.XsLine.XsGeneralLiabilityCoverage.IsActive)
            {
                liabilitiesToBeRemoved.Add("XsGeneralLiabilityCoverage");
            }

            if (!this.ShouldEmployeeBenefitLimitBeDisplayed(policy))
            {
                liabilitiesToBeRemoved.Add("XsEmployeeBenefitsLimit");
            }

            foreach (KeyValuePair<string, MappingDetail> keyRow in dictionaryMADUB1003)
            {
                MappingDetail detail = keyRow.Value;
                if (detail != null && detail.HasFunction == false)
                {
                    if (string.IsNullOrWhiteSpace(detail.Constant))
                    {
                        val = ReflectionHelper.GetPropValue<string>(policy, detail.Field, detail.Format);
                        if (detail.Field == "Policy.XsLine.XsHiredNonOwnedCoverage.CombinedSingleLimit.Text")
                        {
                            if (!string.IsNullOrWhiteSpace(val) && (val.ToLowerInvariant() == "included in gl" || val.ToLowerInvariant() == "included in auto"))
                            {
                                if (!liabilitiesToBeRemoved.Contains("XsHiredNonOwnedCoverage"))
                                {
                                    liabilitiesToBeRemoved.Add("XsHiredNonOwnedCoverage");
                                }
                            }
                        }
                    }
                    else
                    {
                        val = detail.Constant;
                    }
                }

                if (!string.IsNullOrWhiteSpace(val))
                {
                    this.FillBookmarkField(ref documentManager, keyRow.Key, val);
                }

                val = string.Empty;
            }

            ////Delete underlying limit tables
            if (liabilitiesToBeRemoved.Contains("XsAutoLiabilityCoverage"))
            {
                documentManager.RemoveTableWithBookmark("BlankMergeField_6", true);
            }

            if (liabilitiesToBeRemoved.Contains("XsEmployersLiabilityCoverage"))
            {
                documentManager.RemoveTableWithBookmark("BlankMergeField_18", true);
            }

            if (liabilitiesToBeRemoved.Contains("XsHiredNonOwnedCoverage"))
            {
                documentManager.RemoveTableWithBookmark("BlankMergeField_12", true);
            }

            if (liabilitiesToBeRemoved.Contains("XsGeneralLiabilityCoverage"))
            {
                documentManager.RemoveTableWithBookmark("BlankMergeField_1", true);
            }

            if (liabilitiesToBeRemoved.Contains("XsEmployeeBenefitsLimit"))
            {
                documentManager.RemoveTableWithBookmark("BlankMergeField_24", false);
            }
        }

        /// <summary>
        /// Generates MDGL 1012 form
        /// </summary>
        /// <param name="policy">Policy</param>
        /// <param name="documentManager">Document manager</param>
        /// <param name="extraParameters">Extra parameters</param>
        public void GenerateMDGL1012Form(Policy policy, ref IPolicyDocumentManager documentManager, Dictionary<string, object> extraParameters)
        {
            var result = new Dictionary<string, MappingDetail>();

            if (policy.OCPLine != null && policy.OCPLine.RiskUnits != null && policy.OCPLine.RiskUnits.Any())
            {
                var riskUnit = policy.OCPLine.RiskUnits.FirstOrDefault();

                if (riskUnit.Questions != null && riskUnit.Questions.Any())
                {
                    ////Designated contractor
                    var designatedContractorQuestion = riskUnit.Questions.FirstOrDefault(x => x.Code.ToLower().Equals("designated contractor"));
                    if (designatedContractorQuestion != null)
                    {
                        this.FillBookmarkField(ref documentManager, "BlankMergeField_4", designatedContractorQuestion.AnswerValue);
                    }

                    ////Location of operations
                    var locationStreet = riskUnit.Questions.FirstOrDefault(x => x.Code.ToLower().Equals("street question"));
                    var locationCity = riskUnit.Questions.FirstOrDefault(x => x.Code.ToLower().Equals("city question"));
                    var locationState = riskUnit.Questions.FirstOrDefault(x => x.Code.ToLower().Equals("statequestion"));
                    var locationZipCode = riskUnit.Questions.FirstOrDefault(x => x.Code.ToLower().Equals("zipcode question"));
                    var locationOfOperations = this.GetFormattedAddressFromQA(locationStreet, locationCity, locationState, locationZipCode);
                    this.FillBookmarkField(ref documentManager, "BlankMergeField", locationOfOperations);
                }

                this.MDGL1012ClassificationTable(policy, ref documentManager);

                this.FillBookmarkField(ref documentManager, "BlankMergeField_10", string.Format("{0:n0}", policy.OCPLine.Limits.PerOccurrence));
                this.FillBookmarkField(ref documentManager, "BlankMergeField_11", string.Format("{0:n0}", policy.OCPLine.Limits.GeneralAggregate));
            }
        }

        /// <summary>
        /// Generates the i L1201 form.
        /// </summary>
        /// <param name="policy">The policy.</param>
        /// <param name="documentManager">The document manager.</param>
        /// <param name="extraParameters">The extra parameters.</param>
        public void GenerateIL1201Form(Policy policy, ref IPolicyDocumentManager documentManager, Dictionary<string, object> extraParameters)
        {
            var formNormalizedNumber = documentManager.FormNormalizedNumber;
            int? quantityOrder = documentManager.QuantityOrder;
            string coveragePartsAffected = null;

            var document = policy.RollUp.Documents.FirstOrDefault(x => x.NormalizedNumber == formNormalizedNumber && (x.QuantityOrder ?? 0) == quantityOrder);
            var question = document?.Questions.FirstOrDefault(x => x.Code.ToLower().Equals("coveragepartsaffected"));
            var answers = question?.Answers.Where(a => a.IsSelected == true).OrderBy(b => b.Order).Select(c => c.Verbiage).ToArray();

            if (answers?.Length > 0)
            {
                coveragePartsAffected = string.Join("\r\n", answers);
            }

            this.FillBookmarkField(ref documentManager, "BlankMergeField_2", coveragePartsAffected);
        }

        /// <summary>
        /// Generate IM MPFAR form
        /// </summary>
        /// <param name="policy">Passed policy object</param>
        /// <param name="documentManager">Passed document manager</param>
        /// <param name="extraParameters">Dictionary of extra parameters</param>
        public void GenerateIMMPFARForm(Policy policy, ref IPolicyDocumentManager documentManager, Dictionary<string, object> extraParameters)
        {
            //// const string Meil1211FormText = "See MEIL 1211";
            const string Mdil1000FormText = "See MDIL 1000";
            decimal applicableRate = default(decimal);
            var imRiskUnitItem = policy.ImLine.RiskUnits.FirstOrDefault(i => i.Items.Any(r => r.ItemType == ImRateItemType.MiscPropertyEndorsement))
                .Items.Where(i => i.ItemType == ImRateItemType.MiscPropertyEndorsement).FirstOrDefault();
            if (imRiskUnitItem.AgentAdjustedRate != default(decimal))
            {
                applicableRate = imRiskUnitItem.AgentAdjustedRate;
            }
            else if (imRiskUnitItem.UnderwriterAdjustedRate != default(decimal))
            {
                applicableRate = imRiskUnitItem.UnderwriterAdjustedRate;
            }
            else
            {
                applicableRate = imRiskUnitItem.DevelopedRate;
            }

            applicableRate = Math.Round(applicableRate, 2, MidpointRounding.AwayFromZero);

            ImRiskUnit riskUnitMPFar = policy.ImLine.RiskUnits.Where(x => x.ClassCode == "424").First();
            if (riskUnitMPFar != null)
            {
                if (policy.ImLine.RollUp.IsMinimumPremium)
                {
                    this.FillBookmarkField(ref documentManager, "BlankMergeField_2", FormatHelper.FormatValue(policy.ImLine.RollUp.Premium, FormatHelper.CurrencyFormat));
                }
                else
                {
                    this.FillBookmarkField(ref documentManager, "BlankMergeField", FormatHelper.FormatValue(policy.ImLine.RollUp.Premium, FormatHelper.CurrencyFormat));
                }
            }

            this.FillBookmarkField(ref documentManager, "BlankMergeField_1", string.Format("$ {0:n2}", applicableRate));
            //// this.FillBookmarkField(ref documentManager, "BlankMergeField_3", Meil1211FormText);
            this.FillBookmarkField(ref documentManager, "BlankMergeField_4", Mdil1000FormText);
            this.FillBookmarkField(ref documentManager, "BlankMergeField_5", Mdil1000FormText);
            this.FillBookmarkField(ref documentManager, "BlankMergeField_6", Mdil1000FormText);
            this.FillBookmarkField(ref documentManager, "BlankMergeField_7", Mdil1000FormText);
            this.FillBookmarkField(ref documentManager, "BlankMergeField_8", FormatHelper.FormatValue(imRiskUnitItem.TotalInsuredValue, FormatHelper.CurrencyFormat));
            this.FillBookmarkField(ref documentManager, "BlankMergeField_9", FormatHelper.FormatValue(imRiskUnitItem.MaxSingleInsuredValue, FormatHelper.CurrencyFormat));
            this.FillBookmarkField(ref documentManager, "BlankMergeField_10", FormatHelper.FormatValue(policy.ImLine.RiskUnits.FirstOrDefault().PerOccurrenceDeductible.Amount, FormatHelper.CurrencyFormat));
        }

        /// <summary>
        /// Generate IM NR form for IM lines
        /// </summary>
        /// <param name="policy">Passed policy object</param>
        /// <param name="documentManager">Passed document manager</param>
        /// <param name="extraParameters">Dictionary of extra parameters</param>
        public void GenerateIMNRForm(Policy policy, ref IPolicyDocumentManager documentManager, Dictionary<string, object> extraParameters)
        {
            int concessionTrailersSelectedCount = 0;
            int hitchLockWarrantyCount = 0;
            if (policy.ImLine != null)
            {
                foreach (var riskUnit in policy.ImLine.RiskUnits)
                {
                    if (riskUnit.Questions.Where(q => q.Code == "IM Schedule Trailers Concession").Where(s => s.AnswerValue.ToUpperInvariant() == "YES").Any())
                    {
                        concessionTrailersSelectedCount++;
                    }

                    if (riskUnit.Clauses.Select(c => c.Code.ToUpperInvariant() == "TRAILER").Any())
                    {
                        hitchLockWarrantyCount++;
                    }
                }

                this.FillBookmarkField(ref documentManager, "Check2", concessionTrailersSelectedCount > 0 ? "Checked" : string.Empty);
                this.FillBookmarkField(ref documentManager, "Check3", hitchLockWarrantyCount > 0 ? "Checked" : string.Empty);
            }
        }

        /// <summary>
        /// Generate MECP 1227 form for property
        /// </summary>
        /// <param name="policy">Passed policy object</param>
        /// <param name="documentManager">Passed document manager</param>
        /// <param name="extraParameters">Dictionary of extra parameters</param>
        public void GenerateMECP1227Form(Policy policy, ref IPolicyDocumentManager documentManager, Dictionary<string, object> extraParameters)
        {
            decimal applicablePremium = default(decimal);
            bool isIncluded = false;

            var endorsementBreakDown = policy.CfLine.RollUp.OptionalCoverages.Where(o => o.Code == "EB").FirstOrDefault();
            isIncluded = endorsementBreakDown.IsIncluded;
            applicablePremium = endorsementBreakDown.AdjustedPremium.HasValue ? endorsementBreakDown.AdjustedPremium.Value : endorsementBreakDown.Premium;

            this.FillBookmarkField(ref documentManager, "BlankMergeField", isIncluded == true ? "Included" : FormatHelper.FormatValue(applicablePremium, FormatHelper.CurrencyFormat));
        }

        /// <summary>
        /// Generates MDGL 1000 form
        /// </summary>
        /// <param name="policy">Policy</param>
        /// <param name="documentManager">Document manager</param>
        /// <param name="extraParameters">Extra parameters</param>
        public void GenerateMDGL1000Form(Policy policy, ref IPolicyDocumentManager documentManager, Dictionary<string, object> extraParameters)
        {
            if (policy.LiquorLine != null && policy.LiquorLine.Limits != null)
            {
                this.FillBookmarkField(ref documentManager, "RetroDate", "None");
                this.FillBookmarkField(ref documentManager, "BlankMergeField", string.Format("{0:n0}", policy.LiquorLine.Limits.PerOccurrence));
                this.FillBookmarkField(ref documentManager, "BlankMergeField_1", string.Format("{0:n0}", policy.LiquorLine.Limits.GeneralAggregate));

                DocumentBuilder builder = new DocumentBuilder(documentManager.Document);
                this.MDGL1000FormLocations(policy, builder);
                this.MDGL1000Classification(policy, builder);
            }
        }

        /// <summary>
        /// Generate PR MECP 1223 form
        /// </summary>
        /// <param name="policy">Passed policy object</param>
        /// <param name="documentManager">Passed document manager</param>
        /// <param name="extraParameters">Dictionary of extra parameters</param>
        public void GenerateCFMECPForm(Policy policy, ref IPolicyDocumentManager documentManager, Dictionary<string, object> extraParameters)
        {
            const string CoverageCode = "PP";
            decimal applicablePremium = default(decimal);

            var cfRiskUnitItem = policy.CfLine.RollUp.OptionalCoverages.Where(i => i.Code == CoverageCode).FirstOrDefault();
            if (cfRiskUnitItem != null)
            {
                if (cfRiskUnitItem.IsIncluded)
                {
                    this.FillBookmarkField(ref documentManager, "BlankMergeField", "Included");
                }
                else
                {
                    if (cfRiskUnitItem.AdjustedPremium != null)
                    {
                        applicablePremium = (decimal)cfRiskUnitItem.AdjustedPremium;
                    }
                    else
                    {
                        applicablePremium = (decimal)cfRiskUnitItem.Premium;
                    }

                    applicablePremium = Math.Round(applicablePremium, 2, MidpointRounding.AwayFromZero);
                    this.FillBookmarkField(ref documentManager, "BlankMergeField", FormatHelper.FormatValue(applicablePremium, FormatHelper.CurrencyFormat));
                }
            }
        }

        /// <summary>
        /// Generate GL MEGL 1279 form
        /// </summary>
        /// <param name="policy">Passed policy object</param>
        /// <param name="documentManager">Passed document manager</param>
        /// <param name="extraParameters">Dictionary of extra parameters</param>
        public void GenerateGLMEGL1279Form(Policy policy, ref IPolicyDocumentManager documentManager, Dictionary<string, object> extraParameters)
        {
            const string CoverageCode = "AB";
            string applicablePremium = string.Empty;

            OptionalCoverage optCoverageUnitItem = null;
            if (policy.GlLine != null && policy.GlLine.RollUp != null && policy.GlLine.RollUp.OptionalCoverages != null)
            {
                optCoverageUnitItem = policy.GlLine.RollUp.OptionalCoverages.Where(i => i.Code == CoverageCode).FirstOrDefault();
            }
            else if (policy.LiquorLine != null && policy.LiquorLine.RollUp != null && policy.LiquorLine.RollUp.OptionalCoverages != null)
            {
                optCoverageUnitItem = policy.LiquorLine.RollUp.OptionalCoverages.Where(i => i.Code == CoverageCode).FirstOrDefault();
            }

            if (optCoverageUnitItem != null)
            {
                if (optCoverageUnitItem.IsIncluded)
                {
                    applicablePremium = "Included";
                }
                else
                {
                    var premium = (decimal)0;
                    if (optCoverageUnitItem.AdjustedPremium != null)
                    {
                        premium = (decimal)optCoverageUnitItem.AdjustedPremium;
                    }
                    else
                    {
                        premium = optCoverageUnitItem.Premium;
                    }

                    premium = Math.Round(premium, 2, MidpointRounding.AwayFromZero);
                    applicablePremium = FormatHelper.FormatValue(premium, FormatHelper.CurrencyFormat);
                }

                var limitOption = optCoverageUnitItem.LimitOptions.SingleOrDefault(c => c.Id == optCoverageUnitItem.LimitLevelId);
                if (limitOption != null)
                {
                    string aggregateLimit = limitOption.Limits.SingleOrDefault(c => c.LimitTypeCode == "A").Code;
                    string commonCauseLimit = limitOption.Limits.SingleOrDefault(c => c.LimitTypeCode == "EOCC").Code;

                    this.FillBookmarkField(ref documentManager, "BlankMergeField", FormatHelper.FormatValue(Convert.ToDecimal(commonCauseLimit), FormatHelper.CurrencyFormat));
                    this.FillBookmarkField(ref documentManager, "BlankMergeField_1", FormatHelper.FormatValue(Convert.ToDecimal(aggregateLimit), FormatHelper.CurrencyFormat));
                    this.FillBookmarkField(ref documentManager, "BlankMergeField_2", applicablePremium);
                }
            }
        }

        /// <summary>
        /// Generate MDCP 1000 form for property
        /// </summary>
        /// <param name="policy">Passed policy object</param>
        /// <param name="documentManager">Passed document manager</param>
        /// <param name="extraParameters">Extra parameters</param>
        public void GenerateMDCP1000Form(Policy policy, ref IPolicyDocumentManager documentManager, Dictionary<string, object> extraParameters)
        {
            if (policy.CfLine != null && policy.CfLine.RiskUnits.Any())
            {
                DocumentBuilder builder = new DocumentBuilder(documentManager.Document);
                this.FillBlankMergeFieldMDCP1000DescriptionOfPremises(policy.CfLine, builder);
                this.FillBlankMergeFieldMDCP1000CoverageProvided(policy, builder);
                this.FillBlankMergeFieldMDCP1000OptionalCoverages(policy.CfLine, builder);
                this.FillBlankMergeFieldMDCP1000MortgageHolder(policy.CfLine, builder);
                this.FillBlankMergeFieldMDCP1000TotalPremium(policy, builder);
            }
        }

        /// <summary>
        /// Generate MECP 1285 form
        /// </summary>
        /// <param name="policy">Policy</param>
        /// <param name="documentManager">Document manager</param>
        /// <param name="extraParameters">Extra parameters</param>
        public void GenerateMECP1285Form(Policy policy, ref IPolicyDocumentManager documentManager, Dictionary<string, object> extraParameters)
        {
            if (policy.CfLine != null)
            {
                DocumentBuilder builder = new DocumentBuilder(documentManager.Document);

                builder.MoveToBookmark("WindhailTable");

                Aspose.Words.Tables.Table table = builder.StartTable();
                builder.Font.Size = 10;
                builder.Font.Bold = true;
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;

                builder.InsertCell();
                builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(125.15);
                builder.Write("Premises Number");

                builder.InsertCell();
                builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(120);
                builder.Write("Building Number");

                builder.InsertCell();
                builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(258.85);
                builder.Write("Hurricane Deductible Percentage");

                builder.EndRow();
                builder.CellFormat.ClearFormatting();

                builder.Font.Bold = false;

                decimal windhailDeductibleAmount = 0;
                foreach (var riskunit in policy.CfLine.RiskUnits.Where(r => r.StreetAddress.StateCode == "HI" && r.WindHailDeductible != "NONE" && r.IsWindHailExcluded == false && r.IsWindHailPercentage == true))
                {
                    builder.InsertCell();
                    builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(125.15);
                    builder.Write(riskunit.PremisesNumber.ToString());

                    builder.InsertCell();
                    builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(120);
                    builder.Write(riskunit.BuildingNumber.ToString());

                    builder.InsertCell();
                    builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(258.85);

                    Regex pattern = new Regex(@"(?<percent>.+?)PCT(?<amount>.+?)MIN");
                    var matches = pattern.Matches(riskunit.WindHailDeductible);
                    if (matches.Count > 0)
                    {
                        builder.Write(string.Format("{0:n0}%", Convert.ToDecimal(matches[0].Groups["percent"].Value)));

                        if (Convert.ToDecimal(matches[0].Groups["amount"].Value) > windhailDeductibleAmount)
                        {
                            windhailDeductibleAmount = Convert.ToDecimal(matches[0].Groups["amount"].Value);
                        }
                    }

                    builder.EndRow();
                }

                builder.Font.Bold = true;

                builder.InsertCell();
                builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.First;
                builder.Write(string.Format("Minimum Deductible: {0:C0} per occurrence", windhailDeductibleAmount));
                builder.InsertCell();
                builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
                builder.InsertCell();
                builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
                builder.EndRow();

                builder.Font.Bold = false;

                builder.InsertCell();
                builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.First;
                builder.Write("Information required to complete this Schedule, if not shown above, will be shown in the Declarations.");
                builder.InsertCell();
                builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
                builder.InsertCell();
                builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
                builder.EndRow();

                table.StyleOptions = Aspose.Words.Tables.TableStyleOptions.Default2003;
                table.LeftPadding = 3;
                table.RightPadding = 3;
                table.TopPadding = 3;
                table.BottomPadding = 3;
                table.LeftIndent = 2.85;
                builder.EndTable();
            }
        }

        /// <summary>
        /// Generate MECP 1244 form
        /// </summary>
        /// <param name="policy">Policy</param>
        /// <param name="documentManager">Document manager</param>
        /// <param name="extraParameters">Extra parameters</param>
        public void GenerateMECP1244Form(Policy policy, ref IPolicyDocumentManager documentManager, Dictionary<string, object> extraParameters)
        {
            if (policy.CfLine != null)
            {
                Regex pattern = new Regex(@"(?<percent>.+?)PCT(?<amount>.+?)MIN");

                var windHailDeductible = policy.CfLine.RiskUnits.Select(r => pattern.Matches(r.WindHailDeductible)).Where(r => r.Count > 0)
                    .Select(r => new { Percent = r[0].Groups["percent"].Value, Amount = r[0].Groups["amount"].Value })
                    .OrderByDescending(r => r.Percent).OrderByDescending(r => r.Amount).FirstOrDefault();

                this.FillBookmarkField(ref documentManager, "BlankMergeField", string.Format("{0:n0}", Convert.ToDecimal(windHailDeductible.Percent)));
                this.FillBookmarkField(ref documentManager, "BlankMergeField_1", string.Format("{0:n0}", Convert.ToDecimal(windHailDeductible.Amount)));
            }
        }

        /// <summary>
        /// Generate MECP 1207 form
        /// </summary>
        /// <param name="policy">Policy</param>
        /// <param name="documentManager">Document manager</param>
        /// <param name="extraParameters">Extra parameters</param>
        public void GenerateMECP1207Form(Policy policy, ref IPolicyDocumentManager documentManager, Dictionary<string, object> extraParameters)
        {
            if (policy.CfLine != null)
            {
                DocumentBuilder builder = new DocumentBuilder(documentManager.Document);

                builder.MoveToBookmark("TheftSublimitTable");

                Aspose.Words.Tables.Table table = builder.StartTable();
                builder.Font.Size = 10;
                builder.Font.Bold = true;
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;

                builder.InsertCell();
                builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(58.6);
                builder.Write("Premises No.");

                builder.InsertCell();
                builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(54);
                builder.Write("Building No.");

                builder.InsertCell();
                builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(171);
                builder.Write("Theft Sublimit – Each Occurrence");

                builder.InsertCell();
                builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(220.55);
                builder.Write("Specified Covered Property*");

                builder.EndRow();
                builder.CellFormat.ClearFormatting();
                builder.Font.Bold = false;

                foreach (var riskunit in policy.CfLine.RiskUnits.Where(r => r.CauseOfLoss == "SP" && r.TheftSublimit.HasValue))
                {
                    builder.InsertCell();
                    builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(58.6);
                    builder.Write(riskunit.PremisesNumber.ToString());

                    builder.InsertCell();
                    builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(54);
                    builder.Write(riskunit.BuildingNumber.ToString());

                    builder.InsertCell();
                    builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(171);
                    builder.Write(string.Format("${0:n0}", riskunit.TheftSublimit));

                    builder.InsertCell();
                    builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(220.55);
                    builder.Write(string.IsNullOrEmpty(riskunit.TheftSpecifiedProperty) ? string.Empty : riskunit.TheftSpecifiedProperty);

                    builder.EndRow();
                }

                table.Alignment = Aspose.Words.Tables.TableAlignment.Center;
                table.StyleOptions = Aspose.Words.Tables.TableStyleOptions.Default2003;
                table.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(504.15);
                table.LeftPadding = 3;
                table.RightPadding = 3;
                table.TopPadding = 3;
                table.BottomPadding = 3;
                table.LeftIndent = 2.85;
                builder.EndTable();
            }
        }

        /// <summary>
        /// Inserts Vehicle Schedule table for IM MTC 003/004
        /// </summary>
        /// <param name="policy">Passed policy object</param>
        /// <param name="documentManager">Passed document manager</param>
        /// <param name="extraParameters">Dictionary of extra parameters</param>
        public void InsertVehicleScheduleTableIMMTC(Policy policy, ref IPolicyDocumentManager documentManager, Dictionary<string, object> extraParameters)
        {
            var blankMergeFieldName = "VehicleScheduleTable";

            var builder = new DocumentBuilder(documentManager.Document);
            builder.MoveToBookmark(blankMergeFieldName);
            this.FillBookmarkField(ref documentManager, string.Format(blankMergeFieldName), string.Empty);

            builder.Font.Size = 10;
            builder.Font.Bold = true;
            builder.Font.Name = "Arial";

            Aspose.Words.Tables.Table table = builder.StartTable();

            builder.CellFormat.ClearFormatting();
            builder.CellFormat.TopPadding = 3;

            //// populate headings
            var cell = builder.InsertCell();
            cell.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(36);
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
            builder.Write("UNIT");

            builder.InsertCell();
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(45);
            builder.Write("YEAR");

            builder.InsertCell();
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(99);
            builder.Write("MANUFACTURER");

            builder.InsertCell();
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(99);
            builder.Write("IDENTIFICATION NUMBER");

            builder.InsertCell();
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(54);
            builder.Write("BODY TYPE");

            builder.InsertCell();
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(85.5);
            builder.Write("LIMIT OF LIABILITY PER VEHICLE");

            builder.InsertCell();
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(117);
            builder.Write("PREMIUM PER VEHICLE");
            builder.EndRow();

            var document = policy.ImLine.RollUp.Documents.FirstOrDefault(x => x.NormalizedNumber.Equals("IMMTC003") && x.ReturnIsSelected());
            if (document == null)
            {
                document = policy.ImLine.RollUp.Documents.FirstOrDefault(x => x.NormalizedNumber.Equals("IMMTC004") && x.ReturnIsSelected());
            }

            builder.Font.Bold = false;

            var vehicleSchedules = this.GetVehicleSchedule(document.Questions);

            var limitOfLiabilityPerVehicle = policy.ImLine.RiskUnits.FirstOrDefault(x => x.ClassCode.Equals("445"))?.VehicleLiabilityLimit;
            var premiumPreVehicle = policy.ImLine.RiskUnits.FirstOrDefault(x => x.ClassCode.Equals("445"))?.ApplicableRate;

            var rowNumber = 1;
            foreach (var schedule in vehicleSchedules)
            {
                builder.CellFormat.ClearFormatting();
                builder.CellFormat.TopPadding = 3;

                builder.InsertCell();
                builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(36);
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
                builder.Write(string.Format("{0}.", rowNumber));

                builder.InsertCell();
                builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(45);
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                builder.CellFormat.LeftPadding = 5;
                builder.Write(!string.IsNullOrWhiteSpace(schedule.Year) ? schedule.Year : string.Empty);

                builder.InsertCell();
                builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(99);
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                builder.CellFormat.LeftPadding = 5;
                builder.Write(!string.IsNullOrWhiteSpace(schedule.Manufacturer) ? schedule.Manufacturer : string.Empty);

                builder.InsertCell();
                builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(99);
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                builder.CellFormat.LeftPadding = 5;
                builder.Write(!string.IsNullOrWhiteSpace(schedule.IdentificationNumber) ? schedule.IdentificationNumber : string.Empty);

                builder.InsertCell();
                builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(54);
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                builder.CellFormat.LeftPadding = 5;
                builder.Write(!string.IsNullOrWhiteSpace(schedule.BodyType) ? schedule.BodyType : string.Empty);

                builder.InsertCell();
                builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(85.5);
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                builder.CellFormat.LeftPadding = 0;
                builder.CellFormat.RightPadding = 5;
                builder.Write(limitOfLiabilityPerVehicle != null ? string.Format("{0:C0}", limitOfLiabilityPerVehicle) : string.Empty);

                builder.InsertCell();
                builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(117);
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                builder.CellFormat.LeftPadding = 0;
                builder.CellFormat.RightPadding = 5;
                builder.Write(string.Format("$ {0:N2}", premiumPreVehicle));
                builder.EndRow();
                rowNumber++;
            }

            builder.InsertCell();
            builder.CellFormat.LeftPadding = 5;
            builder.CellFormat.RightPadding = 5;
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.First;
            builder.InsertCheckBox("check3", false, 0);
            builder.Write(" IF CHECKED, BLANKET COVERAGE ON ALL VEHICLES OPERATED BY THE ");
            builder.Font.Bold = true;
            builder.Write("INSURED");
            builder.Font.Bold = false;
            builder.Write(", INCLUDING DETACHED TRAILERS(SCHEDULE WAIVED).");
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.EndRow();

            builder.CellFormat.ClearFormatting();

            table.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(535.5);
            table.StyleOptions = Aspose.Words.Tables.TableStyleOptions.Default2003;
            table.LeftPadding = 0;
            table.RightPadding = 0;
            ////table.LeftIndent = 5.4;
            builder.EndTable();
        }

        /// <summary>
        /// Generate MECP 1281 form for property
        /// </summary>
        /// <param name="policy">Passed policy object</param>
        /// <param name="documentManager">Passed document manager</param>
        /// <param name="extraParameters">Extra parameters</param>
        public void GenerateMECP1281Form(Policy policy, ref IPolicyDocumentManager documentManager, Dictionary<string, object> extraParameters)
        {
            List<string> questionCodeList = new List<string> { "PremisesNo", "BuildingNo" };
            List<List<string>> blankMergeFieldList = new List<List<string>>
            {
                new List<string> { string.Empty, "2", "4" },
                new List<string> { "1", "3", "5" },
            };

            ProcessCustomTable(policy, documentManager, blankMergeFieldList, questionCodeList);
        }

        /// <summary>
        /// Generate MECP 1321 form for property
        /// </summary>
        /// <param name="policy">Passed policy object</param>
        /// <param name="documentManager">Passed document manager</param>
        /// <param name="extraParameters">Extra parameters</param>
        public void GenerateMECP1321Form(Policy policy, ref IPolicyDocumentManager documentManager, Dictionary<string, object> extraParameters)
        {
            List<string> questionCodeList = new List<string> { "PremisesNo", "BuildingNo" };
            List<List<string>> blankMergeFieldList = new List<List<string>>
            {
                new List<string> { string.Empty, "2", "4" },
                new List<string> { "1", "3", "5" },
            };

            ProcessCustomTable(policy, documentManager, blankMergeFieldList, questionCodeList);
        }

        /// <summary>
        /// Get the bookmark control from a word  template using bookmark name
        /// </summary>
        /// <param name="range">The range.</param>
        /// <param name="bookmarkName">The bookmark name.</param>
        /// <returns>Bookmark</returns>
        public Bookmark GetBookMarkByName(Range range, string bookmarkName)
        {
            return range.Bookmarks[bookmarkName];
        }

        /// <summary>
        /// Processes the table from a custom function.
        /// </summary>
        /// <param name="policy">The policy.</param>
        /// <param name="documentManager">The document manager.</param>
        /// <param name="blankMergeFieldList">The blank merge field list.</param>
        /// <param name="questionCodeList">The question code list.</param>
        /// <param name="removeGridLines">The remove grid lines.</param>
        /// <param name="rowHandler">The row handler.</param>
        /// <param name="cellHandlers">The cell handlers.</param>
        /// <param name="formatFunctions">The named format function per bookmark</param>
        /// <exception cref="System.Exception">Bookmark not contained in a table.</exception>
        private static void ProcessCustomTable(Policy policy, IPolicyDocumentManager documentManager, List<List<string>> blankMergeFieldList, List<string> questionCodeList, RemoveGridLines removeGridLines = RemoveGridLines.None, RowCallbackType rowHandler = null, List<CellCallbackType> cellHandlers = null, Dictionary<string, Func<string, string>> formatFunctions = null)
        {
            // Get the builder for the document
            var builder = new DocumentBuilder(documentManager.Document);

            // Move to the first bookmark we are looking for
            builder.MoveToBookmark(GetBlankMergeFieldName(blankMergeFieldList, 0, 0));

            // Set this as the table we are going to work
            var table = builder.CurrentNode.GetParentTable();
            if (table == null)
            {
                // this is an error...create the table manually?
                throw new Exception("Bookmark not contained in a table.");
            }

            // Get the questions for this form and quantity order
            var formNumber = documentManager.FormNormalizedNumber;
            var quantityOrder = documentManager.QuantityOrder;
            var questions = policy.RollUp?.Documents?.FirstOrDefault(d => d.NormalizedNumber == formNumber &&
                                                                          d.QuantityOrder.GetValueOrDefault(0) == quantityOrder)?.Questions?.OrderBy(o => o.MultipleRowGroupingNumber)
                .ToList();

            if (questions == null ||
                !questions.Any())
            {
                // No questions returned, nothing to do then....
                return;
            }

            // Get a list of the grouping numbers
            var groupingNumbers = questions.Where(s => s.MultipleRowGroupingNumber > 0)
                                            .Select(s => s.MultipleRowGroupingNumber)
                                            .Distinct()
                                            .ToList();

            // Get the number of rows defined in the blankMergeFieldList
            var numberOfDefinedRows = blankMergeFieldList[0].Count;

            // Get the number of columns defined in the blankMergeFieldList
            var numberOfColumns = blankMergeFieldList.Count;

            if (questionCodeList.Count < numberOfColumns)
            {
                numberOfColumns = questionCodeList.Count;
            }

            // Get the number of grouping numbers which equates to how many rows will be added to the table.
            var numberOfDataRows = groupingNumbers.Count;
            if (numberOfDataRows > numberOfDefinedRows)
            {
                // If there are more rows of data then rows in the table, then we need to find an appropriate "cloned" row
                // Since the first row of the table and the last row of the table tend to have borders on top but not bottom
                // or on the bottom but not the top due to the table having a border but the rows not having one. So get
                // the second row as our "cloned".
                // Subtracting 2 since the list is 0 based and the numberOfDefinedRows is 1 based.   Since we want to get the
                // previous row, subtract 2.
                var secondToLastRowNumber = numberOfDefinedRows - 2;

                // Find the bookmark name for the 1st column in the second row.
                var bookmarkName = GetBlankMergeFieldName(blankMergeFieldList, 0, secondToLastRowNumber);

                // Get the second row
                var secondToLastRow = table.Range?.Bookmarks[bookmarkName]?.BookmarkStart.GetAncestor(NodeType.Row);
                if (secondToLastRow == null)
                {
                    throw new Exception("Invalid bookmark name trying to find table row to clone.");
                }

                // Loop through and add any necessary rows, setting their bookmark names, and adding those names to
                // the list of BlankMergeField names
                for (var lastRowNumber = numberOfDefinedRows; lastRowNumber < numberOfDataRows; lastRowNumber++)
                {
                    // Clone the row
                    var clonedRow = (Row)secondToLastRow.Clone(true);

                    // Loop, through the columns and rename the bookmarks we are going to try to use
                    for (var columnNumber = 0; columnNumber < numberOfColumns; columnNumber++)
                    {
                        // Get the bookmark name for the column in question from the row we cloned.
                        bookmarkName = GetBlankMergeFieldName(blankMergeFieldList, columnNumber, secondToLastRowNumber);

                        // Find that bookmark in the cloned row.
                        var bookmark = clonedRow.Range.Bookmarks[bookmarkName];

                        // Set a new DISTINCT name for the bookmark
                        var newName = $"CustomFunctionAdd_{lastRowNumber}_{columnNumber}";

                        // Set the name on the bookmark
                        bookmark.Name = $"{BlankMergeField}{newName}";

                        blankMergeFieldList[columnNumber].Insert(secondToLastRowNumber + 1, newName);
                    }

                    // Add the "cloned" row to the table
                    table.InsertAfter(clonedRow, secondToLastRow);
                }
            }

            documentManager.Document.UpdateTableLayout();
            documentManager.Document.UpdatePageLayout();

            // Default the current rowNumber
            var rowNumber = 0;

            foreach (var groupingNumber in groupingNumbers)
            {
                var bookmarkName = GetBlankMergeFieldName(blankMergeFieldList, 0, rowNumber);
                var workingRow = table.Range?.Bookmarks[bookmarkName].BookmarkStart.GetCurrentRow();

                if (workingRow == null)
                {
                    continue;
                }

                rowHandler?.Invoke(documentManager, table, workingRow, rowNumber);

                for (var columnNumber = 0; columnNumber < numberOfColumns; columnNumber++)
                {
                    if (blankMergeFieldList[columnNumber][rowNumber] == null)
                    {
                        continue;
                    }

                    bookmarkName = GetBlankMergeFieldName(blankMergeFieldList, columnNumber, rowNumber);
                    var question = questions.FirstOrDefault(q => q.MultipleRowGroupingNumber == groupingNumber &&
                                                                 q.Code == questionCodeList[columnNumber]);
                    if (question == null)
                    {
                        continue;
                    }

                    var bookmark = workingRow.Range.Bookmarks[bookmarkName];
                    bool? replace = true;
                    if (cellHandlers != null &&
                        cellHandlers.Count > columnNumber &&
                        cellHandlers[columnNumber] != null)
                    {
                        replace = cellHandlers[columnNumber]?.Invoke(documentManager, table, workingRow, rowNumber, columnNumber, question, bookmark, bookmarkName);
                    }

                    if (replace.GetValueOrDefault(false))
                    {
                        Func<string, string> formatFunction = null;
                        if (formatFunctions != null &&
                            formatFunctions.ContainsKey(bookmarkName))
                        {
                            formatFunction = formatFunctions[bookmarkName];
                        }

                        documentManager.ReplaceQuestionBookmarkValue(bookmark, question, formatFunction);
                    }
                }

                rowNumber++;
            }
        }

        /// <summary>
        /// Gets a blank merge field name based on row and column
        /// </summary>
        /// <param name="blankMergeFieldList">The list of blank merge field names</param>
        /// <param name="columnNumber">The column number</param>
        /// <param name="workingRowNumber">The working row number</param>
        /// <returns>String containing blank merge field name</returns>
        private static string GetBlankMergeFieldName(List<List<string>> blankMergeFieldList, int columnNumber, int workingRowNumber)
        {
            string bookmarkName = BlankMergeField.Replace("_", string.Empty);
            var number = blankMergeFieldList[columnNumber][workingRowNumber];
            if (!string.IsNullOrWhiteSpace(number))
            {
                bookmarkName += $"_{number}";
            }

            return bookmarkName;
        }

        /// <summary>
        /// Fill Classification and Premium bookmark field
        /// </summary>
        /// <param name="lob">IBaseGlLine</param>
        /// <param name="documentManager">PolicyDocumentManager</param>
        private void FillBlankMergeFieldValueInMDGL1008Form(DecisionModel.Interfaces.IBaseGlLine lob, ref IPolicyDocumentManager documentManager)
        {
            List<List<int>> rowMergeField = new List<List<int>>()
                {
                    { new List<int> { 17, 18, 19, 20, 22, 23, 24, 25, 26 } },
                    { new List<int> { 28, 28, 29, 30, 31, 32, 33, 34, 35 } },
                    { new List<int> { 36, 37, 38, 39, 40, 41, 43, 44, 45 } },
                    { new List<int> { 46, 75, 76, 77, 78, 79, 80, 42, 81 } },
                    { new List<int> { 82, 47, 48, 49, 50, 51, 52, 53, 83 } },
                    { new List<int> { 54, 55, 56, 57, 58, 59, 60, 61, 62 } },
                    { new List<int> { 63, 64, 65, 66, 67, 68, 69, 70, 71 } }
                };

            int loopIndex = 0, totalRow = rowMergeField.Count;

            if (lob.GlRiskUnits.Any())
            {
                int count = lob.GlRiskUnits.Count();
                var riskUnits = lob.GlRiskUnits.ToList();
                for (loopIndex = 0; loopIndex < count && loopIndex < totalRow; loopIndex++)
                {
                    var riskUnit = riskUnits[loopIndex] as BaseGlRiskUnit;
                    if (riskUnit != null)
                    {
                        this.FillClassificationAndPremiumInMDGL1008n1009Form(ref documentManager, rowMergeField[loopIndex], riskUnit);
                    }
                }
            }

            if (loopIndex < totalRow && lob.RollUp.OptionalCoverages.Any(c => c.IsSelected == true))
            {
                var optionalCoverages = lob.RollUp.OptionalCoverages.Where(c => c.IsSelected == true).ToList();

                int countOC = optionalCoverages.Count();
                for (int i = 0; i < countOC && i < totalRow - loopIndex; i++)
                {
                    var optionalCoverage = optionalCoverages[i];
                    if (optionalCoverage != null)
                    {
                        this.FillClassificationAndPremiumInMDGL1008n1009FormOC(ref documentManager, rowMergeField[loopIndex + i], optionalCoverage);
                    }
                }
            }

            this.FillBookmarkField(ref documentManager, "BlankMergeField_85", string.Empty);
            this.FillBookmarkField(ref documentManager, "BlankMergeField_72", string.Format("{0:n0}", lob.RollUp.Premium));
            if (lob.RollUp.Documents.Any(d => d.NormalizedNumber == "MDGL1009"))
            {
                this.FillBookmarkField(ref documentManager, "Check1", "CHECKED");
            }
        }

        /// <summary>
        /// Create a merge field list to contain identifiers for blank merge fields.
        /// </summary>
        /// <param name="startIndex">the start index identifier </param>
        /// <param name="endIndex">end index identifier</param>
        /// <returns>List containing the merge field identifiers</returns>
        private List<int> CreateMergeFieldRowMDGL1009(int startIndex, ref int endIndex)
        {
            List<int> mergeFieldRow = new List<int>(9);
            int strIndex = startIndex;
            endIndex = endIndex + 8;
            while (strIndex <= endIndex)
            {
                if (strIndex == 88 || strIndex == 97)
                {
                    strIndex++;
                    endIndex++;
                }

                mergeFieldRow.Add(strIndex);
                strIndex++;
            }

            endIndex++;
            return mergeFieldRow;
        }

        /// <summary>
        /// Fill Classification and Premium bookmark field
        /// </summary>
        /// <param name="documentManager">PolicyDocumentManager</param>
        /// <param name="list">List</param>
        /// <param name="riskUnit">BaseGlRiskUnit</param>
        /// <param name="isMDGl1008">is 1008 form attached</param>
        /// <returns>returns value of other premium</returns>
        private decimal FillClassificationAndPremiumInMDGL1008n1009Form(ref IPolicyDocumentManager documentManager, List<int> list, BaseGlRiskUnit riskUnit, bool isMDGl1008 = true)
        {
            this.FillBookmarkField(ref documentManager, string.Format("BlankMergeField_{0}", list[0]), string.Empty);
            this.FillBookmarkField(ref documentManager, string.Format("BlankMergeField_{0}", list[1]), string.Format("{0} - {1}", riskUnit.ClassCode, riskUnit.ClassDescription));
            this.FillBookmarkField(ref documentManager, string.Format("BlankMergeField_{0}", list[2]), riskUnit.PremiumBaseUnitMeasure);

            decimal premium = 0m;
            if (riskUnit.IsIncluded)
            {
                this.FillBookmarkField(ref documentManager, string.Format("BlankMergeField_{0}", list[3]), "Included");
            }
            else if (riskUnit.IsIfAny)
            {
                this.FillBookmarkField(ref documentManager, string.Format("BlankMergeField_{0}", list[3]), "If Any");
            }
            else if (riskUnit.PremiumBaseUnitMeasure == "Gross Sales")
            {
                this.FillBookmarkField(ref documentManager, string.Format("BlankMergeField_{0}", list[3]), string.Format("{0:C0}", riskUnit.Exposure));
            }
            else
            {
                this.FillBookmarkField(ref documentManager, string.Format("BlankMergeField_{0}", list[3]), string.Format("{0:n0}", riskUnit.Exposure));
            }

            this.FillBookmarkField(ref documentManager, string.Format("BlankMergeField_{0}", list[4]), string.Empty);
            this.FillBookmarkField(ref documentManager, string.Format("BlankMergeField_{0}", list[5]), string.Empty);

            if (riskUnit.IsIncluded)
            {
                this.FillBookmarkField(ref documentManager, string.Format("BlankMergeField_{0}", list[6]), "Included");
            }
            else if (riskUnit.UnderwriterAdjustedRate != 0)
            {
                string adjrate = string.Format("{0:n2}", Math.Round(riskUnit.UnderwriterAdjustedRate, 2, MidpointRounding.AwayFromZero));
                if (!isMDGl1008)
                {
                    adjrate = string.Concat("$", adjrate);
                }

                this.FillBookmarkField(ref documentManager, string.Format("BlankMergeField_{0}", list[6]), adjrate);
            }
            else if (riskUnit.AgentAdjustedRate != 0)
            {
                string adjrate = string.Format("{0:n2}", Math.Round(riskUnit.AgentAdjustedRate, 2, MidpointRounding.AwayFromZero));
                if (!isMDGl1008)
                {
                    adjrate = string.Concat("$", adjrate);
                }

                this.FillBookmarkField(ref documentManager, string.Format("BlankMergeField_{0}", list[6]), adjrate);
            }
            else
            {
                string adjrate = string.Format("{0:n2}", Math.Round(riskUnit.DevelopedRate, 2, MidpointRounding.AwayFromZero));
                if (!isMDGl1008)
                {
                    adjrate = string.Concat("$", adjrate);
                }

                this.FillBookmarkField(ref documentManager, string.Format("BlankMergeField_{0}", list[6]), adjrate);
            }

            this.FillBookmarkField(ref documentManager, string.Format("BlankMergeField_{0}", list[7]), string.Empty);

            if (riskUnit.IsIncluded || (riskUnit.UnderwriterAdjustedPremium.HasValue && riskUnit.UnderwriterAdjustedPremium == 0) || (riskUnit.AgentAdjustedPremium.HasValue && riskUnit.AgentAdjustedPremium == 0) || (riskUnit.UnderwriterAdjustedPremium.HasValue == false && riskUnit.AgentAdjustedPremium.HasValue == false && riskUnit.Premium == 0))
            {
                this.FillBookmarkField(ref documentManager, string.Format("BlankMergeField_{0}", list[8]), "Included");
            }
            else if (riskUnit.UnderwriterAdjustedPremium.HasValue)
            {
                premium = premium + riskUnit.UnderwriterAdjustedPremium.Value;
                this.FillBookmarkField(ref documentManager, string.Format("BlankMergeField_{0}", list[8]), string.Format("{0:n0}", riskUnit.UnderwriterAdjustedPremium));
            }
            else if (riskUnit.AgentAdjustedPremium.HasValue)
            {
                premium = premium + riskUnit.AgentAdjustedPremium.Value;
                this.FillBookmarkField(ref documentManager, string.Format("BlankMergeField_{0}", list[8]), string.Format("{0:n0}", riskUnit.AgentAdjustedPremium));
            }
            else
            {
                premium = premium + riskUnit.Premium;
                this.FillBookmarkField(ref documentManager, string.Format("BlankMergeField_{0}", list[8]), string.Format("{0:n0}", riskUnit.Premium));
            }

            return premium;
        }

        /// <summary>
        /// Fill Classification and Premium bookmark field for Optional Coverages
        /// </summary>
        /// <param name="documentManager">PolicyDocumentManager</param>
        /// <param name="list">List</param>
        /// <param name="optionalCoverage">OptionalCoverage</param>
        /// <param name="isMDGl1008">is MDGL form 1008 attached</param>
        /// <returns>decimal optional coverage other premium value</returns>
        private decimal FillClassificationAndPremiumInMDGL1008n1009FormOC(ref IPolicyDocumentManager documentManager, List<int> list, OptionalCoverage optionalCoverage, bool isMDGl1008 = true)
        {
            decimal premium = 0m;
            this.FillBookmarkField(ref documentManager, string.Format("BlankMergeField_{0}", list[0]), string.Empty);
            this.FillBookmarkField(ref documentManager, string.Format("BlankMergeField_{0}", list[1]), string.Empty);
            this.FillBookmarkField(ref documentManager, string.Format("BlankMergeField_{0}", list[2]), string.Empty);
            this.FillBookmarkField(ref documentManager, string.Format("BlankMergeField_{0}", list[3]), string.Empty);
            this.FillBookmarkField(ref documentManager, string.Format("BlankMergeField_{0}", list[4]), string.Empty);
            this.FillBookmarkField(ref documentManager, string.Format("BlankMergeField_{0}", list[5]), string.Empty);

            if (optionalCoverage.IsIncluded || (optionalCoverage.AdjustedRate.HasValue && optionalCoverage.AdjustedRate == 0))
            {
                this.FillBookmarkField(ref documentManager, string.Format("BlankMergeField_{0}", list[6]), "Included");
            }
            else if (optionalCoverage.AdjustedRate.HasValue)
            {
                string adjrate = string.Format("{0:n2}", Math.Round(optionalCoverage.AdjustedRate.Value, 2, MidpointRounding.AwayFromZero));
                if (!isMDGl1008)
                {
                    adjrate = string.Concat("$", adjrate);
                }

                this.FillBookmarkField(ref documentManager, string.Format("BlankMergeField_{0}", list[6]), adjrate);
            }
            else if (optionalCoverage.Basis == RateBasis.ClassPremium || optionalCoverage.Basis == RateBasis.LinePremium)
            {
                this.FillBookmarkField(ref documentManager, string.Format("BlankMergeField_{0}", list[6]), string.Format("{0:P0}", optionalCoverage.Rate.Value));
            }
            else
            {
                string adjrate = string.Format("{0:n2}", Math.Round(optionalCoverage.Rate.Value, 2, MidpointRounding.AwayFromZero));
                if (!isMDGl1008)
                {
                    adjrate = string.Concat("$", adjrate);
                }

                this.FillBookmarkField(ref documentManager, string.Format("BlankMergeField_{0}", list[6]), adjrate);
            }

            this.FillBookmarkField(ref documentManager, string.Format("BlankMergeField_{0}", list[7]), string.Empty);

            if (optionalCoverage.IsIncluded || (optionalCoverage.AdjustedPremium.HasValue && optionalCoverage.AdjustedPremium == 0) || (optionalCoverage.AdjustedPremium.HasValue == false && optionalCoverage.Premium == 0))
            {
                this.FillBookmarkField(ref documentManager, string.Format("BlankMergeField_{0}", list[8]), "Included");
            }
            else if (optionalCoverage.AdjustedPremium.HasValue)
            {
                this.FillBookmarkField(ref documentManager, string.Format("BlankMergeField_{0}", list[8]), string.Format("{0:n0}", optionalCoverage.AdjustedPremium));
                premium = premium + optionalCoverage.AdjustedPremium.Value;
            }
            else
            {
                this.FillBookmarkField(ref documentManager, string.Format("BlankMergeField_{0}", list[8]), string.Format("{0:n0}", optionalCoverage.Premium));
                premium = premium + optionalCoverage.Premium;
            }

            return premium;
        }

        /// <summary>
        /// Fill bookmark field value for Optional Coverages
        /// </summary>
        /// <param name="item">OptionalCoverage</param>
        /// <param name="builder">DocumentBuilder</param>
        private void FillBlankMergeFieldValueInMDGL1000n1012FormOC(OptionalCoverage item, DocumentBuilder builder)
        {
            builder.CellFormat.TopPadding = 4;
            builder.CellFormat.BottomPadding = 4;

            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.Write(item.ShortName);
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;

            if (item.IsIncluded)
            {
                builder.Write("Incl.");
            }
            else if (item.Basis == RateBasis.Unit)
            {
                builder.Write("Each");
            }
            else if (item.Basis == RateBasis.Flat)
            {
                builder.Write("Flat");
            }
            else if (item.Basis == RateBasis.ClassPremium || item.Basis == RateBasis.LinePremium)
            {
                builder.Write("Percent of Premium");
            }
            else
            {
                builder.Write(string.Empty);
            }

            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;

            if (item.IsIncluded || (item.AdjustedRate.HasValue && item.AdjustedRate == 0))
            {
                builder.Write("Incl.");
            }
            else if (item.Basis == RateBasis.Unit)
            {
                builder.Write(string.Empty);
            }
            else if (item.Basis == RateBasis.Flat)
            {
                builder.Write("Flat");
            }
            else if (item.Basis == RateBasis.ClassPremium || item.Basis == RateBasis.LinePremium)
            {
                builder.Write(string.Empty);
            }
            else if (item.AdjustedRate.HasValue)
            {
                builder.Write(string.Format("${0:n2}", Math.Round(item.AdjustedRate.Value, 2, MidpointRounding.AwayFromZero)));
            }
            else
            {
                builder.Write(string.Format("${0:n2}", Math.Round(item.Rate.Value, 2, MidpointRounding.AwayFromZero)));
            }

            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;

            if (item.IsIncluded)
            {
                builder.Write("Incl.");
            }
            else if (item.AdjustedPremium.HasValue)
            {
                builder.Write(string.Format("${0:n0}", item.AdjustedPremium.Value));
            }
            else
            {
                builder.Write(string.Format("${0:n0}", item.Premium));
            }

            builder.EndRow();
        }

        /// <summary>
        /// MDGL 1012 Form Classification Table Header
        /// </summary>
        /// <param name="builder">DocumentBuilder</param>
        private void MDGL1012FormClassificationTableHeader(DocumentBuilder builder)
        {
            builder.CellFormat.TopPadding = 4;
            builder.CellFormat.BottomPadding = 4;

            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.First;
            builder.Write("Classification And Premium");
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.EndRow();

            // populate headings
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.None;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPercent(234.15);
            builder.Write("Classification");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPercent(72);
            builder.Write("Code\nNo.");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPercent(58.5);
            builder.Write("Premium\nBase");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPercent(85.5);
            builder.Write("Rate Per\n $1,000 Of Cost");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPercent(94.35);
            builder.Write("Advance\nPremium");
            builder.EndRow();

            builder.CellFormat.ClearFormatting();
            builder.Font.Bold = false;
        }

        /// <summary>
        /// MDGL 1000 Form Locations table
        /// </summary>
        /// <param name="policy">Policy</param>
        /// <param name="builder">DocumentBuilder</param>
        private void MDGL1000Classification(Policy policy, DocumentBuilder builder)
        {
            builder.MoveToBookmark("BlankMergeField_3");

            builder.Font.Size = 10;
            builder.Font.Bold = true;
            Aspose.Words.Tables.Table table = builder.StartTable();

            this.MDGL1000FormClassificationTableHeader(builder);
            if (policy.LiquorLine.RiskUnits.Any())
            {
                foreach (var riskUnit in policy.LiquorLine.RiskUnits)
                {
                    this.FillBlankMergeFieldValueInMDGL1000n1012Form(riskUnit, builder, riskUnit.Exposure);
                }
            }

            if (policy.LiquorLine.RollUp.OptionalCoverages.Any(c => c.Code != "TRIA" && c.IsSelected == true))
            {
                foreach (var optionalCoverage in policy.LiquorLine.RollUp.OptionalCoverages.Where(c => c.Code != "TRIA" && c.IsSelected == true))
                {
                    this.FillBlankMergeFieldValueInMDGL1000n1012FormOC(optionalCoverage, builder);
                }
            }

            this.MDGL1000n1012FormTotalPremium(policy, builder);

            table.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(544.5);
            table.StyleOptions = Aspose.Words.Tables.TableStyleOptions.Default2003;
            table.LeftPadding = 3.05;
            table.RightPadding = 3.05;
            table.LeftIndent = 3.5;
            builder.EndTable();
        }

        /// <summary>
        /// Generate MDGL 1012 tabular form
        /// </summary>
        /// <param name="policy">Policy</param>
        /// <param name="documentManager">IPolicyDocumentManager</param>
        private void MDGL1012ClassificationTable(Policy policy, ref IPolicyDocumentManager documentManager)
        {
            DocumentBuilder builder = new DocumentBuilder(documentManager.Document);
            builder.MoveToBookmark("BlankMergeField_12");

            builder.Font.Size = 10;
            builder.Font.Bold = true;
            Aspose.Words.Tables.Table table = builder.StartTable();

            this.MDGL1012FormClassificationTableHeader(builder);
            if (policy.OCPLine.RiskUnits.Any())
            {
                foreach (var riskUnit in policy.OCPLine.RiskUnits)
                {
                    this.FillBlankMergeFieldValueInMDGL1000n1012Form(riskUnit, builder, riskUnit.Exposure);
                }
            }

            if (policy.OCPLine.RollUp.OptionalCoverages.Any(c => c.Code != "TRIA" && c.IsSelected == true))
            {
                foreach (var optionalCoverage in policy.OCPLine.RollUp.OptionalCoverages.Where(c => c.Code != "TRIA" && c.IsSelected == true))
                {
                    this.FillBlankMergeFieldValueInMDGL1000n1012FormOC(optionalCoverage, builder);
                }
            }

            this.MDGL1000n1012FormTotalPremium(policy, builder, false);

            table.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(544.5);
            table.StyleOptions = Aspose.Words.Tables.TableStyleOptions.Default2003;
            table.LeftPadding = 3.05;
            table.RightPadding = 3.05;
            table.LeftIndent = 3.5;
            builder.EndTable();
        }

        /// <summary>
        /// Fill bookmark field value
        /// </summary>
        /// <param name="item">Risk Unit Item</param>
        /// <param name="builder">DocumentBuilder</param>
        /// <param name="exposure">exposure</param>
        private void FillBlankMergeFieldValueInMDGL1000n1012Form(RiskUnit item, DocumentBuilder builder, long exposure)
        {
            builder.CellFormat.TopPadding = 4;
            builder.CellFormat.BottomPadding = 4;

            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.Write(item.ClassDescription);

            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.Write(item.ClassCode);

            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;

            if (item.IsIncluded)
            {
                builder.Write("Incl.");
            }
            else if (item.IsIfAny)
            {
                builder.Write("If Any");
            }
            else
            {
                builder.Write(string.Format("${0:n0}", exposure));
            }

            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;

            if (item.IsIncluded)
            {
                builder.Write("Incl.");
            }
            else if (item.UnderwriterAdjustedRate != 0)
            {
                builder.Write(string.Format("${0:n2}", Math.Round(item.UnderwriterAdjustedRate, 2, MidpointRounding.AwayFromZero)));
            }
            else if (item.AgentAdjustedRate != 0)
            {
                builder.Write(string.Format("${0:n2}", Math.Round(item.AgentAdjustedRate, 2, MidpointRounding.AwayFromZero)));
            }
            else
            {
                builder.Write(string.Format("${0:n2}", Math.Round(item.DevelopedRate, 2, MidpointRounding.AwayFromZero)));
            }

            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;

            if (item.IsIncluded)
            {
                builder.Write("Incl.");
            }
            else if (item.IsIfAny)
            {
                builder.Write("If Any");
            }
            else if (item.UnderwriterAdjustedPremium.HasValue)
            {
                builder.Write(string.Format("${0:n0}", item.UnderwriterAdjustedPremium));
            }
            else if (item.AgentAdjustedPremium.HasValue)
            {
                builder.Write(string.Format("${0:n0}", item.AgentAdjustedPremium));
            }
            else
            {
                builder.Write(string.Format("${0:n0}", item.Premium));
            }

            builder.EndRow();
        }

        /// <summary>
        /// MDGL 1000 Form Locations table
        /// </summary>
        /// <param name="policy">Policy</param>
        /// <param name="builder">DocumentBuilder</param>
        private void MDGL1000FormLocations(Policy policy, DocumentBuilder builder)
        {
            builder.MoveToBookmark("BlankMergeField_2");

            builder.Font.Size = 12;
            builder.Font.Bold = true;

            Aspose.Words.Tables.Table table = builder.StartTable();
            builder.InsertCell();
            builder.CellFormat.LeftPadding = 2;
            builder.CellFormat.RightPadding = 2;
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.First;
            builder.Write("All Premises You Own, Rent or Occupy");
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.EndRow();

            builder.Font.Size = 10;
            builder.Font.Bold = false;

            // populate headings
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.None;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(49.5);
            builder.CellFormat.Borders.Left.LineStyle = LineStyle.Single;
            builder.CellFormat.Borders.Right.LineStyle = LineStyle.None;
            builder.CellFormat.TopPadding = 3;
            builder.CellFormat.BottomPadding = 3;
            builder.Write("Loc. No.");

            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.Auto;
            builder.CellFormat.Borders.Left.LineStyle = LineStyle.None;
            builder.CellFormat.Borders.Right.LineStyle = LineStyle.Single;
            builder.CellFormat.TopPadding = 3;
            builder.CellFormat.BottomPadding = 3;
            builder.Write("Address of All Premises You Own, Rent or Occupy");
            builder.EndRow();

            builder.CellFormat.ClearFormatting();

            if (policy.LiquorLine.RiskUnits.Any())
            {
                foreach (var riskUnit in policy.LiquorLine.GetDistinctLocations().ToLookup(r => r).OrderBy(r => r.Key.LocationNumber))
                {
                    var location = riskUnit.First();

                    builder.InsertCell();
                    builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                    builder.CellFormat.TopPadding = 3;
                    builder.CellFormat.BottomPadding = 3;
                    builder.Write(location.LocationNumber.ToString());

                    builder.InsertCell();
                    var address = string.Empty;
                    address += !string.IsNullOrWhiteSpace(location.StreetAddress.Line1) ? location.StreetAddress.Line1 + ", " : string.Empty;
                    address += !string.IsNullOrWhiteSpace(location.StreetAddress.Line2) ? location.StreetAddress.Line2 + ", " : string.Empty;
                    address += !string.IsNullOrWhiteSpace(location.StreetAddress.City) ? location.StreetAddress.City + ", " : string.Empty;
                    address += !string.IsNullOrWhiteSpace(location.StreetAddress.StateCode) ? location.StreetAddress.StateCode + ", " : string.Empty;
                    address += !string.IsNullOrWhiteSpace(location.StreetAddress.ZipCode) ? location.StreetAddress.ZipCode : string.Empty;
                    builder.Write(address);

                    builder.EndRow();
                }
            }

            builder.CellFormat.ClearFormatting();

            table.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(544.5);
            table.StyleOptions = Aspose.Words.Tables.TableStyleOptions.Default2003;
            table.LeftPadding = 3.05;
            table.RightPadding = 3.05;
            table.LeftIndent = 2.5;
            builder.EndTable();
        }

        /// <summary>
        /// MDGL 1000 Form Classification Table Header
        /// </summary>
        /// <param name="builder">DocumentBuilder</param>
        private void MDGL1000FormClassificationTableHeader(DocumentBuilder builder)
        {
            builder.CellFormat.TopPadding = 4;
            builder.CellFormat.BottomPadding = 4;

            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.First;
            builder.Write("Classification And Premium");
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.EndRow();

            // populate headings
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.None;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPercent(234.15);
            builder.Write("Classification");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPercent(72);
            builder.Write("Code\nNo.");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPercent(58.5);
            builder.Write("Premium\nBase");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPercent(85.5);
            builder.Write("Rate");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPercent(94.35);
            builder.Write("Advance\nPremium");
            builder.EndRow();

            builder.CellFormat.ClearFormatting();
            builder.Font.Bold = false;
        }

        /// <summary>
        /// MDGL 1000 Form Total Premium
        /// </summary>
        /// <param name="policy">Policy</param>
        /// <param name="builder">DocumentBuilder</param>
        /// <param name="liquorLine">liquor Line or OCP line</param>
        private void MDGL1000n1012FormTotalPremium(Policy policy, DocumentBuilder builder, bool liquorLine = true)
        {
            builder.CellFormat.TopPadding = 4;
            builder.CellFormat.BottomPadding = 4;

            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.None;
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.First;
            builder.Write("Total Premium (Subject To Audit)");
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.None;
            if (liquorLine)
            {
                builder.Write(string.Format("${0}{1}", this.GetApplicablePremium(policy, LineOfBusiness.LL, FormatHelper.CurrencyFormat), policy.LiquorLine.RollUp.IsMinimumPremium ? " MP" : string.Empty));
            }
            else
            {
                builder.Write(string.Format("${0}{1}", this.GetApplicablePremium(policy, LineOfBusiness.OCP, FormatHelper.CurrencyFormat), policy.OCPLine.RollUp.IsMinimumPremium ? " MP" : string.Empty));
            }

            builder.EndRow();
        }

        /// <summary>
        /// Increment Index
        /// </summary>
        /// <param name="index">Index Value</param>
        /// <returns> BlankMergeField</returns>
        private string Incrementindex(ref int index)
        {
            // Bookmark BlankMergeField_37 is not present in document so excluded it from logic and jump to bookmark BlankMergeField_38 from BlankMergeField_36
            const int NeglectCountValue = 36;
            if (index.Equals(NeglectCountValue))
            {
                index = index + 1;
            }

            index = index + 1;
            return "BlankMergeField_" + index;
        }

        /// <summary>
        /// Indicates a value whether Employee Benefits Limit should be displayed or not.
        /// </summary>
        /// <param name="policy">Policy object</param>
        /// <returns>True if Employee Benefits Limit should be displayed else False.</returns>
        private bool ShouldEmployeeBenefitLimitBeDisplayed(Policy policy)
        {
            if (policy.XsLine.XsGeneralLiabilityCoverage.EmployeeBenefitsLimit != null &&
                !string.IsNullOrWhiteSpace(policy.XsLine.XsGeneralLiabilityCoverage.EmployeeBenefitsLimit.Code) &&
                !policy.XsLine.XsGeneralLiabilityCoverage.EmployeeBenefitsLimit.Code.Equals("N/A"))
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Fill bookmark field value
        /// </summary>
        /// <param name="documentManager">PolicyDocumentManager</param>
        /// <param name="bookmarkName">bookmark name</param>
        /// <param name="value">value to be replaced</param>
        private void FillBookmarkField(ref IPolicyDocumentManager documentManager, string bookmarkName, string value)
        {
            Bookmark bookMark = documentManager.GetBookMarkByName(bookmarkName);
            this.FillBookmarkField(ref documentManager, bookMark, value);
        }

        /// <summary>
        /// Fills the bookmark field.
        /// </summary>
        /// <param name="documentManager">The document manager.</param>
        /// <param name="bookmark">The bookmark.</param>
        /// <param name="value">The value.</param>
        private void FillBookmarkField(ref IPolicyDocumentManager documentManager, Bookmark bookmark, string value)
        {
            if (bookmark != null)
            {
                documentManager.ReplaceNodeValue(bookmark, value);
            }
        }

        /// <summary>
        /// Replaces the cell paragraph or bookmark.
        /// </summary>
        /// <param name="documentManager">The document manager.</param>
        /// <param name="cell">The cell.</param>
        /// <param name="value">The value.</param>
        private void ReplaceCellParagraphOrBookmark(ref IPolicyDocumentManager documentManager, Cell cell, string value)
        {
            var paragraph = cell.FirstParagraph;
            if (paragraph == null)
            {
                paragraph = new Paragraph(documentManager.Document);
                cell.AppendChild(paragraph);
            }

            var bookmarks = paragraph.Range.Bookmarks;
            if (bookmarks.Count > 0)
            {
                this.FillBookmarkField(ref documentManager, bookmarks[0], value);
            }
            else
            {
                paragraph.Runs.Clear();
                paragraph.AppendChild(new Run(documentManager.Document, value));
            }
        }

        /// <summary>
        /// Create a dictionary containing mapping details for MADUB 1003
        /// </summary>
        /// <param name="policy">Policy object</param>
        /// <returns>dictionary of mapping details</returns>
        private Dictionary<string, MappingDetail> CreateMADUB1003Dictionary(Policy policy)
        {
            Dictionary<string, MappingDetail> dictionaryMADUB1003 = new Dictionary<string, MappingDetail>();
            dictionaryMADUB1003.Add("BlankMergeField", new MappingDetail() { Constant = "Commercial General Liability" });
            dictionaryMADUB1003.Add("BlankMergeField_1", new MappingDetail() { Field = "Policy.XsLine.XsGeneralLiabilityCoverage.CarrierName" });
            dictionaryMADUB1003.Add("BlankMergeField_2", new MappingDetail() { Field = "Policy.XsLine.XsGeneralLiabilityCoverage.EffectiveDate", Format = FormatHelper.DateFormat });
            dictionaryMADUB1003.Add("BlankMergeField_3", new MappingDetail() { Field = "Policy.XsLine.XsGeneralLiabilityCoverage.ExpirationDate", Format = FormatHelper.DateFormat });
            dictionaryMADUB1003.Add("Dropdown2", new MappingDetail() { Field = "Policy.XsLine.XsGeneralLiabilityCoverage.PerOccurrenceLimit.LimitTypeName" });
            dictionaryMADUB1003.Add("Dropdown3", new MappingDetail() { Field = "Policy.XsLine.XsGeneralLiabilityCoverage.PerOccurrenceLimit.Text", Format = FormatHelper.CurrencyFormat });

            dictionaryMADUB1003.Add("Dropdown4", new MappingDetail() { Field = "Policy.XsLine.XsGeneralLiabilityCoverage.GeneralAggregateLimit.LimitTypeName" });
            dictionaryMADUB1003.Add("Dropdown5", new MappingDetail() { Field = "Policy.XsLine.XsGeneralLiabilityCoverage.GeneralAggregateLimit.Text", Format = FormatHelper.CurrencyFormat });

            dictionaryMADUB1003.Add("Dropdown6", new MappingDetail() { Field = "Policy.XsLine.XsGeneralLiabilityCoverage.ProductOperationsLimit.LimitTypeName" });
            dictionaryMADUB1003.Add("Dropdown7", new MappingDetail() { Field = "Policy.XsLine.XsGeneralLiabilityCoverage.ProductOperationsLimit.Text", Format = FormatHelper.CurrencyFormat });

            dictionaryMADUB1003.Add("Dropdown8", new MappingDetail() { Field = "Policy.XsLine.XsGeneralLiabilityCoverage.PersonalAndAdvertisingLimit.LimitTypeName" });
            dictionaryMADUB1003.Add("Dropdown9", new MappingDetail() { Field = "Policy.XsLine.XsGeneralLiabilityCoverage.PersonalAndAdvertisingLimit.Text", Format = FormatHelper.CurrencyFormat });

            dictionaryMADUB1003.Add("BlankMergeField_6", new MappingDetail() { Constant = "Commercial Automobile Liability" });
            dictionaryMADUB1003.Add("BlankMergeField_7", new MappingDetail() { Field = "Policy.XsLine.XsAutoLiabilityCoverage.CarrierName" });

            dictionaryMADUB1003.Add("BlankMergeField_8", new MappingDetail() { Field = "Policy.XsLine.XsAutoLiabilityCoverage.EffectiveDate", Format = FormatHelper.DateFormat });
            dictionaryMADUB1003.Add("BlankMergeField_9", new MappingDetail() { Field = "Policy.XsLine.XsAutoLiabilityCoverage.ExpirationDate", Format = FormatHelper.DateFormat });

            dictionaryMADUB1003.Add("Dropdown11", new MappingDetail() { Field = "Policy.XsLine.XsAutoLiabilityCoverage.CombinedSingleLimit.LimitTypeName" });
            dictionaryMADUB1003.Add("Dropdown12", new MappingDetail() { Field = "Policy.XsLine.XsAutoLiabilityCoverage.CombinedSingleLimit.Text", Format = FormatHelper.CurrencyFormat });

            dictionaryMADUB1003.Add("BlankMergeField_12", new MappingDetail() { Constant = "Hired and Non-Owned Automobile Liability" });
            dictionaryMADUB1003.Add("BlankMergeField_13", new MappingDetail() { Field = "Policy.XsLine.XsHiredNonOwnedCoverage.CarrierName" });

            dictionaryMADUB1003.Add("BlankMergeField_14", new MappingDetail() { Field = "Policy.XsLine.XsHiredNonOwnedCoverage.EffectiveDate", Format = FormatHelper.DateFormat });
            dictionaryMADUB1003.Add("BlankMergeField_15", new MappingDetail() { Field = "Policy.XsLine.XsHiredNonOwnedCoverage.ExpirationDate", Format = FormatHelper.DateFormat });

            dictionaryMADUB1003.Add("Dropdown20", new MappingDetail() { Field = "Policy.XsLine.XsHiredNonOwnedCoverage.CombinedSingleLimit.LimitTypeName" });
            dictionaryMADUB1003.Add("Dropdown21", new MappingDetail() { Field = "Policy.XsLine.XsHiredNonOwnedCoverage.CombinedSingleLimit.Text", Format = FormatHelper.CurrencyFormat });

            dictionaryMADUB1003.Add("BlankMergeField_18", new MappingDetail() { Constant = "Employers Liability" });
            dictionaryMADUB1003.Add("BlankMergeField_19", new MappingDetail() { Field = "Policy.XsLine.XsEmployersLiabilityCoverage.CarrierName" });

            dictionaryMADUB1003.Add("BlankMergeField_20", new MappingDetail() { Field = "Policy.XsLine.XsEmployersLiabilityCoverage.EffectiveDate", Format = FormatHelper.DateFormat });
            dictionaryMADUB1003.Add("BlankMergeField_21", new MappingDetail() { Field = "Policy.XsLine.XsEmployersLiabilityCoverage.ExpirationDate", Format = FormatHelper.DateFormat });

            dictionaryMADUB1003.Add("Dropdown29", new MappingDetail() { Field = "Policy.XsLine.XsEmployersLiabilityCoverage.EachAccident.LimitTypeName" });
            dictionaryMADUB1003.Add("Dropdown30", new MappingDetail() { Field = "Policy.XsLine.XsEmployersLiabilityCoverage.EachAccident.Text", Format = FormatHelper.CurrencyFormat });

            dictionaryMADUB1003.Add("Dropdown31", new MappingDetail() { Field = "Policy.XsLine.XsEmployersLiabilityCoverage.PolicyLimitAggregate.LimitTypeName" });
            dictionaryMADUB1003.Add("Dropdown32", new MappingDetail() { Field = "Policy.XsLine.XsEmployersLiabilityCoverage.PolicyLimitAggregate.Text", Format = FormatHelper.CurrencyFormat });

            dictionaryMADUB1003.Add("Dropdown33", new MappingDetail() { Field = "Policy.XsLine.XsEmployersLiabilityCoverage.EachEmployee.LimitTypeName" });
            dictionaryMADUB1003.Add("Dropdown34", new MappingDetail() { Field = "Policy.XsLine.XsEmployersLiabilityCoverage.EachEmployee.Text", Format = FormatHelper.CurrencyFormat });

            if (this.ShouldEmployeeBenefitLimitBeDisplayed(policy))
            {
                dictionaryMADUB1003.Add("BlankMergeField_24", new MappingDetail() { Constant = "Employee Benefits Liability" });
                dictionaryMADUB1003.Add("BlankMergeField_25", new MappingDetail() { Field = "Policy.XsLine.XsGeneralLiabilityCoverage.CarrierName" });

                dictionaryMADUB1003.Add("BlankMergeField_26", new MappingDetail() { Field = "Policy.XsLine.XsGeneralLiabilityCoverage.EffectiveDate", Format = FormatHelper.DateFormat });
                dictionaryMADUB1003.Add("BlankMergeField_27", new MappingDetail() { Field = "Policy.XsLine.XsGeneralLiabilityCoverage.ExpirationDate", Format = FormatHelper.DateFormat });

                dictionaryMADUB1003.Add("Dropdown38", new MappingDetail() { Constant = "Each Employee" });
                dictionaryMADUB1003.Add("Dropdown40", new MappingDetail() { Constant = "Aggregate" });

                if (policy.XsLine.XsGeneralLiabilityCoverage.EmployeeBenefitsLimit.Code.Equals("Other"))
                {
                    dictionaryMADUB1003.Add("Dropdown39", new MappingDetail() { Constant = "Other" });
                    dictionaryMADUB1003.Add("Dropdown41", new MappingDetail() { Constant = "Other" });
                }
                else if (policy.XsLine.XsGeneralLiabilityCoverage.EmployeeBenefitsLimit.Code.Equals("1000000"))
                {
                    dictionaryMADUB1003.Add("Dropdown39", new MappingDetail() { Constant = "$1,000,000" });
                    dictionaryMADUB1003.Add("Dropdown41", new MappingDetail() { Constant = "$1,000,000" });
                }
            }

            return dictionaryMADUB1003;
        }

        /// <summary>
        /// Returns formatted address for the passed in Questions
        /// </summary>
        /// <param name="street">Question for Street</param>
        /// <param name="city">Question for City</param>
        /// <param name="state">Question for State</param>
        /// <param name="zipCode">Question for ZipCode</param>
        /// <returns>Formatted Address</returns>
        private string GetFormattedAddressFromQA(Question street, Question city, Question state, Question zipCode)
        {
            var address = string.Empty;
            address += street != null && !string.IsNullOrWhiteSpace(street.AnswerValue) ? street.AnswerValue + ", " : string.Empty;
            address += city != null && !string.IsNullOrWhiteSpace(city.AnswerValue) ? city.AnswerValue + ", " : string.Empty;
            address += state != null && !string.IsNullOrWhiteSpace(state.AnswerValue) ? state.AnswerValue + " " : string.Empty;
            address += zipCode != null && !string.IsNullOrWhiteSpace(zipCode.AnswerValue) ? zipCode.AnswerValue : string.Empty;
            return address;
        }

        /// <summary>
        /// Gets the applicable premium for the line of business based on agent and underwriter override
        /// </summary>
        /// <param name="policy">Passed policy</param>
        /// <param name="lob">Line of Business</param>
        /// <param name="currencyFormat">Premium Currency Format</param>
        /// <returns>Applicable premium text</returns>
        private string GetApplicablePremium(Policy policy, LineOfBusiness lob, string currencyFormat = FormatHelper.CurrencyFormatWithDecimal)
        {
            decimal applicablePremium = default(decimal);
            var line = policy.GetLine(lob);
            if (line != null)
            {
                if (line.AgentAdjustedPremium != null && line.AgentAdjustedPremium.Value != 0)
                {
                    applicablePremium = line.AgentAdjustedPremium.Value;
                }
                else if (line.UnderwriterAdjustedPremium != null && line.UnderwriterAdjustedPremium.Value != 0)
                {
                    applicablePremium = line.UnderwriterAdjustedPremium.Value;
                }
                else
                {
                    applicablePremium = Math.Round(line.RollUp.Premium, 0, MidpointRounding.AwayFromZero);
                }
            }

            return applicablePremium != default(decimal) ? FormatHelper.FormatValue(applicablePremium, currencyFormat) : "Not Covered";
        }

        /// <summary>
        /// Fill bookmark field value
        /// </summary>
        /// <param name="item">Risk Unit Item</param>
        /// <param name="builder">DocumentBuilder</param>
        private void FillBlankMergeFieldMDCP1000DescriptionOfPremises(CfLine item, DocumentBuilder builder)
        {
            builder.MoveToBookmark("BlankMergeField_1");

            builder.Font.Size = 12;
            builder.Font.Name = "Arial";
            builder.Font.Bold = true;

            Aspose.Words.Tables.Table table = builder.StartTable();

            this.AddDescriptionOfPremisesRowHeader(builder);

            if (item.RiskUnits.Any())
            {
                var list = item.RiskUnits.OrderBy(x => x.PremisesNumber).ThenBy(y => y.BuildingNumber).ToList();
                if (list.Any())
                {
                    foreach (var riskUnit in list)
                    {
                        this.AddDescriptionOfPremisesRows(riskUnit, builder);
                    }
                }
            }

            table.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(550.00);
            table.StyleOptions = Aspose.Words.Tables.TableStyleOptions.Default2003;
            table.LeftPadding = 3.05;
            table.RightPadding = 3.05;
            table.LeftIndent = 3.5;
            builder.EndTable();
        }

        /// <summary>
        /// Fill bookmark field value
        /// </summary>
        /// <param name="policy">Policy</param>
        /// <param name="builder">DocumentBuilder</param>
        private void FillBlankMergeFieldMDCP1000CoverageProvided(Policy policy, DocumentBuilder builder)
        {
            var term = string.Empty;
            builder.MoveToBookmark("BlankMergeField_2");

            builder.Font.Size = 12;
            builder.Font.Name = "Arial";
            builder.Font.Bold = true;

            Aspose.Words.Tables.Table table = builder.StartTable();

            this.AddCoverageProvidedRowHeader(builder, true);
            this.AddCoverageColumnHeader(builder);
            term = this.CalculateTermCode(policy);

            if (policy.CfLine != null && policy.CfLine.RiskUnits.Any())
            {
                var list = policy.CfLine.RiskUnits.OrderBy(x => x.PremisesNumber).ThenBy(y => y.BuildingNumber).ToList();
                if (list.Any())
                {
                    foreach (var riskUnit in list)
                    {
                        this.AddCoverageRows(riskUnit, builder, term);
                    }
                }

                this.AddCoverageProvidedRowHeader(builder);
            }

            table.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(550.00);
            table.StyleOptions = Aspose.Words.Tables.TableStyleOptions.Default2003;
            table.LeftPadding = 3.05;
            table.RightPadding = 3.05;
            table.LeftIndent = 3.5;
            builder.EndTable();
        }

        /// <summary>
        /// Fill bookmark field value
        /// </summary>
        /// <param name="item">Risk Unit Item</param>
        /// <param name="builder">DocumentBuilder</param>
        private void FillBlankMergeFieldMDCP1000OptionalCoverages(CfLine item, DocumentBuilder builder)
        {
            builder.MoveToBookmark("BlankMergeField_3");

            builder.Font.Size = 12;
            builder.Font.Bold = true;

            Aspose.Words.Tables.Table table = builder.StartTable();

            this.AddCoverageProvidedRowHeader(builder, false);
            this.AddCoverageColumnHeader(builder);
            this.AddOptionalCoveragesRows(item, builder);
            this.AddCoverageProvidedRowHeader(builder);

            table.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(550.00);
            table.StyleOptions = Aspose.Words.Tables.TableStyleOptions.Default2003;
            table.LeftPadding = 3.05;
            table.RightPadding = 3.05;
            table.LeftIndent = 3.5;
            builder.EndTable();
        }

        /// <summary>
        /// Fill bookmark field value
        /// </summary>
        /// <param name="item">Risk Unit Item</param>
        /// <param name="builder">DocumentBuilder</param>
        private void FillBlankMergeFieldMDCP1000MortgageHolder(CfLine item, DocumentBuilder builder)
        {
            builder.MoveToBookmark("BlankMergeField_4");

            builder.Font.Size = 12;
            builder.Font.Bold = true;

            Aspose.Words.Tables.Table table = builder.StartTable();

            this.AddMortgageHolderTableHeaderRows(builder);

            if (item != null)
            {
                this.AddMortgageHolderRows(item, builder);
            }

            table.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(544.5);
            table.StyleOptions = Aspose.Words.Tables.TableStyleOptions.Default2003;
            table.LeftPadding = 3.05;
            table.RightPadding = 3.05;
            table.LeftIndent = 3.5;
            builder.EndTable();
        }

        /// <summary>
        /// Fill bookmark field value
        /// </summary>
        /// <param name="policy">Policy</param>
        /// <param name="builder">DocumentBuilder</param>
        private void FillBlankMergeFieldMDCP1000TotalPremium(Policy policy, DocumentBuilder builder)
        {
            CfLine item = policy.CfLine;
            builder.MoveToBookmark("BlankMergeField_5");

            builder.Font.Size = 12;
            builder.Font.Bold = true;

            Aspose.Words.Tables.Table table = builder.StartTable();

            builder.Font.Bold = true;
            builder.CellFormat.TopPadding = 4;
            builder.CellFormat.BottomPadding = 4;
            builder.CellFormat.LeftPadding = 8;
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.First;
            builder.Write("FORMS AND ENDORSEMENTS");
            builder.Font.Bold = false;
            builder.Write(": SEE FORMS SCHEDULE – MDIL 1001");
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.EndRow();
            builder.CellFormat.ClearFormatting();

            if (item != null && item.RollUp != null)
            {
                builder.Font.Bold = true;
                builder.CellFormat.TopPadding = 4;
                builder.CellFormat.BottomPadding = 4;
                builder.CellFormat.LeftPadding = 8;

                builder.InsertCell();
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.First;
                builder.Write("TOTAL PREMIUM FOR THIS COVERAGE PART:        ");
                builder.Font.Bold = false;
                builder.Write(string.Format("${0}{1}", this.GetApplicablePremium(policy, LineOfBusiness.CF, FormatHelper.CurrencyFormat), policy.CfLine.RollUp.IsMinimumPremium ? " MP" : string.Empty));

                builder.InsertCell();
                builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;

                builder.InsertCell();
                builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
                builder.EndRow();
            }

            table.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(544.5);
            table.StyleOptions = Aspose.Words.Tables.TableStyleOptions.Default2003;
            table.LeftPadding = 3.05;
            table.RightPadding = 3.05;
            table.LeftIndent = 3.5;
            builder.EndTable();
        }

        /// <summary>
        /// Fill bookmark field value
        /// </summary>
        /// <param name="builder">DocumentBuilder</param>
        private void AddDescriptionOfPremisesRowHeader(DocumentBuilder builder)
        {
            builder.CellFormat.TopPadding = 4;
            builder.CellFormat.BottomPadding = 4;
            builder.CellFormat.LeftPadding = 8;
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.First;
            builder.Write("DESCRIPTION OF PREMISES");
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.EndRow();
            builder.CellFormat.ClearFormatting();

            builder.Font.Bold = true;
            builder.Font.Size = 10;
            builder.CellFormat.TopPadding = 4;
            builder.CellFormat.BottomPadding = 4;
            builder.CellFormat.LeftPadding = 8;
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(50.00);
            builder.Write("Prem.\nNo.");

            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(50.00);
            builder.Write("Bldg.\nNo.");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(150.00);
            builder.Write("Location Address");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(50.00);
            builder.Write("No. of\nStories");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(50.00);
            builder.Write("Year\nBuilt");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(150.00);
            builder.Write("Occupancy");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(50.00);
            builder.Write("Construction");
            builder.EndRow();
            builder.CellFormat.ClearFormatting();
        }

        /// <summary>
        /// Fill bookmark field value
        /// </summary>
        /// <param name="item">CfRiskUnit</param>
        /// <param name="builder">DocumentBuilder</param>
        private void AddDescriptionOfPremisesRows(CfRiskUnit item, DocumentBuilder builder)
        {
            builder.Font.Bold = false;
            builder.Font.Size = 10;
            builder.CellFormat.TopPadding = 4;
            builder.CellFormat.BottomPadding = 4;
            builder.CellFormat.LeftPadding = 8;

            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(50.00);
            builder.Write(item.PremisesNumber.ToString());

            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(50.00);
            builder.Write(item.BuildingNumber.ToString());

            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(150.00);
            var address = string.Empty;
            address += !string.IsNullOrWhiteSpace(item.StreetAddress.Line1) ? item.StreetAddress.Line1 + ", " : string.Empty;
            address += !string.IsNullOrWhiteSpace(item.StreetAddress.Line2) ? item.StreetAddress.Line2 + ", " : string.Empty;
            address += !string.IsNullOrWhiteSpace(item.StreetAddress.City) ? item.StreetAddress.City + ", " : string.Empty;
            address += !string.IsNullOrWhiteSpace(item.StreetAddress.StateCode) ? item.StreetAddress.StateCode + " " : string.Empty;
            address += !string.IsNullOrWhiteSpace(item.StreetAddress.ZipCode) ? item.StreetAddress.ZipCode : string.Empty;
            builder.Write(address);

            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(50.00);
            builder.Write(!string.IsNullOrEmpty(item.NumberStories.ToString()) ? item.NumberStories.ToString() : string.Empty);

            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(50.00);
            builder.Write(item.YearBuilt.ToString());

            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(150.00);
            if (!string.IsNullOrEmpty(item.Occupancy))
            {
                builder.Write(item.Occupancy);
            }
            else
            {
                builder.Write(string.Empty);
            }

            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(50.00);
            var constructionType = string.Empty;
            if (item.ConstructionTypes.Any())
            {
                var selectedConstructionType = item.ConstructionTypes.Where(x => x.Code == item.ConstructionType).First();
                if (selectedConstructionType != null)
                {
                    constructionType = selectedConstructionType.Description;
                }
            }

            builder.Write(constructionType);
            builder.EndRow();
            builder.CellFormat.ClearFormatting();

            builder.Font.Size = 10;
            builder.CellFormat.TopPadding = 4;
            builder.CellFormat.BottomPadding = 4;
            builder.CellFormat.LeftPadding = 8;

            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.First;
            builder.Font.Bold = true;
            builder.Write("Class Code: ");
            builder.Font.Bold = false;
            builder.Write(item.ClassCode);
            builder.Font.Bold = true;
            builder.Write("     Class Description: ");
            builder.Font.Bold = false;
            builder.Write(item.ClassDescription);

            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;

            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;

            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;

            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;

            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;

            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.EndRow();
            builder.CellFormat.ClearFormatting();
        }

        /// <summary>
        /// Fill bookmark field value
        /// </summary>
        /// <param name="builder">DocumentBuilder</param>
        /// <param name="isCoverageProvideTable">Coverage Provider table</param>
        private void AddCoverageProvidedRowHeader(DocumentBuilder builder, bool? isCoverageProvideTable = null)
        {
            builder.CellFormat.TopPadding = 4;
            builder.CellFormat.BottomPadding = 4;
            builder.CellFormat.LeftPadding = 8;
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.First;

            if (isCoverageProvideTable == null)
            {
                builder.Font.Size = 10;
                builder.Font.Bold = false;
                builder.Write("*AA-Agreed Amount	    *ACV-Actual Cash Value	    **If Extra Expense Coverage, Limits On Loss Payment     *RC-Replacement Cost");
            }
            else if (isCoverageProvideTable.Value)
            {
                builder.Font.Size = 12;
                builder.Font.Bold = true;
                builder.Write("COVERAGES PROVIDED");
                builder.Font.Bold = false;
                builder.Font.Size = 10;
                builder.Write("– Insurance at the described premises applies only for coverages for which a limit of insurance is shown.");
            }
            else if (!isCoverageProvideTable.Value)
            {
                builder.Font.Size = 12;
                builder.Font.Bold = true;
                builder.Write("OPTIONAL COVERAGES");
                builder.Font.Size = 10;
                builder.Font.Bold = false;
                builder.Write(" – Applicable only when entries are made in the schedule below.");
            }

            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.EndRow();
            builder.CellFormat.ClearFormatting();
        }

        /// <summary>
        /// Fill bookmark field value
        /// </summary>
        /// <param name="builder">DocumentBuilder</param>
        private void AddCoverageColumnHeader(DocumentBuilder builder)
        {
            builder.Font.Size = 10;
            builder.Font.Bold = true;
            builder.CellFormat.TopPadding = 4;
            builder.CellFormat.BottomPadding = 4;
            builder.CellFormat.LeftPadding = 8;
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(40.00);
            builder.Write("Prem.\nNo.");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(40.00);
            builder.Write("Bldg.\nNo.");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(150.00);
            builder.Write("Coverages");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(60.00);
            builder.Write("Limit of\nInsurance");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(60.00);
            builder.Write("Covered\nCauses Of\nLoss");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(60.00);
            builder.Write("Valuation*");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(60.00);
            builder.Write("Coinsurance**");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(40.00);
            builder.Write("Rates");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(40.00);
            builder.Write("Rate\nTerm");
            builder.EndRow();
            builder.CellFormat.ClearFormatting();
        }

        /// <summary>
        /// Fill bookmark field value
        /// </summary>
        /// <param name="item">CfRiskUnit</param>
        /// <param name="builder">DocumentBuilder</param>
        /// <param name="term">Policy term</param>
        private void AddCoverageRows(CfRiskUnit item, DocumentBuilder builder, string term)
        {
            if (item.Coverages.Any())
            {
                builder.Font.Size = 10;
                builder.Font.Bold = false;
                var causeofLoss = string.Empty;
                foreach (var coverage in item.Coverages)
                {
                    builder.CellFormat.TopPadding = 4;
                    builder.CellFormat.BottomPadding = 4;
                    builder.CellFormat.LeftPadding = 8;

                    builder.InsertCell();
                    builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                    builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(40.00);
                    builder.Write(item.PremisesNumber.ToString());

                    builder.InsertCell();
                    builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                    builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(40.00);
                    builder.Write(item.BuildingNumber.ToString());

                    builder.InsertCell();
                    builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                    builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(150.00);
                    builder.Write(coverage.Description);

                    builder.InsertCell();
                    builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                    builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(60.00);
                    builder.Write(string.Format("${0:N0}", coverage.Limit));

                    builder.InsertCell();
                    builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                    builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(60.00);
                    this.GetCauseOfLose(item, out causeofLoss);
                    builder.Write(causeofLoss);

                    builder.InsertCell();
                    builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                    builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(60.00);
                    if (coverage.Valuation.Equals("AC"))
                    {
                        builder.Write("ACV");
                    }
                    else
                    {
                        builder.Write(coverage.Valuation);
                    }

                    int outputValue = 0;
                    builder.InsertCell();
                    builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                    builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(60.00);
                    bool successfullyParsed = int.TryParse(coverage.Coinsurance, out outputValue);
                    if (successfullyParsed)
                    {
                        builder.Write(string.Format("{0}%", coverage.Coinsurance));
                    }
                    else
                    {
                        if (coverage.Coinsurance.Contains("-"))
                        {
                            builder.Write(coverage.Coinsurance);
                        }
                        else
                        {
                            builder.Write(string.Format("{0} mo. limit", coverage.Coinsurance));
                        }
                    }

                    builder.InsertCell();
                    builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                    builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(40.00);
                    if (coverage.UnderwriterAdjustedRate > 0)
                    {
                        builder.Write(string.Format("{0:n3}", coverage.UnderwriterAdjustedRate));
                    }
                    else if (coverage.AgentAdjustedRate > 0)
                    {
                        builder.Write(string.Format("{0:n3}", coverage.AgentAdjustedRate));
                    }
                    else
                    {
                        builder.Write(string.Format("{0:n3}", coverage.RateAmount));
                    }

                    builder.InsertCell();
                    builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                    builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(40.00);
                    builder.Write(term);
                    builder.EndRow();
                }
            }
        }

        /// <summary>
        /// Fill bookmark field value
        /// </summary>
        /// <param name="item">CfRiskUnit</param>
        /// <param name="val">Get Causes of Loss value</param>
        private void GetCauseOfLose(CfRiskUnit item, out string val)
        {
            val = string.Empty;
            if (item.CausesOfLoss.Any())
            {
                var causeOfLose = item.CausesOfLoss.Where(x => x.Code == item.CauseOfLoss).First();
                if (causeOfLose != null)
                {
                    switch (causeOfLose.Code)
                    {
                        case "BA":
                            if (!item.IsWindHailExcluded)
                            {
                                val = "Basic";
                            }
                            else if (item.IsWindHailExcluded)
                            {
                                val = "Basic x-wind";
                            }

                            break;
                        case "NT":
                            if (!item.IsWindHailExcluded)
                            {
                                val = "Special x-theft";
                            }
                            else if (item.IsWindHailExcluded)
                            {
                                val = "Special x-wind and x-theft";
                            }

                            break;
                        case "BR":
                            if (!item.IsWindHailExcluded)
                            {
                                val = "Broad";
                            }
                            else if (item.IsWindHailExcluded)
                            {
                                val = "Broad x-wind";
                            }

                            break;
                        case "SP":
                            if (!item.IsWindHailExcluded)
                            {
                                val = "Special";
                            }
                            else if (item.IsWindHailExcluded)
                            {
                                val = "Special x-wind";
                            }

                            break;
                        default:
                            break;
                    }
                }
            }
        }

        /// <summary>
        /// Fill bookmark field value
        /// </summary>
        /// <param name="item">CfRiskUnit</param>
        /// <param name="builder">DocumentBuilder</param>
        private void AddOptionalCoveragesRows(CfLine item, DocumentBuilder builder)
        {
            if (item.RollUp != null && item.RollUp.OptionalCoverages.Any())
            {
                foreach (var coverage in item.RollUp.OptionalCoverages.Where(c => c.Code != "TRIA" && c.IsSelected == true))
                {
                    builder.Font.Bold = false;
                    builder.CellFormat.TopPadding = 4;
                    builder.CellFormat.BottomPadding = 4;
                    builder.CellFormat.LeftPadding = 8;

                    builder.InsertCell();
                    builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                    builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(40.00);
                    builder.Write("All");

                    builder.InsertCell();
                    builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                    builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(40.00);
                    builder.Write("All");

                    builder.InsertCell();
                    builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                    builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(150.00);
                    builder.Write(string.Format("{0}", coverage.ShortName));

                    builder.InsertCell();
                    builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                    builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(60.00);
                    builder.Write(string.Empty);

                    builder.InsertCell();
                    builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                    builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(60.00);
                    string[] causeOfLoss = coverage.ShortName.Split('-');
                    if (causeOfLoss != null && causeOfLoss.Any())
                    {
                        if (!string.IsNullOrEmpty(causeOfLoss[1].ToString()))
                        {
                            builder.Write(string.Format("See {0}", causeOfLoss[1].ToString()));
                        }
                    }

                    builder.InsertCell();
                    builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                    builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(60.00);
                    builder.Write(string.Empty);

                    builder.InsertCell();
                    builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                    builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(60.00);
                    builder.Write(string.Empty);

                    builder.InsertCell();
                    builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                    builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(40.00);
                    builder.Write(string.Empty);

                    builder.InsertCell();
                    builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                    builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(40.00);
                    builder.Write(string.Empty);
                    builder.EndRow();
                }
            }
        }

        /// <summary>
        /// Fill bookmark field value
        /// </summary>
        /// <param name="builder">DocumentBuilder</param>
        private void AddMortgageHolderTableHeaderRows(DocumentBuilder builder)
        {
            builder.Font.Bold = true;
            builder.Font.Size = 12;
            builder.CellFormat.TopPadding = 4;
            builder.CellFormat.BottomPadding = 4;
            builder.CellFormat.LeftPadding = 8;
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.First;
            builder.Write("MORTGAGEHOLDERS");
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.EndRow();
            builder.CellFormat.ClearFormatting();

            builder.Font.Bold = true;
            builder.Font.Size = 10;
            builder.CellFormat.TopPadding = 4;
            builder.CellFormat.BottomPadding = 4;
            builder.CellFormat.LeftPadding = 8;
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(50.00);
            builder.Write("Prem.\nNo.");

            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(50.00);
            builder.Write("Bldg.\nNo.");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(444.00);
            builder.Write("Mortgageholder Name And Mailing Address");
            builder.EndRow();
            builder.CellFormat.ClearFormatting();
        }

        /// <summary>
        /// Fill bookmark field value
        /// </summary>
        /// <param name="item">CfRiskUnit</param>
        /// <param name="builder">DocumentBuilder</param>
        private void AddMortgageHolderRows(CfLine item, DocumentBuilder builder)
        {
            if (item.RollUp.Documents != null && item.RollUp.Documents.Any())
            {
                var documentQuestions = item.RollUp.Documents.Where(x => x.NormalizedNumber == "MDCP1000")?.FirstOrDefault().Questions.GroupBy(p => p.MultipleRowGroupingNumber);
                foreach (var questionGroup in documentQuestions)
                {
                    builder.Font.Bold = false;
                    builder.Font.Size = 10;
                    builder.CellFormat.TopPadding = 4;
                    builder.CellFormat.BottomPadding = 4;
                    builder.CellFormat.LeftPadding = 8;

                    builder.InsertCell();
                    builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                    builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(50.00);
                    var premNo = questionGroup.Where(x => x.Code == "PremNo")?.First().AnswerValue;
                    builder.Write(!string.IsNullOrEmpty(premNo) ? premNo : string.Empty);

                    builder.InsertCell();
                    builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                    builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(50.00);
                    var bldgNo = questionGroup.Where(x => x.Code == "BldgNo")?.First().AnswerValue;
                    builder.Write(!string.IsNullOrEmpty(bldgNo) ? bldgNo : string.Empty);

                    builder.InsertCell();
                    builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                    builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(444.00);
                    var mortgageholder = questionGroup.Where(x => x.Code == "Mortgageholder")?.First().AnswerValue;
                    builder.Write(!string.IsNullOrEmpty(mortgageholder) ? mortgageholder : string.Empty);

                    builder.EndRow();
                }

                builder.Font.Bold = true;
                builder.Font.Size = 12;
                builder.CellFormat.TopPadding = 4;
                builder.CellFormat.BottomPadding = 4;
                builder.CellFormat.LeftPadding = 8;
                builder.InsertCell();
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(50.00);
                builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.First;
                builder.CellFormat.VerticalMerge = Aspose.Words.Tables.CellMerge.First;
                builder.Write("DEDUCTIBLE");
                builder.InsertCell();
                builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(50.00);
                builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
                builder.CellFormat.VerticalMerge = Aspose.Words.Tables.CellMerge.First;
                builder.InsertCell();
                builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(444.00);
                builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
                builder.CellFormat.VerticalMerge = Aspose.Words.Tables.CellMerge.First;
                builder.EndRow();
                builder.CellFormat.ClearFormatting();

                builder.Font.Bold = false;
                builder.Font.Size = 10;
                builder.CellFormat.TopPadding = 4;
                builder.CellFormat.BottomPadding = 4;
                builder.CellFormat.LeftPadding = 8;

                this.FillMortagageDeductiblesForMDCP1000(item, builder);
                builder.CellFormat.ClearFormatting();
            }
        }

        /// <summary>
        /// Creates and returns a collection of Vehicle schedule from question collection
        /// </summary>
        /// <param name="questions">Collection of questions</param>
        /// <returns>Collection of VehicleSchedule</returns>
        private List<IMMTCVehicleSchedule> GetVehicleSchedule(List<Question> questions)
        {
            var result = new List<IMMTCVehicleSchedule>();

            var multipleGroupingNumbers = questions.Select(x => x.MultipleRowGroupingNumber).Distinct().ToList();
            multipleGroupingNumbers.Sort();

            var yearQuestions = questions.Where(x => x.Code.Equals("VehicleYear")).Select(x => x).OrderBy(x => x.MultipleRowGroupingNumber).ToList();
            var manufacturerQuestions = questions.Where(x => x.Code.Equals("VehicleMfg")).Select(x => x).OrderBy(x => x.MultipleRowGroupingNumber).ToList();
            var identificationNumberQuestions = questions.Where(x => x.Code.Equals("VehicleID")).Select(x => x).OrderBy(x => x.MultipleRowGroupingNumber).ToList();
            var bodyTypeQuestions = questions.Where(x => x.Code.Equals("VehicleBodyType")).Select(x => x).OrderBy(x => x.MultipleRowGroupingNumber).ToList();
            var vehicleBodyTypeOthers = questions.Where(x => x.Code.Equals("VehicleBodyTypeOther")).Select(x => x).OrderBy(x => x.MultipleRowGroupingNumber).ToList();
            var vehicleMfgOthers = questions.Where(x => x.Code.Equals("VehicleMfgOther")).Select(x => x).OrderBy(x => x.MultipleRowGroupingNumber).ToList();

            for (var i = 0; i < yearQuestions.Count(); i++)
            {
                var groupingNumber = yearQuestions[i].MultipleRowGroupingNumber;
                var year = yearQuestions[i];
                var manufacturer = manufacturerQuestions.FirstOrDefault(x => x.MultipleRowGroupingNumber == groupingNumber);
                var identificationNumber = identificationNumberQuestions.FirstOrDefault(x => x.MultipleRowGroupingNumber == groupingNumber);
                var bodyType = bodyTypeQuestions.FirstOrDefault(x => x.MultipleRowGroupingNumber == groupingNumber);
                var vehicleBodyTypeOther = vehicleBodyTypeOthers.FirstOrDefault(x => x.MultipleRowGroupingNumber == groupingNumber);
                var vehicleMfgOther = vehicleMfgOthers.FirstOrDefault(x => x.MultipleRowGroupingNumber == groupingNumber);

                var schedule = new IMMTCVehicleSchedule();

                schedule.Year = this.GetQuestionValue(year);
                schedule.Manufacturer = manufacturer.AnswerValue == "995" ? vehicleMfgOther.AnswerValue : this.GetQuestionValue(manufacturer);
                schedule.IdentificationNumber = this.GetQuestionValue(identificationNumber);
                schedule.BodyType = bodyType.AnswerValue == "1015" ? vehicleBodyTypeOther.AnswerValue : this.GetQuestionValue(bodyType);
                schedule.MultipleRowGroupingNumber = groupingNumber;

                result.Add(schedule);
            }

            return result;
        }

        /// <summary>
        /// Returns the value of a question based on DisplayFormat
        /// </summary>
        /// <param name="question">Question</param>
        /// <returns>Answer</returns>
        private string GetQuestionValue(Question question)
        {
            var result = string.Empty;

            if (question != null)
            {
                switch (question.DisplayFormat)
                {
                    case DisplayFormat.TextBox:
                        result = question.AnswerValue;
                        break;
                    case DisplayFormat.DropDown:
                        var selectedOption = question.Answers.FirstOrDefault(x => x.Value.Equals(question.AnswerValue));
                        if (selectedOption != null)
                        {
                            result = selectedOption.Verbiage;
                        }

                        break;
                    default:
                        result = question.AnswerValue;
                        break;
                }
            }

            return result;
        }

        /// <summary>
        /// Fill bookmark field value
        /// </summary>
        /// <param name="policy">Policy</param>
        /// <param name="builder">DocumentBuilder</param>
        /// <param name="documentManager">Document Manager</param>
        /// <param name="normalizeNumber">Form normalize number</param>
        private void FillBlankMergeFieldMECP1281MECP1321(Policy policy, DocumentBuilder builder, ref IPolicyDocumentManager documentManager, string normalizeNumber)
        {
            if (policy != null && policy.CfLine != null && policy.CfLine.RiskUnits.Any())
            {
                List<string> blankMergeFieldList = new List<string> { string.Empty, "2", "4" };

                builder.MoveToBookmark("BlankMergeField");
                var table = builder.CurrentNode.GetParentTable();
                if (table == null)
                {
                    // this is an error...create the table manually?
                    throw new Exception();
                }

                var documentQuestions = policy.CfLine.RollUp.Documents.Where(x => x.NormalizedNumber == normalizeNumber)?.FirstOrDefault().Questions.Where(y => y.MultipleRowGroupingNumber > 0).GroupBy(p => p.MultipleRowGroupingNumber);
                var rowNumber = 1;
                foreach (var questionGroup in documentQuestions)
                {
                    Row workingRow;
                    ////if (rowNumber <= blankMergeFieldList[0].Count)
                    if (rowNumber <= blankMergeFieldList.Count)
                    {
                        ////var number = blankMergeFieldList[0][rowNumber - 1];
                        var number = blankMergeFieldList[rowNumber - 1];
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

                    var premNo = questionGroup.Where(x => x.Code == "PremisesNo")?.First().AnswerValue;
                    var buildingNo = questionGroup.Where(x => x.Code == "BuildingNo")?.First().AnswerValue;

                    this.ReplaceCellParagraphOrBookmark(ref documentManager, workingRow.Cells[0], !string.IsNullOrEmpty(premNo) ? premNo : string.Empty);
                    this.ReplaceCellParagraphOrBookmark(ref documentManager, workingRow.Cells[1], !string.IsNullOrEmpty(buildingNo) ? buildingNo : string.Empty);
                    rowNumber++;
                }
            }
        }

        /// <summary>
        /// Get term code for policy
        /// </summary>
        /// <param name="policy">policy</param>
        /// <returns>string</returns>
        private string CalculateTermCode(Policy policy)
        {
            var termCode = string.Empty;
            if (policy != null)
            {
                var expirationDate = policy.ExpirationDate.Value;
                if (policy.EffectiveDate.AddYears(1).Date.Equals(expirationDate.Date))
                {
                    termCode = "an";
                }
                else
                {
                    termCode = "term";
                }
            }

            return termCode;
        }

        /// <summary>
        /// Fill bookmark field value
        /// </summary>
        /// <param name="item">CfRiskUnit</param>
        /// <param name="builder">DocumentBuilder</param>
        private void FillMortagageDeductiblesForMDCP1000(CfLine item, DocumentBuilder builder)
        {
            List<string> formNormalizeNumbers = new List<string> { "CP0320", "CP0321", "CP0325", "CP1044", "CP1121", "MECP1222", "MECP1223", "MECP1231", "MECP1244", "MECP1285" };
            if (item != null && item.RiskUnits.Any())
            {
                var documents = item.RollUp?.Documents?.Where(d => formNormalizeNumbers.Contains(d.NormalizedNumber) && d.ReturnIsSelected());
                var riskUnits = item.RiskUnits.GroupBy(x => x.Deductible).Count();

                var isChecked = documents.Any() ? true : false;
                var normalizeNumber = documents.Any() ? ("See " + string.Join(", ", documents.Select(p => p.DisplayNumber.ToString()))) : string.Empty;
                if (riskUnits == 1)
                {
                    if ((item.RiskUnits.Where(x => x.CauseOfLoss != "SP" && x.IsWindHailExcluded).Count() == item.RiskUnits.Count()) || (item.RiskUnits.Where(x => x.CauseOfLoss == "SP" && x.WindHailDeductible == "NONE" && x.TheftDeductible == "NONE").Count() == item.RiskUnits.Count()))
                    {
                        this.MapDeductibleValuesInBlankMergeField(builder, isChecked, normalizeNumber, item.RiskUnits[0].Deductible);
                    }
                    else if (item.RiskUnits.Where(x => x.CauseOfLoss == "SP" && (x.WindHailDeductible != "NONE" || x.TheftDeductible != "NONE")).Count() == item.RiskUnits.Count() && !string.IsNullOrEmpty(normalizeNumber))
                    {
                        this.MapDeductibleValuesInBlankMergeField(builder, isChecked, normalizeNumber, item.RiskUnits.Count == 1 ? item.RiskUnits[0].Deductible : (int?)null);
                    }
                    else
                    {
                        this.MapDeductibleValuesInBlankMergeField(builder, isChecked, normalizeNumber, item.RiskUnits.Count == 1 ? item.RiskUnits[0].Deductible : (int?)null);
                    }
                }
                else if (riskUnits > 1)
                {
                    this.MapDeductibleValuesInBlankMergeField(builder, isChecked, normalizeNumber);
                }
            }
        }

        /// <summary>
        /// Fill bookmark field value
        /// </summary>
        /// <param name="builder">DocumentBuilder</param>
        /// <param name="exceptionBoolVal">Exception checkbox bool value</param>
        /// <param name="commaSeparatedNormalizeNo">comma Separated Normalize number</param>
        /// <param name="deductible">Deductible value</param>
        private void MapDeductibleValuesInBlankMergeField(DocumentBuilder builder, bool exceptionBoolVal, string commaSeparatedNormalizeNo, int? deductible = null)
        {
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(50.00);
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.First;
            builder.Write((deductible.HasValue ? string.Format("${0:n0}", deductible.Value) : "$").PadRight(25));

            builder.InsertCheckBox("check1", true, 0);
            builder.Write("  Per occurrence  ");

            builder.InsertCheckBox("check2", false, 0);
            builder.Write("Per Location  ");

            builder.InsertCheckBox("check3", false, 0);
            builder.Write("Per Building  ");

            builder.InsertCheckBox("check4", exceptionBoolVal, 0);
            builder.Write("Exception:  ");
            builder.Write(commaSeparatedNormalizeNo);

            builder.InsertCell();
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(50.00);
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;

            builder.InsertCell();
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(444.00);
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;

            builder.EndRow();
        }
    }
}