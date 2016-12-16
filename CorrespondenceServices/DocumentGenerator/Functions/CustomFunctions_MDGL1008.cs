// <copyright file="CustomFunctions_MDGL1008.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>

namespace Mkl.WebTeam.DocumentGenerator.Functions
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using Aspose.Words;
    using DecisionModel.Models.Policy;
    using DecisionModel.Representations;
    using Helpers;
    using SubmissionShared.Enumerations;

    /// <summary>
    /// Partial class containing custom function for MDGL1008
    /// </summary>
    public partial class CustomFunctions
    {
        /// <summary>
        /// Generates MDGL 1008 form
        /// </summary>
        /// <param name="policy">Policy</param>
        /// <param name="documentManager">Document manager</param>
        /// <param name="extraParameters">Extra parameters</param>
        public void GenerateMDGL1008Form(Policy policy, ref IPolicyDocumentManager documentManager, Dictionary<string, object> extraParameters)
        {
            if (policy.GlLine != null && policy.GlLine.Limits != null)
            {
                this.FillLimitsOfInsuranceInMDGL1008Form(policy.GlLine.Limits, ref documentManager);
                DocumentBuilder builder = new DocumentBuilder(documentManager.Document);
                this.FillLocationInBlankMergeFieldValueForMDGL1008Form(policy, builder);
                this.MDGL1008Classification(policy, builder);
            }
            else if (policy.SpecialEventLine != null && policy.SpecialEventLine.Limits != null)
            {
                this.FillLimitsOfInsuranceInMDGL1008Form(policy.SpecialEventLine.Limits, ref documentManager);
                DocumentBuilder builder = new DocumentBuilder(documentManager.Document);
                this.FillLocationInBlankMergeFieldValueForMDGL1008Form(policy, builder);
                this.MDGL1008Classification(policy, builder);
            }
        }

        /// <summary>
        /// Fill bookmark field value
        /// </summary>
        /// <param name="limits">GlLimit</param>
        /// <param name="documentManager">PolicyDocumentManager</param>
        private void FillLimitsOfInsuranceInMDGL1008Form(GlLimit limits, ref IPolicyDocumentManager documentManager)
        {
            this.FillBookmarkField(ref documentManager, "RetroDate", "None");

            this.FillBookmarkField(ref documentManager, "BlankMergeField_1", string.Format("{0:n0}", limits.GeneralAggregate));

            if (limits.ProductsCopsAggregate == "Included" || limits.ProductsCopsAggregate == "Excluded")
            {
                this.FillBookmarkField(ref documentManager, "BlankMergeField_2", limits.ProductsCopsAggregate);
            }
            else
            {
                this.FillBookmarkField(ref documentManager, "BlankMergeField_2", string.Format("{0:n0}", FormatHelper.FormatValue(Convert.ToDecimal(limits.ProductsCopsAggregate), FormatHelper.CurrencyFormat)));
            }

            if (limits.PersonalAdvertising == "Excluded")
            {
                this.FillBookmarkField(ref documentManager, "BlankMergeField_3", limits.PersonalAdvertising);
            }
            else
            {
                this.FillBookmarkField(ref documentManager, "BlankMergeField_3", string.Format("{0:n0}", FormatHelper.FormatValue(Convert.ToDecimal(limits.PersonalAdvertising), FormatHelper.CurrencyFormat)));
            }

            this.FillBookmarkField(ref documentManager, "BlankMergeField_4", string.Format("{0:n0}", limits.PerOccurrence));

            if (limits.DamageRentedPremises == "Excluded")
            {
                this.FillBookmarkField(ref documentManager, "BlankMergeField_5", limits.DamageRentedPremises);
            }
            else
            {
                this.FillBookmarkField(ref documentManager, "BlankMergeField_5", string.Format("{0:n0}", FormatHelper.FormatValue(Convert.ToDecimal(limits.DamageRentedPremises), FormatHelper.CurrencyFormat)));
            }

            if (limits.MedicalExpense == "Excluded")
            {
                this.FillBookmarkField(ref documentManager, "BlankMergeField_6", limits.MedicalExpense);
            }
            else
            {
                this.FillBookmarkField(ref documentManager, "BlankMergeField_6", string.Format("{0:n0}", FormatHelper.FormatValue(Convert.ToDecimal(limits.MedicalExpense), FormatHelper.CurrencyFormat)));
            }
        }

        /// <summary>
        /// Fill location details bookmark field
        /// </summary>
        /// <param name="policy">IBaseGlLine</param>
        /// <param name="builder">PolicyDocumentManager</param>
        private void FillLocationInBlankMergeFieldValueForMDGL1008Form(Policy policy, DocumentBuilder builder)
        {
            builder.MoveToBookmark("BlankMergeField_7");

            builder.Font.Size = 12;
            builder.Font.Bold = true;

            Aspose.Words.Tables.Table table = builder.StartTable();
            builder.InsertCell();
            builder.CellFormat.LeftPadding = 2;
            builder.CellFormat.RightPadding = 2;
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.First;
            builder.Write("ALL PREMISES YOU OWN, RENT OR OCCUPY");
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
            builder.CellFormat.Borders.Right.LineStyle = LineStyle.None;
            builder.CellFormat.TopPadding = 3;
            builder.CellFormat.BottomPadding = 3;
            builder.Write("Loc. No.");
            builder.InsertCell();

            builder.CellFormat.ClearFormatting();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.Auto;
            builder.CellFormat.Borders.Left.LineStyle = LineStyle.None;
            builder.CellFormat.TopPadding = 3;
            builder.CellFormat.BottomPadding = 3;
            builder.Write("ADDRESS OF ALL PREMISES YOU OWN, RENT OR OCCUPY");
            builder.EndRow();

            builder.CellFormat.ClearFormatting();

            List<RiskUnitLocationResponse> locationsList = null;
            if (policy.LobOrder.Any(x => x.Code == LineOfBusiness.GL))
            {
                locationsList = policy.GlLine.GetDistinctLocations().OrderBy(t => t.LocationNumber).ToList<RiskUnitLocationResponse>();
            }
            else if (policy.LobOrder.Any(x => x.Code == LineOfBusiness.SpecialEvent))
            {
                locationsList = policy.SpecialEventLine.GetDistinctLocations().OrderBy(t => t.LocationNumber).ToList<RiskUnitLocationResponse>();
            }

            if (locationsList.Any())
            {
                foreach (var location in locationsList)
                {
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

            table.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(568.5);
            table.StyleOptions = Aspose.Words.Tables.TableStyleOptions.Default2003;
            table.LeftPadding = 0.5;
            table.LeftIndent = -3;
            builder.EndTable();
        }

        /// <summary>
        /// MDGL 1000 Form Locations table
        /// </summary>
        /// <param name="policy">Policy</param>
        /// <param name="builder">DocumentBuilder</param>
        private void MDGL1008Classification(Policy policy, DocumentBuilder builder)
        {
            builder.MoveToBookmark("BlankMergeField_17");

            builder.Font.Size = 10;
            builder.Font.Bold = true;
            Aspose.Words.Tables.Table table = builder.StartTable();

            this.MDGL1008FormClassificationTableHeader(builder);

            List<BaseGlRiskUnit> riskUnits = null;
            string productsCopsAggregate = string.Empty;

            if (policy.LobOrder.Any(x => x.Code == LineOfBusiness.GL))
            {
                riskUnits = policy.GlLine.RiskUnits.ToList<BaseGlRiskUnit>();
                productsCopsAggregate = this.GetProdCopsValue(policy.GlLine.Limits);
            }
            else if (policy.LobOrder.Any(x => x.Code == LineOfBusiness.SpecialEvent))
            {
                riskUnits = policy.SpecialEventLine.RiskUnits.ToList<BaseGlRiskUnit>();
                productsCopsAggregate = this.GetProdCopsValue(policy.SpecialEventLine.Limits);
            }

            if (riskUnits.Any())
            {
                foreach (var riskUnit in riskUnits)
                {
                    this.FillBlankMergeFieldValueInMDGL1008FormClassificationPremium(productsCopsAggregate, riskUnit, builder);

                    if (policy.LobOrder.Any(x => x.Code == LineOfBusiness.GL))
                    {
                        foreach (GlRiskUnit exposure in (riskUnit as GlRiskUnit).AdditionalExposures)
                        {
                            exposure.LocationNumber = riskUnit.LocationNumber;
                            this.FillBlankMergeFieldValueInMDGL1008FormClassificationPremium(productsCopsAggregate, exposure, builder);
                        }
                    }
                    else if (policy.LobOrder.Any(x => x.Code == LineOfBusiness.SpecialEvent))
                    {
                        foreach (GlRiskUnit exposure in (riskUnit as SpecialEventRiskUnit).AdditionalExposures)
                        {
                            exposure.LocationNumber = riskUnit.LocationNumber;
                            this.FillBlankMergeFieldValueInMDGL1008FormClassificationPremium(productsCopsAggregate, exposure, builder);
                        }
                    }
                }
            }

            if (policy.LobOrder.Any(x => x.Code == LineOfBusiness.GL))
            {
                if (policy.GlLine.RollUp.OptionalCoverages.Any(c => c.Code != "TRIA" && c.IsSelected == true))
                {
                    foreach (var optionalCoverage in policy.GlLine.RollUp.OptionalCoverages.Where(c => c.Code != "TRIA" && c.IsSelected == true))
                    {
                        this.FillBlankMergeFieldValueInMDGL1008FormOC(optionalCoverage, builder);
                    }
                }
            }
            else if (policy.LobOrder.Any(x => x.Code == LineOfBusiness.SpecialEvent))
            {
                if (policy.SpecialEventLine.RollUp.OptionalCoverages.Any(c => c.Code != "TRIA" && c.IsSelected == true))
                {
                    foreach (var optionalCoverage in policy.SpecialEventLine.RollUp.OptionalCoverages.Where(c => c.Code != "TRIA" && c.IsSelected == true))
                    {
                        this.FillBlankMergeFieldValueInMDGL1008FormOC(optionalCoverage, builder);
                    }
                }
            }

            this.MDGL1008FormTotalPremium(policy, builder);

            table.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(568.5);
            table.StyleOptions = Aspose.Words.Tables.TableStyleOptions.Default2003;
            table.LeftPadding = 0.5;
            table.LeftIndent = -3;
            builder.EndTable();
        }

        /// <summary>
        /// MDGL 1008 Form Classification Table Header
        /// </summary>
        /// <param name="builder">DocumentBuilder</param>
        private void MDGL1008FormClassificationTableHeader(DocumentBuilder builder)
        {
            builder.CellFormat.TopPadding = 4;
            builder.CellFormat.BottomPadding = 4;

            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.First;
            builder.Write("CLASSIFICATION AND PREMIUM");
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

            // populate headings
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.VerticalMerge = Aspose.Words.Tables.CellMerge.First;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPercent(95);
            builder.Write("Loc. No");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.VerticalMerge = Aspose.Words.Tables.CellMerge.First;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPercent(234.15);
            builder.Write("Code No. Classification");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.VerticalMerge = Aspose.Words.Tables.CellMerge.First;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPercent(58.5);
            builder.Write("Rating Basis");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.VerticalMerge = Aspose.Words.Tables.CellMerge.First;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPercent(85.5);
            builder.Write("Premium Basis");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.VerticalMerge = Aspose.Words.Tables.CellMerge.First;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPercent(100);
            builder.Write("Other Basis");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.VerticalMerge = Aspose.Words.Tables.CellMerge.None;
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.First;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPercent(100);
            builder.Write("Rate");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.VerticalMerge = Aspose.Words.Tables.CellMerge.None;
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPercent(100);
            builder.Write(string.Empty);
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.VerticalMerge = Aspose.Words.Tables.CellMerge.None;
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.First;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPercent(100);
            builder.Write("Advance Premium");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.VerticalMerge = Aspose.Words.Tables.CellMerge.None;
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPercent(100);
            builder.Write(string.Empty);
            builder.EndRow();

            builder.CellFormat.ClearFormatting();

            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.VerticalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPercent(95);
            builder.Write(string.Empty);
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.VerticalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPercent(234.15);
            builder.Write(string.Empty);
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.VerticalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPercent(58.5);
            builder.Write(string.Empty);
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.VerticalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPercent(85.5);
            builder.Write(string.Empty);
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.VerticalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPercent(100);
            builder.Write(string.Empty);
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.VerticalMerge = Aspose.Words.Tables.CellMerge.None;
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.None;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPercent(100);
            builder.Write("Pr/Co");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.VerticalMerge = Aspose.Words.Tables.CellMerge.None;
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.None;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPercent(100);
            builder.Write("All Other");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.VerticalMerge = Aspose.Words.Tables.CellMerge.None;
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.None;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPercent(100);
            builder.Write("Pr/Co");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.VerticalMerge = Aspose.Words.Tables.CellMerge.None;
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.None;
            builder.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPercent(100);
            builder.Write("All Other");
            builder.EndRow();

            builder.CellFormat.ClearFormatting();
            builder.Font.Bold = false;
        }

        /// <summary>
        /// Fill bookmark field value
        /// </summary>
        /// <param name="productCopsAggregate">Products Cops aggregate</param>
        /// <param name="item">Risk Unit Item</param>
        /// <param name="builder">DocumentBuilder</param>
        private void FillBlankMergeFieldValueInMDGL1008FormClassificationPremium(string productCopsAggregate, BaseGlRiskUnit item, DocumentBuilder builder)
        {
            decimal premium = 0m;
            builder.CellFormat.TopPadding = 4;
            builder.CellFormat.BottomPadding = 4;

            ////Loc. No.
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.Write(item.LocationNumber.HasValue ? item.LocationNumber.Value.ToString() : string.Empty);

            ////Code No. Classification
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.Write(string.Format("{0} {1}", item.ClassCode, item.ClassDescription));

            ////Rating Basis
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.Write(item.PremiumBase);

            ////Premium Basis
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
                builder.Write(string.Format("{0:n0}", item.Exposure));
            }

            ////Other Basis
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.Write(string.Empty);

            ////Pr-Co (Rate)
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.Write(productCopsAggregate);

            ////All Other (Rate)
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            if (item.IsIncluded)
            {
                builder.Write("Incl.");
            }
            else if (item.UnderwriterAdjustedRate != 0)
            {
                string adjrate = string.Format("${0:n2}", Math.Round(item.UnderwriterAdjustedRate, 2, MidpointRounding.AwayFromZero));
                builder.Write(adjrate);
            }
            else if (item.AgentAdjustedRate != 0)
            {
                string adjrate = string.Format("${0:n2}", Math.Round(item.AgentAdjustedRate, 2, MidpointRounding.AwayFromZero));
                builder.Write(adjrate);
            }
            else
            {
                string adjrate = string.Format("${0:n2}", Math.Round(item.DevelopedRate, 2, MidpointRounding.AwayFromZero));
                builder.Write(adjrate);
            }

            ////Pr-Co (Advance Premium)
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.Write(productCopsAggregate);

            ////All Other (Advance Premium)
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            if (item.IsIfAny)
            {
                builder.Write("If Any");
            }
            else if (item.IsIncluded)
            {
                builder.Write("Incl.");
            }
            else if (item.UnderwriterAdjustedPremium.HasValue)
            {
                premium = premium + item.UnderwriterAdjustedPremium.Value;
                builder.Write(string.Format("${0:n0}", item.UnderwriterAdjustedPremium));
            }
            else if (item.AgentAdjustedPremium.HasValue)
            {
                premium = premium + item.AgentAdjustedPremium.Value;
                builder.Write(string.Format("${0:n0}", item.AgentAdjustedPremium));
            }
            else
            {
                premium = premium + item.Premium;
                builder.Write(string.Format("${0:n0}", item.Premium));
            }

            builder.EndRow();
        }

        /// <summary>
        /// Fill bookmark field value for Optional Coverages
        /// </summary>
        /// <param name="item">OptionalCoverage</param>
        /// <param name="builder">DocumentBuilder</param>
        private void FillBlankMergeFieldValueInMDGL1008FormOC(OptionalCoverage item, DocumentBuilder builder)
        {
            builder.CellFormat.TopPadding = 4;
            builder.CellFormat.BottomPadding = 4;

            ////Loc. No
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.Write(string.Empty);

            ////Code No. Classification
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.Write(item.ShortName);

            ////Rating Basis & Premium Basis
            if (item.Basis == RateBasis.Unit)
            {
                builder.InsertCell();
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                builder.Write("Each");
                builder.InsertCell();
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                builder.Write(item.IsIncluded ? "Incl." : "Each");
            }
            else if (item.Basis == RateBasis.Flat)
            {
                builder.InsertCell();
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                builder.Write("Flat");
                builder.InsertCell();
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                builder.Write(item.IsIncluded ? "Incl." : "Flat");
            }
            else if (item.Basis == RateBasis.ClassPremium || item.Basis == RateBasis.LinePremium)
            {
                builder.InsertCell();
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                builder.Write("Percent of Premium");
                builder.InsertCell();
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                builder.Write(item.IsIncluded ? "Incl." : string.Empty);
            }
            else
            {
                builder.InsertCell();
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                builder.Write(string.Empty);

                builder.InsertCell();
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                builder.Write(item.IsIncluded ? "Incl." : string.Empty);
            }

            ////Other Basis
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.Write(string.Empty);

            ////Pr-Co (Rate)
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.Write(string.Empty);

            ////All Other (Rate)
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
            else if (item.Basis == RateBasis.ClassPremium || item.Basis == RateBasis.LinePremium)
            {
                builder.Write(string.Empty);
            }
            else if (item.Basis == RateBasis.Flat)
            {
                builder.Write("Flat");
            }
            else if (item.AdjustedRate.HasValue)
            {
                string adjrate = string.Format("${0:n2}", Math.Round(item.AdjustedRate.Value, 2, MidpointRounding.AwayFromZero));
                builder.Write(adjrate);
            }
            else
            {
                string adjrate = string.Format("${0:n2}", Math.Round(item.Rate.Value, 2, MidpointRounding.AwayFromZero));
                builder.Write(adjrate);
            }

            ////Pr-Co (Advance Premium)
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.Write(string.Empty);

            ////All Other (Advance Premium)
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            if (item.IsIncluded)
            {
                builder.Write("Incl.");
            }
            else if (item.AdjustedPremium.HasValue)
            {
                builder.Write(string.Format("${0:n0}", item.AdjustedPremium));
            }
            else
            {
                builder.Write(string.Format("${0:n0}", item.Premium));
            }

            builder.EndRow();
        }

        /// <summary>
        /// MDGL 1008 Form Total Premium
        /// </summary>
        /// <param name="policy">Policy</param>
        /// <param name="builder">DocumentBuilder</param>
        private void MDGL1008FormTotalPremium(Policy policy, DocumentBuilder builder)
        {
            StringBuilder val = new StringBuilder();
            val.AppendFormat("*(a) Area  *(c) Total Cost  *(m) Admissions  *(p) Payroll  *(s) Gross Sales  (u) Units *(r) Gross Receipts  (e) Each  (o) Other: {0}\n", string.Empty);
            val.AppendLine("Premium Basis identified with a “*” is per 1000 of selected basis.");

            string premium = string.Empty;
            if (policy.LobOrder.Any(x => x.Code == LineOfBusiness.GL))
            {
                premium = string.Format("${0}{1}", this.GetApplicablePremium(policy, LineOfBusiness.GL, FormatHelper.CurrencyFormat), policy.GlLine.RollUp.IsMinimumPremium ? " MP" : string.Empty);
            }
            else if (policy.LobOrder.Any(x => x.Code == LineOfBusiness.SpecialEvent))
            {
                premium = string.Format("${0}{1}", this.GetApplicablePremium(policy, LineOfBusiness.SpecialEvent, FormatHelper.CurrencyFormat), policy.SpecialEventLine.RollUp.IsMinimumPremium ? " MP" : string.Empty);
            }

            builder.CellFormat.TopPadding = 4;
            builder.CellFormat.BottomPadding = 4;

            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.First;
            builder.Write(val.ToString());
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

            builder.Font.Bold = true;
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.First;
            builder.Write("Total Advance Premium");

            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.Previous;
            builder.Write(string.Empty);

            builder.Font.Bold = false;
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.HorizontalMerge = Aspose.Words.Tables.CellMerge.None;
            builder.Write(premium);
            builder.EndRow();
        }

        /// <summary>
        /// Returns the ProdCops aggregate value
        /// </summary>
        /// <param name="item">BaseGlRiskUnit</param>
        /// <returns>String</returns>
        private string GetProdCopsValue(GlLimit item)
        {
            if (item.ProductsCopsAggregate.EqualsIgnoreCase("Excluded"))
            {
                return "Excl.";
            }
            else if (item.ProductsCopsAggregate.EqualsIgnoreCase("Included"))
            {
                return "Incl.";
            }

            return string.Empty;
        }
    }
}
