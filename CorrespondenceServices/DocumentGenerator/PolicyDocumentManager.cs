// <copyright file="PolicyDocumentManager.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>

namespace Mkl.WebTeam.DocumentGenerator
{
    using System;
    using System.Collections.Generic;
    using System.Drawing;
    using System.IO;
    using System.Linq;
    using Aspose.Words;
    using Aspose.Words.Markup;
    using Aspose.Words.Tables;
    using DecisionModel.Interfaces;
    using DecisionModel.Models.Policy;
    using Mkl.WebTeam.DocumentGenerator.Helpers;
    using Mkl.WebTeam.SubmissionShared.Enumerations;

    /// <summary>
    /// This is an example of a class to utilize the Aspose library for WORD Doc interactions
    /// </summary>
    public class PolicyDocumentManager : DocumentManager, IPolicyDocumentManager
    {
        /// <summary>
        /// The limit type dictionary
        /// </summary>
        private Dictionary<string, string> limitTypeDict = new Dictionary<string, string>
            {
                { "PO", "Per Occurrence" },
                { "GA", "General Aggregate" },
                { "POA", "Products/Completed Operations Aggregate" },
                { "CSL", "Combined Single Limit" },
                { "EA", "Each Accident" },
                { "PLA", "Policy Limit Aggregate" },
                { "EE", "Each Employee" },
                { "Included_GL", "Included in General Liability" },
                { "Included_Auto", "Included in Auto Liability" },
                { "na", "n/a" }
            };

        /// <summary>
        /// Initializes a new instance of the <see cref="PolicyDocumentManager"/> class. merging all docs in the docPath array and adding them to.
        /// memory stream to work with
        /// </summary>
        /// <param name="docPath">Array of document paths.</param>
        /// <param name="id">Transaction Id</param>
        /// <exception cref="System.ArgumentException">Specified document does not exist. Please verify path is for a valid document.;docPath</exception>
        public PolicyDocumentManager(string[] docPath, string id)
            : base(docPath, id)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PolicyDocumentManager"/> class.
        /// </summary>
        /// <param name="docStream">Stream containing the document template</param>
        /// <param name="id">Transaction Id</param>
        public PolicyDocumentManager(Stream docStream, string id)
             : base(docStream, id)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PolicyDocumentManager"/> class.
        /// </summary>
        /// <param name="docPath">The document path.</param>
        /// <param name="id">Transaction Id</param>
        /// <exception cref="System.ArgumentException">Specified document does not exist. Please verify path is for a valid document.;docPath</exception>
        public PolicyDocumentManager(string docPath, string id)
            : base(docPath, id)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PolicyDocumentManager"/> class.
        /// </summary>
        /// <param name="doc">The document.</param>
        public PolicyDocumentManager(Aspose.Words.Document doc)
            : base(doc)
        {
        }

        /// <summary>
        /// Builds the unordered list.
        /// </summary>
        /// <param name="list">The list.</param>
        /// <param name="control">The control.</param>
        /// <param name="isBind">if set to <c>true</c> [is bind].</param>
        /// <exception cref="System.Exception">Specified control must be a Paragraph or Inline control, or must come after a Paragraph or Inline control.</exception>
        public void BuildUnorderedList(IEnumerable<Subjectivity> list, StructuredDocumentTag control, bool isBind = false)
        {
            DocumentBuilder builder = new DocumentBuilder(this.WordDoc);
            if (!(control.PreviousSibling is Paragraph))
            {
                throw new Exception("Specified control must be a Paragraph or Inline control, or must come after a Paragraph or Inline control.");
            }

            builder.MoveTo(control.PreviousSibling);
            builder.StartBookmark(control.Tag);
            builder.EndBookmark(control.Tag);

            if (list == null || list.Count() <= 0)
            {
                // if we had no subjectivities we could write a message here such as
                // builder.Writeln("No " + control.GetText());
            }
            else
            {
                builder.InsertBreak(BreakType.ParagraphBreak);
                builder.Write(isBind ? control.GetText().Replace("quote", "binder") : control.GetText());
                Aspose.Words.Lists.List blist = this.WordDoc.Lists.Add(Aspose.Words.Lists.ListTemplate.BulletDefault);
                Aspose.Words.Lists.ListLevel level1 = blist.ListLevels[0];
                builder.ListFormat.List = blist;
                builder.Font.Bold = false;
                foreach (Subjectivity line in list)
                {
                    if (line.Text != null)
                    {
                        builder.Writeln(this.StripHTML(line.Text.ToString()));
                    }
                    else if (line.Name != null)
                    {
                        builder.Writeln(line.Name.ToString());
                    }
                }

                builder.ListFormat.List = null;
            }
        }

        /// <summary>
        /// Inserts the buildings table.
        /// </summary>
        /// <param name="buildings">The buildings.</param>
        /// <param name="control">The control.</param>
        /// <exception cref="Exception">Specified control must be a Paragraph or Inline control, or must come after a Paragraph or Inline control.</exception>
        public void InsertBuildingsTable(IOrderedEnumerable<CfRiskUnit> buildings, StructuredDocumentTag control)
        {
            DocumentBuilder builder = new DocumentBuilder(this.WordDoc);
            if (!(control.PreviousSibling is Paragraph))
            {
                throw new Exception("Specified control must be a Paragraph or Inline control, or must come after a Paragraph or Inline control.");
            }

            builder.MoveTo(control.PreviousSibling);
            builder.StartBookmark(control.Tag);
            builder.EndBookmark(control.Tag);
            foreach (CfRiskUnit building in buildings)
            {
                Table table = this.BuildBuildingsTable(building, builder);
                builder.InsertBreak(BreakType.ParagraphBreak);
                table = this.BuildPremiseStructuresTable(building, builder);
                builder.InsertBreak(BreakType.ParagraphBreak);
            }
        }

        /// <summary>
        /// Inserts the clauses.
        /// </summary>
        /// <param name="clauses">The clauses.</param>
        /// <param name="control">The control.</param>
        /// <param name="documents">The documents.</param>
        /// <param name="useHeading">if set to <c>true</c> [use heading].</param>
        /// <exception cref="Exception">Specified control must be a Paragraph or Inline control, or must come after a Paragraph or Inline control.</exception>
        public void InsertClauses(IEnumerable<Clause> clauses, StructuredDocumentTag control, IEnumerable<DecisionModel.Models.Policy.Document> documents, bool useHeading = true)
        {
            DocumentBuilder builder = new DocumentBuilder(this.WordDoc);
            if (!(control.PreviousSibling is Paragraph))
            {
                throw new Exception("Specified control must be a Paragraph or Inline control, or must come after a Paragraph or Inline control.");
            }

            builder.MoveTo(control.PreviousSibling);
            builder.StartBookmark(control.Tag);
            builder.EndBookmark(control.Tag);

            this.BuildClauses(clauses, documents, builder);
        }

        /// <summary>
        /// Inserts the coverage options.
        /// </summary>
        /// <param name="clauses">The clauses.</param>
        /// <param name="control">The control.</param>
        /// <param name="documents">The documents.</param>
        /// <param name="optCov">The optional coverages.</param>
        /// <param name="useHeading">if set to <c>true</c> [use heading].</param>
        /// <exception cref="Exception">Specified control must be a Paragraph or Inline control, or must come after a Paragraph or Inline control.</exception>
        public void InsertCoverageOptions(IEnumerable<Clause> clauses, StructuredDocumentTag control, IEnumerable<DecisionModel.Models.Policy.Document> documents, IEnumerable<OptionalCoverage> optCov, bool useHeading = true)
        {
            DocumentBuilder builder = new DocumentBuilder(this.WordDoc);
            if (!(control.PreviousSibling is Paragraph))
            {
                throw new Exception("Specified control must be a Paragraph or Inline control, or must come after a Paragraph or Inline control.");
            }

            builder.MoveTo(control.PreviousSibling);
            builder.StartBookmark(control.Tag);
            builder.EndBookmark(control.Tag);

            Table table = this.BuildCoverageOptionsTable(clauses, documents, optCov, builder);
        }

        /// <summary>
        /// Inserts the coverages table.
        /// </summary>
        /// <param name="coverages">The coverages.</param>
        /// <param name="control">The control.</param>
        /// <param name="glOccurLimit">The gl occur limit.</param>
        /// <param name="glDeductible">The gl deductible.</param>
        /// <param name="showLimit">if set to <c>true</c> [show limit].</param>
        /// <exception cref="System.Exception">Specified control must be a Paragraph or Inline control, or must come after a Paragraph or Inline control.</exception>
        public void InsertCoveragesTable(IEnumerable<OptionalCoverage> coverages, StructuredDocumentTag control, int glOccurLimit, int glDeductible, bool showLimit = true)
        {
            DocumentBuilder builder = new DocumentBuilder(this.WordDoc);
            if (!(control.PreviousSibling is Paragraph))
            {
                throw new Exception("Specified control must be a Paragraph or Inline control, or must come after a Paragraph or Inline control.");
            }

            builder.MoveTo(control.PreviousSibling);
            builder.StartBookmark(control.Tag);
            builder.EndBookmark(control.Tag);
            builder.Font.Bold = control.ContentsFont.Bold;
            builder.InsertBreak(BreakType.ParagraphBreak);
            builder.Writeln(control.GetText());
            Table table = this.BuildCoverageListTable(coverages, builder, glOccurLimit, glDeductible, showLimit);
            builder.InsertBreak(BreakType.ParagraphBreak);
        }

        /// <summary>
        /// Inserts the exposures table.
        /// </summary>
        /// <param name="risks">The risks.</param>
        /// <param name="control">The control.</param>
        /// <exception cref="Exception">Specified control must be a Paragraph or Inline control, or must come after a Paragraph or Inline control.</exception>
        public void InsertExposuresTable(IEnumerable<IBaseGlRiskUnit> risks, StructuredDocumentTag control)
        {
            DocumentBuilder builder = new DocumentBuilder(this.WordDoc);
            if (!(control.PreviousSibling is Paragraph))
            {
                throw new Exception("Specified control must be a Paragraph or Inline control, or must come after a Paragraph or Inline control.");
            }

            string[] objArray = control.Tag.Split('.');
            builder.MoveTo(control.PreviousSibling);
            builder.StartBookmark(control.Tag);
            builder.EndBookmark(control.Tag);
            builder.Font.Bold = control.ContentsFont.Bold;
            builder.InsertBreak(BreakType.ParagraphBreak);
            builder.Writeln(control.GetText());
            Table table = this.BuildExposures(risks, builder, objArray[1]);
        }

        /// <summary>
        /// Inserts the im items table.
        /// </summary>
        /// <param name="imRisk">The im risk.</param>
        /// <param name="control">The control.</param>
        /// <param name="useHeading">if set to <c>true</c> [use heading].</param>
        /// <exception cref="Exception">Specified control must be a Paragraph or Inline control, or must come after a Paragraph or Inline control.</exception>
        public void InsertImItemsTable(ImRiskUnit imRisk, StructuredDocumentTag control, bool useHeading = true)
        {
            DocumentBuilder builder = new DocumentBuilder(this.WordDoc);
            if (!(control.PreviousSibling is Paragraph))
            {
                throw new Exception("Specified control must be a Paragraph or Inline control, or must come after a Paragraph or Inline control.");
            }

            string header = control.GetText().Trim();
            builder.MoveTo(control.PreviousSibling);
            builder.StartBookmark(control.Tag);
            builder.EndBookmark(control.Tag);

            if (imRisk.ImClassType == IMClassType.MotorTruckCargo)
            {
                Table table = this.BuildMtcCommoditiesTable(imRisk, builder);
            }
            else if (imRisk.ImClassType == IMClassType.ContractorsEquipment)
            {
                ImRatingItem item = imRisk.Items.Find(x => x.ItemType == ImRateItemType.ScheduledEquipment && x.TotalInsuredValue > 0);
                if (item != null)
                {
                    builder.InsertBreak(BreakType.ParagraphBreak);
                    Table table = this.BuildImSchedEquipTable(item, builder);
                }

                if (imRisk.Items.Any(x => (x.ItemType == ImRateItemType.MiscUnscheduledTools || x.ItemType == ImRateItemType.EmployeesTools || x.ItemType == ImRateItemType.EquipmentLeasedRented) && x.TotalInsuredValue > 0))
                {
                    builder.InsertBreak(BreakType.ParagraphBreak);
                    builder.Bold = true;
                    builder.Writeln("Optional Coverage Extensions:");
                    Table table = this.BuildCoverageExtTable(imRisk, builder);
                }
            }
            else if (imRisk.ImClassType == IMClassType.MiscellaneousPropertyFloater)
            {
                ImRatingItem item = imRisk.Items.Find(x => x.ItemType == ImRateItemType.MiscPropertyEndorsement);
                if (item != null)
                {
                    Table table = this.BuildMiscPropEndorsementTable(imRisk, builder);
                }
            }
        }

        /// <summary>
        /// Inserts the xs layers table.
        /// </summary>
        /// <param name="layers">The layers.</param>
        /// <param name="control">The control.</param>
        /// <param name="useHeading">if set to <c>true</c> [use heading].</param>
        /// <exception cref="System.Exception">Specified control must be a Paragraph or Inline control, or must come after a Paragraph or Inline control.</exception>
        public void InsertLayerTable(List<XsLayer> layers, StructuredDocumentTag control, bool useHeading = true)
        {
            DocumentBuilder builder = new DocumentBuilder(this.WordDoc);
            if (!(control.PreviousSibling is Paragraph))
            {
                throw new Exception("Specified control must be a Paragraph or Inline control, or must come after a Paragraph or Inline control.");
            }

            string header = control.GetText().Trim();
            builder.MoveTo(control.PreviousSibling);
            builder.StartBookmark(control.Tag);
            builder.EndBookmark(control.Tag);

            if (layers == null || layers.Count <= 0)
            {
                if (header != string.Empty)
                {
                    // if No Documents then show nothing
                    // builder.Writeln("No " + header);
                }
            }
            else
            {
                if (header != string.Empty)
                {
                    builder.Writeln(header);
                }

                Table table = this.BuildLayerTable(layers, builder);
            }
        }

        /// <summary>
        /// Inserts the link list table.
        /// </summary>
        /// <param name="docs">The docs.</param>
        /// <param name="control">The control.</param>
        /// <param name="useHeading">if set to <c>true</c> [use heading].</param>
        /// <param name="isBindLetter">Indicates if the letter is a bind letter</param>
        /// <exception cref="Exception">Specified control must be a Paragraph or Inline control, or must come after a Paragraph or Inline control.</exception>
        public void InsertLinkListTable(List<DecisionModel.Models.Policy.Document> docs, StructuredDocumentTag control, bool useHeading = true, bool isBindLetter = false)
        {
            DocumentBuilder builder = new DocumentBuilder(this.WordDoc);
            if (!(control.PreviousSibling is Paragraph))
            {
                throw new Exception("Specified control must be a Paragraph or Inline control, or must come after a Paragraph or Inline control.");
            }

            string header = control.GetText().Trim();
            builder.MoveTo(control.PreviousSibling);
            builder.StartBookmark(control.Tag);
            builder.EndBookmark(control.Tag);

            if (docs == null || docs.Count <= 0)
            {
                if (header != string.Empty)
                {
                    // if No Documents then show nothing
                    // builder.Writeln("No " + header);
                }
            }
            else
            {
                if (header != string.Empty)
                {
                    builder.Writeln(header);
                }

                if (docs.Where(d => d.IsDisplayFormChangeCue == true).Count() > 0)
                {
                    builder.Font.Bold = true;
                    builder.Write("+");
                    builder.Font.Bold = false;
                    builder.Write(":  indicates that form or edition is new for this renewal term");
                    builder.Writeln(" ");
                    builder.Writeln(" ");
                }

                Table table = this.BuildLinkListTable(docs, builder, isBindLetter);
            }
        }

        /// <summary>
        /// Inserts the name amount table.
        /// </summary>
        /// <param name="fees">The fees.</param>
        /// <param name="control">The control.</param>
        /// <exception cref="Exception">Specified control must be a Paragraph or Inline control, or must come after a Paragraph or Inline control.</exception>
        public void InsertNameAmountTable(IOrderedEnumerable<TaxFee> fees, StructuredDocumentTag control)
        {
            DocumentBuilder builder = new DocumentBuilder(this.WordDoc);
            if (!(control.PreviousSibling is Paragraph))
            {
                throw new Exception("Specified control must be a Paragraph or Inline control, or must come after a Paragraph or Inline control.");
            }

            bool bold = builder.Font.Bold;
            builder.MoveTo(control.PreviousSibling);
            builder.StartBookmark(control.Tag);
            builder.EndBookmark(control.Tag);
            builder.InsertBreak(BreakType.ParagraphBreak);
            builder.Font.Bold = true;
            builder.Write(control.GetText());
            Table table = this.BuildNameAmountTable(fees, builder);
            builder.InsertBreak(BreakType.ParagraphBreak);
            builder.Font.Bold = bold;
        }

        /// <summary>
        /// Inserts the excess underlying coverage table.
        /// </summary>
        /// <param name="limits">The limits.</param>
        /// <param name="coverage">The coverage.</param>
        /// <param name="control">The control.</param>
        /// <exception cref="System.Exception">Specified control must be a Paragraph or Inline control, or must come after a Paragraph or Inline control.</exception>
        public void InsertUnderlyingCoverageTable(List<Limit> limits, XsUnderlyingCoverage coverage, StructuredDocumentTag control)
        {
            DocumentBuilder builder = new DocumentBuilder(this.WordDoc);
            if (!(control.PreviousSibling is Paragraph))
            {
                throw new Exception("Specified control must be a Paragraph or Inline control, or must come after a Paragraph or Inline control.");
            }

            builder.MoveTo(control.PreviousSibling);
            builder.StartBookmark(control.Tag);
            builder.EndBookmark(control.Tag);
            builder.Font.Bold = control.ContentsFont.Bold;
            builder.InsertBreak(BreakType.ParagraphBreak);
            builder.Writeln(control.GetText());
            Table table = this.BuildUnderlyingCoverageTable(limits, coverage, builder);
            builder.InsertBreak(BreakType.ParagraphBreak);
        }

        /// <summary>
        /// Insert XS auto liability vehicleType.
        /// </summary>
        /// <param name="vehicleTypes">The vehicle types.</param>
        /// <param name="control">The control.</param>
        public void InsertXSAutoLiabilityVehicleType(IEnumerable<XsAutoLiabilityVehicleType> vehicleTypes, StructuredDocumentTag control)
        {
            DocumentBuilder builder = new DocumentBuilder(this.WordDoc);
            builder.MoveTo(control.PreviousSibling);
            builder.StartBookmark(control.Tag);
            builder.EndBookmark(control.Tag);
            builder.Write("Auto exposure:");
            var isValTabs = vehicleTypes.Any(x => x.VehicleType.ToString().Count() > 20);

            foreach (var vehicle in vehicleTypes)
            {
                if (vehicle.Quantity > 0)
                {
                    var valTabs = "\t";
                    var vehicleName = string.Concat(vehicle.VehicleType.ToString().Select(x => char.IsUpper(x) ? " " + x : x.ToString())).Trim();
                    if (vehicleName.Count() < 20 && isValTabs)
                    {
                        valTabs = "\t\t";
                    }

                    builder.Write("\t" + vehicleName + valTabs + vehicle.Quantity + " unit(s) \v");
                }
            }
        }

        /// <summary>
        /// Inserts the exposures table.
        /// </summary>
        /// <param name="warranties">The warranties.</param>
        /// <param name="control">The control.</param>
        /// <exception cref="System.Exception">Specified control must be a Paragraph or Inline control, or must come after a Paragraph or Inline control.</exception>
        public void InsertWarrantiesTable(IEnumerable<Warranty> warranties, StructuredDocumentTag control)
        {
            DocumentBuilder builder = new DocumentBuilder(this.WordDoc);
            if (!(control.PreviousSibling is Paragraph))
            {
                throw new Exception("Specified control must be a Paragraph or Inline control, or must come after a Paragraph or Inline control.");
            }

            builder.MoveTo(control.PreviousSibling);
            builder.StartBookmark(control.Tag);
            builder.EndBookmark(control.Tag);
            builder.Font.Bold = true;
            builder.Write("Warranties:");
            builder.InsertBreak(BreakType.LineBreak);
            builder.Font.Bold = false;
            builder.Write(control.GetText());
            Table table = this.BuildWarranties(warranties, builder);
        }

        /// <summary>
        /// Inserts the policy form table
        /// </summary>
        /// <param name="documents">Collection of documents</param>
        /// <param name="control">Book mark</param>
        public void InsertPolicyFormTable(Dictionary<dynamic, Dictionary<string, DecisionModel.Models.Policy.Document>> documents, Aspose.Words.Bookmark control)
        {
            this.ReplaceNodeValue(control, string.Empty);
            DocumentBuilder builder = new DocumentBuilder(this.WordDoc);
            builder.MoveToBookmark(control.Name);
            builder.StartBookmark(control.Name);
            builder.EndBookmark(control.Name);
            var table = this.BuildPolicyFormTable(documents, builder);
        }

        /// <summary>
        /// Inserts the policy form table
        /// </summary>
        /// <param name="notes">Underwriter to agent notes</param>
        /// <param name="control">Book mark</param>
        public void InsertAgentToUnderwriterNotes(string notes, StructuredDocumentTag control)
        {
            DocumentBuilder builder = new DocumentBuilder(this.WordDoc);
            if (control.PreviousSibling != null)
            {
                builder.MoveTo(control.PreviousSibling);
            }

            if (string.IsNullOrWhiteSpace(notes))
            {
                notes = string.Empty;
            }

            builder.StartBookmark(control.Tag);
            builder.EndBookmark(control.Tag);

            builder.InsertBreak(BreakType.ParagraphBreak);
            builder.InsertHtml(notes);
        }

        /// <summary>
        /// Builds the mailing address
        /// </summary>
        /// <param name="streetAddress">streetAddress</param>
        /// <param name="control">Document Control</param>
        public void InsertMailingAddress(Address streetAddress, StructuredDocumentTag control)
        {
            if (streetAddress == null)
            {
                return;
            }

            DocumentBuilder builder = new DocumentBuilder(this.WordDoc);
            if (!(control.PreviousSibling is Paragraph))
            {
                throw new Exception("Specified control must be a Paragraph or Inline control, or must come after a Paragraph or Inline control.");
            }

            if (!string.IsNullOrWhiteSpace(streetAddress.Line1) || !string.IsNullOrWhiteSpace(streetAddress.Line2) || !string.IsNullOrWhiteSpace(streetAddress.City) || !string.IsNullOrWhiteSpace(streetAddress.StateCode) || !string.IsNullOrWhiteSpace(streetAddress.ZipCode))
            {
                builder.MoveTo(control.PreviousSibling);
                builder.StartBookmark(control.Tag);
                builder.EndBookmark(control.Tag);
                builder.InsertBreak(BreakType.ParagraphBreak);
            }

            string header = control.GetText();
            bool displayHeader = false;
            if (header == " \r")
            {
                header = "Mailing Address: ";
            }
            else if (header.Length > 1 && header.Substring(header.Length - 1) == "\r")
            {
                header = header.Substring(0, header.Length - 1);
            }

            if (!string.IsNullOrWhiteSpace(streetAddress.Line1))
            {
                builder.Write(header + "\t" + streetAddress.Line1);
                displayHeader = true;
            }

            if (!string.IsNullOrWhiteSpace(streetAddress.Line2))
            {
                if (displayHeader == false)
                {
                    builder.Write(header + "\t" + streetAddress.Line2);
                    displayHeader = true;
                }
                else
                {
                    builder.Write("\v\t\t\t" + streetAddress.Line2);
                }
            }

            if (!string.IsNullOrWhiteSpace(streetAddress.City))
            {
                if (displayHeader == false)
                {
                    builder.Write(header + "\t" + streetAddress.City);
                    displayHeader = true;
                }
                else
                {
                    builder.Write("\v\t\t\t" + streetAddress.City);
                }
            }

            if (!string.IsNullOrWhiteSpace(streetAddress.StateCode))
            {
                if (string.IsNullOrWhiteSpace(streetAddress.City))
                {
                    if (displayHeader == false)
                    {
                        builder.Write(header + "\t" + streetAddress.StateCode);
                        displayHeader = true;
                    }
                    else
                    {
                        builder.Write("\v\t\t\t" + streetAddress.StateCode);
                    }
                }
                else
                {
                    builder.Write(", " + streetAddress.StateCode);
                }
            }

            if (!string.IsNullOrWhiteSpace(streetAddress.ZipCode))
            {
                if (string.IsNullOrWhiteSpace(streetAddress.City))
                {
                    if (displayHeader == false)
                    {
                        builder.Write(header + "\t" + streetAddress.ZipCode);
                        displayHeader = true;
                    }
                    else
                    {
                        builder.Write(" " + streetAddress.ZipCode);
                    }
                }
                else
                {
                    builder.Write(" " + streetAddress.ZipCode);
                }
            }
        }

        /// <summary>
        /// Builds the buildings table.
        /// </summary>
        /// <param name="building">The building.</param>
        /// <param name="builder">The builder.</param>
        /// <returns>Table</returns>
        private Table BuildBuildingsTable(CfRiskUnit building, DocumentBuilder builder)
        {
            Table table = builder.StartTable();
            table.SetBorders(LineStyle.Single, 1, Color.BlueViolet);
            bool bold = builder.Font.Bold;
            ParagraphAlignment parAlign = builder.ParagraphFormat.Alignment;
            double firstColWidth = 120;

            // override settings
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.Font.Bold = true;

            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(firstColWidth);
            builder.Write("Location " + this.CheckStringForNulls(building.PremisesNumber) + ", Building " + this.CheckStringForNulls(building.BuildingNumber));
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.Auto;
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.EndRow();

            if (!string.IsNullOrWhiteSpace(building.CustomDescription))
            {
                builder.InsertCell();
                builder.Font.Bold = true;
                builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(firstColWidth);
                builder.Write("Description");
                builder.InsertCell();
                builder.CellFormat.PreferredWidth = PreferredWidth.Auto;
                builder.CellFormat.HorizontalMerge = CellMerge.First;
                builder.Font.Bold = false;
                builder.Write(building.CustomDescription);
                builder.InsertCell();
                builder.CellFormat.HorizontalMerge = CellMerge.Previous;
                builder.InsertCell();
                builder.CellFormat.HorizontalMerge = CellMerge.Previous;
                builder.EndRow();
            }

            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(firstColWidth);
            builder.Font.Bold = true;
            builder.Write("Address");
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.Auto;
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Font.Bold = false;
            builder.Write(this.CheckStringForNulls(building.StreetAddress.Line1) + " ");
            builder.Write((building.StreetAddress.Line2 == null) ? string.Empty : building.StreetAddress.Line2.ToString() + " ");
            builder.Write(this.CheckStringForNulls(building.StreetAddress.City) + ", ");
            builder.Write(this.CheckStringForNulls(building.StreetAddress.StateCode) + " ");
            builder.Write(this.CheckStringForNulls(building.StreetAddress.ZipCode));
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            builder.EndRow();

            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(firstColWidth);
            builder.Font.Bold = true;
            builder.Write("Occupancy Class");
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.Auto;
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Font.Bold = false;
            builder.Write(this.CheckStringForNulls(building.ClassCode) + " - " + this.CheckStringForNulls(building.ClassDescription));
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            builder.EndRow();

            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(firstColWidth);
            builder.Font.Bold = true;
            builder.Write("Causes of Loss");
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.Auto;
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.Font.Bold = false;
            if (string.IsNullOrWhiteSpace(building.CauseOfLoss))
            {
                builder.Write("n/a");
            }
            else
            {
                builder.Write(building.CausesOfLoss.Find(x => x.Code == building.CauseOfLoss).Description.ToString());
            }

            if (building.IsWindHailExcluded)
            {
                builder.InsertCell();
                builder.CellFormat.PreferredWidth = PreferredWidth.Auto;
                builder.CellFormat.HorizontalMerge = CellMerge.First;
                builder.Font.Bold = false;
                builder.Write("excluding the perils of wind and hail");
                builder.InsertCell();
                builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            }
            else
            {
                builder.InsertCell();
                builder.CellFormat.HorizontalMerge = CellMerge.None;
                builder.InsertCell();
                builder.CellFormat.HorizontalMerge = CellMerge.None;
            }

            builder.EndRow();

            List<CfCoverage> outdoorCovs = building.Coverages.FindAll(x => x.CoverageTypeCode == "FENCE"
                || x.CoverageTypeCode == "OTHER"
                || x.CoverageTypeCode == "SIGN"
                || x.CoverageTypeCode == "POOLSPA"
                || x.CoverageTypeCode == "CANPUMP"
                || x.CoverageTypeCode == "TGLASS");
            if (outdoorCovs.Count > 0)
            {
                string outdoorStr = string.Empty;
                foreach (CfCoverage cov in outdoorCovs)
                {
                    outdoorStr += " " + cov.Description + ";";
                }

                if (outdoorStr.Length > 1)
                {
                    outdoorStr = outdoorStr.Remove(outdoorStr.Length - 1);
                }

                builder.InsertCell();
                builder.CellFormat.HorizontalMerge = CellMerge.None;
                builder.Font.Bold = true;
                builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(firstColWidth);
                builder.Write("Except");
                builder.InsertCell();
                builder.CellFormat.PreferredWidth = PreferredWidth.Auto;
                builder.CellFormat.HorizontalMerge = CellMerge.First;
                builder.Font.Bold = false;
                if (building.IsWindHailExcluded)
                {
                    builder.Write("Basic X VMM, X-Wind & Hail for" + outdoorStr);
                }
                else
                {
                    builder.Write("Basic X VMM for" + outdoorStr);
                }

                builder.InsertCell();
                builder.CellFormat.HorizontalMerge = CellMerge.Previous;
                builder.InsertCell();
                builder.CellFormat.HorizontalMerge = CellMerge.Previous;
                builder.EndRow();
            }

            if ((building.IsWindHailExcluded || string.IsNullOrWhiteSpace(building.WindHailDeductible) || building.WindHailDeductible == "NONE") && string.IsNullOrWhiteSpace(building.TheftDeductible) && string.IsNullOrWhiteSpace(building.GlassDeductible))
            {
                builder.InsertCell();
                builder.CellFormat.HorizontalMerge = CellMerge.None;
                builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(firstColWidth);
                builder.Font.Bold = true;
                builder.Write("Deductible");
                builder.InsertCell();
                builder.CellFormat.PreferredWidth = PreferredWidth.Auto;
                builder.CellFormat.HorizontalMerge = CellMerge.First;
                builder.Font.Bold = true;
                builder.Write("All perils");
                builder.Font.Bold = false;
                builder.Write("\t$" + Convert.ToDouble(this.CheckStringForNulls(building.Deductible, "00")).ToString("N0"));
                builder.InsertCell();
                builder.CellFormat.HorizontalMerge = CellMerge.Previous;
                builder.InsertCell();
                builder.CellFormat.HorizontalMerge = CellMerge.Previous;
                builder.EndRow();
            }
            else
            {
                builder.InsertCell();
                builder.CellFormat.HorizontalMerge = CellMerge.None;
                builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(firstColWidth);
                builder.Font.Bold = true;
                builder.Write("Deductible");
                builder.InsertCell();
                builder.CellFormat.PreferredWidth = PreferredWidth.Auto;
                builder.CellFormat.HorizontalMerge = CellMerge.First;
                builder.Font.Bold = true;
                builder.Write("AOP");
                builder.Font.Bold = false;
                builder.Write("\t\t$" + Convert.ToDouble(this.CheckStringForNulls(building.Deductible, "00")).ToString("N0"));
                builder.InsertCell();
                builder.CellFormat.HorizontalMerge = CellMerge.Previous;
                builder.InsertCell();
                builder.CellFormat.HorizontalMerge = CellMerge.Previous;
                builder.EndRow();

                if (!string.IsNullOrWhiteSpace(building.GlassDeductible))
                {
                    builder.InsertCell();
                    builder.CellFormat.HorizontalMerge = CellMerge.None;
                    builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(firstColWidth);
                    builder.InsertCell();
                    builder.CellFormat.PreferredWidth = PreferredWidth.Auto;
                    builder.CellFormat.HorizontalMerge = CellMerge.First;
                    builder.Font.Bold = true;
                    builder.Write("Glass");
                    builder.Font.Bold = false;
                    if (building.GlassDeductible == "NONE")
                    {
                        builder.Write("\t\tAOP");
                    }
                    else
                    {
                        builder.Write("\t\t$" + Convert.ToDouble(this.CheckStringForNulls(building.GlassDeductible, "00")).ToString("N0"));
                    }

                    builder.InsertCell();
                    builder.CellFormat.HorizontalMerge = CellMerge.Previous;
                    builder.InsertCell();
                    builder.CellFormat.HorizontalMerge = CellMerge.Previous;
                    builder.EndRow();
                }

                if (!string.IsNullOrWhiteSpace(building.TheftDeductible))
                {
                    builder.InsertCell();
                    builder.CellFormat.HorizontalMerge = CellMerge.None;
                    builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(firstColWidth);
                    builder.InsertCell();
                    builder.CellFormat.PreferredWidth = PreferredWidth.Auto;
                    builder.CellFormat.HorizontalMerge = CellMerge.First;
                    builder.Font.Bold = true;
                    builder.Write("Theft");
                    builder.Font.Bold = false;
                    if (building.TheftDeductible == "NONE")
                    {
                        builder.Write("\t\tAOP");
                    }
                    else
                    {
                        builder.Write("\t\t$" + Convert.ToDouble(this.CheckStringForNulls(building.TheftDeductible, "00")).ToString("N0"));
                    }

                    builder.InsertCell();
                    builder.CellFormat.HorizontalMerge = CellMerge.Previous;
                    builder.InsertCell();
                    builder.CellFormat.HorizontalMerge = CellMerge.Previous;
                    builder.EndRow();
                }

                if (!building.IsWindHailExcluded && (!string.IsNullOrWhiteSpace(building.WindHailDeductible) && building.WindHailDeductible != "NONE"))
                {
                    builder.InsertCell();
                    builder.CellFormat.HorizontalMerge = CellMerge.None;
                    builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(firstColWidth);
                    builder.InsertCell();
                    builder.CellFormat.PreferredWidth = PreferredWidth.Auto;
                    builder.CellFormat.HorizontalMerge = CellMerge.First;
                    builder.Font.Bold = true;
                    builder.Write("Wind/hail");
                    builder.Font.Bold = false;
                    if (building.WindHailDeductible == "Other")
                    {
                        if (building.OtherWindHailDeductible.IsNumeric())
                        {
                            builder.Write("\t$" + Convert.ToDouble(this.CheckStringForNulls(building.OtherWindHailDeductible, "00")).ToString("N0"));
                        }
                        else
                        {
                            builder.Write("\t" + this.CheckStringForNulls(building.OtherWindHailDeductible, "$00"));
                        }
                    }
                    else
                    {
                        builder.Write("\t" + building.WindHailDeductibles.Find(x => x.Code == building.WindHailDeductible).Description.ToString());
                    }

                    builder.InsertCell();
                    builder.CellFormat.HorizontalMerge = CellMerge.Previous;
                    builder.InsertCell();
                    builder.CellFormat.HorizontalMerge = CellMerge.Previous;
                    builder.EndRow();
                }
            }

            if (!string.IsNullOrWhiteSpace(Convert.ToString(building.TheftSublimit)) && building.TheftSublimit != 0)
            {
                builder.InsertCell();
                builder.CellFormat.HorizontalMerge = CellMerge.None;
                builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(firstColWidth);
                builder.Font.Bold = true;
                builder.Write("Theft Sublimit");
                builder.InsertCell();
                builder.CellFormat.PreferredWidth = PreferredWidth.Auto;
                builder.CellFormat.HorizontalMerge = CellMerge.First;
                builder.Font.Bold = false;
                builder.Write("$" + Convert.ToDouble(this.CheckStringForNulls(building.TheftSublimit, "00")).ToString("N0"));
                builder.InsertCell();
                builder.CellFormat.HorizontalMerge = CellMerge.Previous;
                builder.InsertCell();
                builder.CellFormat.HorizontalMerge = CellMerge.Previous;
                builder.EndRow();
            }

            if (!string.IsNullOrWhiteSpace(building.TheftSpecifiedProperty))
            {
                builder.InsertCell();
                builder.CellFormat.HorizontalMerge = CellMerge.None;
                builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(firstColWidth);
                builder.Font.Bold = true;
                builder.Write("Spec Cov Prop");
                builder.InsertCell();
                builder.CellFormat.HorizontalMerge = CellMerge.First;
                builder.CellFormat.PreferredWidth = PreferredWidth.Auto;
                builder.Font.Bold = false;
                builder.Write(building.TheftSpecifiedProperty);
                builder.InsertCell();
                builder.CellFormat.HorizontalMerge = CellMerge.Previous;
                builder.InsertCell();
                builder.CellFormat.HorizontalMerge = CellMerge.Previous;
                builder.EndRow();
            }

            builder.EndTable();
            table.ClearBorders();

            builder.Font.Bold = bold;
            builder.ParagraphFormat.Alignment = parAlign;
            return table;
        }

        /// <summary>
        /// Builds the clauses.
        /// </summary>
        /// <param name="clauses">The clauses.</param>
        /// <param name="documents">The documents.</param>
        /// <param name="builder">The builder.</param>
        private void BuildClauses(IEnumerable<Clause> clauses, IEnumerable<DecisionModel.Models.Policy.Document> documents, DocumentBuilder builder)
        {
            Aspose.Words.Lists.List blist = this.WordDoc.Lists.Add(Aspose.Words.Lists.ListTemplate.BulletDefault);
            Aspose.Words.Lists.ListLevel level1 = blist.ListLevels[0];
            builder.Font.Bold = false;
            string docName = string.Empty;
            foreach (var document in documents)
            {
                if (document.ReturnIsSelected())
                {
                    var list = clauses.Where(x => x.Documents.Contains(document.NormalizedNumber));
                    if (list != null && list.Count() > 0)
                    {
                        builder.Writeln("Per form " + document.DisplayNumber + " the following terms, limitations, conditions, exclusions and extensions will apply:");
                        builder.ListFormat.List = blist;
                        foreach (var line in list)
                        {
                            builder.Writeln(line.Name.ToString());
                        }

                        builder.ListFormat.List = null;
                        builder.InsertBreak(BreakType.ParagraphBreak);
                    }
                }
            }
        }

        /// <summary>
        /// Builds the coverage ext table.
        /// </summary>
        /// <param name="imRisk">The im risk.</param>
        /// <param name="builder">The builder.</param>
        /// <returns>Table</returns>
        private Table BuildCoverageExtTable(ImRiskUnit imRisk, DocumentBuilder builder)
        {
            // item.TotalInsuredValue,item.RateBasis, item.Premium
            Table table = builder.StartTable();
            double fntSize = builder.Font.Size;
            bool bold = builder.Font.Bold;
            ParagraphAlignment parAlign = builder.ParagraphFormat.Alignment;

            // override for heading
            builder.Font.Bold = true;
            builder.Font.Size = 9;

            // set headings
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.Write("Equipment type");
            builder.InsertCell();
            builder.Write("Any one item limit");
            builder.InsertCell();
            builder.Write("Any one occurrence limit");
            builder.InsertCell();
            builder.Write("Rate");
            builder.InsertCell();
            builder.Write("Premium");
            builder.EndRow();

            // set to table vals
            builder.Font.Bold = false;

            ImRatingItem item;
            item = imRisk.Items.Find(x => x.TotalInsuredValue > 0 && x.ItemType == ImRateItemType.MiscUnscheduledTools);
            this.BuildCoverageExtRow(item, "Miscellaneous unscheduled tools", builder);
            item = imRisk.Items.Find(x => x.TotalInsuredValue > 0 && x.ItemType == ImRateItemType.EmployeesTools);
            this.BuildCoverageExtRow(item, "Employee's tools", builder);
            item = imRisk.Items.Find(x => x.TotalInsuredValue > 0 && x.ItemType == ImRateItemType.EquipmentLeasedRented);
            this.BuildCoverageExtRow(item, "Equipment leased or rented from others", builder);

            builder.EndTable();

            // reset to original vals
            builder.Font.Bold = bold;
            builder.ParagraphFormat.Alignment = parAlign;
            builder.Font.Size = fntSize;
            return table;
        }

        /// <summary>
        /// Builds the coverage ext row.
        /// </summary>
        /// <param name="item">The item.</param>
        /// <param name="equipType">Type of the equip.</param>
        /// <param name="builder">The builder.</param>
        private void BuildCoverageExtRow(ImRatingItem item, string equipType, DocumentBuilder builder)
        {
            if (item != null)
            {
                builder.InsertCell();
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(600);
                builder.Write(equipType);

                builder.InsertCell();
                builder.Write("$" + Convert.ToDouble(this.CheckStringForNulls(item.MaxSingleInsuredValue, "0")).ToString("N0"));

                builder.InsertCell();
                builder.Write("$" + Convert.ToDouble(this.CheckStringForNulls(item.TotalInsuredValue, "0")).ToString("N0"));

                builder.InsertCell();
                builder.Write("$" + Convert.ToDouble(this.CheckStringForNulls(item.ApplicableRate, "0")).ToString("N2"));

                builder.InsertCell();
                builder.Write("$" + Convert.ToDouble(item.Premium).ToString("N0"));

                builder.EndRow();
            }
        }

        /// <summary>
        /// Builds the coverage list table.
        /// </summary>
        /// <param name="coverages">The coverages.</param>
        /// <param name="builder">The builder.</param>
        /// <param name="glOccurLimit">The gl occur limit.</param>
        /// <param name="glDeductible">The gl deductible.</param>
        /// <param name="showLimit">if set to <c>true</c> [show limit].</param>
        /// <returns>
        /// Table
        /// </returns>
        private Table BuildCoverageListTable(IEnumerable<OptionalCoverage> coverages, DocumentBuilder builder, int glOccurLimit, int glDeductible, bool showLimit = true)
        {
            Table table = builder.StartTable();
            double fntSize = builder.Font.Size;
            bool bold = builder.Font.Bold;
            ParagraphAlignment parAlign = builder.ParagraphFormat.Alignment;

            // override for heading
            builder.Font.Bold = true;
            builder.Font.Size = 9;

            // set headings
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.Write("Coverage");
            if (showLimit)
            {
                builder.InsertCell();
                builder.CellFormat.HorizontalMerge = CellMerge.First;
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                builder.Write("Limit");
                builder.InsertCell();
                builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            }

            // builder.InsertCell();
            // builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            // builder.Write("Rating Basis");
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.Write("Qty.");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.Write("Premium");
            builder.EndRow();

            // set to table vals
            builder.Font.Bold = false;

            foreach (OptionalCoverage coverage in coverages)
            {
                // Coverage
                builder.InsertCell();
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(600);
                builder.Write(coverage.Name.ToString());

                // Limit
                if (showLimit)
                {
                    if (coverage.LimitLevelId == null)
                    {
                        builder.InsertCell();

                        // Anfisa: Remove later when data is loaded and rule is created for handling Miscellaneous Professional Liability Coverage
                        if (glOccurLimit != 0 && coverage.Code == "PL")
                        {
                            builder.CellFormat.Borders.Right.LineStyle = LineStyle.None;
                            builder.CellFormat.HorizontalMerge = CellMerge.None;
                            builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(500);
                            builder.Write("Each Wrongful Act");
                            builder.InsertBreak(BreakType.LineBreak);
                            builder.Write("Aggregate");

                            builder.CellFormat.Borders.Left.LineStyle = LineStyle.None;
                            builder.InsertCell();
                            builder.CellFormat.HorizontalMerge = CellMerge.None;
                            builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(200);
                            builder.Write("$" + glOccurLimit.ToString());
                            builder.InsertBreak(BreakType.LineBreak);
                            builder.Write("$" + glOccurLimit.ToString());
                        }
                        else
                        {
                            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                            builder.CellFormat.HorizontalMerge = CellMerge.First;
                            builder.Write("n/a");
                            builder.InsertCell();
                            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
                        }
                    }
                    else
                    {
                        builder.InsertCell();
                        builder.CellFormat.Borders.Right.LineStyle = LineStyle.None;
                        builder.CellFormat.HorizontalMerge = CellMerge.None;
                        builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(500);
                        Dictionary<string, string> limitOpt = new Dictionary<string, string>();
                        List<Limit> limits = coverage.LimitOptions.Find(x => x.Id == coverage.LimitLevelId).Limits.OrderBy(x => x.LimitOptionDisplayOrder).ToList();
                        foreach (Limit limit in limits)
                        {
                            limitOpt.Add(limit.LimitTypeName, limit.Text);
                        }

                        if (coverage.Code == "LDB")
                        {
                            limitOpt.Add("Retro Date", "Inception");
                        }
                        else if (coverage.Code == "LPWEE")
                        {
                            limitOpt.Add("Deductible", "$" + Convert.ToDouble(glDeductible).ToString("N0"));
                        }

                        bool firstLine = true;
                        foreach (var limit in limitOpt)
                        {
                            if (!firstLine)
                            {
                                builder.InsertBreak(BreakType.LineBreak);
                            }

                            firstLine = false;
                            builder.Write(limit.Key);
                        }

                        builder.CellFormat.Borders.Left.LineStyle = LineStyle.None;
                        builder.InsertCell();
                        builder.CellFormat.HorizontalMerge = CellMerge.None;
                        builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(200);
                        firstLine = true;
                        foreach (var limit in limitOpt)
                        {
                            if (!firstLine)
                            {
                                builder.InsertBreak(BreakType.LineBreak);
                            }

                            firstLine = false;
                            builder.Write(limit.Value);
                        }
                    }
                }

                // Quantity
                builder.InsertCell();
                builder.CellFormat.ClearFormatting();
                builder.CellFormat.HorizontalMerge = CellMerge.None;
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(200);
                if (coverage.Basis.ToString() == "Unit")
                {
                    builder.Write(string.IsNullOrWhiteSpace(coverage.Quantity.ToString()) ? "0" : Convert.ToInt16(coverage.Quantity).ToString("N0"));
                }
                else
                {
                    builder.Write("n/a");
                }

                // Premium
                builder.InsertCell();
                builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(200);
                if (coverage.IsIncluded)
                {
                    builder.Write("Included");
                }
                else
                {
                    builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                    if (coverage.IsIncluded)
                    {
                        builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                        builder.Write("Included");
                    }
                    else if (string.IsNullOrWhiteSpace(coverage.AdjustedPremium.ToString()) && !(coverage.AdjustedPremium > 0))
                    {
                        builder.Write("$" + Convert.ToDouble(coverage.Premium).ToString("N0"));
                    }
                    else
                    {
                        builder.Write("$" + Convert.ToDouble(coverage.AdjustedPremium).ToString("N0"));
                    }
                }

                builder.EndRow();
            }

            builder.EndTable();

            // reset to original vals
            builder.Font.Bold = bold;
            builder.ParagraphFormat.Alignment = parAlign;
            builder.Font.Size = fntSize;
            return table;
        }

        /// <summary>
        /// Builds the coverage options table.
        /// </summary>
        /// <param name="clauses">The clauses.</param>
        /// <param name="documents">The documents.</param>
        /// <param name="optCov">The optional coverages.</param>
        /// <param name="builder">The builder.</param>
        /// <returns>Table</returns>
        private Table BuildCoverageOptionsTable(IEnumerable<Clause> clauses, IEnumerable<DecisionModel.Models.Policy.Document> documents, IEnumerable<OptionalCoverage> optCov, DocumentBuilder builder)
        {
            Table table = builder.StartTable();
            double fntSize = builder.Font.Size;
            bool bold = builder.Font.Bold;
            LineStyle lineStyle = builder.CellFormat.Borders.LineStyle;
            builder.CellFormat.Borders.LineStyle = LineStyle.None;
            ParagraphAlignment parAlign = builder.ParagraphFormat.Alignment;

            // override for heading
            builder.Font.Bold = true;
            builder.Font.Size = 9;

            // set headings
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.Write("Coverage Options");
            builder.InsertCell();
            builder.Write("Premium");
            builder.EndRow();

            // set to table vals
            builder.Font.Bold = false;

            foreach (var document in documents)
            {
                if (document.ReturnIsSelected())
                {
                    var list = clauses.Where(x => x.Documents.Contains(document.NormalizedNumber));
                    if (list != null && list.Count() > 0)
                    {
                        foreach (var line in list)
                        {
                            OptionalCoverage coverage = optCov.FirstOrDefault(x => x.Code == line.Code);
                            builder.InsertCell();
                            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(600);
                            builder.Write("Per form " + this.CheckStringForNulls(document.DisplayNumber) + " " + this.CheckStringForNulls(line.Name));

                            builder.InsertCell();

                            // We should not be coverage.CustomLimitLevel here. This is just a band aid until a more
                            // premenant filx. Such as adding a text field with the dialog to the clause itself
                            string premium = coverage == null ? "Included" : "Included " + coverage.CustomLimitLevel;
                            builder.Write(premium);
                            builder.EndRow();
                        }
                    }
                }
            }

            builder.EndTable();

            // reset to original vals
            builder.CellFormat.Borders.LineStyle = lineStyle;
            builder.Font.Bold = bold;
            builder.ParagraphFormat.Alignment = parAlign;
            builder.Font.Size = fntSize;
            if (table.Rows.Count == 1)
            {
                table.Remove();
            }

            return table;
        }

        /// <summary>
        /// Builds the exposure row.
        /// </summary>
        /// <param name="locationNumber">The location number of the exposure.</param>
        /// <param name="stateCode">The state code.</param>
        /// <param name="territory">The territory.</param>
        /// <param name="classCode">The class code.</param>
        /// <param name="classDescription">The class description.</param>
        /// <param name="customDescription">The custom description.</param>
        /// <param name="premiumBase">The premium base.</param>
        /// <param name="isIncluded">if set to <c>true</c> [is included].</param>
        /// <param name="isIfAny">Indicates if the exposure is rated on an if any basis</param>
        /// <param name="exposure">The exposure.</param>
        /// <param name="rate">The rate.</param>
        /// <param name="premium">The premium.</param>
        /// <param name="builder">The builder.</param>
        private void BuildExposureRow(
            int? locationNumber,
            string stateCode,
            string territory,
            string classCode,
            string classDescription,
            string customDescription,
            string premiumBase,
            bool isIncluded,
            bool isIfAny,
            long exposure,
            decimal rate,
            decimal premium,
            DocumentBuilder builder)
        {
            // Location
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(50);
            builder.Write(this.CheckStringForNulls(locationNumber).ToString());

            // State-Territory
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(50);
            builder.Write(this.CheckStringForNulls(stateCode).ToString() + " - " + this.CheckStringForNulls(territory).ToString());

            // Class Code
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(40);
            builder.Write(this.CheckStringForNulls(classCode).ToString());

            // Description
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = PreferredWidth.Auto;
            builder.Write(this.CheckStringForNulls(classDescription).ToString());
            if (!string.IsNullOrWhiteSpace(customDescription))
            {
                builder.Write(" (" + customDescription.ToString() + ')');
            }

            // Rating Basis
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(50);
            builder.Write(this.CheckStringForNulls(premiumBase).ToString());

            // Exposure
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(50);
            if (isIncluded)
            {
                builder.Write("Included");
            }
            else if (isIfAny)
            {
                builder.Write("If any");
            }
            else
            {
                builder.Write(Convert.ToDouble(this.CheckStringForNulls(exposure)).ToString("N0"));
            }

            // Rate
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(30);
            if (isIncluded)
            {
                builder.Write("n/a");
            }
            else
            {
                builder.Write(Convert.ToDouble(this.CheckStringForNulls(rate)).ToString("N2"));
            }

            // Premium
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(50);
            if (isIncluded)
            {
                builder.Write("Included");
            }
            else
            {
                builder.Write("$" + Convert.ToDouble(this.CheckStringForNulls(premium)).ToString("N0"));
            }

            builder.EndRow();
        }

        /// <summary>
        /// Builds the exposures.
        /// </summary>
        /// <param name="coverages">The coverages.</param>
        /// <param name="builder">The builder.</param>
        /// <param name="lob">The lob.</param>
        /// <returns>
        /// Table
        /// </returns>
        private Table BuildExposures(IEnumerable<IBaseGlRiskUnit> coverages, DocumentBuilder builder, string lob)
        {
            Table table = builder.StartTable();
            double fntSize = builder.Font.Size;
            bool bold = builder.Font.Bold;
            ParagraphAlignment parAlign = builder.ParagraphFormat.Alignment;

            builder.Font.Size = 9;

            // override for heading
            builder.Font.Bold = true;

            // loop through columns and set headings
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(50);
            builder.Write("Loc");
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(50);
            builder.Write("State - Territory");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(40);
            builder.Write("Class Code");
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.Auto;
            builder.Write("Description");
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(50);
            builder.Write("Rating Basis");
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(50);
            builder.Write("Exposure");
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(30);
            builder.Write("Rate");
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(50);
            builder.Write("Premium");
            builder.EndRow();

            // reset to original vals
            builder.Font.Bold = false;
            builder.ParagraphFormat.Alignment = parAlign;

            // build exposure rows based on lob type
            if (lob == "GlLine")
            {
                // Commented lines below are for a bug fix that is not ready to be released
                // IEnumerable<GlRiskUnit> includedExposures = new List<GlRiskUnit>();
                foreach (GlRiskUnit coverage in coverages)
                {
                    // IEnumerable<GlRiskUnit> displayExposures = new List<GlRiskUnit>();
                    this.BuildExposureRow(coverage.LocationNumber, coverage.StateCode, coverage.Territory, coverage.ClassCode, coverage.ClassDescription, coverage.CustomDescription, coverage.PremiumBase, coverage.IsIncluded, coverage.IsIfAny, coverage.Exposure, coverage.ApplicableRate, coverage.Premium, builder);
                    if (coverage.AdditionalExposures.Count > 0)
                    {
                        // displayExposures = coverage.AdditionalExposures.Except(includedExposures, new RiskUnitComparer());
                        // includedExposures = includedExposures.Concat(displayExposures.Where(x => x.IsIncluded));
                        // foreach (GlRiskUnit exposure in displayExposures)
                        foreach (GlRiskUnit exposure in coverage.AdditionalExposures)
                        {
                            this.BuildExposureRow(coverage.LocationNumber, exposure.StateCode, exposure.Territory, exposure.ClassCode, exposure.ClassDescription, exposure.CustomDescription, exposure.PremiumBase, exposure.IsIncluded, exposure.IsIfAny, exposure.Exposure, exposure.ApplicableRate, exposure.Premium, builder);
                        }
                    }
                }
            }
            else if (lob == "SpecialEventLine")
            {
                foreach (SpecialEventRiskUnit coverage in coverages)
                {
                    this.BuildExposureRow(coverage.LocationNumber, coverage.StateCode, coverage.Territory, coverage.ClassCode, coverage.ClassDescription, coverage.CustomDescription, coverage.PremiumBase, coverage.IsIncluded, coverage.IsIfAny, coverage.Exposure, coverage.ApplicableRate, coverage.Premium, builder);
                    if (coverage.AdditionalExposures.Count > 0)
                    {
                        foreach (GlRiskUnit exposure in coverage.AdditionalExposures)
                        {
                            this.BuildExposureRow(coverage.LocationNumber, exposure.StateCode, exposure.Territory, exposure.ClassCode, exposure.ClassDescription, exposure.CustomDescription, exposure.PremiumBase, exposure.IsIncluded, exposure.IsIfAny, exposure.Exposure, exposure.ApplicableRate, exposure.Premium, builder);
                        }
                    }
                }
            }
            else if (lob == "OCPLine")
            {
                foreach (OCPRiskUnit coverage in coverages)
                {
                    this.BuildExposureRow(null, coverage.StateCode, coverage.Territory, coverage.ClassCode, coverage.ClassDescription, coverage.CustomDescription, coverage.PremiumBase, coverage.IsIncluded, coverage.IsIfAny, coverage.Exposure, coverage.ApplicableRate, coverage.Premium, builder);
                }
            }
            else if (lob == "LiquorLine")
            {
                foreach (LiquorRiskUnit coverage in coverages)
                {
                    this.BuildExposureRow(coverage.LocationNumber, coverage.StateCode, coverage.Territory, coverage.ClassCode, coverage.ClassDescription, coverage.CustomDescription, coverage.PremiumBase, coverage.IsIncluded, coverage.IsIfAny, coverage.Exposure, coverage.ApplicableRate, coverage.Premium, builder);
                }
            }

            builder.EndTable();
            return table;
        }

        /// <summary>
        /// Builds the im schedule equip table.
        /// </summary>
        /// <param name="item">The item.</param>
        /// <param name="builder">The builder.</param>
        /// <returns>Table</returns>
        private Table BuildImSchedEquipTable(ImRatingItem item, DocumentBuilder builder)
        {
            Table table = builder.StartTable();
            double fntSize = builder.Font.Size;
            bool bold = builder.Font.Bold;
            ParagraphAlignment parAlign = builder.ParagraphFormat.Alignment;

            // override for heading
            builder.Font.Bold = true;
            builder.Font.Size = 9;

            // set headings
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.Write("Equipment type");
            builder.InsertCell();
            builder.Write("Any one occurrence limit");
            builder.InsertCell();
            builder.Write("Rate");
            builder.InsertCell();
            builder.Write("Premium");
            builder.EndRow();

            // set to table vals
            builder.Font.Bold = false;

            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(600);
            builder.Write("Scheduled equipment");

            builder.InsertCell();
            builder.Write("$" + Convert.ToDouble(this.CheckStringForNulls(item.TotalInsuredValue, "0")).ToString("N0"));

            builder.InsertCell();
            builder.Write("$" + Convert.ToDouble(this.CheckStringForNulls(item.ApplicableRate, "0")).ToString("N2"));

            builder.InsertCell();
            builder.Write("$" + Convert.ToDouble(item.Premium).ToString("N0"));

            builder.EndRow();

            builder.EndTable();

            // reset to original vals
            builder.Font.Bold = bold;
            builder.ParagraphFormat.Alignment = parAlign;
            builder.Font.Size = fntSize;
            return table;
        }

        /// <summary>
        /// Builds the layers table.
        /// </summary>
        /// <param name="layers">The layers.</param>
        /// <param name="builder">The builder.</param>
        /// <returns>Table</returns>
        private Table BuildLayerTable(List<XsLayer> layers, DocumentBuilder builder)
        {
            Table table = builder.StartTable();
            bool bold = builder.Font.Bold;
            ParagraphAlignment parAlign = builder.ParagraphFormat.Alignment;
            Color fntColor = builder.Font.Color;

            builder.Font.Bold = false;
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;

            // header row
            builder.InsertCell();
            builder.Write("Excess Limit");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.Write("Premium (excluding Terrorism)");
            builder.InsertBreak(BreakType.LineBreak);
            builder.Font.Italic = true;
            builder.Write("Taxes & fees will vary");
            builder.Font.Italic = false;
            builder.EndRow();

            foreach (XsLayer layer in layers.OrderBy(x => x.Layer))
            {
                builder.InsertCell();
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(130);
                builder.Write("$" + Convert.ToDouble(this.CheckStringForNulls(layer.Layer, "00")).ToString("N0") + " xs primary");

                builder.InsertCell();
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                builder.CellFormat.PreferredWidth = PreferredWidth.Auto;
                builder.Write("$" + Convert.ToDouble(this.CheckStringForNulls(layer.LayerTotalPremium, "00")).ToString("N0"));
                builder.EndRow();
            }

            builder.EndTable();
            builder.Font.Bold = bold;
            builder.ParagraphFormat.Alignment = parAlign;
            builder.Font.Color = fntColor;

            return table;
        }

        /// <summary>
        /// Builds the link list table.
        /// </summary>
        /// <param name="docs">The docs.</param>
        /// <param name="builder">The builder.</param>
        /// <param name="isBindLetter">Indicates if the letter is a bind letter</param>
        /// <returns>Table</returns>
        private Table BuildLinkListTable(List<DecisionModel.Models.Policy.Document> docs, DocumentBuilder builder, bool isBindLetter)
        {
            Table table = builder.StartTable();
            builder.CellFormat.Borders.LineStyle = LineStyle.None;
            table.ClearBorders();
            bool bold = builder.Font.Bold;
            ParagraphAlignment parAlign = builder.ParagraphFormat.Alignment;
            Color fntColor = builder.Font.Color;
            Underline fntUnderline = builder.Font.Underline;
            double borders = builder.CellFormat.Borders.LineWidth;
            bool wraptext = builder.CellFormat.WrapText;

            builder.CellFormat.Borders.LineWidth = 0;
            builder.Font.Bold = false;
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;

            int docOrder = 0;
            foreach (DecisionModel.Models.Policy.Document doc in docs.OrderBy(x => (x.Order == 0)).ThenBy(x => x.Order).ThenBy(x => x.DisplayNumber))
            {
                if (doc.ReturnIsSelected())
                {
                    if (doc.Order / 1000 != docOrder && doc.Order != 0)
                    {
                        builder.InsertCell();
                        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(130);
                        builder.CellFormat.Width = 130;
                        builder.InsertCell();
                        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(350);
                        builder.CellFormat.Width = 350;
                        builder.EndRow();
                        docOrder = doc.Order / 1000;
                    }

                    builder.InsertCell();
                    builder.Font.Italic = false;
                    builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(130);
                    builder.CellFormat.Width = 130;
                    builder.CellFormat.WrapText = false;
                    builder.Font.Color = Color.FromArgb(185, 31, 41);
                    builder.Font.Underline = Underline.Single;

                    // uncomment next line to display order:
                    // builder.Write(doc.Order.ToString() + " ::: ");
                    var formDownloadURL = "https://www.markelonline.com/forms/download/"
                        + this.CheckStringForNulls(doc.NormalizedNumber)
                        + this.CheckStringForNulls(doc.EditionDate.ToUniversalTime().ToString("MMyy"));

                    ////builder.InsertHyperlink(this.CheckStringForNulls(doc.DisplayNumber).ToString() + " " + doc.EditionDate.ToString("MM' 'yy"), "https://www.markelonline.com/form/Pages/FormViewer.aspx?BU=" + this.CheckStringForNulls(doc.AttachmentType).ToString() + "&FormNumber=" + this.CheckStringForNulls(doc.DisplayNumber).Replace(" ", "+"), false);
                    builder.InsertHyperlink(this.CheckStringForNulls(doc.DisplayNumber).ToString() + " " + doc.EditionDate.ToString("MM' 'yy"), formDownloadURL, false);

                    builder.InsertCell();
                    builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(350);
                    builder.CellFormat.Width = 350;
                    builder.CellFormat.WrapText = wraptext;
                    builder.Font.Color = fntColor;
                    builder.Font.Underline = Underline.None;
                    builder.Write(this.CheckStringForNulls(doc.Title).ToString());
                    if (doc.IsDisplayFormChangeCue)
                    {
                        builder.Font.Bold = true;
                        builder.Write(" +");
                        builder.Font.Bold = false;
                    }

                    builder.EndRow();

                    if (doc.Questions != null && doc.Questions.Count > 0)
                    {
                        foreach (var question in doc.Questions.OrderBy(o => o.MultipleRowGroupingNumber == 0 ? o.Order : doc.Questions.Where(q => q.MultipleRowGroupingNumber == o.MultipleRowGroupingNumber).Min(a => a.Order)).ThenBy(b => b.MultipleRowGroupingNumber).ThenBy(r => r.Row).ThenBy(c => c.Column))
                        {
                            if ((isBindLetter && !question.IsHiddenOnBindLetter)
                                || (!isBindLetter && !question.IsHiddenOnQuoteLetter))
                            {
                                var answer = question.AnswerValue ?? string.Empty;
                                switch (question.DisplayFormat)
                                {
                                    case DisplayFormat.DatePicker:
                                        DateTime date = DateTime.MinValue;
                                        if (DateTime.TryParse(answer, out date))
                                        {
                                            answer = date.ToShortDateString();
                                        }

                                        break;
                                    case DisplayFormat.DropDown:
                                    case DisplayFormat.RadioButton:
                                        if (question.Answers != null && question.Answers.Any())
                                        {
                                            var selectedAnswer = question.Answers.FirstOrDefault(a => a.Value == answer);
                                            if (selectedAnswer != null)
                                            {
                                                answer = selectedAnswer.Verbiage;
                                            }
                                        }

                                        break;
                                    case DisplayFormat.CheckBox:
                                        answer = string.Empty;
                                        if (question.Answers != null && question.Answers.Any())
                                        {
                                            var selectedAnswers = question.Answers.Where(a => a.IsSelected).Select(v => v.Verbiage);
                                            foreach (var checkboxAnswer in selectedAnswers)
                                            {
                                                answer += checkboxAnswer + ", ";
                                            }

                                            // Remove the ending comma
                                            if (answer.Length > 1)
                                            {
                                                answer = answer.Substring(0, answer.Length - 2);
                                            }
                                        }

                                        break;
                                    case DisplayFormat.StaticText:
                                        answer = question.Verbiage;
                                        break;
                                }

                                if (answer != string.Empty)
                                {
                                    builder.InsertCell();
                                    builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(130);
                                    builder.CellFormat.Width = 130;
                                    builder.Write(" ");

                                    builder.InsertCell();
                                    builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(350);
                                    builder.CellFormat.Width = 350;
                                    builder.Font.Color = fntColor;
                                    builder.Font.Underline = Underline.None;
                                    builder.Font.Italic = true;

                                    if (question.DisplayFormat == DisplayFormat.StaticText)
                                    {
                                        builder.Write(this.CheckStringForNulls(question.Verbiage?.TrimEnd(':')));
                                    }
                                    else
                                    {
                                        builder.Write(this.CheckStringForNulls(question.Verbiage?.TrimEnd(':')) + ":     " + this.CheckStringForNulls(answer));
                                    }

                                    builder.EndRow();
                                }
                            }
                        }
                    }
                }
            }

            builder.CellFormat.Borders.LineWidth = borders;
            builder.EndTable();
            builder.Font.Bold = bold;
            builder.ParagraphFormat.Alignment = parAlign;
            builder.Font.Color = fntColor;
            builder.Font.Underline = fntUnderline;

            return table;
        }

        /// <summary>
        /// Builds the miscellaneous property endorsement table.
        /// </summary>
        /// <param name="imRisk">The im risk.</param>
        /// <param name="builder">The builder.</param>
        /// <returns>Table</returns>
        private Table BuildMiscPropEndorsementTable(ImRiskUnit imRisk, DocumentBuilder builder)
        {
            ImRatingItem item = imRisk.Items.Find(x => x.ItemType == ImRateItemType.MiscPropertyEndorsement);
            Table table = builder.StartTable();
            double fntSize = builder.Font.Size;
            bool bold = builder.Font.Bold;
            ParagraphAlignment parAlign = builder.ParagraphFormat.Alignment;

            // override for heading
            builder.Font.Bold = true;
            builder.Font.Size = 9;

            // set headings
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.Write("Any one occurrence limit");
            builder.InsertCell();
            builder.Write("Max limit for any one item in any one occurrence");
            builder.InsertCell();
            builder.Write("Deductible per occurrence");
            builder.InsertCell();
            builder.Write("Rate");
            builder.InsertCell();
            builder.Write("Premium");
            builder.EndRow();

            // set to table vals
            builder.Font.Bold = false;

            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(600);
            builder.Write("$" + Convert.ToDouble(this.CheckStringForNulls(item.TotalInsuredValue, "0")).ToString("N0"));

            builder.InsertCell();
            builder.Write("$" + Convert.ToDouble(this.CheckStringForNulls(item.MaxSingleInsuredValue, "0")).ToString("N0"));

            builder.InsertCell();
            builder.Write("$" + Convert.ToDouble(this.CheckStringForNulls(imRisk.PerOccurrenceDeductible.Amount, "0")).ToString("N0"));

            builder.InsertCell();
            builder.Write("$" + Convert.ToDouble(this.CheckStringForNulls(item.ApplicableRate, "0")).ToString("N2"));

            builder.InsertCell();
            builder.Write("$" + Convert.ToDouble(imRisk.Premium).ToString("N0"));

            builder.EndRow();

            builder.EndTable();

            // reset to original vals
            builder.Font.Bold = bold;
            builder.ParagraphFormat.Alignment = parAlign;
            builder.Font.Size = fntSize;
            return table;
        }

        /// <summary>
        /// Builds the MTC commodities table.
        /// </summary>
        /// <param name="imRisk">The im risk.</param>
        /// <param name="builder">The builder.</param>
        /// <returns>Table</returns>
        private Table BuildMtcCommoditiesTable(ImRiskUnit imRisk, DocumentBuilder builder)
        {
            Table table = builder.StartTable();
            double fntSize = builder.Font.Size;
            bool bold = builder.Font.Bold;
            ParagraphAlignment parAlign = builder.ParagraphFormat.Alignment;

            // override for heading
            builder.Font.Bold = true;
            builder.Font.Size = 9;

            // set headings
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(30);
            builder.Write("# of Units");
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(350);
            builder.Write("Commodity");
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(90);
            builder.Write("Limit of Liability per Vehicle");
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(50);
            builder.Write("Rate per unit");
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.Auto;
            builder.Write("Total Premium");
            builder.EndRow();

            // set to table vals
            builder.Font.Bold = false;

            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.Write(this.CheckStringForNulls(imRisk.NumberOfVehicles, "0"));

            builder.InsertCell();
            IEnumerable<ImRatingItem> commodities = imRisk.Commodities.FindAll(x => x.IsSelected).OrderBy(x => x.Description);
            foreach (ImRatingItem item in commodities)
            {
                builder.Writeln("- " + item.Description);
            }

            builder.InsertCell();
            builder.Write("$" + Convert.ToDouble(this.CheckStringForNulls(imRisk.VehicleLiabilityLimit, "0")).ToString("N0"));

            builder.InsertCell();
            builder.Write("$" + Convert.ToDouble(this.CheckStringForNulls(imRisk.ApplicableRate, "0")).ToString("N2"));

            builder.InsertCell();
            builder.Write("$" + Convert.ToDouble(imRisk.Premium).ToString("N0"));

            builder.EndRow();

            builder.EndTable();

            // reset to original vals
            builder.Font.Bold = bold;
            builder.ParagraphFormat.Alignment = parAlign;
            builder.Font.Size = fntSize;
            return table;
        }

        /// <summary>
        /// Builds the name amount table.
        /// </summary>
        /// <param name="fees">The fees.</param>
        /// <param name="builder">The builder.</param>
        /// <returns>Table</returns>
        private Table BuildNameAmountTable(IOrderedEnumerable<TaxFee> fees, DocumentBuilder builder)
        {
            Table table = builder.StartTable();
            double fntSize = builder.Font.Size;
            bool bold = builder.Font.Bold;
            ParagraphAlignment parAlign = builder.ParagraphFormat.Alignment;

            builder.Font.Bold = false;
            builder.ParagraphFormat.Alignment = parAlign;

            if (fees != null)
            {
                foreach (TaxFee fee in fees)
                {
                    builder.InsertCell();
                    builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(255);
                    builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                    builder.Write(fee.Name.ToString());

                    builder.InsertCell();
                    builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
                    builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                    if (fee.Amount == 0)
                    {
                        builder.Write("n/a");
                    }
                    else
                    {
                        builder.Write("$" + Convert.ToDouble(this.CheckStringForNulls(fee.Amount, "00")).ToString("N2"));
                    }

                    builder.EndRow();
                }
            }

            if (table.Rows.Count > 0)
            {
                table.PreferredWidth = PreferredWidth.FromPoints(355);
            }

            builder.EndTable();
            table.ClearBorders();
            table.PreferredWidth = PreferredWidth.FromPoints(355);

            // reset to original vals
            builder.Font.Bold = bold;
            builder.ParagraphFormat.Alignment = parAlign;
            builder.Font.Size = fntSize;
            return table;
        }

        /// <summary>
        /// Builds the premise structures table.
        /// </summary>
        /// <param name="building">The building.</param>
        /// <param name="builder">The builder.</param>
        /// <returns>Table</returns>
        private Table BuildPremiseStructuresTable(CfRiskUnit building, DocumentBuilder builder)
        {
            Table table = builder.StartTable();

            // store original values
            double fntSize = builder.Font.Size;
            bool bold = builder.Font.Bold;
            double border = builder.CellFormat.Borders.LineWidth;
            ParagraphAlignment parAlign = builder.ParagraphFormat.Alignment;

            builder.Font.Size = 9;
            builder.Font.Bold = true;

            // populate headings
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.Write("Coverage Type");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.Write("Limit");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.Write("Coinsurance");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.Write("Valuation");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.Write("Rate");
            builder.InsertCell();
            builder.Write("Premium");
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.EndRow();

            builder.Font.Bold = false;
            foreach (CfCoverage coverage in building.Coverages)
            {
                // Coverage
                builder.InsertCell();
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                builder.Write(this.CheckStringForNulls(coverage.Description));

                // Limit
                builder.InsertCell();
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                builder.Write("$" + Convert.ToDouble(this.CheckStringForNulls(coverage.Limit, "00")).ToString("N0"));

                // CoInsurance
                builder.InsertCell();
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                if (coverage.Coinsurance == null)
                {
                    builder.Write("n/a");
                }
                else
                {
                    if (coverage.Coinsurance.IsNumeric())
                    {
                        builder.Write(this.CheckStringForNulls(coverage.Coinsurance, ".00") + "%");
                    }
                    else
                    {
                        builder.Write(coverage.Coinsurance);
                    }
                }

                // Valuation
                builder.InsertCell();
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                string valuation = this.CheckStringForNulls(coverage.Valuation);
                valuation = valuation == "AC" ? "ACV" : valuation;
                builder.Write(valuation);

                // Rate
                builder.InsertCell();
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

                var rate = coverage.AgentAdjustedRate == 0 ? coverage.UnderwriterAdjustedRate == 0 ? coverage.RateAmount : coverage.UnderwriterAdjustedRate : coverage.AgentAdjustedRate;
                builder.Write(Convert.ToDecimal(this.CheckStringForNulls(rate, "0.000")).ToString("N3"));

                // Premium
                builder.InsertCell();
                builder.Write("$" + Convert.ToDouble(this.CheckStringForNulls(coverage.Premium, "00")).ToString("N0"));
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

                builder.EndRow();
            }

            // reset to original vals
            builder.EndTable();
            builder.Font.Bold = bold;
            builder.ParagraphFormat.Alignment = parAlign;
            builder.Font.Size = fntSize;
            return table;
        }

        /// <summary>
        /// Builds the underlying coverage limit table.
        /// </summary>
        /// <param name="limits">The limits.</param>
        /// <param name="coverage">The coverage.</param>
        /// <param name="builder">The builder.</param>
        /// <returns>Table</returns>
        private Table BuildUnderlyingCoverageTable(List<Limit> limits, XsUnderlyingCoverage coverage, DocumentBuilder builder)
        {
            Table table = builder.StartTable();
            double fntSize = builder.Font.Size;
            bool bold = builder.Font.Bold;
            ParagraphAlignment parAlign = builder.ParagraphFormat.Alignment;

            // set first row
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(120);
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.Write("Carrier");
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.CellFormat.PreferredWidth = PreferredWidth.Auto;
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.Write(this.CheckStringForNulls(coverage.CarrierName));
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            builder.EndRow();

            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(120);
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.Write("Policy Period: ");
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.CellFormat.PreferredWidth = PreferredWidth.Auto;
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.Write(coverage.EffectiveDate.ToString("MM/dd/yyyy") + " to " + coverage.ExpirationDate.ToString("MM/dd/yyyy"));
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            builder.EndRow();

            // Coverage
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(120);
            if (limits.Count > 1)
            {
                builder.Write("Limits");
            }
            else
            {
                builder.Write("Limit");
            }

            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = PreferredWidth.Auto;
            bool firstLine = true;
            foreach (var limit in limits)
            {
                if (!firstLine)
                {
                    builder.InsertBreak(BreakType.LineBreak);
                }

                if (limit.Code.IsNumeric())
                {
                    builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(80);

                    if (!this.HasRenderedUnderlyingCoverageTextConditionally(limit, builder))
                    {
                        builder.Write("$" + Convert.ToDouble(this.CheckStringForNulls(limit.Code, "00")).ToString("N0"));
                    }

                    builder.CellFormat.Borders.Right.LineStyle = LineStyle.None;
                }
                else
                {
                    builder.CellFormat.HorizontalMerge = CellMerge.First;
                    string limitText;
                    if (this.limitTypeDict.TryGetValue(this.CheckStringForNulls(limit.Code, "na"), out limitText))
                    {
                        builder.Write(this.limitTypeDict[this.CheckStringForNulls(limit.Code, "na")]);
                        builder.CellFormat.Borders.Right.LineStyle = LineStyle.Single;
                    }
                    else
                    {
                        builder.Write(limit.Code);
                        builder.CellFormat.Borders.Right.LineStyle = LineStyle.None;
                    }
                }

                firstLine = false;
            }

            builder.InsertCell();
            builder.CellFormat.Borders.Left.LineStyle = LineStyle.None;
            builder.CellFormat.Borders.Right.LineStyle = LineStyle.Single;
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            builder.CellFormat.PreferredWidth = PreferredWidth.Auto;
            firstLine = true;
            foreach (var limit in limits)
            {
                if (!firstLine)
                {
                    builder.InsertBreak(BreakType.LineBreak);
                }

                if (limit.Code.IsNumeric() || !this.limitTypeDict.ContainsKey(limit.Code))
                {
                    builder.Write(this.CheckStringForNulls(limit.LimitTypeName, "na"));
                    builder.CellFormat.HorizontalMerge = CellMerge.None;
                }
                else
                {
                    builder.CellFormat.HorizontalMerge = CellMerge.Previous;
                    builder.CellFormat.Borders.Right.LineStyle = LineStyle.Single;
                }

                firstLine = false;
            }

            builder.EndRow();

            builder.EndTable();

            // reset to original vals
            builder.Font.Bold = bold;
            builder.ParagraphFormat.Alignment = parAlign;
            builder.Font.Size = fntSize;
            return table;
        }

        /// <summary>
        /// Render underlying coverage text conditionally
        /// </summary>
        /// <param name="limit">Limit object</param>
        /// <param name="builder">Document builder</param>
        /// <returns>True is text was written to Document builder else False.</returns>
        private bool HasRenderedUnderlyingCoverageTextConditionally(Limit limit, DocumentBuilder builder)
        {
            var result = false;

            if (limit.LimitTypeCode.Equals("EBL", StringComparison.InvariantCultureIgnoreCase) &&
                        limit.Code.Equals("1000000"))
            {
                builder.Write("$1,000,000 each employee/$1,000,000 aggregate");
                result = true;
            }
            else if (limit.LimitTypeCode.Equals("EBL", StringComparison.InvariantCultureIgnoreCase) &&
                        limit.Code.Equals("2000000"))
            {
                builder.Write("2,000,000 each employee/2,000,000 aggregate");
                result = true;
            }

            return result;
        }

        /// <summary>
        /// Builds the exposures.
        /// </summary>
        /// <param name="warranties">The warranties.</param>
        /// <param name="builder">The builder.</param>
        /// <returns>Table</returns>
        private Table BuildWarranties(IEnumerable<Warranty> warranties, DocumentBuilder builder)
        {
            Table table = builder.StartTable();
            double fntSize = builder.Font.Size;
            bool bold = builder.Font.Bold;
            ParagraphAlignment parAlign = builder.ParagraphFormat.Alignment;

            builder.Font.Size = 9;

            // override for heading
            builder.Font.Bold = true;

            // loop through columns and set headings
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(50);
            builder.Write("Symbol");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.PreferredWidth = PreferredWidth.Auto;
            builder.Write("Description");
            builder.EndRow();

            // reset to original vals
            builder.Font.Bold = false;
            builder.ParagraphFormat.Alignment = parAlign;
            foreach (Warranty warranty in warranties)
            {
                builder.InsertCell();
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(50);
                builder.Write(warranty.WarrantyType.ToString());
                builder.InsertCell();
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                builder.CellFormat.PreferredWidth = PreferredWidth.Auto;
                builder.Write(warranty.Text.ToString());
                builder.EndRow();
            }

            builder.EndTable();
            return table;
        }

        /// <summary>
        /// Builds the policy form table
        /// </summary>
        /// <param name="documents">Collection of documents</param>
        /// <param name="builder">Document builder</param>
        /// <returns>Table</returns>
        private Table BuildPolicyFormTable(Dictionary<dynamic, Dictionary<string, DecisionModel.Models.Policy.Document>> documents, DocumentBuilder builder)
        {
            var table = builder.StartTable();

            ////Table formatting
            builder.CellFormat.Borders.LineStyle = LineStyle.None;
            table.ClearBorders();
            builder.CellFormat.TopPadding = 0.2;
            builder.CellFormat.BottomPadding = 0.2;

            // Write the line of business documents out to the form
            foreach (var key in documents.Keys.OrderBy(o => o.Order))
            {
                builder.InsertCell();
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(230);
                builder.EndRow();
                builder.InsertCell();
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(230);
                builder.Bold = true;
                builder.Write(key.Title);
                builder.Bold = false;
                builder.EndRow();
                builder.InsertCell();
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(230);
                builder.EndRow();

                foreach (var document in ((IEnumerable<DecisionModel.Models.Policy.Document>)documents[key].Values).OrderBy(d => d.Order == 0 ? 99999 : d.Order).ThenBy(d => d.NormalizedNumber))
                {
                    ////Form number
                    builder.InsertCell();
                    builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                    builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(230);
                    builder.Write(document.DisplayNumber + " " + document.EditionDate.ToString("MM yy"));

                    ////Form title
                    builder.InsertCell();
                    builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                    builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(700);
                    builder.Write(document.Title);

                    builder.EndRow();
                }
            }

            builder.EndTable();
            return table;
        }
    }
}