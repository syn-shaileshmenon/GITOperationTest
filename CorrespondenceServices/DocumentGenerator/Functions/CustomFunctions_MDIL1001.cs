// <copyright file="CustomFunctions_MDIL1001.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>

namespace Mkl.WebTeam.DocumentGenerator.Functions
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using DecisionModel.Models.Policy;
    using SubmissionShared.Enumerations;

    /// <summary>
    /// The custom functions class
    /// </summary>
    public partial class CustomFunctions
    {
        /// <summary>
        /// Generates the MDI L1001 0811 form.
        /// </summary>
        /// <param name="policy">The policy.</param>
        /// <param name="documentManager">The document manager.</param>
        /// <param name="extraParameters">The extra parameters.</param>
        public void GenerateMDIL1001_0811Form(Policy policy, ref IPolicyDocumentManager documentManager, Dictionary<string, object> extraParameters)
        {
            var formsByLob = new Dictionary<dynamic, Dictionary<string, Document>>();
            var commonDocuments = policy.RollUp.Documents.Where(d => d.IsCommonForm() && d.ReturnIsSelected() && d.Section != DocumentSection.Application).ToList();
            var commonFormsDictionary = new Dictionary<string, Document>();
            commonDocuments.ForEach(d => commonFormsDictionary.Add(GetDocumentKey(d), d));
            dynamic lobKey = new LobKey();
            lobKey.Title = "COMMON";
            lobKey.Order = 1;
            formsByLob.Add(lobKey, commonFormsDictionary);

            // Group and dedupe the forms within each line of business.
            // This is required to pull GL, OCP, liquor, and special events into one list.
            foreach (var lob in policy.LobOrder)
            {
                lobKey = new LobKey();
                switch (lob.Code)
                {
                    case LineOfBusiness.CF:
                        lobKey.Title = "PROPERTY";
                        lobKey.Order = 3;
                        formsByLob.Add(lobKey, new Dictionary<string, Document>());
                        break;
                    case LineOfBusiness.IM:
                        lobKey.Title = "INLAND MARINE";
                        lobKey.Order = 4;
                        formsByLob.Add(lobKey, new Dictionary<string, Document>());
                        break;
                    case LineOfBusiness.XS:
                        lobKey.Title = "EXCESS LIABILITY";
                        lobKey.Order = 5;
                        formsByLob.Add(lobKey, new Dictionary<string, Document>());
                        break;
                    default:
                        lobKey.Title = "GENERAL LIABILITY";
                        lobKey.Order = 2;
                        if (!formsByLob.ContainsKey(lobKey))
                        {
                            formsByLob.Add(lobKey, new Dictionary<string, Document>());
                        }

                        break;
                }

                foreach (var document in policy.GetLine(lob.Code).RollUp.Documents.Where(d => !d.IsCommonForm() && d.ReturnIsSelected() && d.Section != DocumentSection.Application))
                {
                    var key = GetDocumentKey(document);
                    if (!formsByLob[lobKey].ContainsKey(key))
                    {
                        formsByLob[lobKey].Add(key, document);
                    }
                }
            }

            documentManager.InsertPolicyFormTable(formsByLob, documentManager.GetBookMarkByName("PolicyForm"));
        }

        /// <summary>
        /// Gets the document key.
        /// </summary>
        /// <param name="document">The document.</param>
        /// <returns>The key</returns>
        private static string GetDocumentKey(Document document)
        {
            return document.NormalizedNumber + document.EditionDate.ToString("MMyy") + document.QuantityOrder;
        }

        /// <summary>
        /// The line of business key
        /// </summary>
        /// <seealso cref="System.IComparable" />
        public class LobKey : IComparable
        {
            /// <summary>
            /// Gets or sets the title.
            /// </summary>
            public string Title { get; set; }

            /// <summary>
            /// Gets or sets the order.
            /// </summary>
            public int Order { get; set; }

            /// <summary>
            /// Compares the current instance with another object of the same type and returns an integer that indicates whether the current instance precedes, follows, or occurs in the same position in the sort order as the other object.
            /// </summary>
            /// <param name="obj">An object to compare with this instance.</param>
            /// <returns>
            /// A value that indicates the relative order of the objects being compared. The return value has these meanings: Value Meaning Less than zero This instance precedes <paramref name="obj" /> in the sort order. Zero This instance occurs in the same position in the sort order as <paramref name="obj" />. Greater than zero This instance follows <paramref name="obj" /> in the sort order.
            /// </returns>
            public int CompareTo(object obj)
            {
                if (obj == null)
                {
                    return -1;
                }

                if (((LobKey)obj).Title.Equals(this.Title))
                {
                    return 0;
                }

                return -1;
            }

            /// <summary>
            /// Returns a hash code for this instance.
            /// </summary>
            /// <returns>
            /// A hash code for this instance, suitable for use in hashing algorithms and data structures like a hash table.
            /// </returns>
            public override int GetHashCode()
            {
                return this.Title != null ? this.Title.GetHashCode() : base.GetHashCode();
            }

            /// <summary>
            /// Determines whether the specified <see cref="System.Object" />, is equal to this instance.
            /// </summary>
            /// <param name="obj">The <see cref="System.Object" /> to compare with this instance.</param>
            /// <returns>
            ///   <c>true</c> if the specified <see cref="System.Object" /> is equal to this instance; otherwise, <c>false</c>.
            /// </returns>
            public override bool Equals(object obj)
            {
                return this.CompareTo(obj) == 0;
            }
        }
    }
}
