// ***********************************************************************
// Assembly         : Mkl.WebTeam.DocumentGenerator
// Author           : rsteelea
// Created          : 11-08-2016
//
// Last Modified By : rsteelea
// Last Modified On : 11-08-2016
// ***********************************************************************
// <copyright file="TableHelper.cs" company="Markel">
// Copyright (c) Markel. All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************
namespace Mkl.WebTeam.DocumentGenerator.Helpers
{
    using Aspose.Words;
    using Aspose.Words.Tables;

    /// <summary>
    /// Class TableHelper.
    /// </summary>
    public static class TableHelper
    {
        /// <summary>
        /// Gets the parent table.
        /// </summary>
        /// <param name="currentNode">The current node.</param>
        /// <returns>Table.</returns>
        public static Table GetParentTable(this Node currentNode)
        {
            while (!(currentNode is Document))
            {
                var cell = currentNode as Cell;
                if (cell != null)
                {
                    return cell.ParentRow.ParentTable;
                }

                var row = currentNode as Row;
                if (row != null)
                {
                    return row.ParentTable;
                }

                currentNode = currentNode.ParentNode;
            }

            return null;
        }

        /// <summary>
        /// Gets the current row.
        /// </summary>
        /// <param name="currentNode">The current node.</param>
        /// <returns>Row.</returns>
        public static Row GetCurrentRow(this Node currentNode)
        {
            while (!(currentNode is Document))
            {
                var cell = currentNode as Cell;
                if (cell != null)
                {
                    return cell.ParentRow;
                }

                var row = currentNode as Row;
                if (row != null)
                {
                    return row;
                }

                currentNode = currentNode.ParentNode;
            }

            return null;
        }
    }
}
