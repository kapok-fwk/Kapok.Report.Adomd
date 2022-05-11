using System.Data;
using System.Diagnostics;
using Microsoft.AnalysisServices.AdomdClient;

namespace Kapok.Report.Adomd;

public static class AdomdCellSetExtension
{
    /// <summary>
    /// Converts the CellSet into a DataTable.
    ///
    /// The CellSet must have at least one axis with one tuple.
    /// The CellSet can have a second axis which is then written as first columns.
    /// </summary>
    /// <param name="cellSet"></param>
    /// <param name="tableName"></param>
    /// <returns>
    /// Returns a new DataTable object or null when the CellSet is empty (e.g. it has no axis).
    /// </returns>
    public static DataTable? ToDataTable(this CellSet cellSet, string tableName)
    {
        if (cellSet == null) throw new ArgumentNullException(nameof(cellSet));
            
        if (cellSet.Axes.Count == 0)
            return null; // no data in the query

        Axis columnAxis = cellSet.Axes[0];
        Axis? rowAxis = cellSet.Axes.Count >= 2 ? cellSet.Axes[1] : null;
        var cells = cellSet.Cells;
            
        if (cellSet.Axes.Count > 2)
            throw new ArgumentException("The CellSet has more than two axis.", nameof(cellSet));

        int noOfColumnLines = 0;
        if (columnAxis != null && columnAxis.Set.Tuples.Count > 0)
            noOfColumnLines = columnAxis.Set.Tuples[0].Members.Count;

        int noOfRowLines = 0;
        if (rowAxis != null && rowAxis.Set.Tuples.Count > 0)
            noOfRowLines = rowAxis.Set.Tuples[0].Members.Count;

        int noOfColumns = columnAxis?.Set.Tuples.Count ?? 0;
        int noOfRows = rowAxis?.Set.Tuples.Count ?? 0;

        if (noOfColumnLines > 1)
            throw new ArgumentException("The CellSet has more than two tuples in the column axis. The CellSet can not be converted to a DataTable.", nameof(cellSet));

        var dataTable = new DataTable(tableName);

        for (int n = 0; n < noOfRowLines; n++)
        {
            // make sure that the first columns where the rows are shown don't have an visible header,
            // so we create here a fake header
            dataTable.Columns.Add(new string(' ', 1 + n), typeof(string));
        }

        // write all columns:

        for (var cIndex = 0; cIndex < noOfColumns; cIndex++)
        {
#pragma warning disable 8602
            var column = columnAxis.Set.Tuples[cIndex];
#pragma warning restore 8602
            Debug.Assert(noOfColumnLines == column.Members.Count);

            // write dimensions on 'column' axis

            for (int c = 0; c < noOfColumnLines; c++) // NOTE: this will currently only run through once, see exception catch above
            {
                Type type = typeof(object);

                if (column.Members.Count == 1)
                {
                    // get the object type from the first line with a value
                    for (int rIndex = 0; rIndex < noOfRows; rIndex++)
                    {
                        object? value = cells[cIndex, rIndex].ReadValue(errorValue: null);
                        if (value != null)
                        {
                            type = value.GetType();
                            break;
                        }
                    }
                }

                dataTable.Columns.Add(column.Members[noOfColumnLines].Caption, type);
            }
        }

        // write all rows & cell data:

        for (int rIndex = 0; rIndex < noOfRows; rIndex++)
        {
#pragma warning disable 8602
            var row = rowAxis.Set.Tuples[rIndex];
#pragma warning restore 8602
            Debug.Assert(noOfRowLines == row.Members.Count);

            var newRow = dataTable.NewRow();

            // write dimensions on 'row' axis

            for (int r = 0; r < noOfRowLines; r++)
            {
                newRow[r] = row.Members[r].Caption;
            }

            // write cell data

            for (int cIndex = 0; cIndex < (columnAxis?.Set.Tuples.Count ?? 0); cIndex++)
            {
                newRow[noOfRowLines + cIndex] = cells[cIndex, rIndex].ReadValue(errorValue: DBNull.Value);
            }

            dataTable.Rows.Add(newRow);
        }

        return dataTable;
    }
}