using System.Diagnostics;
using System.Globalization;
using Microsoft.AnalysisServices.AdomdClient;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using Kapok.Report.Adomd.ExcelStyling;

namespace Kapok.Report.Adomd;

public static class AdomdExcelHelper
{
    // TODO: implement that the the ReadValue error is written into the excel sheet; use optimistic behavior/conversation to default value only when activated (e.g. with a parameter)

    /// <summary>
    /// Writes the CellSet to a worksheet table.
    /// 
    /// When the CellSet has on the column axis one tuple, a excel table is created;
    /// otherwise, a named address range is created.
    /// The CellSet can not have more than two axis.
    /// </summary>
    /// <param name="worksheet"></param>
    /// <param name="cellSet"></param>
    /// <param name="tableName"></param>
    /// <param name="tableStyle"></param>
    /// <param name="columnStart"></param>
    /// <param name="rowStart"></param>
    /// <param name="title">
    /// If a title is given it is added in the first column of the table
    /// with the same style as the default header.
    /// 
    /// When no row axis is given, a empty row axis will be added (to have the space for the first column)
    /// </param>
    /// <param name="columnDynamicCaptions"></param>
    /// <param name="rowDynamicCaptions"></param>
    /// <param name="createRowLevelGroups">
    /// If activated and members on row axis have a higher 'LevelDepth' than 1, they will be grouped by the excel grouping function.
    /// </param>
    /// <returns>
    /// Returns the size of the table on the worksheet in a tuple:
    /// (int noOfColumns, int noOfRows)
    /// </returns>
    public static (int, int) CellSetToWorksheet(
        ExcelWorksheet worksheet,
        CellSet cellSet,
        string tableName,
        ExcelContingencyTableStyle? tableStyle = null,
        int columnStart = 1, int rowStart = 1,
        string? title = null,
        List<AdomdAxisDynamicCaption>? columnDynamicCaptions = null,
        List<AdomdAxisDynamicCaption>? rowDynamicCaptions = null,
        bool createRowLevelGroups = false
    )
    {
        if (worksheet == null) throw new ArgumentNullException(nameof(worksheet));
        if (cellSet == null) throw new ArgumentNullException(nameof(cellSet));

        Axis? columnAxis = cellSet.Axes.Count > 0 ? cellSet.Axes[0] : null;
        Axis? rowAxis = cellSet.Axes.Count >= 2 ? cellSet.Axes[1] : null;
        var cells = cellSet.Cells;

        if (cellSet.Axes.Count > 2)
            throw new NotSupportedException("In excel we can not write more than two axis!");


        int noOfColumnLines = 0;
        if (columnAxis != null && columnAxis.Set.Tuples.Count > 0)
            noOfColumnLines = columnAxis.Set.Tuples[0].Members.Count;

        int noOfRowLines = 0;
        if (rowAxis != null && rowAxis.Set.Tuples.Count > 0)
            noOfRowLines = rowAxis.Set.Tuples[0].Members.Count;

        if (noOfRowLines == 0 && title != null)
            noOfRowLines = 1;

        int noOfColumns = columnAxis?.Set.Tuples.Count ?? 0;
        int noOfRows = rowAxis?.Set.Tuples.Count ?? 0;

        if (title != null)
        {
            if (noOfRowLines == 0)
                noOfRowLines = 1;

            // set table title
            worksheet.Cells[rowStart, columnStart].Value = title;
        }

        if (noOfColumnLines == 0)
        {
            // no column exist --> do nothing here ...
        }
        else if (noOfColumnLines == 1)
        {
            // add a excel table if possible
            var table = worksheet.Tables.Add(new ExcelAddressBase(
                rowStart,
                columnStart,
                rowStart + noOfColumnLines - 1 + noOfRows,
                columnStart + noOfRowLines - 1 + noOfColumns
            ), tableName);

            for (int n = (title == null ? 0 : 1); n < noOfRowLines; n++)
            {
                // make sure that the first columns where the rows are shown don't have an visible header,
                // so we create here a fake header
                worksheet.Cells[rowStart, columnStart + n].Value = new string(' ', 1 + n);
            }

            table.TableStyle = tableStyle == null
                ? TableStyles.Light1
                : TableStyles.None;

            if (noOfRowLines > 0)
                table.ShowFirstColumn = true;

            if (tableStyle?.ShowFilter != null)
                table.ShowFilter = tableStyle.ShowFilter.Value;
        }
        else
        {
            // here we where not able to add a table, so, we want to create at least a range for the data (including columns and rows!)
            var range = worksheet.Cells[
                rowStart,
                columnStart,
                rowStart + noOfColumnLines - 1 + noOfRows,
                columnStart + noOfRowLines - 1 + noOfColumns
            ];
            worksheet.Names.Add(tableName, range);

            if (tableStyle == null)
                tableStyle = ExcelContingencyTableStyle.Default;
        }

        // write all columns:

        var groupingStart = new List<int>();

        for (var cIndex = 0; cIndex < noOfColumns; cIndex++)
        {
#pragma warning disable 8602
            var column = columnAxis.Set.Tuples[cIndex];
#pragma warning restore 8602
            Debug.Assert(noOfColumnLines == column.Members.Count);

            // write dimensions on 'column' axis

            for (int c = 0; c < noOfColumnLines; c++)
            {
                var columnMember = column.Members[c];

                if (cIndex > 0 && c < noOfColumnLines - 1)
                {
                    var prevColumn = columnAxis.Set.Tuples[cIndex - 1];
                    if (prevColumn.Members[c].UniqueName == columnMember.UniqueName)
                    {
                        // has been merged -> do not write member caption
                        continue;
                    }
                }

                var cell = worksheet.Cells[
                    rowStart + c,
                    columnStart + noOfRowLines + cIndex
                ];

                var columnDynamicCaption = columnDynamicCaptions?.FirstOrDefault(adc => adc.ApplyOnTupleLevel == c + 1);
                if (columnDynamicCaption != null)
                {
                    if (columnMember.MemberProperties.Find(columnDynamicCaption.CaptionPropertyName) != null)
                    {
                        cell.Value = columnMember.MemberProperties[columnDynamicCaption.CaptionPropertyName].Value;
                    }
                    else
                    {
                        // Fallback, property does not exit
                        cell.Value = columnMember.Caption;
                    }
                }
                else
                {
                    cell.Value = columnMember.Caption;
                }

                // pre-look if the column should be merged
                if (c < noOfColumnLines - 1 && cIndex < noOfColumns - 1)
                {
                    int n;
                    for (n = cIndex; n < noOfColumns; n++)
                    {
                        var nextColumn = columnAxis.Set.Tuples[n];
                        if (nextColumn.Members[c].UniqueName != columnMember.UniqueName)
                            break;
                    }
                    n--;

                    if (n != cIndex)
                    {
                        // the column on a higher level is equal --> we merge the column cells in excel
                        var range = worksheet.Cells[
                            rowStart + c,
                            columnStart + noOfRowLines + cIndex,
                            rowStart + c,
                            columnStart + noOfRowLines + n
                        ];
                        range.Merge = true;
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }
                }
            }
        }

        groupingStart.Clear();

        // write all rows & cell data:

        for (int rIndex = 0; rIndex < noOfRows; rIndex++)
        {
#pragma warning disable 8602
            var row = rowAxis.Set.Tuples[rIndex];
#pragma warning restore 8602
            Debug.Assert(noOfRowLines == row.Members.Count);

            // write dimensions on 'row' axis

            for (int r = 0; r < noOfRowLines; r++)
            {
                var rowMember = row.Members[r];

                var cell = worksheet.Cells[
                    rowStart + noOfColumnLines + rIndex,
                    columnStart + r
                ];

                var rowDynamicCaption = rowDynamicCaptions?.FirstOrDefault(adc => adc.ApplyOnTupleLevel == r + 1);
                if (rowDynamicCaption != null)
                {
                    if (rowMember.MemberProperties.Find(rowDynamicCaption.CaptionPropertyName) != null)
                    {
                        cell.Value = rowMember.MemberProperties[rowDynamicCaption.CaptionPropertyName].Value;
                    }
                    else
                    {
                        // Fallback, property does not exit
                        cell.Value = rowMember.Caption;
                    }
                }
                else
                {
                    cell.Value = rowMember.Caption;
                }

                if (rowMember.LevelDepth > 1)
                {
                    cell.Value = new string(' ', 3 * rowMember.LevelDepth) + cell.Value;
                }

                if (createRowLevelGroups)
                {
                    while (groupingStart.Count + 1 < rowMember.LevelDepth)
                    {
                        groupingStart.Add(rIndex);
                    }

                    while (groupingStart.Count + 1 > rowMember.LevelDepth)
                    {
                        // add grouping
                        for (int gr = groupingStart[groupingStart.Count - 1]; gr < rIndex; gr++)
                        {
                            var wsRow = worksheet.Row(rowStart + noOfColumnLines + gr);
                            wsRow.OutlineLevel += 1;
                        }

                        // remove last
                        groupingStart.RemoveAt(groupingStart.Count - 1);
                    }
                }
            }

            // write cell data

            for (int cIndex = 0; cIndex < noOfColumns; cIndex++)
            {
                var cell = worksheet.Cells[
                    rowStart + noOfColumnLines + rIndex,
                    columnStart + noOfRowLines + cIndex
                ];
                var sourceCell = cells[cIndex, rIndex];

                WriteCell(cell, sourceCell);
            }
        }

        // run column with auto-fit over the whole table

        worksheet.Cells[
            rowStart,
            columnStart,
            rowStart + noOfColumnLines - 1 + noOfRows,
            columnStart + noOfRowLines - 1 + noOfColumns
        ].AutoFitColumns();

        // apply table style

        if (tableStyle != null)
        {
            if (tableStyle.ColumnAxisStyle != null && noOfColumns > 0)
            {
                var columns = worksheet.Cells[
                    rowStart,
                    columnStart + (title == null ? noOfRowLines : 0),
                    rowStart + noOfColumnLines - 1,
                    columnStart + noOfRowLines - 1 + noOfColumns
                ];
                tableStyle.ColumnAxisStyle.Apply(columns, worksheet);
            }

            if (tableStyle.RowAxisStyle != null && noOfRows > 0)
            {
                var rows = worksheet.Cells[
                    rowStart + noOfColumnLines,
                    columnStart,
                    rowStart + noOfColumnLines - 1 + noOfRows,
                    columnStart + noOfRowLines - 1
                ];
                tableStyle.RowAxisStyle.Apply(rows, worksheet);
            }

            if (tableStyle.CellAreaStyle != null && noOfColumns > 0 && noOfRows > 0)
            {
                var cellRange = worksheet.Cells[
                    rowStart + noOfColumnLines,
                    columnStart + noOfRowLines,
                    rowStart + noOfColumnLines - 1 + noOfRows,
                    columnStart + noOfRowLines - 1 + noOfColumns
                ];
                tableStyle.CellAreaStyle.Apply(cellRange, worksheet);
            }

            // retrieves a column grouping based on the defined tuple level
            int[] ReadTupleGrouping(TupleCollection tuples, int tupleLevel)
            {
                List<int> grouping = new List<int>();
                grouping.Add(0);
                string? lastUniqueName = null;
                foreach (var tuple in tuples)
                {
                    var uniqueName = tuple.Members[tupleLevel - 1].UniqueName;

                    if (lastUniqueName == null || uniqueName == lastUniqueName)
                    {
                        grouping[grouping.Count - 1]++;
                    }
                    else
                    {
                        grouping.Add(1);
                    }

                    lastUniqueName = uniqueName;
                }

                return grouping.ToArray();
            }

            void ExtendedGroupingApply(ExcelRange range, ExcelWorksheet w, ExcelAxisCellStyleTemplate axisCellStyle, TupleCollection tuples, int tupleLines)
            {
                bool applyOnAllTuples = false;
                int[]? grouping = null;

                if (axisCellStyle.ApplyOnTupleLevel == null)
                {
                    applyOnAllTuples = true;
                }
                else if (axisCellStyle.ApplyOnTupleLevel.Value == tupleLines)
                {
                    applyOnAllTuples = true;
                }
                else if (axisCellStyle.ApplyOnTupleLevel.Value < tupleLines)
                {
                    grouping = ReadTupleGrouping(tuples, axisCellStyle.ApplyOnTupleLevel.Value);
                }

                if (applyOnAllTuples || grouping != null)
                {
                    if (axisCellStyle.ApplyOnTuple == null || axisCellStyle.ApplyOnTuple.Length == 0)
                    {
                        axisCellStyle.Apply(range, w, applyOnAllTuples ? null : grouping);
                    }
                    else
                    {
                        for (var tupleIndex = 0; tupleIndex < tuples.Count; tupleIndex++)
                        {
                            var tuple = tuples[tupleIndex];
                            var membersCount = axisCellStyle.ApplyOnTuple.Length > tuple.Members.Count
                                ? axisCellStyle.ApplyOnTuple.Length
                                : tuple.Members.Count;

                            bool found = false;
                            for (int m = 0; m < membersCount; m++)
                            {
                                if (axisCellStyle.ApplyOnTuple.Length - 1 < m)
                                {
                                    found = true;
                                    break;
                                }

                                if (tuple.Members.Count - 1 < m)
                                {
                                    found = false;
                                    break;
                                }

                                if (tuple.Members[m].UniqueName == axisCellStyle.ApplyOnTuple[m])
                                {
                                    found = true;
                                }
                                else
                                {
                                    found = false;
                                    break;
                                }
                            }

                            if (found)
                            {
                                ExcelRange tupleRange;

                                if (axisCellStyle is ExcelColumnAxisCellStyleTemplate)
                                {
                                    tupleRange = w.Cells[
                                        range.Start.Row,
                                        range.Start.Column + tupleIndex,
                                        range.End.Row,
                                        range.Start.Column + tupleIndex
                                    ];

                                    if (axisCellStyle.DynamicCaptionFromTupleMemberAdomdProperty != null)
                                    {
                                        for (int i = 0;
                                             i < tupleRange.End.Row - tupleRange.Start.Row &&
                                             i < axisCellStyle.DynamicCaptionFromTupleMemberAdomdProperty.Length;
                                             i++)
                                        {
                                            var cell = w.Cells[tupleRange.Start.Row + i, tupleRange.Start.Column];
                                            cell.Value =
                                                axisCellStyle.DynamicCaptionFromTupleMemberAdomdProperty[i];
                                        }
                                    }
                                }
                                else if (axisCellStyle is ExcelRowAxisCellStyleTemplate)
                                {
                                    tupleRange = w.Cells[
                                        range.Start.Row + tupleIndex,
                                        range.Start.Column,
                                        range.Start.Row + tupleIndex,
                                        range.End.Column
                                    ];
                                }
                                else
                                {
                                    throw new NotSupportedException(
                                        "Expect a axis cell style of type ExcelColumnAxisCellStyleTemplate or ExcelRowAxisCellStyleTemplate");
                                }

                                axisCellStyle.Apply(tupleRange, w);
                            }
                        }
                    }
                }
            }

            if (tableStyle.ColumnAxisCellStyle != null && tableStyle.ColumnAxisCellStyle.Count > 0 && columnAxis != null &&
                noOfColumns > 0 && noOfRows > 0)
            {
                foreach (var stylingLevel in
                         from s in tableStyle.ColumnAxisCellStyle
                         group s by new {s.ApplyOnTupleLevel, s.ApplyOnTuple}
                         into sg
                         select new
                         {
                             Tuple = sg.Key.ApplyOnTuple,
                             TupleLevel = sg.Key.ApplyOnTupleLevel,
                             CellStyles = sg.ToArray()
                         })
                {
                    if (stylingLevel.CellStyles.Length == 1)
                    {
                        // apply to whole table
                        var axisCellStyle = stylingLevel.CellStyles[0];

                        var cellRange = worksheet.Cells[
                            rowStart + (axisCellStyle.ExtendStyleToHeader ? 0 : noOfColumnLines),
                            columnStart + noOfRowLines,
                            rowStart + noOfColumnLines - 1 + noOfRows,
                            columnStart + noOfRowLines - 1 + noOfColumns
                        ];

                        ExtendedGroupingApply(cellRange, worksheet,
                            axisCellStyle: axisCellStyle,
                            tuples: columnAxis.Set.Tuples,
                            tupleLines: noOfColumnLines);
                    }
                    else
                    {
                        // for-each each row, apply repeatingly
                        for (int cIndex = 0; cIndex < noOfColumns; cIndex++)
                        {
                            var axisCellStyle = stylingLevel.CellStyles[cIndex % stylingLevel.CellStyles.Length];

                            var rowCellRange = worksheet.Cells[
                                rowStart + (axisCellStyle.ExtendStyleToHeader ? 0 : noOfColumnLines),
                                columnStart + noOfRowLines + cIndex,
                                rowStart + noOfColumnLines - 1 + noOfRowLines,
                                columnStart + noOfRowLines + cIndex
                            ];

                            ExtendedGroupingApply(rowCellRange, worksheet,
                                axisCellStyle: axisCellStyle,
                                tuples: columnAxis.Set.Tuples,
                                tupleLines: noOfColumnLines);
                        }
                    }
                }
            }

            if (tableStyle.RowAxisCellStyle != null && tableStyle.RowAxisCellStyle.Count > 0 && rowAxis != null &&
                noOfColumns > 0 && noOfRows > 0)
            {
                foreach (var stylingLevel in
                         from s in tableStyle.RowAxisCellStyle
                         group s by new { s.ApplyOnTupleLevel, s.ApplyOnTuple }
                         into sg
                         select new
                         {
                             Tuple = sg.Key.ApplyOnTuple,
                             TupleLevel = sg.Key.ApplyOnTupleLevel,
                             CellStyles = sg.ToArray()
                         })
                {
                    if (stylingLevel.CellStyles.Length == 1)
                    {
                        // apply to whole table
                        var axisCellStyle = stylingLevel.CellStyles[0];
                            
                        var cellRange = worksheet.Cells[
                            rowStart + noOfColumnLines,
                            columnStart + (axisCellStyle.ExtendStyleToHeader ? 0 : noOfRowLines),
                            rowStart + noOfColumnLines - 1 + noOfRows,
                            columnStart + noOfRowLines - 1 + noOfColumns
                        ];

                        ExtendedGroupingApply(cellRange, worksheet,
                            axisCellStyle: axisCellStyle,
                            tuples: rowAxis.Set.Tuples,
                            tupleLines: noOfRowLines);
                    }
                    else
                    {
                        // for-each each row, apply repeatingly
                        for (int rIndex = 0; rIndex < noOfRows; rIndex++)
                        {
                            var axisCellStyle = stylingLevel.CellStyles[rIndex % stylingLevel.CellStyles.Length];

                            var rowCellRange = worksheet.Cells[
                                rowStart + noOfColumnLines + rIndex,
                                columnStart + (axisCellStyle.ExtendStyleToHeader ? 0 : noOfRowLines),
                                rowStart + noOfColumnLines + rIndex,
                                columnStart + noOfRowLines - 1 + noOfColumns
                            ];

                            ExtendedGroupingApply(rowCellRange, worksheet,
                                axisCellStyle: axisCellStyle,
                                tuples: rowAxis.Set.Tuples,
                                tupleLines: noOfColumnLines);
                        }
                    }
                }
            }

            if (tableStyle.SetFreezePane)
            {
                worksheet.View.FreezePanes(
                    rowStart + noOfColumnLines,
                    columnStart + noOfRowLines
                );
            }
        }

        return (
            noOfRowLines - 1 + noOfColumns,
            noOfColumnLines - 1 + noOfRows
        );
    }

    private static void WriteCell(ExcelRange cell, Cell sourceCell, CultureInfo? cultureInfo = null)
    {
        cultureInfo ??= CultureInfo.CurrentUICulture;

        cell.Value = sourceCell.ReadValue(errorValue: null);

        /*
         * NOTE: I do not have a sample to test this e.g. with action type MDACTION_TYPE_URL,
         * so, I will disable this here to make nothing crash one day
        var propertyActionType = sourceCell.CellProperties.Find("ACTION_TYPE");
        if (propertyActionType != null)
        {
            var value = propertyActionType.Value;
            if (value != null)
            {
                //throw new NotImplementedException();
            }
        }*/

        /*var propertyLanguage = sourceCell.CellProperties.Find("LANGUAGE");
        if (propertyLanguage != null)
        {
            // TBD, convert this to 'culture info' to be used to e.g. use different format translations (currency symbols, 'True/False' translation etc.);
            // we ignore this here because 
        }*/

        var propertyFormatString = sourceCell.CellProperties.Find("FORMAT_STRING");
        if (propertyFormatString != null)
        {
            var value = propertyFormatString.Value;
            if (value != null)
            {
                var format = value.ToString();

                // parse standard MDX formats to standard Excel formats
                switch (format)
                {
                    // Excel build-in formats:
                    // 0   General
                    // 1   0
                    // 2   0.00
                    // 3   #,##0
                    // 4   #,##0.00
                    // 9   0%
                    // 10  0.00%
                    // 11  0.00E+00
                    // 12  # ?/?
                    // 13  # ??/??
                    // 14  mm-dd-yy
                    // 15  d-mmm-yy
                    // 16  d-mmm
                    // 17  mmm-yy
                    // 18  h:mm AM/PM
                    // 19  h:mm:ss AM/PM
                    // 20  h:mm
                    // 21  h:mm:ss
                    // 22  m/d/yy h:mm
                    // 37  #,##0 ;(#,##0)
                    // 38  #,##0 ;[Red](#,##0)
                    // 39  #,##0.00;(#,##0.00)
                    // 40  #,##0.00;[Red](#,##0.00)
                    // 45  mm:ss
                    // 46  [h]:mm:ss
                    // 47  mmss.0
                    // 48  ##0.0E+0
                    // 49  @

                    // see as well: https://docs.microsoft.com/en-us/analysis-services/multidimensional-models/mdx/mdx-cell-properties-format-string-contents?view=asallproducts-allversions
                    case "General":
                    case "Number":
                        format = "General";  // Excel Build-in style 0
                        break;
                    case "Currency":
                        var ri = new RegionInfo(cultureInfo.LCID);
                        format = $"#.##0,00 {ri.ISOCurrencySymbol}";
                        break;
                    case "Fixed":
                        format = "0.00"; // Excel Build-in style 2
                        break;
                    case "Standard":
                        format = "#,##0.00"; // Excel Build-in style 4
                        break;
                    case "Percent":
                        format = "0.00%"; // Excel Build-in style 10
                        break;
                    case "Scientific":
                        format = "0.00E+00"; // Excel Build-in style 11
                        break;
                    case "Yes/No":
                        format = Resources.ExcelHelper.ResourceManager.GetString("FormatString_YesNo", cultureInfo);
                        break;
                    case "True/False":
                        format = Resources.ExcelHelper.ResourceManager.GetString("FormatString_TrueFalse", cultureInfo);
                        break;
                    case "On/Off":
                        format = Resources.ExcelHelper.ResourceManager.GetString("FormatString_OnOff", cultureInfo);
                        break;
                }

                if (format != null)
                    cell.Style.Numberformat.Format = format;
            }
        }
        else // NOTE: this should only happen when mode "pretty" is activated! if not, we guess the user prefers to get the 'real data' instead of a nice value in a string
        {
            var propertyFormattedValue = sourceCell.CellProperties.Find("FORMATTED_VALUE");
            if (propertyFormattedValue != null)
            {
                // NOTE: when "FORMAT_STRING" is not give, but "FORMATTED_VALUE" is given, we assume that the user wants to see the 
                cell.Value = sourceCell.FormattedValue; // this is a shortcut to propertyFormattedValue.Value
            }
        }

        var propertyBackColor = sourceCell.CellProperties.Find("BACK_COLOR");
        if (propertyBackColor != null)
        {
            var value = propertyBackColor.Value;
            if (value != null)
            {
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(MdxParsing.ColorFromRgbInteger((uint)value));
            }
        }

        var propertyFontFlag = sourceCell.CellProperties.Find("FONT_FLAGS");
        if (propertyFontFlag != null)
        {
            var value = propertyFontFlag.Value;
            if (value != null)
            {
                bool isBold = ((int)value & MdxParsing.MDFF_BOLD) == MdxParsing.MDFF_BOLD;
                bool isItalic = ((int)value & MdxParsing.MDFF_ITALIC) == MdxParsing.MDFF_ITALIC;
                bool isUnderline = ((int)value & MdxParsing.MDFF_UNDERLINE) == MdxParsing.MDFF_UNDERLINE;
                bool isStrikeout = ((int)value & MdxParsing.MDFF_STRIKEOUT) == MdxParsing.MDFF_STRIKEOUT;

                if (isBold)
                    cell.Style.Font.Bold = true;
                if (isItalic)
                    cell.Style.Font.Italic = true;
                if (isUnderline)
                {
                    cell.Style.Font.UnderLine = true;
                    cell.Style.Font.UnderLineType = ExcelUnderLineType.Single;
                }
                if (isStrikeout)
                    cell.Style.Font.Strike = true;
            }
        }

        var propertyFontSize = sourceCell.CellProperties.Find("FONT_SIZE");
        if (propertyFontSize != null)
        {
            var value = propertyFontSize.Value;
            if (value != null)
            {
                cell.Style.Font.Size = (ushort)value;
            }
        }

        var propertyName = sourceCell.CellProperties.Find("FONT_NAME");
        if (propertyName != null)
        {
            var value = propertyName.Value;
            if (value != null)
            {
                cell.Style.Font.Name = value.ToString();
            }
        }

        var propertyForeColor = sourceCell.CellProperties.Find("FORE_COLOR");
        if (propertyForeColor != null)
        {
            var value = propertyForeColor.Value;
            if (value != null)
            {
                cell.Style.Font.Color.SetColor(MdxParsing.ColorFromRgbInteger((uint)value));
            }
        }
    }
}