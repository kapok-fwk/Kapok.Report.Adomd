using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Text;
using Microsoft.AnalysisServices.AdomdClient;

namespace Kapok.Report.Adomd;

public static class AdomdHtmlHelper
{
    public static string ToHtml(this AdomdReportDataSet dataSet)
    {
        if (dataSet == null) throw new ArgumentNullException(nameof(dataSet));

        if (dataSet.CellSet == null)
            throw new ArgumentException($"The ADOMD data set must be executed before calling extension method {nameof(ToHtml)}", nameof(dataSet));

        using var textWriter = new StringWriter();
        CellSetToHtmlTable(textWriter, cellSet: dataSet.CellSet);

        return textWriter.ToString();
    }

    public static void CellSetToHtmlTable(
        TextWriter htmlWriter,
        CellSet cellSet,
        string? title = null,
        List<AdomdAxisDynamicCaption>? columnDynamicCaptions = null,
        List<AdomdAxisDynamicCaption>? rowDynamicCaptions = null)
    {
        Axis? columnAxis = cellSet.Axes.Count > 0 ? cellSet.Axes[0] : null;
        Axis? rowAxis = cellSet.Axes.Count >= 2 ? cellSet.Axes[1] : null;
        var cells = cellSet.Cells;

        if (cellSet.Axes.Count > 2)
            throw new NotSupportedException("In html we can not write more than two axis!");

        int noOfColumnLines = 0;
        if (columnAxis != null && columnAxis.Set.Tuples.Count > 0)
            noOfColumnLines = columnAxis.Set.Tuples[0].Members.Count;

        int noOfRowLines = 0;
        if (rowAxis != null && rowAxis.Set.Tuples.Count > 0)
            noOfRowLines = rowAxis.Set.Tuples[0].Members.Count;

        int noOfColumns = columnAxis?.Set.Tuples.Count ?? 0;
        int noOfRows = rowAxis?.Set.Tuples.Count ?? 1;

        htmlWriter.Write("<table>");

        if (title != null)
        {
            htmlWriter.Write($"<caption>{title}</caption>");
        }

        htmlWriter.Write("<thead>");

        // write all columns:

        var groupingStart = new List<int>();

        for (int c = 0; c < noOfColumnLines; c++) 
        {
            htmlWriter.Write("<tr>");

            // add empty 'th' tags for row column lines
            if (noOfRowLines > 0)
            {
                htmlWriter.Write(new StringBuilder().Insert(0, "<th>&nbsp;</th>", noOfRowLines));
            }

            // write dimensions on 'column' axis

            for (var cIndex = 0; cIndex < noOfColumns; cIndex++)
            {
#pragma warning disable 8602
                var column = columnAxis.Set.Tuples[cIndex];
#pragma warning restore 8602
                Debug.Assert(noOfColumnLines == column.Members.Count);

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

                int columnSpan = 0;

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

                        columnSpan = n - cIndex + 1;
                    }
                }

                if (columnSpan > 0)
                {
                    htmlWriter.Write($"<th colspan=\"{columnSpan}\">");
                }
                else
                {
                    htmlWriter.Write("<th>");
                }

                string headerName;

                var columnDynamicCaption = columnDynamicCaptions?.FirstOrDefault(adc => adc.ApplyOnTupleLevel == c + 1);
                if (columnDynamicCaption != null)
                {
                    if (columnMember.Properties.Find(columnDynamicCaption.CaptionPropertyName) != null)
                    {
                        headerName = columnMember.Properties[columnDynamicCaption.CaptionPropertyName].ToString();
                    }
                    else
                    {
                        // Fallback, property does not exit
                        headerName = columnMember.Caption;
                    }
                }
                else
                {
                    headerName = columnMember.Caption;
                }

                htmlWriter.Write(headerName);

                htmlWriter.Write("</th>");
            }

            htmlWriter.Write("</tr>");
        }

        groupingStart.Clear();

        htmlWriter.Write("</thead>");
        htmlWriter.Write("<tbody>");

        // write all rows & cell data:

        for (int rIndex = 0; rIndex < noOfRows; rIndex++)
        {
            htmlWriter.Write("<tr>");

            if (rowAxis != null)
            {

                var row = rowAxis.Set.Tuples[rIndex];
                Debug.Assert(noOfRowLines == row?.Members.Count);

                // write dimensions on 'row' axis

                for (int r = 0; r < noOfRowLines; r++)
                {
                    htmlWriter.Write("<th>");

                    var rowMember = row.Members[r];

                    string headerName;

                    var rowDynamicCaption = rowDynamicCaptions?.FirstOrDefault(adc => adc.ApplyOnTupleLevel == r + 1);
                    if (rowDynamicCaption != null)
                    {
                        if (rowMember.Properties.Find(rowDynamicCaption.CaptionPropertyName) != null)
                        {
                            headerName = rowMember.Properties[rowDynamicCaption.CaptionPropertyName].ToString();
                        }
                        else
                        {
                            // Fallback, property does not exit
                            headerName = rowMember.Caption;
                        }
                    }
                    else
                    {
                        headerName = rowMember.Caption;
                    }

                    if (row.Members[r].LevelDepth > 1)
                    {
                        // NOTE: we use here '&ensp;', which is equal to two times space
                        const string levelIndentChar = "&ensp;";

                        // here we repeat the level indent per level depth at the beginning of the string
                        headerName = new StringBuilder(rowMember.LevelDepth + headerName.Length)
                            .Insert(0, levelIndentChar, rowMember.LevelDepth)
                            .Append(headerName)
                            .ToString();
                    }

                    htmlWriter.Write(headerName);

                    /*if (createRowLevelGroups)
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
                    }*/

                    htmlWriter.Write("</th>");
                }
            }

            // write cell data

            for (int cIndex = 0; cIndex < noOfColumns; cIndex++)
            {
                htmlWriter.Write("<td>");

                var sourceCell = rowAxis == null ? cells[cIndex] : cells[cIndex, rIndex];

                htmlWriter.Write(
                    WriteCell(sourceCell)
                );

                htmlWriter.Write("</td>");
            }

            htmlWriter.Write("</tr>");
        }

        htmlWriter.Write("</tbody>");

        htmlWriter.Write("</table>");
    }

    private static string WriteCell(Cell sourceCell, CultureInfo? cultureInfo = null)
    {
        cultureInfo ??= CultureInfo.CurrentUICulture;

        var cellValue = sourceCell.ReadValue(errorValue: null);

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

        string? valueAsHtml = null;

        var propertyFormatString = sourceCell.CellProperties.Find("FORMAT_STRING");
        if (propertyFormatString != null && cellValue != null)
        {
            var value = propertyFormatString.Value;
            if (value != null)
            {
                var cellValueType = cellValue.GetType();
                var format = value.ToString();

                // parse standard MDX formats to standard Excel formats
                switch (format)
                {
                    // see as well: https://docs.microsoft.com/en-us/analysis-services/multidimensional-models/mdx/mdx-cell-properties-format-string-contents?view=asallproducts-allversions
                    case "General":
                    case "Number":
                        if (cellValueType == typeof(int))
                        {
                            valueAsHtml = ((int) cellValue).ToString("", cultureInfo);
                        }
                        else if (cellValueType == typeof(float))
                        {
                            valueAsHtml = ((float)cellValue).ToString("###0.", cultureInfo);
                        }
                        else
                        {
                            // TODO: implement optimistic behavior here
                            throw new NotSupportedException($"Unexpected cell value type: {cellValueType}");
                        }
                        break;
                    case "Currency":
                        if (cellValueType == typeof(int))
                        {
                            valueAsHtml = ((int)cellValue).ToString("C", cultureInfo);
                        }
                        else if (cellValueType == typeof(float))
                        {
                            valueAsHtml = ((float)cellValue).ToString("C", cultureInfo);
                        }
                        else
                        {
                            // TODO: implement optimistic behavior here
                            throw new NotSupportedException($"Unexpected cell value type: {cellValueType}");
                        }
                        break;
                    case "Fixed":
                        if (cellValueType == typeof(int))
                        {
                            valueAsHtml = ((int)cellValue).ToString("0.00", cultureInfo);
                        }
                        else if (cellValueType == typeof(float))
                        {
                            valueAsHtml = ((float)cellValue).ToString("0.00", cultureInfo);
                        }
                        else
                        {
                            // TODO: implement optimistic behavior here
                            throw new NotSupportedException($"Unexpected cell value type: {cellValueType}");
                        }
                        break;
                    case "Standard":
                        if (cellValueType == typeof(int))
                        {
                            valueAsHtml = ((int)cellValue).ToString("#,##0.00", cultureInfo);
                        }
                        else if (cellValueType == typeof(float))
                        {
                            valueAsHtml = ((float)cellValue).ToString("#,##0.00", cultureInfo);
                        }
                        else
                        {
                            // TODO: implement optimistic behavior here
                            throw new NotSupportedException($"Unexpected cell value type: {cellValueType}");
                        }
                        break;
                    case "Percent":
                        if (cellValueType == typeof(int))
                        {
                            valueAsHtml = ((int)cellValue).ToString("P", cultureInfo);
                        }
                        else if (cellValueType == typeof(float))
                        {
                            valueAsHtml = ((float)cellValue).ToString("P", cultureInfo);
                        }
                        else
                        {
                            // TODO: implement optimistic behavior here
                            throw new NotSupportedException($"Unexpected cell value type: {cellValueType}");
                        }
                        break;
                    case "Scientific":
                        if (cellValueType == typeof(int))
                        {
                            valueAsHtml = ((int)cellValue).ToString("E", cultureInfo);
                        }
                        else if (cellValueType == typeof(float))
                        {
                            valueAsHtml = ((float)cellValue).ToString("E", cultureInfo);
                        }
                        else
                        {
                            // TODO: implement optimistic behavior here
                            throw new NotSupportedException($"Unexpected cell value type: {cellValueType}");
                        }
                        break;
                    case "Yes/No":
                    case "True/False":
                    case "On/Off":
                        bool? valueAsBoolean;
                        if (cellValueType == typeof(bool))
                        {
                            valueAsBoolean = (bool) cellValue;
                        }
                        else if (cellValueType == typeof(int))
                        {
                            valueAsBoolean = (int) cellValue != 0;
                        }
                        else if (cellValueType == typeof(float))
                        {
                            valueAsBoolean = (float) cellValue != 0;
                        }
                        else
                        {
                            // TODO: implement optimistic behavior here
                            throw new NotSupportedException($"Unexpected cell value type: {cellValueType}");
                        }

                        if (!valueAsBoolean.HasValue)
                        {
                            valueAsHtml = string.Empty;
                        }
                        else if (valueAsBoolean.Value == true)
                        {
                            string resourceKey;
                            switch (format)
                            {
                                case "Yes/No":
                                    resourceKey = "FormatString_YesNo_Yes";
                                    break;
                                case "True/False":
                                    resourceKey = "FormatString_TrueFalse_True";
                                    break;
                                case "On/Off":
                                    resourceKey = "FormatString_OnOff_On";
                                    break;
                                default:
                                    throw new NotSupportedException(); // internal programming error
                            }

                            valueAsHtml = Resources.HtmlHelper.ResourceManager.GetString(resourceKey, cultureInfo);
                        }
                        else
                        {
                            string resourceKey;
                            switch (format)
                            {
                                case "Yes/No":
                                    resourceKey = "FormatString_YesNo_No";
                                    break;
                                case "True/False":
                                    resourceKey = "FormatString_TrueFalse_False";
                                    break;
                                case "On/Off":
                                    resourceKey = "FormatString_OnOff_Off";
                                    break;
                                default:
                                    throw new NotSupportedException(); // internal programming error
                            }

                            valueAsHtml = Resources.HtmlHelper.ResourceManager.GetString(resourceKey, cultureInfo);
                        }
                        break;
                    default:
                        throw new NotImplementedException("TODO transform use custom format not implemented yet");
                }
            }
        }
        else // NOTE: this should only happen when mode "pretty" is activated! if not, we guess the user prefers to get the 'real data' instead of a nice value in a string
        {
            var propertyFormattedValue = sourceCell.CellProperties.Find("FORMATTED_VALUE");
            if (propertyFormattedValue != null)
            {
                // NOTE: when "FORMAT_STRING" is not give, but "FORMATTED_VALUE" is given, we assume that the user wants to see the 
                valueAsHtml = sourceCell.FormattedValue; // this is a shortcut to propertyFormattedValue.Value
            }
        }

        valueAsHtml ??= cellValue?.ToString() ?? string.Empty;

        var styleString = new StringBuilder();

        var propertyBackColor = sourceCell.CellProperties.Find("BACK_COLOR");
        if (propertyBackColor != null)
        {
            var value = propertyBackColor.Value;
            if (value != null)
            {
                if (styleString.Length > 0)
                    styleString.Append(';');

                styleString.Append("background-color: ");
                styleString.Append(ColorTranslator.ToHtml(MdxParsing.ColorFromRgbInteger((uint)value)));
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
                {
                    if (styleString.Length > 0)
                        styleString.Append(';');

                    styleString.Append("font-weight: bold");
                }

                if (isItalic)
                {
                    if (styleString.Length > 0)
                        styleString.Append(';');

                    styleString.Append("font-style: italic");
                }

                if (isUnderline || isStrikeout)
                {
                    if (styleString.Length > 0)
                        styleString.Append(';');

                    styleString.Append("text-decoration:");

                    if (isUnderline)
                        styleString.Append(" underline");
                    if (isStrikeout)
                        styleString.Append(" line-through");
                }
            }
        }

        var propertyFontSize = sourceCell.CellProperties.Find("FONT_SIZE");
        if (propertyFontSize != null)
        {
            var value = propertyFontSize.Value;
            if (value != null)
            {
                if (styleString.Length > 0)
                    styleString.Append(';');

                styleString.Append("background-color: ");
                styleString.Append((ushort)value);
                styleString.Append("px");
            }
        }

        var propertyName = sourceCell.CellProperties.Find("FONT_NAME");
        if (propertyName != null)
        {
            var value = propertyName.Value;
            if (value != null)
            {
                string fontName = value.ToString();

                if (fontName.Contains(' '))
                    fontName = $"\"{fontName}\"";

                if (styleString.Length > 0)
                    styleString.Append(';');

                styleString.Append("background-color: ");
                styleString.Append(fontName);
            }
        }

        var propertyForeColor = sourceCell.CellProperties.Find("FORE_COLOR");
        if (propertyForeColor != null)
        {
            var value = propertyForeColor.Value;
            if (value != null)
            {
                if (styleString.Length > 0)
                    styleString.Append(';');

                styleString.Append("color: ");
                styleString.Append(ColorTranslator.ToHtml(MdxParsing.ColorFromRgbInteger((uint)value)));
            }
        }

        return styleString.Length > 0
            ? $"<span style=\"{styleString}\">{valueAsHtml}</span>"
            : valueAsHtml;
    }
}