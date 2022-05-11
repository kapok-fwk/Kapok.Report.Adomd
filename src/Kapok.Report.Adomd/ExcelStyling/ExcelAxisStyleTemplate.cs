using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Kapok.Report.Adomd.ExcelStyling;

public class ExcelAxisStyleTemplate : ICloneable
{
    public bool? FontBold { get; set; }
    public Color[]? FontColors { get; set; }
    public ExcelBorderRangeTemplate? Border { get; set; }
    public Color[]? BackgroundColors { get; set; }
    public ExcelHorizontalAlignment? HorizontalAlignment { get; set; }
    public ExcelVerticalAlignment? VerticalAlignment { get; set; }

    public void Apply(ExcelRange range, ExcelWorksheet worksheet)
    {
        // apply style which applies to the whole table
        Apply(range.Style);

        // apply style which applies to each column line
        for (int cIndex = 0; cIndex < range.Columns; cIndex++)
        {
            var lineRange = worksheet.Cells[
                range.Start.Row,
                range.Start.Column,
                range.End.Row,
                range.Start.Column + cIndex
            ];

            Apply(lineRange.Style, cIndex);
        }

        Border?.Apply(range, worksheet);
    }

    private void Apply(ExcelStyle excelStyle)
    {
        if (FontBold.HasValue)
            excelStyle.Font.Bold = FontBold.Value;

        if (HorizontalAlignment.HasValue)
            excelStyle.HorizontalAlignment = HorizontalAlignment.Value;
        if (VerticalAlignment.HasValue)
            excelStyle.VerticalAlignment = VerticalAlignment.Value;
    }

    private void Apply(ExcelStyle excelStyle, int line)
    {
        if (BackgroundColors != null && BackgroundColors.Length > 0)
        {
            excelStyle.Fill.PatternType = ExcelFillStyle.Solid;
            excelStyle.Fill.BackgroundColor.SetColor(BackgroundColors[line % BackgroundColors.Length]);
        }

        if (FontColors != null && FontColors.Length > 0)
            excelStyle.Font.Color.SetColor(FontColors[line % FontColors.Length]);
    }

    public virtual object Clone()
    {
        var newObject = (ExcelAxisStyleTemplate)MemberwiseClone();
        if (newObject.Border != null)
            newObject.Border = (ExcelBorderRangeTemplate)newObject.Border.Clone();
        return newObject;
    }
}