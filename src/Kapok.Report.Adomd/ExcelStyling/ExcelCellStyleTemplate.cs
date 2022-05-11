using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Kapok.Report.Adomd.ExcelStyling;

public class ExcelCellStyleTemplate : ICloneable
{
    public bool? FontBold { get; set; }
    public Color? BackgroundColor { get; set; }
    public ExcelHorizontalAlignment? HorizontalAlignment { get; set; }
    public ExcelVerticalAlignment? VerticalAlignment { get; set; }
    public ExcelBorderRangeTemplate? Border { get; set; }
    public string? Format { get; set; }

    public void Apply(ExcelRange range, ExcelWorksheet worksheet)
    {
        Apply(range.Style);

        Border?.Apply(range, worksheet);
    }

    private void Apply(ExcelStyle excelStyle)
    {
        if (FontBold != null)
            excelStyle.Font.Bold = FontBold.Value;

        if (BackgroundColor.HasValue)
        {
            excelStyle.Fill.PatternType = ExcelFillStyle.Solid;
            excelStyle.Fill.BackgroundColor.SetColor(BackgroundColor.Value);
        }

        if (HorizontalAlignment.HasValue)
            excelStyle.HorizontalAlignment = HorizontalAlignment.Value;
        if (VerticalAlignment.HasValue)
            excelStyle.VerticalAlignment = VerticalAlignment.Value;
        if (Format != null)
            excelStyle.Numberformat.Format = Format;
    }

    public virtual object Clone()
    {
        var newObject = (ExcelCellStyleTemplate)MemberwiseClone();
        if (newObject.Border != null)
            newObject.Border = (ExcelBorderRangeTemplate)newObject.Border.Clone();
        return newObject;
    }
}