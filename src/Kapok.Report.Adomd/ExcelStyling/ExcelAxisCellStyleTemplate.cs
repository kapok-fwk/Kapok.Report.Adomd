using System.ComponentModel.DataAnnotations;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Kapok.Report.Adomd.ExcelStyling;

public abstract class ExcelAxisCellStyleTemplate : ICloneable
{
    /// <summary>
    /// The tuple level on which the style will apply.
    ///
    /// When you use just one tuple, this has no change.  When you use multiple tuples, this gives you the opportunity to
    /// make borders just on e.g. the first tuple level in the cell range.
    ///
    /// The style will be applied with the range of the tuple level.
    /// </summary>
    [Range(1, int.MaxValue)]
    public int? ApplyOnTupleLevel { get; set; }

    /// <summary>
    /// Apply the style on a defined tuple.
    ///
    /// For MDX, you need to use the unique name of the members of the tuple
    /// </summary>
    public string[]? ApplyOnTuple { get; set; }

    /// <summary>
    /// Extends the range to the column/row header as well.
    /// </summary>
    public bool ExtendStyleToHeader { get; set; }

    public Color? BackgroundColor { get; set; }

    public bool? FontBold { get; set; }
    public Color? FontColor { get; set; }
    public float? FontSize { get; set; }

    public ExcelHorizontalAlignment? HorizontalAlignment { get; set; }
    public ExcelVerticalAlignment? VerticalAlignment { get; set; }
    public string? Format { get; set; }

    public abstract void Apply(ExcelRange range, ExcelWorksheet worksheet, int[]? tupleGrouping = null);

    /// <summary>
    /// Shows an alternative caption from another ADOMD dimension property.
    /// </summary>
    public string[]? DynamicCaptionFromTupleMemberAdomdProperty { get; set; }

    protected void Apply(ExcelStyle excelStyle)
    {
        if (BackgroundColor.HasValue)
        {
            excelStyle.Fill.PatternType = ExcelFillStyle.Solid;
            excelStyle.Fill.BackgroundColor.SetColor(BackgroundColor.Value);
        }

        if (FontBold.HasValue)
            excelStyle.Font.Bold = FontBold.Value;
        if (FontColor.HasValue)
            excelStyle.Font.Color.SetColor(FontColor.Value);
        if (FontSize.HasValue)
            excelStyle.Font.Size = FontSize.Value;

        if (HorizontalAlignment.HasValue)
            excelStyle.HorizontalAlignment = HorizontalAlignment.Value;
        if (VerticalAlignment.HasValue)
            excelStyle.VerticalAlignment = VerticalAlignment.Value;
        if (Format != null)
            excelStyle.Numberformat.Format = Format;
    }

    public virtual object Clone()
    {
        return MemberwiseClone();
    }
}