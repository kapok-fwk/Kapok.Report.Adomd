using OfficeOpenXml;

namespace Kapok.Report.Adomd.ExcelStyling;

public class ExcelColumnAxisCellStyleTemplate : ExcelAxisCellStyleTemplate
{
    public ExcelBorderRangeTemplate? Border { get; set; }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="range"></param>
    /// <param name="worksheet"></param>
    /// <param name="tupleGrouping">
    /// Contains the definition how cells are grouped on the defined level in 'ApplyOnTupleLevel';
    ///
    /// The array is a list of all tuples in the defined tuple level;
    /// the integer gives the number of columns in the specific tuple.
    /// </param>
    public override void Apply(ExcelRange range, ExcelWorksheet worksheet, int[]? tupleGrouping = null)
    {
        Apply(range.Style);

        Border?.Apply(range, worksheet, verticalGrouping: tupleGrouping);
    }

    public override object Clone()
    {
        var newObject = (ExcelColumnAxisCellStyleTemplate)base.Clone();
        if (newObject.Border != null)
            newObject.Border = (ExcelBorderRangeTemplate)newObject.Border.Clone();
        return newObject;
    }
}