namespace Kapok.Report.Adomd.ExcelStyling;

public partial class ExcelContingencyTableStyle : ICloneable
{
    /// <summary>
    /// Styles the column axis
    /// </summary>
    public ExcelAxisStyleTemplate? ColumnAxisStyle { get; set; }

    /// <summary>
    /// Styles the row axis
    /// </summary>
    public ExcelAxisStyleTemplate? RowAxisStyle { get; set; }

    /// <summary>
    /// Styles the whole cell area style
    /// </summary>
    public ExcelCellStyleTemplate? CellAreaStyle { get; set; }

    /// <summary>
    /// Styles the cells based on the column dimension.
    ///
    /// When using more than one style, it is applied repeatingly; e.g. usable to color each second column differently.
    /// </summary>
    public List<ExcelColumnAxisCellStyleTemplate>? ColumnAxisCellStyle { get; set; }

    /// <summary>
    /// Styles the cells based on the row dimension.
    ///
    /// When using more than one style, it is applied repeatingly; e.g. usable to color each second column differently.
    /// </summary>
    public List<ExcelRowAxisCellStyleTemplate>? RowAxisCellStyle { get; set; }

    /// <summary>
    /// If the excel freeze pane shall be set. The freeze pane will be set
    /// on the cell where the cell area of the contingency table starts.
    /// </summary>
    public bool SetFreezePane { get; set; }

    /// <summary>
    /// By default a filter is added to the header of the table (when there exist not multiple header!)
    /// 
    /// With this property you have the option to change this behavior.
    /// </summary>
    public bool? ShowFilter { get; set; }

    public object Clone()
    {
        var newObject = (ExcelContingencyTableStyle) MemberwiseClone();
        if (newObject.ColumnAxisStyle != null)
            newObject.ColumnAxisStyle = (ExcelAxisStyleTemplate)newObject.ColumnAxisStyle.Clone();
        if (newObject.RowAxisStyle != null)
            newObject.RowAxisStyle = (ExcelAxisStyleTemplate)newObject.RowAxisStyle.Clone();
        if (newObject.CellAreaStyle != null)
            newObject.CellAreaStyle = (ExcelCellStyleTemplate)newObject.CellAreaStyle.Clone();
        if (newObject.ColumnAxisCellStyle != null)
            newObject.ColumnAxisCellStyle = new List<ExcelColumnAxisCellStyleTemplate>(newObject.ColumnAxisCellStyle);
        if (newObject.RowAxisCellStyle != null)
            newObject.RowAxisCellStyle = new List<ExcelRowAxisCellStyleTemplate>(newObject.RowAxisCellStyle);
        return newObject;
    }
}