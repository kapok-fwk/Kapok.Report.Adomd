using OfficeOpenXml.Style;

namespace Kapok.Report.Adomd.ExcelStyling;

/// <summary>
/// A border style applied to a single cell
/// </summary>
public class ExcelBorderCellTemplate : ICloneable
{
    /// <summary>Left border style</summary>
    public ExcelBorderItemTemplate? Left { get; set; }
    /// <summary>Right border style</summary>
    public ExcelBorderItemTemplate? Right { get; set; }
    /// <summary>Top border style</summary>
    public ExcelBorderItemTemplate? Top { get; set; }
    /// <summary>Bottom border style</summary>
    public ExcelBorderItemTemplate? Bottom { get; set; }
    /// <summary>0Diagonal border style</summary>
    public ExcelBorderItemTemplate? Diagonal { get; set; }

    /// <summary>
    /// A diagonal from the bottom left to top right of the cell
    /// </summary>
    public bool DiagonalUp { get; set; }
    /// <summary>
    /// A diagonal from the top left to bottom right of the cell
    /// </summary>
    public bool DiagonalDown { get; set; }

    public virtual void Apply(Border border)
    {
        Left?.Apply(border.Left);
        Right?.Apply(border.Right);
        Top?.Apply(border.Top);
        Bottom?.Apply(border.Bottom);
        if (Diagonal != null)
        {
            Diagonal.Apply(border.Diagonal);
            border.DiagonalUp = DiagonalUp;
            border.DiagonalDown = DiagonalDown;
        }
    }

    public virtual object Clone()
    {
        var newObject = (ExcelBorderCellTemplate)MemberwiseClone();
        if (newObject.Left != null)
            newObject.Left = (ExcelBorderItemTemplate)newObject.Left.Clone();
        if (newObject.Right != null)
            newObject.Right = (ExcelBorderItemTemplate) newObject.Right.Clone();
        if (newObject.Top != null)
            newObject.Top = (ExcelBorderItemTemplate)newObject.Top.Clone();
        if (newObject.Bottom != null)
            newObject.Bottom = (ExcelBorderItemTemplate)newObject.Bottom.Clone();
        if (newObject.Diagonal != null)
            newObject.Diagonal = (ExcelBorderItemTemplate)newObject.Diagonal.Clone();
        return newObject;
    }
}