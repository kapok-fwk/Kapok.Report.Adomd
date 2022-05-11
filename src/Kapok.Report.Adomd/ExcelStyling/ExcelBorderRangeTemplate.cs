using OfficeOpenXml;

namespace Kapok.Report.Adomd.ExcelStyling;

/// <summary>
/// A border style applied to a range
/// </summary>
public class ExcelBorderRangeTemplate : ICloneable
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

    public ExcelBorderItemTemplate? Horizontal { get; set; }
    public ExcelBorderItemTemplate? Vertical { get; set; }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="range"></param>
    /// <param name="worksheet"></param>
    /// <param name="horizontalGrouping">
    /// By default, when using Horizontal parameter, it applies to all columns.
    /// 
    /// This parameter gives the opportunity to define groupings, so e.g.
    /// column A and B belong together, column C is separate. The border should
    /// only appear between B and C. In this case you have to pass
    /// new int[] {2,1}.
    /// So, you define an array with the number of groups. The integer is then
    /// the number of columns in this group.
    /// </param>
    /// <param name="verticalGrouping">
    /// By default, when using Vertical parameter, it applies to all rows.
    /// 
    /// This parameter gives the opportunity to define groupings, so e.g.
    /// row A and B belong together, row C is separate. The border should
    /// only appear between B and C. In this case you have to pass
    /// new int[] {2,1}.
    /// So, you define an array with the number of groups. The integer is then
    /// the number of rows in this group.
    /// </param>
    public void Apply(ExcelRange range, ExcelWorksheet worksheet, int[]? horizontalGrouping = null, int[]? verticalGrouping = null)
    {
        if (Left != null)
        {
            var leftRange = worksheet.Cells[
                range.Start.Row,
                range.Start.Column,
                range.End.Row,
                range.Start.Column
            ];
            Left.Apply(leftRange.Style.Border.Left);
        }

        if (Right != null)
        {
            var rightRange = worksheet.Cells[
                range.Start.Row,
                range.End.Column,
                range.End.Row,
                range.End.Column
            ];
            Right.Apply(rightRange.Style.Border.Right);
        }

        if (Top != null)
        {
            var topRange = worksheet.Cells[
                range.Start.Row,
                range.Start.Column,
                range.Start.Row,
                range.End.Column
            ];
            Top.Apply(topRange.Style.Border.Top);
        }

        if (Bottom != null)
        {
            var bottomRange = worksheet.Cells[
                range.End.Row,
                range.Start.Column,
                range.End.Row,
                range.End.Column
            ];
            Bottom.Apply(bottomRange.Style.Border.Bottom);
        }

        if (Diagonal != null)
        {
            // diagonal is applied to all cells, as it is in Excel
            Diagonal.Apply(range.Style.Border.Diagonal);
            range.Style.Border.DiagonalUp = DiagonalUp;
            range.Style.Border.DiagonalDown = DiagonalDown;
        }

        if (Horizontal != null)
        {
            if (range.Rows > 1)
            {
                if (horizontalGrouping == null)
                {
                    var horizontalRange = worksheet.Cells[
                        range.Start.Row,
                        range.Start.Column,
                        range.End.Row - 1,
                        range.End.Column
                    ];
                    Horizontal?.Apply(horizontalRange.Style.Border.Bottom);
                }
                else
                {
                    int precedingRows = 0;
                    for (int g = 0; g < horizontalGrouping.Length - 1; g++)
                    {
                        var horizontalRange = worksheet.Cells[
                            range.Start.Row + precedingRows + horizontalGrouping[g] - 1,
                            range.Start.Column,
                            range.Start.Row + precedingRows + horizontalGrouping[g] - 1,
                            range.End.Column
                        ];
                        Horizontal?.Apply(horizontalRange.Style.Border.Bottom);

                        precedingRows += horizontalGrouping[g];
                    }
                }
            }
        }

        if (Vertical != null)
        {
            // draw on the first column the vertical lines
            if (range.Columns > 1)
            {
                if (verticalGrouping == null)
                {
                    var verticalRange = worksheet.Cells[
                        range.Start.Row,
                        range.Start.Column,
                        range.End.Row,
                        range.End.Column - 1
                    ];
                    Vertical?.Apply(verticalRange.Style.Border.Right);
                }
                else
                {
                    int precedingColumns = 0;
                    for (int g = 0; g < verticalGrouping.Length - 1; g++)
                    {
                        var verticalRange = worksheet.Cells[
                            range.Start.Row,
                            range.Start.Column + precedingColumns + verticalGrouping[g] - 1,
                            range.End.Row,
                            range.Start.Column + precedingColumns + verticalGrouping[g] - 1
                        ];
                        Vertical?.Apply(verticalRange.Style.Border.Right);

                        precedingColumns += verticalGrouping[g];
                    }
                }
            }
        }
    }

    public object Clone()
    {
        var newObject = (ExcelBorderRangeTemplate)MemberwiseClone();
        if (newObject.Left != null)
            newObject.Left = (ExcelBorderItemTemplate)newObject.Left.Clone();
        if (newObject.Right != null)
            newObject.Right = (ExcelBorderItemTemplate)newObject.Right.Clone();
        if (newObject.Top != null)
            newObject.Top = (ExcelBorderItemTemplate)newObject.Top.Clone();
        if (newObject.Bottom != null)
            newObject.Bottom = (ExcelBorderItemTemplate)newObject.Bottom.Clone();
        if (newObject.Diagonal != null)
            newObject.Diagonal = (ExcelBorderItemTemplate)newObject.Diagonal.Clone();
        if (newObject.Horizontal != null)
            newObject.Horizontal = (ExcelBorderItemTemplate)newObject.Horizontal.Clone();
        if (newObject.Vertical != null)
            newObject.Vertical = (ExcelBorderItemTemplate)newObject.Vertical.Clone();
        return newObject;
    }
}