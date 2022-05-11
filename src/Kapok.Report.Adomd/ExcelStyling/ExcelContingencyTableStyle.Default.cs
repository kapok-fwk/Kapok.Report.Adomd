using System.Drawing;
using OfficeOpenXml.Style;

namespace Kapok.Report.Adomd.ExcelStyling;

public partial class ExcelContingencyTableStyle
{
    /// <summary>
    /// The default style when no style was selected.
    /// </summary>
    public static ExcelContingencyTableStyle Default
    {
        get
        {
            return new ExcelContingencyTableStyle
            {
                ColumnAxisStyle = new ExcelAxisStyleTemplate
                {
                    FontBold = true,
                    BackgroundColors = new[] { Color.FromArgb(255, 230, 153) },
                    HorizontalAlignment = ExcelHorizontalAlignment.Center,
                    Border = new ExcelBorderRangeTemplate
                    {
                        Bottom = new ExcelBorderItemTemplate
                        {
                            Color = Color.Black,
                            Style = ExcelBorderStyle.Thin
                        },
                        Left = new ExcelBorderItemTemplate
                        {
                            Color = Color.Black,
                            Style = ExcelBorderStyle.Thin
                        },
                        Right = new ExcelBorderItemTemplate
                        {
                            Color = Color.Black,
                            Style = ExcelBorderStyle.Thin
                        },
                        Top = new ExcelBorderItemTemplate
                        {
                            Color = Color.Black,
                            Style = ExcelBorderStyle.Thin
                        },
                        Horizontal = new ExcelBorderItemTemplate
                        {
                            Color = Color.Black,
                            Style = ExcelBorderStyle.Thin
                        },
                        Vertical = new ExcelBorderItemTemplate
                        {
                            Color = Color.Black,
                            Style = ExcelBorderStyle.Thin
                        }
                    }
                },
                RowAxisStyle = new ExcelAxisStyleTemplate
                {
                    BackgroundColors = new[] { Color.FromArgb(255, 242, 204) },
                    Border = new ExcelBorderRangeTemplate
                    {
                        Bottom = new ExcelBorderItemTemplate
                        {
                            Color = Color.Black,
                            Style = ExcelBorderStyle.Thin
                        },
                        Left = new ExcelBorderItemTemplate
                        {
                            Color = Color.Black,
                            Style = ExcelBorderStyle.Thin
                        },
                        Right = new ExcelBorderItemTemplate
                        {
                            Color = Color.Black,
                            Style = ExcelBorderStyle.Thin
                        },
                        Top = new ExcelBorderItemTemplate
                        {
                            Color = Color.Black,
                            Style = ExcelBorderStyle.Thin
                        }
                    }
                },
                CellAreaStyle = new ExcelCellStyleTemplate
                {
                    Border = new ExcelBorderRangeTemplate
                    {
                        Bottom = new ExcelBorderItemTemplate
                        {
                            Color = Color.Black,
                            Style = ExcelBorderStyle.Thin
                        },
                        Left = new ExcelBorderItemTemplate
                        {
                            Color = Color.Black,
                            Style = ExcelBorderStyle.Thin
                        },
                        Right = new ExcelBorderItemTemplate
                        {
                            Color = Color.Black,
                            Style = ExcelBorderStyle.Thin
                        },
                        Top = new ExcelBorderItemTemplate
                        {
                            Color = Color.Black,
                            Style = ExcelBorderStyle.Thin
                        }
                    }
                },
                ColumnAxisCellStyle = new List<ExcelColumnAxisCellStyleTemplate>
                {
                    new ExcelColumnAxisCellStyleTemplate
                    {
                        Border = new ExcelBorderRangeTemplate
                        {
                            Vertical = new ExcelBorderItemTemplate
                            {
                                Style = ExcelBorderStyle.None
                            }
                        }
                    },
                    new ExcelColumnAxisCellStyleTemplate
                    {
                        ApplyOnTupleLevel = 1,
                        Border = new ExcelBorderRangeTemplate
                        {
                            Vertical = new ExcelBorderItemTemplate
                            {
                                Color = Color.Black,
                                Style = ExcelBorderStyle.Thin
                            }
                        }
                    }
                },
                RowAxisCellStyle = new List<ExcelRowAxisCellStyleTemplate>
                {
                    new ExcelRowAxisCellStyleTemplate
                    {
                        BackgroundColor = Color.FromArgb(189, 215, 238)
                    },
                    new ExcelRowAxisCellStyleTemplate
                    {
                        BackgroundColor = Color.FromArgb(221, 235, 247)
                    }
                }
            };
        }
    }
}