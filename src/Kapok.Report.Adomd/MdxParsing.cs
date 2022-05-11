using System.Drawing;
using Microsoft.AnalysisServices.AdomdClient;

namespace Kapok.Report.Adomd;

internal static class MdxParsing
{
    public static Color ColorFromRgbInteger(uint rgb)
    {
        int red = (int)(rgb & 0x0000FF);
        int green = (int)((rgb & 0x00FF00) >> 8);
        int blue = (int)((rgb & 0xFF0000) >> 16);

        return Color.FromArgb(red, green, blue);
    }

    // ReSharper disable IdentifierTypo
    // ReSharper disable InconsistentNaming

    // Source of constants: https://docs.microsoft.com/en-us/analysis-services/multidimensional-models/mdx/mdx-cell-properties-using-cell-properties?view=asallproducts-allversions

    // property ACTION_TYPE; source: https://docs.microsoft.com/en-us/previous-versions/sql/sql-server-2012/ms126032%28v%3dsql.110%29
    public const int MDACTION_TYPE_URL = 0x01;
    public const int MDACTION_TYPE_HTML = 0x02;
    public const int MDACTION_TYPE_STATEMENT = 0x04;
    public const int MDACTION_TYPE_DATASET = 0x08;
    public const int MDACTION_TYPE_ROWSET = 0x10;
    public const int MDACTION_TYPE_COMMANDLINE = 0x20;
    public const int MDACTION_TYPE_PROPRIETARY = 0x40;
    public const int MDACTION_TYPE_REPORT = 0x80;
    public const int MDACTION_TYPE_DRILLTHROUGH = 0x100;

    // property FONT_FLAGS
    public const int MDFF_BOLD = 1;
    public const int MDFF_ITALIC = 2;
    public const int MDFF_UNDERLINE = 4;
    public const int MDFF_STRIKEOUT = 8;

    // ReSharper restore InconsistentNaming
    // ReSharper restore IdentifierTypo

    /// <summary>
    /// An optimistic approach reading cell values. In case of an AdomdErrorResponseException exception,
    /// the value given in <param name="errorValue">errorValue</param> is used.
    /// </summary>
    /// <param name="cell"></param>
    /// <param name="errorValue"></param>
    /// <returns></returns>
    public static object? ReadValue(this Cell cell, object? errorValue = null)
    {
        try
        {
            return cell.Value;
        }
        catch (AdomdErrorResponseException)
        {
            // when an exception occured e.g. in a calculated measure, we cast this
            // error here over to the default errorValue parameter
            return errorValue;
        }
    }
}