using System.Drawing;
using OfficeOpenXml.Style;

namespace Kapok.Report.Adomd.ExcelStyling;

public class ExcelBorderItemTemplate : ICloneable
{
    public ExcelBorderStyle? Style { get; set; }
    public Color? Color { get; set; }

    public void Apply(ExcelBorderItem borderItem)
    {
        if (Style != null)
            borderItem.Style = Style.Value;
        if (Color != null)
            borderItem.Color.SetColor(Color.Value);
    }

    public object Clone()
    {
        return MemberwiseClone();
    }
}