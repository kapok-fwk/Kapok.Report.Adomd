namespace Kapok.Report.Adomd;

public class AdomdAxisDynamicCaption
{
    /// <summary>
    /// Says to which level of the tuples it applies.
    /// </summary>
    public int ApplyOnTupleLevel { get; set; }

    /// <summary>
    /// Let you return the value of another property then the member caption property.
    /// </summary>
    public string CaptionPropertyName { get; set; } = "CAPTION";
}