using System.Collections;
using Microsoft.AnalysisServices.AdomdClient;

namespace Kapok.Report.Adomd;

/// <summary>
/// Provides an iterator for an ADOMD MDX query:
///
/// It executes the ADOMD MDX query and iterates over the first members in the first axis of the query.
/// </summary>
public class ReportParameterAdomdIterationSource : IEnumerable<object>
{
    protected readonly AdomdConnection Connection;
    protected readonly AdomdReportDataSet DataSet;
    protected readonly IReportResourceProvider? ResourceProvider;
    protected readonly IReadOnlyDictionary<string, object>? Parameters;
    protected readonly string? MemberPropertyName;

    /// <summary>
    /// Initiates a iteration source from a ADOMD MDX query.
    /// </summary>
    /// <param name="connection">
    /// The ADOMD connection.
    /// </param>
    /// <param name="dataSet">
    /// The ADOMD data set which will be queried.
    /// </param>
    /// <param name="memberPropertyName">
    /// (Optional) The member property which shall be returned. When not specified, the CAPTION
    /// property is returned.
    /// 
    /// NOTE: Make sure you define in the MDX query on the first axis that the member property is
    /// returned. You do this with 'DIMENSION PROPERTIES [DimensionName].[HierarchyName].[AttributeOrLevelName].[PropertyName]'.
    /// </param>
    /// <param name="resourceProvider">
    /// The resource provider, will be required when <c>dataSet.ExecuteQuery(..)</c> requires a resource, e.g. because <see cref="AdomdReportDataSet.MdxQueryResourceName"/> is used.
    /// </param>
    /// <param name="parameters">
    /// The parameters to be passed to <c>dataSet.ExecuteQuery(..)</c>.
    /// </param>
    public ReportParameterAdomdIterationSource(AdomdConnection connection, AdomdReportDataSet dataSet,
        string? memberPropertyName = default,
        IReportResourceProvider? resourceProvider = default,
        IReadOnlyDictionary<string, object>? parameters = default)
    {
        Connection = connection;
        DataSet = dataSet;
        ResourceProvider = resourceProvider;
        MemberPropertyName = memberPropertyName;
        Parameters = parameters;
    }

    private CellSet QueryExecuteCellSet()
    {
        DataSet.ExecuteQuery(Connection, Parameters, ResourceProvider);

#pragma warning disable 8603
        return DataSet.CellSet;
#pragma warning restore 8603
    }

    /// <summary>
    /// 
    /// </summary>
    /// <exception cref="NotSupportedException">
    /// The iteration throws a 'NotSupportedException' exception when the member property
    /// was not given for a member in the query.
    /// </exception>
    /// <returns></returns>
    public IEnumerator<object> GetEnumerator()
    {
        var cellSet = QueryExecuteCellSet();
        if (cellSet.Axes.Count == 0)
            yield break;

        foreach (var tuple in cellSet.Axes[0].Set.Tuples)
        {
            var member = tuple.Members[0];

            if (MemberPropertyName != null)
            {
                var memberProperty = member.MemberProperties.Find(MemberPropertyName);
                if (memberProperty == null)
                    throw new NotSupportedException(
                        $"The member property '{MemberPropertyName}' was not found at the member with caption '{member.Caption}'.");

                yield return memberProperty.Value;
            }
            else
            {
                yield return member.Caption;
            }
        }
    }

    #region IEnumerator

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion
}