using System.Collections;
using System.Data;
using Kapok.Report.Model;
using Microsoft.AnalysisServices.AdomdClient;

namespace Kapok.Report.Adomd;

public class AdomdReportDataSet : IDbReportDataSet
{
    public string? DataSourceName { get; set; }

    /// <summary>
    /// The MDX Query to be called.
    /// </summary>
    public string? MdxQuery { get; set; }

    /// <summary>
    /// The resource name for parameter 'MdxQuery'
    /// </summary>
    public string? MdxQueryResourceName { get; set; }

    public IEnumerator GetEnumerator()
    {
        throw new NotImplementedException();
    }

    /// <summary>
    /// The result ADOMD cell set from the MDX query.
    /// </summary>
    public CellSet? CellSet { get; private set; }

    public void ExecuteQuery(IDbConnection connection, IReadOnlyDictionary<string, object?>? parameters = default, IReportResourceProvider? resourceProvider = default)
    {
        string mdxQuery;

        if (MdxQueryResourceName != null)
        {
            if (resourceProvider == null)
            {
                throw new ArgumentException($"The DataSet uses a resource but {nameof(resourceProvider)} was not given", nameof(resourceProvider));
            }
            mdxQuery = System.Text.Encoding.Default.GetString(resourceProvider[MdxQueryResourceName].Data ?? Array.Empty<byte>());
        }
        else if (MdxQuery != null)
        {
            mdxQuery = MdxQuery;
        }
        else
        {
            throw new NotSupportedException($"Could not determine MDX query from DataSet. You need to set property {MdxQuery} or {MdxQueryResourceName}.");
        }

        ExecuteQuery(connection, mdxQuery, parameters);
    }

    public void ExecuteQuery(IDbConnection connection, ReportParameterCollection parameters, IReportResourceProvider? resourceProvider = default)
    {
        ExecuteQuery(connection, parameters.ToDictionary(), resourceProvider);
    }

    private void ExecuteQuery(IDbConnection connection, string mdxQuery, IReadOnlyDictionary<string, object?>? parameters)
    {
        if (connection == null) throw new ArgumentNullException(nameof(connection));

        var command = new AdomdCommand(mdxQuery, (AdomdConnection)connection);

        if (parameters != null)
            foreach (var reportParameter in parameters)
            {
                command.Parameters.Add(new AdomdParameter(reportParameter.Key, reportParameter.Value));
            }

        bool handleConnection = connection.State == ConnectionState.Closed;

        try
        {
            if (handleConnection)
                connection.Open();

            CellSet = command.ExecuteCellSet();
        }
        finally
        {
            if (handleConnection)
                connection.Close();
        }
    }
}