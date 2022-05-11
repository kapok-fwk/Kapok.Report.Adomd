using System.Data;
using Microsoft.AnalysisServices.AdomdClient;

namespace Kapok.Report.Adomd;

/// <summary>
/// A ADOMD connection report data source.
/// </summary>
public class AdomdReportDataSource : DbReportDataSource
{
    public override IDbConnection CreateNewConnection()
    {
        return new AdomdConnection(ConnectionString);
    }
}