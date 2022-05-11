using Xunit;

namespace Kapok.Report.Adomd.UnitTest;

public class UnitTest
{
    [Fact]
    public void TestBasicMdxQuery()
    {
        var dataSource = new AdomdReportDataSource
        {
            Name = "LocalTest",
            ConnectionString = "Data Source=localhost;Catalog=DWH"
        };

        var dataSet = new AdomdReportDataSet
        {
            DataSourceName = "LocalTest",

            // A query to selecte all databases
            MdxQuery = @"SELECT * FROM $system.dbschema_catalogs"
        };

        using (var connection = dataSource.CreateNewConnection())
        {
            Assert.Null(dataSet.CellSet);

            dataSet.ExecuteQuery(connection);

            Assert.NotNull(dataSet.CellSet);
        }
    }
}