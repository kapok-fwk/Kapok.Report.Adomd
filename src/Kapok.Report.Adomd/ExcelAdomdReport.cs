using System.Diagnostics;
using Kapok.Report.Adomd.ExcelStyling;
using Kapok.Report.Model;
using Microsoft.AnalysisServices.AdomdClient;
using OfficeOpenXml;

namespace Kapok.Report.Adomd;

public abstract class ExcelAdomdReport : ExcelReport
{
    public ExcelContingencyTableStyle? TableStyle { get; set; }

    public override ExcelWorksheet Build(ReportEngine reportEngine, ExcelWorkbook workbook)
    {
        var adomdDataSet = (AdomdReportDataSet)DataSets.FirstOrDefault(ds => ds.Value is AdomdReportDataSet).Value;
        if (adomdDataSet == null)
            throw new NotSupportedException($"The ExcelAdomdReport does not have a DataSet assignable to type {typeof(AdomdReportDataSet).FullName}");
        if (adomdDataSet.DataSourceName == null)
            throw new NotSupportedException($"The property {nameof(adomdDataSet.DataSourceName)} of the first DataSet assignable to type {typeof(AdomdReportDataSet).FullName} is not set");

        // TODO: implement a global connection pool handling
        AdomdConnection? connection = null;
        bool newConnection = false;

        ExcelWorksheet worksheet;

        try
        {
            if (connection == null)
            {
                var dataSource = reportEngine.GetDataSource(adomdDataSet.DataSourceName);
                if (!(dataSource is AdomdReportDataSource adomdDataSource))
                    throw new NotSupportedException($"The report data source must be of type {nameof(AdomdReportDataSource)}");

                connection = (AdomdConnection) adomdDataSource.CreateNewConnection();
                connection.Open();
                newConnection = true;
            }

            worksheet = base.Build(reportEngine, workbook);

            WriteToExcelWorksheet(worksheet, connection, adomdDataSet);
        }
        finally
        {
            if (newConnection)
                connection?.Close();
        }

        return worksheet;
    }

    protected virtual void WriteToExcelWorksheet(ExcelWorksheet worksheet, AdomdConnection connection, AdomdReportDataSet dataSet)
    {
        AdomdQueryToExcelWorksheet(
            connection,
            dataSet: dataSet,
            resourceProvider: Resources,
            reportParameters: Parameters,
            worksheet: worksheet,
            tableName: $"{worksheet.Name.Replace(" ", "")}_Table1",
            tableStyle: TableStyle,
            columnStart: 1,
            rowStart: 5
        );
    }

    public static (int, int) AdomdQueryToExcelWorksheet(AdomdConnection connection, AdomdReportDataSet dataSet,
        IReportResourceProvider? resourceProvider, ReportParameterCollection reportParameters,
        ExcelWorksheet worksheet, string tableName, ExcelContingencyTableStyle? tableStyle, int columnStart, int rowStart,
        string? title = null,
        List<AdomdAxisDynamicCaption>? columnDynamicCaptions = null,
        List<AdomdAxisDynamicCaption>? rowDynamicCaptions = null,
        bool createRowLevelGroups = false)
    {
        dataSet.ExecuteQuery(connection, reportParameters, resourceProvider);

        // TODO: need a optimistic behavior implementation here + a warning on report-creation written to the log
        Debug.Assert(dataSet.CellSet?.Axes.Count == 2);

        // write contingency table
        return AdomdExcelHelper.CellSetToWorksheet(worksheet,
            cellSet: dataSet.CellSet,
            tableName: tableName,
            tableStyle: tableStyle,
            columnStart: columnStart,
            rowStart: rowStart,
            title: title,
            columnDynamicCaptions: columnDynamicCaptions,
            rowDynamicCaptions: rowDynamicCaptions,
            createRowLevelGroups: createRowLevelGroups);
    }
}