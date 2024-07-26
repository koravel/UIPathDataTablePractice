using Microsoft.Office.Interop.Excel;
using System.IO;
using Syncfusion.XlsIO;

namespace ExcelPractice
{
    public class ChartEditor
    {
        
//        void ApplyCustomType(ISheetRef rrRef, IWorkbookQuickHandle qq) {
//            Microsoft.Office.Interop.Excel.Range rng = new Microsoft.Office.Interop.Excel.Range();
//            string builtInType = "Line - Column";
//            XlChartType customChartType = XlChartType.xlLine;
//            Chart chart = new Chart();
//            rrRef.
//            chart.SetSourceData(rrRef.Address)
//            chart.ApplyCustomType(customChartType, builtInType);
//            IChartRef q;
//            ExcelChartAction chac;
//            var charttt = qq.Sheet[""].Chart[""];
            
//            Worksheet ws = new WorksheetQuickHandle("");
//            ws.
//        }
        
//        void qwer() {
//            using (ExcelEngine excelEngine = new ExcelEngine())
//            {
//                IApplication application = excelEngine.Excel;
//                application.DefaultVersion = ExcelVersion.Excel2016;
 
//                Open existing workbook with data entered
//                Assembly assembly = typeof(Program).GetTypeInfo().Assembly;
//                Stream fileStream = assembly.GetManifestResourceStream("ChartSample.InputTemplate.xlsx");
//                IWorkbook workbook = application.Workbooks.Open(fileStream);
//                IWorksheet worksheet = workbook.Worksheets[0];
 
//                Initialize chart and assign data
//                IChartShape chart = worksheet.Charts.Add();
//                chart.DataRange = worksheet["A1:C13"];
 
//                Apply chart elements
//                Set Chart Title
//                chart.ChartTitle = "Combination Chart";
//                chart.IsSeriesInRows = false;
 
//                Set Legend
//                chart.HasLegend = true;
//                chart.Legend.Position = ExcelLegendPosition.Bottom;
 
//                Set Serie type
//                IChartSerie serie1 = chart.Series[0];
//                IChartSerie serie2 = chart.Series[1];
//                serie1.SerieType = ExcelChartType.Column_Clustered;
//                serie2.SerieType = ExcelChartType.Line;
//                serie2.UsePrimaryAxis = false;
 
//                set Data labels
//                serie1.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;                     
 
//                Positioning the chart in the worksheet
//                chart.TopRow = 8;
//                chart.LeftColumn = 1;
//                chart.BottomRow = 23;
//                chart.RightColumn = 8;
 
//                Saving and closing the workbook
//                Stream stream = File.Create("Output.xlsx");
//                workbook.SaveAs(stream);
//            }
//        }
        
        public void AddWorksheetToExcelWorkbook(string fullFilename,string worksheetName)
        {
            Application xlApp = new Application();
            Workbook xlWorkbook = xlApp.Workbooks.Open(fullFilename);
            Worksheet xlWorksheet = (Worksheet)xlWorkbook.Sheets[worksheetName];
            
            ChartObject chart = (ChartObject)xlWorksheet.ChartObjects();
            string builtInType = "Line - Column";
            XlChartType customChartType = (XlChartType)XlChartGallery.xlBuiltIn;
            chart.Chart.ApplyCustomType(customChartType, builtInType);
        }
        
        public int CreateChart(string fileName, string sheetName, string dataRangeStr)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2016;
 
                //Open existing workbook with data entered
                //Assembly assembly = typeof(Program).GetTypeInfo().Assembly;
                FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite);
                IWorkbook workbook = application.Workbooks.Open(fileStream);
                IWorksheet worksheet = workbook.Worksheets[sheetName];
 
                //Initialize chart and assign data
                IChartShape chart = worksheet.Charts.Add();
                chart.DataRange = worksheet[dataRangeStr];
 
                //Apply chart elements
                //Set Chart Title
                chart.ChartTitle = "Combination Chart";
                chart.IsSeriesInRows = false;
 
                //Set Legend
                chart.HasLegend = true;
                chart.Legend.Position = ExcelLegendPosition.Bottom;
 
                //Set Serie type
                IChartSerie serie1 = chart.Series[0];
                IChartSerie serie2 = chart.Series[1];
                serie1.SerieType = ExcelChartType.Column_Clustered;
                serie2.SerieType = ExcelChartType.Line;
                serie2.UsePrimaryAxis = false;
 
                //set Data labels
                serie1.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;                     
 
                //Positioning the chart in the worksheet
                chart.TopRow = 8;
                chart.LeftColumn = 1;
                chart.BottomRow = 23;
                chart.RightColumn = 8;
                
                //Saving and closing the workbook
                //Stream stream = File.OpenWrite(fileName);
                workbook.SaveAs(fileStream);
                fileStream.Close();
                return 0;
            }
        }
    }
}