using System.Diagnostics;
using System.Reflection;
using Microsoft.AspNetCore.Mvc;
using Practice_NPOI.Models;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.Util;


namespace Practice_NPOI.Controllers;

public class HomeController : Controller
{
    private List<TableModel> GetData()
    {
        List<TableModel> data = new List<TableModel>
        {
            new TableModel { Field1 = "Value1", Field2 = "Value2" }
        };

        return data;
    }

    [HttpGet("api/downladexcel")]
    public IActionResult DownloadExcel()
    {
        IWorkbook workbook = new HSSFWorkbook();
        ISheet sheet = workbook.CreateSheet("История обслуживания");

        List<TableModel> data = GetData();

        PropertyInfo[] properties = typeof(TableModel).GetProperties();
        
        IRow headerRow = sheet.CreateRow(0);
        for (int i = 0; i < properties.Length; i++)
        {
            ICell cell = headerRow.CreateCell(i);
            cell.SetCellValue(properties[i].Name);
        }
        
        for (int i = 0; i < data.Count; i++)
        {
            TableModel record = data[i];
            IRow row = sheet.CreateRow(i + 1);

            for (int j = 0; j < properties.Length; j++)
            {
                ICell cell = row.CreateCell(j);
                object value = properties[j].GetValue(record);
                cell.SetCellValue(value?.ToString() ?? string.Empty);
            }
        }

        using (MemoryStream stream = new MemoryStream())
        {
            workbook.Write(stream);
            var content = stream.ToArray();

            return File(content, "application/vnd.ms-excel", "История обслуживания.xls");
        }
    }
}

public class TableModel
{
    public string Field1 { get; set; }
    public string Field2 { get; set; }
}    

