using Microsoft.AspNetCore.Mvc;
using System.Data;
using System.Data.OleDb;

namespace Excel_Data_API_Application.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelDataController : ControllerBase
    {
        private readonly string filePath = "details.xlsx"; 

        [HttpGet]
        public IActionResult GetAllData()
        {
            var data = ReadExcelFile();
            return Ok(data);
        }

        [HttpGet("{Trade_Id}")]
        public IActionResult GetDataById(string TradeId)
        {
            var data = ReadExcelFile();

            var row = data.FirstOrDefault(r => r.ContainsKey("Trade Id") && r["Trade Id"] == TradeId);

            if (row == null)
                return NotFound($"No record found with Trade Id = {TradeId}");

            return Ok(row);
        }

        private List<Dictionary<string, string>> ReadExcelFile()
        {
            DataTable dataTable = new DataTable();

            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath};Extended Properties='Excel 12.0;HDR=YES;'";

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                string query = "SELECT * FROM [Sheet1$]"; 
                OleDbDataAdapter adapter = new OleDbDataAdapter(query, connection);
                adapter.Fill(dataTable);
            }


            var result = new List<Dictionary<string, string>>();
            foreach (DataRow row in dataTable.Rows)
            {
                var rowData = new Dictionary<string, string>();
                foreach (DataColumn column in dataTable.Columns)
                {
                    rowData[column.ColumnName] = row[column].ToString();
                }
                result.Add(rowData);
            }

            return result;
        }
    }
}
