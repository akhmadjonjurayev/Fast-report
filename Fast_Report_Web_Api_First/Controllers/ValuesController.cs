using Fast_Report_Web_Api_First.Model;
using FastReport;
using FastReport.Barcode;
using FastReport.Data;
using FastReport.DataVisualization.Charting;
using FastReport.Export.Html;
using FastReport.Export.Image;
using FastReport.Export.Pdf;
using FastReport.Export.RichText;
using FastReport.MSChart;
using FastReport.Table;
using FastReport.Utils;
using FastReport.Web;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Fast_Report_Web_Api_First.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ValuesController : ControllerBase
    {
        [Obsolete]
        private readonly IHostingEnvironment _hostingEnvironment;

        [Obsolete]
        public ValuesController(IHostingEnvironment hostingEnvironment)
        {
            _hostingEnvironment = hostingEnvironment;
        }
        [HttpGet("getPdf")]
        [Obsolete]
        public IActionResult GetPdf()
        {
            string message;
            using (var stream = new MemoryStream())
            {
                try
                {
                    using (var dataSet = new DataSet())
                    {
                        RegisteredObjects.AddConnection(typeof(JsonDataConnection));
                        string webRootPath = _hostingEnvironment.WebRootPath;
                        string reportPath = (webRootPath + @"\App_Data\" + "Simple List.frx");
                        Config.WebMode = true;
                        using (var report = new Report())
                        {
                            ReportPage page = new ReportPage();
                            page.CreateUniqueName();
                            page.TopMargin = 10.0f;
                            page.LeftMargin = 10.0f;
                            page.RightMargin = 10.0f;
                            page.BottomMargin = 10.0f;
                            page.Border.Lines = BorderLines.All;
                            report.Pages.Add(page);
                           
                            report.Load(reportPath);
                           // report.RegisterData(dataSet);
                            report.Dictionary.RegisterData(Person(), "person", true);
                            PDFExport pdf = new PDFExport
                            {
                                AllowCopy = true,
                                AllowPrint = true,
                                Author = "germes.lc",
                            };
                            //RTFExport word = new RTFExport()
                            //{
                            //    AllowOpenAfter = true
                            //};
                            FastReport.Export.OoXML.Word2007Export word = new FastReport.Export.OoXML.Word2007Export()
                            {
                                AllowOpenAfter = true
                            };
                            FastReport.Export.BIFF8.Excel2003Document excel = new FastReport.Export.BIFF8.Excel2003Document()
                            {
                                
                            };
                            
                            BarcodeObject bc = report.FindObject("Barcode2") as BarcodeObject;
                            bc.Text = Guid.NewGuid().ToString();
                            MSChartObject chart = report.FindObject("MSChart1") as MSChartObject;
                            ChartDraw(chart);
                            var table11 = report.FindObject("Table1");
                            TableObject table = report.FindObject("Table1") as TableObject;
                            
                            CreateTable(table);
                            report.Prepare();
                            report.Export(word, stream);
                            var word_mime = "application/msword";
                            var pdf_mime = "application/pdf";
                            var excel_mime = "application/vnd.ms-excel";
                            return File(stream.ToArray(), word_mime, "report.docx");
                        }
                    }
                }
                catch(Exception ex)
                {
                    message = ex.Message;
                }
                finally
                {
                    stream.Close();
                }
            }
            return Ok(message);
        }
        private void CreateTable(TableObject table)
        {
            table.Name = "Table1";
            table.RowCount = 10;
            for (int i = 1; i < table.RowCount; i++)
            {
                for (int j = 0; j < table.ColumnCount; j++)
                {
                    table[j, i].Text = (10 * i + j + 1).ToString();
                    table[j, i].Border.Lines = BorderLines.All;
                }
                table.Rows[i].Height = 26;
            }   
        }
        private void ChartDraw(MSChartObject MSChart1)
        {
            MSChart1.DeleteSeries(0);
            MSChart1.AddSeries(SeriesChartType.Bubble);
            MSChart1.Series[0].SeriesSettings.Points.Clear();
            MSChart1.Series[0].SeriesSettings.Points.AddXY("Bob", 8);
            MSChart1.Series[0].SeriesSettings.Points.AddXY("Damir", 10);
            MSChart1.Series[0].SeriesSettings.Points.AddXY("Anna", 9);
            MSChart1.Chart.Legends[0].Enabled = false;
            MSChart1.Series[0].SeriesSettings.Label = "#VALY";
            MSChart1.Height = 300;
            MSChart1.Width = 300;
        }
        static DataTable GetTable<TEntity>(IEnumerable<TEntity> table, string name) where TEntity : class
        {
            var offset = 78;
            DataTable result = new DataTable(name);
            PropertyInfo[] infos = typeof(TEntity).GetProperties();
            foreach (PropertyInfo info in infos)
            {
                if (info.PropertyType.IsGenericType
                && info.PropertyType.GetGenericTypeDefinition() 
                == typeof(Nullable<>))
                {
                    result.Columns.Add(new DataColumn(info.Name,
                        Nullable.GetUnderlyingType(info.PropertyType)));
                }
                else
                {
                    result.Columns.Add(new DataColumn(info.Name, 
                        info.PropertyType));
                }
            }
            foreach (var el in table)
            {
                DataRow row = result.NewRow();
                foreach (PropertyInfo info in infos)
                {
                    if (info.PropertyType.IsGenericType &&
                        info.PropertyType.GetGenericTypeDefinition()
                        == typeof(Nullable<>))
                    {
                        object t = info.GetValue(el);
                        if (t == null)
                        {
                            t = Activator.CreateInstance(
                                Nullable.GetUnderlyingType(info.PropertyType));
                        }

                        row[info.Name] = t;
                    }
                    else
                    {
                        if (info.PropertyType == typeof(byte[]))
                        {
                            //Fix for Image issue.
                            var imageData = (byte[])info.GetValue(el);
                            var bytes = new byte[imageData.Length - offset];
                            Array.Copy(imageData, offset, bytes, 0, bytes.Length);
                            row[info.Name] = bytes;
                        }
                        else
                        {
                            row[info.Name] = info.GetValue(el);
                        }
                    }
                }
                result.Rows.Add(row);
            }

            return result;
        }
        private List<Person> Person()
        {
            var file = System.IO.File.ReadAllBytes(@"C:\Users\WebDeveloper\Pictures\Saved Pictures\20.jpg");
            return new List<Person>()
            { new Person() { Id = 1, firstName = "name 1", lastName = "name 2", birthday = DateTime.Now, address = "address 1", phone = "998909009090", picture = file,QrCode = Guid.NewGuid().ToString(),
            html = "<i><b>World</b></i>"} };
        }
        private List<Person> ForArrayList()
        {
            return new List<Person>()
            {
                new Person(){Id = 2},
                new Person(){Id = 3},
                new Person(){Id = 4},
                new Person(){Id = 2}
            };
        }
    }
}
