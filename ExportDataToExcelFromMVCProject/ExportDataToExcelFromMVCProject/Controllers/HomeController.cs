using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using OfficeOpenXml;

namespace ExportDataToExcelFromMVCProject.Controllers
{
    public class HomeController : Controller
    {
        MVCExcelTutorialEntities _db = new MVCExcelTutorialEntities();
        // GET: Home
        public ActionResult Index()
        {
            List<EmployeeInfoViewModel> empList = _db.EmployeeInfoes.Select(t => new EmployeeInfoViewModel
            {

                EmployeeID = t.EmployeeID,
                EmployeeName = t.EmployeeName,
                Email = t.Email,
                Phone = t.Phone,
                Experience = t.Experience

            }).ToList();
            return View(empList);
        }

        public void ExportToExcel()
        {

            List<EmployeeInfoViewModel> empList = _db.EmployeeInfoes.Select(t => new EmployeeInfoViewModel
            {

                EmployeeID = t.EmployeeID,
                EmployeeName = t.EmployeeName,
                Email = t.Email,
                Phone = t.Phone,
                Experience = t.Experience

            }).ToList();

            ExcelPackage exclPckg = new ExcelPackage();

            ExcelWorksheet ws = exclPckg.Workbook.Worksheets.Add("DailyReport");
            ws.Cells["A1"].Value = "Communication";
            ws.Cells["B1"].Value = "Com1";

            ws.Cells["A2"].Value = "Report";
            ws.Cells["B2"].Value = "Report1";

            ws.Cells["A3"].Value = "Communication";
            ws.Cells["B3"].Value = string.Format("{0:dd MMM yyyy} at {0:H: mm tt}", DateTimeOffset.Now);

            ws.Cells["A6"].Value = "EmployeeId";
            ws.Cells["B6"].Value = "EmployeeName";
            ws.Cells["C6"].Value = "Email";
            ws.Cells["D6"].Value = "Phone";
            ws.Cells["E6"].Value = "Experience";


            int rowStart = 7;
            foreach (var item in empList)
            {
                if (item.Experience < 5)
                {
                    ws.Row(rowStart).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    ws.Row(rowStart).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(string.Format("red")));
                }
                else
                {
                    ws.Row(rowStart).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    ws.Row(rowStart).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(string.Format("green")));
                }

                ws.Cells[string.Format("A{0}", rowStart)].Value = item.EmployeeID;
                ws.Cells[string.Format("B{0}", rowStart)].Value = item.EmployeeName;
                ws.Cells[string.Format("C{0}", rowStart)].Value = item.Email;
                ws.Cells[string.Format("D{0}", rowStart)].Value = item.Phone;
                ws.Cells[string.Format("E{0}", rowStart)].Value = item.Experience;
                rowStart++;
            }

            ws.Cells["A:AZ"].AutoFitColumns();
            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment:filename=" + "ExcelReport.xlsx");
            Response.BinaryWrite(exclPckg.GetAsByteArray());
            Response.End();
        }
    }
}