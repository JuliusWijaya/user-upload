using OfficeOpenXml;
using OfficeOpenXml.Style;
using PRE_TEST.Models;
using Rotativa;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace PRE_TEST.Controllers
{
    public class DownloadController : Controller
    {
        private readonly SIAKADEntities db = new SIAKADEntities();
        // GET: Download
        public ActionResult DownloadUser()
        {
            var users = db.users.AsQueryable();
            return View(users);
        }

        public ActionResult ExportPdf()
        {
            var getUsers = db.users.ToList();

            var pdf = new ViewAsPdf("user", getUsers)
            {
                FileName = $"user.pdf",
                CustomSwitches = "--page-size A4 --orientation Portrait"
            };

            return pdf;
        }

        [HttpGet]
        public ActionResult ExportExcel()
        {
            var reporting = db.users.ToList();

            var stream = new MemoryStream();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(stream))
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                // Add column headers with light blue background
                var headerRow = worksheet.Cells[1, 1, 1, 6];
                headerRow.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                headerRow.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                headerRow.Style.Font.Color.SetColor(Color.Black);
                headerRow.Style.Font.Size = 13;
                headerRow.Style.Font.Bold = true;
                headerRow.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                // Set column headers
                worksheet.Cells[1, 1].Value = "NO";
                worksheet.Cells[1, 2].Value = "NAME";
                worksheet.Cells[1, 3].Value = "JK";
                worksheet.Cells[1, 4].Value = "EMAIL";
                worksheet.Cells[1, 5].Value = "NO TELP";
                worksheet.Cells[1, 6].Value = "ADDRESS";


                // Set borders for the header row
                for (int col = 1; col <= 6; col++)
                {
                    var cell = worksheet.Cells[1, col];
                    cell.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    cell.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    cell.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    cell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                }

                int row = 2;
                foreach (var item in reporting)
                {
                    worksheet.Cells[row, 1].Value = item.id;
                    worksheet.Cells[row, 2].Value = item.name;
                    worksheet.Cells[row, 3].Value = item.jk;
                    worksheet.Cells[row, 4].Value = item.email;
                    worksheet.Cells[row, 5].Value = item.no_telp;
                    worksheet.Cells[row, 6].Value = item.address;

                    // Set borders for data row
                    for (int col = 1; col <= 6; col++)
                    {
                        var cell = worksheet.Cells[row, col];
                        cell.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        cell.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        cell.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        cell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    }

                    row++;
                }

                worksheet.Cells.AutoFitColumns();
                package.Save();
            }

            stream.Position = 0;
            string fileName = $"List_User_{DateTime.Now.ToString("ddMMyyyyHHmmss")}.xlsx";
            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }
    }
}