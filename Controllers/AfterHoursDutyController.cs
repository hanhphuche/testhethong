using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using OfficeOpenXml;

public class AfterHoursDutyController : Controller
{
    [HttpGet]
    public ActionResult Upload()
    {
        return View();
    }

    [HttpGet]
    public FileResult Template()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using (var package = new ExcelPackage())
        {
            var ws = package.Workbook.Worksheets.Add("AfterHoursDuty");
            string[] headers = { "DutyDate", "StaffCode", "FullName", "Department", "StartTime", "EndTime", "DutyType", "Notes" };
            for (int i = 0; i < headers.Length; i++)
            {
                ws.Cells[1, i + 1].Value = headers[i];
            }
            ws.Cells[2, 1].Value = DateTime.Today.ToString("yyyy-MM-dd");
            ws.Cells[2, 2].Value = "NV001";
            ws.Cells[2, 3].Value = "Nguyen Van A";
            ws.Cells[2, 4].Value = "IT";
            ws.Cells[2, 5].Value = "18:00";
            ws.Cells[2, 6].Value = "22:00";
            ws.Cells[2, 7].Value = "Ngoai gio";
            ws.Cells[2, 8].Value = "Ghi chú";

            var stream = new MemoryStream();
            package.SaveAs(stream);
            stream.Position = 0;
            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "AfterHoursDuty_Template.xlsx");
        }
    }

    [HttpPost]
    public async Task<ActionResult> Import()
    {
        if (Request.Files.Count == 0)
        {
            return Json(new { success = false, message = "Vui lòng chọn file Excel" });
        }

        HttpPostedFileBase file = Request.Files[0];
        if (file == null || file.ContentLength == 0)
        {
            return Json(new { success = false, message = "File không hợp lệ" });
        }

        string extension = Path.GetExtension(file.FileName).ToLower();
        if (extension != ".xlsx" && extension != ".xls")
        {
            return Json(new { success = false, message = "Chỉ hỗ trợ file Excel (.xlsx, .xls)" });
        }

        string uploadDir = Server.MapPath("~/Uploads/AfterHoursDuty/");
        if (!Directory.Exists(uploadDir))
        {
            Directory.CreateDirectory(uploadDir);
        }
        string savedPath = Path.Combine(uploadDir, $"{DateTime.Now:yyyyMMddHHmmss}_{Path.GetFileName(file.FileName)}");
        file.SaveAs(savedPath);

        try
        {
            var service = new AfterHoursDutyImportService();
            var result = await service.ImportAsync(savedPath);
            return Json(new
            {
                success = true,
                message = $"Đã xử lý {result.SavedRows}/{result.ValidRows}/{result.TotalRows} dòng",
                total = result.TotalRows,
                valid = result.ValidRows,
                saved = result.SavedRows,
                errors = result.Rows.Where(r => !r.IsValid).Select(r => new { row = r.RowNumber, errors = r.Errors }).Take(50)
            });
        }
        catch (Exception ex)
        {
            return Json(new { success = false, message = ex.Message });
        }
    }
}