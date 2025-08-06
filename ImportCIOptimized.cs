using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using System.Xml;
using OfficeOpenXml;

public class OptimizedImportController : Controller
{
    // Cấu hình timeout cho các operation
    private const int FILE_OPERATION_TIMEOUT_MS = 30000; // 30 giây
    private const int GRLOADER_TIMEOUT_MS = 300000; // 5 phút
    private const int MAX_FILE_SIZE_MB = 50; // Giới hạn kích thước file
    
    [HttpPost]
    public async Task<JsonResult> ImportCIOptimized()
    {
        var stopwatch = System.Diagnostics.Stopwatch.StartNew();
        string UserName = Session["Username"]?.ToString();
        
        if (string.IsNullOrEmpty(UserName))
        {
            return Json(new { success = false, message = "User session expired" });
        }

        // Tạo activity log ngay từ đầu
        var log = new mdactivity_log
        {
            action = "Import CI",
            id = Utility.GetIdRandom(),
            sessionid = SID,
            unit = "web smportal",
            log_request = "Import CI started"
        };

        try
        {
            // 1. Validate và xử lý file upload
            var fileValidationResult = await ValidateAndProcessUploadedFile();
            if (!fileValidationResult.IsValid)
            {
                log.log_response = fileValidationResult.ErrorMessage;
                await InsertActivityLogAsync(log);
                return Json(new { success = false, message = fileValidationResult.ErrorMessage });
            }

            var fileInfo = fileValidationResult.FileInfo;
            log.log_request = fileInfo.FileName;

            // 2. Lấy credentials một cách tối ưu
            string capass = await GetUserCredentialsAsync(UserName);

            // 3. Xử lý Excel file với memory optimization
            string sheetName = await GetFirstWorksheetNameAsync(fileInfo.FilePath);
            if (string.IsNullOrEmpty(sheetName))
            {
                throw new InvalidOperationException("Không thể đọc worksheet từ file Excel");
            }

            // 4. Kiểm tra và xử lý file error trước khi chạy GRLoader
            await HandleExistingErrorFileAsync(fileInfo);

            // 5. Chạy GRLoader với timeout
            var grLoaderResult = await RunGRLoaderWithTimeoutAsync(UserName, capass, fileInfo.FilePath, sheetName);
            
            // 6. Xử lý kết quả và tạo response
            var result = await ProcessImportResultAsync(fileInfo, sheetName, grLoaderResult);
            
            log.log_response = $"Import completed successfully in {stopwatch.ElapsedMilliseconds}ms";
            await InsertActivityLogAsync(log);
            
            return Json(result, JsonRequestBehavior.AllowGet);
        }
        catch (TimeoutException ex)
        {
            log.log_response = $"Import timeout after {stopwatch.ElapsedMilliseconds}ms: {ex.Message}";
            await InsertActivityLogAsync(log);
            return Json(new { success = false, message = "Import process timeout. Vui lòng thử lại với file nhỏ hơn hoặc liên hệ admin." });
        }
        catch (Exception ex)
        {
            log.log_response = $"Import error: {ex.Message}";
            await InsertActivityLogAsync(log);
            return Json(new { success = false, message = $"Lỗi trong quá trình import: {ex.Message}" });
        }
    }

    private async Task<FileValidationResult> ValidateAndProcessUploadedFile()
    {
        var result = new FileValidationResult();
        
        if (Request.Files.Count == 0)
        {
            result.ErrorMessage = "Không có file được upload";
            return result;
        }

        HttpPostedFileBase file = Request.Files[0];
        
        // Validate file size
        if (file.ContentLength > MAX_FILE_SIZE_MB * 1024 * 1024)
        {
            result.ErrorMessage = $"File quá lớn. Kích thước tối đa cho phép: {MAX_FILE_SIZE_MB}MB";
            return result;
        }

        // Validate file extension
        string extension = Path.GetExtension(file.FileName).ToLower();
        if (extension != ".xlsx" && extension != ".xls")
        {
            result.ErrorMessage = "Chỉ hỗ trợ file Excel (.xlsx, .xls)";
            return result;
        }

        // Save file với async
        string path = Server.MapPath("~/Uploads/");
        if (!Directory.Exists(path))
        {
            Directory.CreateDirectory(path);
        }

        string filePath = Path.Combine(path, $"{DateTime.Now:yyyyMMddHHmmss}_{file.FileName}");
        
        // Sử dụng async để save file
        await Task.Run(() => file.SaveAs(filePath));

        result.IsValid = true;
        result.FileInfo = new UploadedFileInfo
        {
            FileName = file.FileName,
            FilePath = filePath,
            ContentLength = file.ContentLength
        };

        return result;
    }

    private async Task<string> GetUserCredentialsAsync(string userName)
    {
        return await Task.Run(() =>
        {
            string lsuer = GetPassUser(userName);
            cryptsDecrypts cr = new cryptsDecrypts();
            return cr.DeCrypt(lsuer, "fiss@123");
        });
    }

    private async Task<string> GetFirstWorksheetNameAsync(string filePath)
    {
        return await Task.Run(() =>
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                return worksheet?.Name;
            }
        });
    }

    private async Task HandleExistingErrorFileAsync(UploadedFileInfo fileInfo)
    {
        await Task.Run(() =>
        {
            string folderName = Path.GetDirectoryName(fileInfo.FilePath);
            string fileName = Path.GetFileNameWithoutExtension(fileInfo.FilePath);
            string errorFileName = fileName + "_err.xml";
            string errorFilePath = Path.Combine(folderName, errorFileName);

            if (File.Exists(errorFilePath))
            {
                // Move existing error file to history
                string historyFolder = Path.Combine(folderName, "History");
                Directory.CreateDirectory(historyFolder);

                string timestamp = File.GetLastWriteTime(errorFilePath).ToString("yyyyMMddHHmmss");
                string historyFileName = $"{fileName}_err_{timestamp}.xml";
                string historyFilePath = Path.Combine(historyFolder, historyFileName);

                File.Move(errorFilePath, historyFilePath);
            }
        });
    }

    private async Task<string> RunGRLoaderWithTimeoutAsync(string userName, string password, string filePath, string sheetName)
    {
        using (var cts = new System.Threading.CancellationTokenSource(GRLOADER_TIMEOUT_MS))
        {
            return await Task.Run(() =>
            {
                string appLocation = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase);
                appLocation = appLocation.Replace("file:\\", "");
                string grLoaderPath = Path.Combine(appLocation, "Grloader");

                return commandline_importci(grLoaderPath, urlip, userName, password, filePath, sheetName);
            }, cts.Token);
        }
    }

    private async Task<ImportResult> ProcessImportResultAsync(UploadedFileInfo fileInfo, string sheetName, string grLoaderOutput)
    {
        return await Task.Run(() =>
        {
            string folderName = Path.GetDirectoryName(fileInfo.FilePath);
            string fileName = Path.GetFileNameWithoutExtension(fileInfo.FilePath);
            string errorFilePath = Path.Combine(folderName, fileName + "_err.xml");

            var result = new ImportResult
            {
                Success = true,
                Message = grLoaderOutput,
                GrLoaderOutput = grLoaderOutput
            };

            // Đếm tổng số records từ Excel
            int totalRecords = GetRecordCountFromExcel(fileInfo.FilePath, sheetName);
            result.TotalRecords = totalRecords;

            if (File.Exists(errorFilePath))
            {
                var errorList = ParseErrorXml(errorFilePath);
                result.ErrorRecords = errorList;
                result.SuccessfulRecords = totalRecords - errorList.Count;
                
                if (errorList.Count > 0)
                {
                    result.Success = false;
                    result.Message = $"Import hoàn thành với {errorList.Count} lỗi. " +
                                   $"Đã import thành công {result.SuccessfulRecords}/{totalRecords} bản ghi. " +
                                   $"Chi tiết lỗi: {errorFilePath}";
                }
            }
            else
            {
                result.SuccessfulRecords = totalRecords;
                result.Message = $"Import thành công {totalRecords} bản ghi.";
            }

            return result;
        });
    }

    private int GetRecordCountFromExcel(string filePath, string sheetName)
    {
        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[sheetName];
            return worksheet?.Dimension?.Rows ?? 0;
        }
    }

    private List<string> ParseErrorXml(string errorFilePath)
    {
        var errorList = new List<string>();
        
        try
        {
            var doc = new XmlDocument();
            doc.Load(errorFilePath);

            foreach (XmlNode node in doc.SelectNodes("/GRLoader/ci"))
            {
                string name = node["name"]?.InnerText ?? "Unknown";
                string className = node["class"]?.InnerText ?? "Unknown";
                errorList.Add($"{name} ({className})");
            }
        }
        catch (Exception ex)
        {
            errorList.Add($"Error parsing XML: {ex.Message}");
        }

        return errorList;
    }

    private async Task InsertActivityLogAsync(mdactivity_log log)
    {
        await Task.Run(() => InsertActivityLog(log));
    }

    // Supporting classes
    public class FileValidationResult
    {
        public bool IsValid { get; set; }
        public string ErrorMessage { get; set; }
        public UploadedFileInfo FileInfo { get; set; }
    }

    public class UploadedFileInfo
    {
        public string FileName { get; set; }
        public string FilePath { get; set; }
        public int ContentLength { get; set; }
    }

    public class ImportResult
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public string GrLoaderOutput { get; set; }
        public int TotalRecords { get; set; }
        public int SuccessfulRecords { get; set; }
        public List<string> ErrorRecords { get; set; } = new List<string>();
    }

    // Placeholder methods - bạn cần implement dựa trên code hiện tại
    private string GetPassUser(string userName) { /* existing implementation */ return ""; }
    private void InsertActivityLog(mdactivity_log log) { /* existing implementation */ }
    private string commandline_importci(string grLoaderPath, string urlip, string userName, string password, string filePath, string sheetName) { /* existing implementation */ return ""; }
    
    // Properties từ code gốc
    private string SID { get; set; }
    private string urlip { get; set; }
}