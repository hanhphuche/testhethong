using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using OfficeOpenXml;

public class ImportCIService
{
    private readonly ILogger _logger;
    
    public ImportCIService(ILogger logger = null)
    {
        _logger = logger ?? new DefaultLogger();
    }

    /// <summary>
    /// Xử lý import CI với batch processing để tránh timeout
    /// </summary>
    public async Task<ImportResult> ProcessImportWithBatchingAsync(
        string filePath, 
        string sheetName, 
        string userName, 
        string password,
        CancellationToken cancellationToken = default)
    {
        var result = new ImportResult();
        var stopwatch = System.Diagnostics.Stopwatch.StartNew();
        
        try
        {
            _logger.Info($"Starting batch import for file: {filePath}");
            
            // 1. Đọc và validate Excel data
            var excelData = await ReadExcelDataAsync(filePath, sheetName, cancellationToken);
            result.TotalRecords = excelData.Count;
            
            if (excelData.Count == 0)
            {
                result.Success = false;
                result.Message = "Không có dữ liệu để import";
                return result;
            }
            
            // 2. Chia data thành các batch nhỏ
            var batches = CreateBatches(excelData, ImportCIConfiguration.ChunkSize);
            _logger.Info($"Created {batches.Count} batches with chunk size {ImportCIConfiguration.ChunkSize}");
            
            // 3. Xử lý từng batch
            var allErrors = new List<string>();
            int successfulRecords = 0;
            
            for (int i = 0; i < batches.Count; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                
                _logger.Info($"Processing batch {i + 1}/{batches.Count}");
                
                var batchResult = await ProcessBatchWithRetryAsync(
                    batches[i], 
                    userName, 
                    password, 
                    i + 1,
                    cancellationToken);
                
                successfulRecords += batchResult.SuccessfulRecords;
                allErrors.AddRange(batchResult.ErrorRecords);
                
                // Thêm delay nhỏ giữa các batch để giảm tải
                if (i < batches.Count - 1)
                {
                    await Task.Delay(1000, cancellationToken);
                }
            }
            
            // 4. Tổng hợp kết quả
            result.SuccessfulRecords = successfulRecords;
            result.ErrorRecords = allErrors;
            result.Success = allErrors.Count == 0;
            result.ProcessingTimeMs = stopwatch.ElapsedMilliseconds;
            
            if (result.Success)
            {
                result.Message = $"Import thành công {successfulRecords}/{result.TotalRecords} bản ghi trong {stopwatch.ElapsedMilliseconds}ms";
            }
            else
            {
                result.Message = $"Import hoàn thành với {allErrors.Count} lỗi. " +
                               $"Thành công: {successfulRecords}/{result.TotalRecords} bản ghi";
            }
            
            _logger.Info($"Batch import completed: {result.Message}");
            return result;
        }
        catch (OperationCanceledException)
        {
            result.Success = false;
            result.Message = "Import process was cancelled";
            _logger.Warn("Import process was cancelled");
            return result;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.Message = $"Lỗi trong quá trình import: {ex.Message}";
            _logger.Error($"Import failed: {ex}");
            return result;
        }
    }
    
    /// <summary>
    /// Đọc dữ liệu từ Excel với memory optimization
    /// </summary>
    private async Task<List<DataRow>> ReadExcelDataAsync(string filePath, string sheetName, CancellationToken cancellationToken)
    {
        return await Task.Run(() =>
        {
            var dataList = new List<DataRow>();
            
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[sheetName];
                if (worksheet == null)
                {
                    throw new InvalidOperationException($"Worksheet '{sheetName}' not found");
                }
                
                var dataTable = new DataTable();
                
                // Đọc header
                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    var headerValue = worksheet.Cells[1, col].Value?.ToString() ?? $"Column{col}";
                    dataTable.Columns.Add(headerValue);
                }
                
                // Đọc dữ liệu từng dòng để tiết kiệm memory
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    
                    var dataRow = dataTable.NewRow();
                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        dataRow[col - 1] = worksheet.Cells[row, col].Value?.ToString() ?? "";
                    }
                    dataList.Add(dataRow);
                }
            }
            
            return dataList;
        }, cancellationToken);
    }
    
    /// <summary>
    /// Chia dữ liệu thành các batch nhỏ
    /// </summary>
    private List<List<DataRow>> CreateBatches(List<DataRow> data, int batchSize)
    {
        var batches = new List<List<DataRow>>();
        
        for (int i = 0; i < data.Count; i += batchSize)
        {
            var batch = data.Skip(i).Take(batchSize).ToList();
            batches.Add(batch);
        }
        
        return batches;
    }
    
    /// <summary>
    /// Xử lý một batch với retry logic
    /// </summary>
    private async Task<BatchResult> ProcessBatchWithRetryAsync(
        List<DataRow> batch, 
        string userName, 
        string password, 
        int batchNumber,
        CancellationToken cancellationToken)
    {
        var maxRetries = ImportCIConfiguration.MaxRetryAttempts;
        var retryDelay = ImportCIConfiguration.RetryDelayMs;
        
        for (int attempt = 1; attempt <= maxRetries; attempt++)
        {
            try
            {
                _logger.Info($"Processing batch {batchNumber}, attempt {attempt}/{maxRetries}");
                
                // Tạo file tạm cho batch này
                var tempFilePath = await CreateTempFileForBatchAsync(batch, batchNumber, cancellationToken);
                
                try
                {
                    // Chạy GRLoader cho batch này
                    var grLoaderResult = await RunGRLoaderForBatchAsync(
                        tempFilePath, 
                        userName, 
                        password,
                        cancellationToken);
                    
                    // Xử lý kết quả
                    var batchResult = await ProcessBatchResultAsync(tempFilePath, batch.Count);
                    
                    _logger.Info($"Batch {batchNumber} completed successfully on attempt {attempt}");
                    return batchResult;
                }
                finally
                {
                    // Cleanup temp file
                    if (File.Exists(tempFilePath))
                    {
                        File.Delete(tempFilePath);
                    }
                }
            }
            catch (Exception ex) when (attempt < maxRetries)
            {
                _logger.Warn($"Batch {batchNumber} failed on attempt {attempt}: {ex.Message}. Retrying in {retryDelay}ms...");
                await Task.Delay(retryDelay, cancellationToken);
            }
        }
        
        // Nếu tất cả attempts đều fail
        _logger.Error($"Batch {batchNumber} failed after {maxRetries} attempts");
        return new BatchResult
        {
            SuccessfulRecords = 0,
            ErrorRecords = batch.Select((_, index) => $"Batch {batchNumber} Row {index + 1}: Failed after {maxRetries} attempts").ToList()
        };
    }
    
    /// <summary>
    /// Tạo file Excel tạm cho một batch
    /// </summary>
    private async Task<string> CreateTempFileForBatchAsync(List<DataRow> batch, int batchNumber, CancellationToken cancellationToken)
    {
        return await Task.Run(() =>
        {
            var tempPath = Path.GetTempPath();
            var tempFileName = $"ImportCI_Batch_{batchNumber}_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
            var tempFilePath = Path.Combine(tempPath, tempFileName);
            
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Import");
                
                // Thêm header
                if (batch.Count > 0)
                {
                    var firstRow = batch[0];
                    for (int col = 0; col < firstRow.Table.Columns.Count; col++)
                    {
                        worksheet.Cells[1, col + 1].Value = firstRow.Table.Columns[col].ColumnName;
                    }
                    
                    // Thêm dữ liệu
                    for (int row = 0; row < batch.Count; row++)
                    {
                        for (int col = 0; col < batch[row].ItemArray.Length; col++)
                        {
                            worksheet.Cells[row + 2, col + 1].Value = batch[row][col]?.ToString();
                        }
                    }
                }
                
                package.SaveAs(new FileInfo(tempFilePath));
            }
            
            return tempFilePath;
        }, cancellationToken);
    }
    
    /// <summary>
    /// Chạy GRLoader cho một batch với timeout
    /// </summary>
    private async Task<string> RunGRLoaderForBatchAsync(string filePath, string userName, string password, CancellationToken cancellationToken)
    {
        using (var timeoutCts = new CancellationTokenSource(ImportCIConfiguration.GRLoaderTimeoutMs))
        using (var combinedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, timeoutCts.Token))
        {
            return await Task.Run(() =>
            {
                // Implement GRLoader call here
                // This should be the same as your existing commandline_importci method
                return "GRLoader completed successfully"; // Placeholder
            }, combinedCts.Token);
        }
    }
    
    /// <summary>
    /// Xử lý kết quả của một batch
    /// </summary>
    private async Task<BatchResult> ProcessBatchResultAsync(string filePath, int totalRecords)
    {
        return await Task.Run(() =>
        {
            var result = new BatchResult();
            
            // Tìm file error tương ứng
            var folderName = Path.GetDirectoryName(filePath);
            var fileName = Path.GetFileNameWithoutExtension(filePath);
            var errorFilePath = Path.Combine(folderName, fileName + "_err.xml");
            
            if (File.Exists(errorFilePath))
            {
                var errorList = ParseErrorXml(errorFilePath);
                result.ErrorRecords = errorList;
                result.SuccessfulRecords = totalRecords - errorList.Count;
                
                // Cleanup error file
                File.Delete(errorFilePath);
            }
            else
            {
                result.SuccessfulRecords = totalRecords;
                result.ErrorRecords = new List<string>();
            }
            
            return result;
        });
    }
    
    /// <summary>
    /// Parse error XML file
    /// </summary>
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
}

/// <summary>
/// Kết quả xử lý một batch
/// </summary>
public class BatchResult
{
    public int SuccessfulRecords { get; set; }
    public List<string> ErrorRecords { get; set; } = new List<string>();
}

/// <summary>
/// Enhanced ImportResult with additional metrics
/// </summary>
public class ImportResult
{
    public bool Success { get; set; }
    public string Message { get; set; }
    public string GrLoaderOutput { get; set; }
    public int TotalRecords { get; set; }
    public int SuccessfulRecords { get; set; }
    public List<string> ErrorRecords { get; set; } = new List<string>();
    public long ProcessingTimeMs { get; set; }
    public int BatchCount { get; set; }
}

/// <summary>
/// Simple logger interface and implementation
/// </summary>
public interface ILogger
{
    void Info(string message);
    void Warn(string message);
    void Error(string message);
}

public class DefaultLogger : ILogger
{
    public void Info(string message)
    {
        System.Diagnostics.Debug.WriteLine($"INFO: {DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}");
    }
    
    public void Warn(string message)
    {
        System.Diagnostics.Debug.WriteLine($"WARN: {DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}");
    }
    
    public void Error(string message)
    {
        System.Diagnostics.Debug.WriteLine($"ERROR: {DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}");
    }
}