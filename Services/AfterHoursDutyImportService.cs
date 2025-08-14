using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using OfficeOpenXml;

public class AfterHoursDutyImportService
{
    private static readonly string[] RequiredHeaders = new[]
    {
        "DutyDate","StaffCode","FullName","Department","StartTime","EndTime","DutyType","Notes"
    };

    public async Task<AfterHoursDutyImportResult> ImportAsync(string filePath)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        return await Task.Run(() =>
        {
            var result = new AfterHoursDutyImportResult();
            var rowResults = new List<AfterHoursDutyImportRowResult>();
            var validRecords = new List<AfterHoursDuty>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                if (worksheet == null)
                {
                    throw new InvalidOperationException("Không tìm thấy worksheet trong file Excel");
                }

                // Build header map
                var headerMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    var header = worksheet.Cells[1, col].Value?.ToString()?.Trim() ?? string.Empty;
                    if (!string.IsNullOrEmpty(header))
                    {
                        headerMap[NormalizeHeader(header)] = col;
                    }
                }

                // Validate required headers
                var missingHeaders = RequiredHeaders.Where(h => !headerMap.ContainsKey(h)).ToList();
                if (missingHeaders.Any())
                {
                    throw new InvalidOperationException($"Thiếu cột bắt buộc: {string.Join(", ", missingHeaders)}");
                }

                // Read rows
                int lastRow = worksheet.Dimension.End.Row;
                for (int row = 2; row <= lastRow; row++)
                {
                    var rowResult = new AfterHoursDutyImportRowResult { RowNumber = row, IsValid = true };

                    // Skip empty rows
                    bool isEmpty = IsRowEmpty(worksheet, row);
                    if (isEmpty)
                    {
                        continue;
                    }

                    var record = new AfterHoursDuty();
                    // DutyDate
                    var dutyDateObj = worksheet.Cells[row, headerMap["DutyDate"]].Value;
                    if (!TryParseDate(dutyDateObj, out DateTime dutyDate))
                    {
                        rowResult.IsValid = false;
                        rowResult.Errors.Add("DutyDate không hợp lệ (yyyy-MM-dd hoặc dd/MM/yyyy)");
                    }
                    else
                    {
                        record.DutyDate = dutyDate.Date;
                    }

                    // StaffCode
                    record.StaffCode = (worksheet.Cells[row, headerMap["StaffCode"]].Value?.ToString() ?? string.Empty).Trim();
                    if (string.IsNullOrWhiteSpace(record.StaffCode))
                    {
                        rowResult.IsValid = false;
                        rowResult.Errors.Add("StaffCode bắt buộc");
                    }

                    // FullName
                    record.FullName = (worksheet.Cells[row, headerMap["FullName"]].Value?.ToString() ?? string.Empty).Trim();
                    if (string.IsNullOrWhiteSpace(record.FullName))
                    {
                        rowResult.IsValid = false;
                        rowResult.Errors.Add("FullName bắt buộc");
                    }

                    // Department
                    record.Department = (worksheet.Cells[row, headerMap["Department"]].Value?.ToString() ?? string.Empty).Trim();

                    // StartTime
                    var startObj = worksheet.Cells[row, headerMap["StartTime"]].Value;
                    if (!TryParseTime(startObj, out TimeSpan startTime))
                    {
                        rowResult.IsValid = false;
                        rowResult.Errors.Add("StartTime không hợp lệ (HH:mm)");
                    }
                    else
                    {
                        record.StartTime = startTime;
                    }

                    // EndTime
                    var endObj = worksheet.Cells[row, headerMap["EndTime"]].Value;
                    if (!TryParseTime(endObj, out TimeSpan endTime))
                    {
                        rowResult.IsValid = false;
                        rowResult.Errors.Add("EndTime không hợp lệ (HH:mm)");
                    }
                    else
                    {
                        record.EndTime = endTime;
                    }

                    // DutyType
                    record.DutyType = (worksheet.Cells[row, headerMap["DutyType"]].Value?.ToString() ?? string.Empty).Trim();

                    // Notes
                    record.Notes = (worksheet.Cells[row, headerMap["Notes"]].Value?.ToString() ?? string.Empty).Trim();

                    // Business validations
                    if (rowResult.IsValid)
                    {
                        if (record.EndTime <= record.StartTime)
                        {
                            rowResult.IsValid = false;
                            rowResult.Errors.Add("EndTime phải lớn hơn StartTime");
                        }
                    }

                    if (rowResult.IsValid)
                    {
                        rowResult.Data = record;
                        validRecords.Add(record);
                    }

                    rowResults.Add(rowResult);
                }
            }

            // Save valid records (replace with real persistence)
            int saved = SaveRecords(validRecords);

            result.TotalRows = rowResults.Count;
            result.ValidRows = rowResults.Count(r => r.IsValid);
            result.SavedRows = saved;
            result.Rows = rowResults;
            return result;
        });
    }

    private static string NormalizeHeader(string header)
    {
        // Map common localized headers to canonical names
        string h = (header ?? string.Empty).Trim();
        switch (h.ToLowerInvariant())
        {
            case "ngày":
            case "date":
            case "dutydate":
                return "DutyDate";
            case "manv":
            case "mãnv":
            case "staffcode":
            case "staff code":
                return "StaffCode";
            case "họ tên":
            case "họ và tên":
            case "fullname":
            case "full name":
                return "FullName";
            case "phòng ban":
            case "department":
                return "Department";
            case "bắt đầu":
            case "start":
            case "starttime":
            case "start time":
                return "StartTime";
            case "kết thúc":
            case "end":
            case "endtime":
            case "end time":
                return "EndTime";
            case "loại trực":
            case "dutytype":
            case "duty type":
                return "DutyType";
            case "ghi chú":
            case "notes":
                return "Notes";
            default:
                return h.Replace(" ", string.Empty);
        }
    }

    private static bool IsRowEmpty(OfficeOpenXml.ExcelWorksheet worksheet, int row)
    {
        int startCol = worksheet.Dimension.Start.Column;
        int endCol = worksheet.Dimension.End.Column;
        for (int col = startCol; col <= endCol; col++)
        {
            var val = worksheet.Cells[row, col].Value;
            if (val != null && !string.IsNullOrWhiteSpace(val.ToString()))
            {
                return false;
            }
        }
        return true;
    }

    private static bool TryParseDate(object value, out DateTime date)
    {
        date = default;
        if (value == null)
        {
            return false;
        }
        if (value is DateTime dt)
        {
            date = dt.Date;
            return true;
        }
        // Excel numeric date
        if (double.TryParse(value.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out double oa))
        {
            try
            {
                date = DateTime.FromOADate(oa).Date;
                return true;
            }
            catch { }
        }
        var s = value.ToString().Trim();
        string[] formats = { "yyyy-MM-dd", "dd/MM/yyyy", "M/d/yyyy", "d/M/yyyy", "dd-MM-yyyy", "MM/dd/yyyy" };
        if (DateTime.TryParseExact(s, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
        {
            date = dt.Date;
            return true;
        }
        if (DateTime.TryParse(s, out dt))
        {
            date = dt.Date;
            return true;
        }
        return false;
    }

    private static bool TryParseTime(object value, out TimeSpan time)
    {
        time = default;
        if (value == null)
        {
            return false;
        }
        if (value is TimeSpan ts)
        {
            time = ts;
            return true;
        }
        if (value is DateTime dt)
        {
            time = dt.TimeOfDay;
            return true;
        }
        // Excel numeric time (fraction of day)
        if (double.TryParse(value.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out double fraction))
        {
            if (fraction >= 0 && fraction < 1)
            {
                time = TimeSpan.FromDays(fraction);
                return true;
            }
        }
        var s = value.ToString().Trim();
        string[] formats = { "HH:mm", "H:mm", "HH:mm:ss", "H:mm:ss" };
        if (DateTime.TryParseExact(s, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsed))
        {
            time = parsed.TimeOfDay;
            return true;
        }
        if (TimeSpan.TryParse(s, out ts))
        {
            time = ts;
            return true;
        }
        return false;
    }

    private static int SaveRecords(List<AfterHoursDuty> records)
    {
        // Replace with actual persistence to your database
        // For now assume all valid records are saved successfully
        return records.Count;
    }
}