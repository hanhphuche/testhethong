using System;
using System.Collections.Generic;

public class AfterHoursDuty
{
    public DateTime DutyDate { get; set; }
    public string StaffCode { get; set; }
    public string FullName { get; set; }
    public string Department { get; set; }
    public TimeSpan StartTime { get; set; }
    public TimeSpan EndTime { get; set; }
    public string DutyType { get; set; }
    public string Notes { get; set; }
}

public class AfterHoursDutyImportRowResult
{
    public int RowNumber { get; set; }
    public bool IsValid { get; set; }
    public List<string> Errors { get; set; } = new List<string>();
    public AfterHoursDuty Data { get; set; }
}

public class AfterHoursDutyImportResult
{
    public int TotalRows { get; set; }
    public int ValidRows { get; set; }
    public int SavedRows { get; set; }
    public List<AfterHoursDutyImportRowResult> Rows { get; set; } = new List<AfterHoursDutyImportRowResult>();
}