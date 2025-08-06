using System;
using System.Configuration;

public static class ImportCIConfiguration
{
    // Timeout configurations
    public static int FileOperationTimeoutMs => 
        GetConfigValue("FileOperationTimeout", 30000); // 30 seconds

    public static int GRLoaderTimeoutMs => 
        GetConfigValue("GRLoaderTimeout", 300000); // 5 minutes

    public static int DatabaseTimeoutMs => 
        GetConfigValue("DatabaseTimeout", 60000); // 1 minute

    // File handling configurations
    public static int MaxFileSizeMB => 
        GetConfigValue("MaxFileSizeMB", 50);

    public static int MaxBatchSize => 
        GetConfigValue("MaxBatchSize", 1000);

    public static string[] AllowedFileExtensions => 
        GetConfigValue("AllowedFileExtensions", ".xlsx,.xls").Split(',');

    // Performance configurations
    public static bool EnableAsyncProcessing => 
        GetConfigValue("EnableAsyncProcessing", "true").ToLower() == "true";

    public static bool EnableMemoryOptimization => 
        GetConfigValue("EnableMemoryOptimization", "true").ToLower() == "true";

    public static int ChunkSize => 
        GetConfigValue("ChunkSize", 500);

    // Retry configurations
    public static int MaxRetryAttempts => 
        GetConfigValue("MaxRetryAttempts", 3);

    public static int RetryDelayMs => 
        GetConfigValue("RetryDelayMs", 5000);

    // Logging configurations
    public static bool EnableDetailedLogging => 
        GetConfigValue("EnableDetailedLogging", "false").ToLower() == "true";

    public static string LogLevel => 
        GetConfigValue("LogLevel", "Info");

    private static int GetConfigValue(string key, int defaultValue)
    {
        var value = ConfigurationManager.AppSettings[key];
        return int.TryParse(value, out int result) ? result : defaultValue;
    }

    private static string GetConfigValue(string key, string defaultValue)
    {
        return ConfigurationManager.AppSettings[key] ?? defaultValue;
    }
}

// Extension class cho Web.config
/*
Thêm vào Web.config:

<appSettings>
  <!-- Import CI Timeout Settings -->
  <add key="FileOperationTimeout" value="30000" />
  <add key="GRLoaderTimeout" value="300000" />
  <add key="DatabaseTimeout" value="60000" />
  
  <!-- File Handling Settings -->
  <add key="MaxFileSizeMB" value="50" />
  <add key="MaxBatchSize" value="1000" />
  <add key="AllowedFileExtensions" value=".xlsx,.xls" />
  
  <!-- Performance Settings -->
  <add key="EnableAsyncProcessing" value="true" />
  <add key="EnableMemoryOptimization" value="true" />
  <add key="ChunkSize" value="500" />
  
  <!-- Retry Settings -->
  <add key="MaxRetryAttempts" value="3" />
  <add key="RetryDelayMs" value="5000" />
  
  <!-- Logging Settings -->
  <add key="EnableDetailedLogging" value="false" />
  <add key="LogLevel" value="Info" />
</appSettings>
*/