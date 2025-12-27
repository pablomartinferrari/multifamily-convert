using System;
using System.Collections.Generic;

namespace ExcelProcessor.Models
{
    public class ProcessingRequest
    {
        public string SiteUrl { get; set; } = string.Empty;
        public List<string> FileUrls { get; set; } = new List<string>();
        public string UserId { get; set; } = string.Empty;
        public string JobNumber { get; set; } = string.Empty;
        public string FileType { get; set; } = string.Empty; // "Units" or "Common Areas"
    }

    public class ProcessingResponse
    {
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public List<GeneratedReport> GeneratedReports { get; set; } = new List<GeneratedReport>();
        public List<string> Errors { get; set; } = new List<string>();
    }

    public class GeneratedReport
    {
        public string ReportType { get; set; } = string.Empty;
        public string FileName { get; set; } = string.Empty;
        public string Url { get; set; } = string.Empty;
    }

    public class XrfShot
    {
        public int Reading { get; set; }
        public string Component { get; set; } = string.Empty;
        public string Side { get; set; } = string.Empty;
        public string Color { get; set; } = string.Empty;
        public string Substrate { get; set; } = string.Empty;
        public string Condition { get; set; } = string.Empty;
        public string RoomNumber { get; set; } = string.Empty;
        public string RoomType { get; set; } = string.Empty;
        public string Floor { get; set; } = string.Empty;
        public string Result { get; set; } = string.Empty; // "Pos" or "Neg"
        public double Pbc { get; set; } // Lead content
        
        public bool IsPositive => Result.Equals("Pos", StringComparison.OrdinalIgnoreCase) || Pbc >= 1.0;
        public bool IsCalibration => Component.Equals("CALIBRATE", StringComparison.OrdinalIgnoreCase);
    }

    public class ComponentSummary
    {
        public string Component { get; set; } = string.Empty;
        public int Count { get; set; }
        public double NegativePercentage { get; set; }
        public double PositivePercentage { get; set; }
        public string LeadContent { get; set; } = string.Empty; // "Positive" or "Negative"
    }

    public class ProcessingResults
    {
        public List<ComponentSummary> AveragedResults { get; set; } = new();
        public List<ComponentSummary> UniformResults { get; set; } = new();
        public List<XrfShot> ConflictingResults { get; set; } = new();
    }
}


