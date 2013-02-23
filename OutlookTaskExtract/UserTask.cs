using CsvHelper.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OutlookTaskExtract
{
  /// <summary>
  /// Represents an Exchange task. This class uses CsvHelper 
  /// filed attributes to help format column headers.
  /// </summary>
  public class UserTask
  {
    public String Subject { get; set; }

    [CsvField(Name = "Start Date")]
    public DateTime? StartDate { get; set; }

    [CsvField(Name = "Due Date")]
    public DateTime? DueDate { get; set; }

    [CsvField(Name = "Reminder On/Off")]
    public Boolean ReminderOn { get; set; }

    [CsvField(Name = "Reminder Date")]
    public DateTime ReminderDate { get; set; }

    [CsvField(Name = "Reminder Time")]
    public Int32 ReminderMinutes { get; set; }

    [CsvField(Name = "Date Completed")]
    public DateTime? DateCompleted { get; set; }

    [CsvField(Name = "% Complete")]
    public Double PercentComplete { get; set; }

    [CsvField(Name = "Total Work")]
    public Int32? TotalWork { get; set; }

    [CsvField(Name = "Acutal Work")]
    public Int32? AcutalWork { get; set; }

    [CsvField(Name = "Billing Information")]
    public String BillingInfo { get; set; }

    public String Categories { get; set; }
    public String Contacts { get; set; }
    public String Mileage { get; set; }
    public String Notes { get; set; }
    public String Priority { get; set; }

    [CsvField(Name = "Private")]
    public Boolean IsPrivate { get; set; }

    public String Role { get; set; }

    [CsvField(Name = "Schedule+ Priority")]
    public String SchedulePriority { get; set; }

    public String Sensitivity { get; set; }
    public String Status { get; set; }
  }
}
