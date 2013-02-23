using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OutlookTaskExtract;
using System.IO;
using CsvHelper;
using System.Collections.Generic;
using Microsoft.Exchange.WebServices.Data;

namespace OutlookTests
{
  [TestClass]
  public class OutlookTaskTests
  {
    [TestMethod]
    public void readUsersCSV()
    {
      // read csv into list
      String csvFile = AppDomain.CurrentDomain.BaseDirectory + @"\Resources\Users.csv";
      var csv = new CsvReader(new StreamReader(csvFile));
      var usersList = csv.GetRecords<User>();
      Int32 count = 0;
      List<String> users = new List<string>();

      foreach (User user in usersList)
      {
        users.Add(user.Email);
        StringAssert.Contains(user.Email, "@");
        StringAssert.Contains(user.Email, ".");
      }

      // assert list size of 50
      Assert.AreEqual(users.Count, 50);
    }

    [TestMethod]
    public void testUser()
    {
      User user = new User();
      user.Email = String.Empty;

      Assert.AreEqual(user.Email, String.Empty);
    }

    [TestMethod]
    public void testTask()
    {
      OutlookTaskExtract.UserTask task = new OutlookTaskExtract.UserTask();
      task.Subject = String.Empty;
      task.StartDate = DateTime.MaxValue;
      task.DueDate = DateTime.MaxValue;
      task.ReminderOn = true;
      task.ReminderDate = DateTime.MaxValue;
      task.ReminderMinutes = 0;
      task.DateCompleted = DateTime.MaxValue;
      task.PercentComplete = 0;
      task.TotalWork = 0;
      task.AcutalWork = 0;
      task.BillingInfo = String.Empty;
      task.Categories = String.Empty;
      task.Contacts = String.Empty;
      task.Mileage = String.Empty;
      task.Notes = String.Empty;
      task.Sensitivity = String.Empty;
      task.Status = String.Empty;
      task.Priority = String.Empty;
      task.IsPrivate = false;
      task.Role = String.Empty;
      task.SchedulePriority = String.Empty;

      Assert.AreEqual(task.Subject, String.Empty);
      Assert.AreEqual(task.StartDate, DateTime.MaxValue);
      Assert.AreEqual(task.DueDate, DateTime.MaxValue);
      Assert.AreEqual(task.ReminderOn, true);
      Assert.AreEqual(task.ReminderDate, DateTime.MaxValue);
      Assert.AreEqual(task.ReminderMinutes, 0);
      Assert.AreEqual(task.DateCompleted, DateTime.MaxValue);
      Assert.AreEqual(task.PercentComplete, 0);
      Assert.AreEqual(task.TotalWork, 0);
      Assert.AreEqual(task.AcutalWork, 0);
      Assert.AreEqual(task.BillingInfo, String.Empty);
      Assert.AreEqual(task.Categories, String.Empty);
      Assert.AreEqual(task.Contacts, String.Empty);
      Assert.AreEqual(task.Mileage, String.Empty);
      Assert.AreEqual(task.Notes, String.Empty);
      Assert.AreEqual(task.Priority, String.Empty);
      Assert.AreEqual(task.IsPrivate, false);
      Assert.AreEqual(task.Role, String.Empty);
      Assert.AreEqual(task.SchedulePriority, String.Empty);
      Assert.AreEqual(task.Sensitivity, String.Empty);
      Assert.AreEqual(task.Status, String.Empty);
    }

    [TestMethod]
    public void testFormProperties()
    {
      MainForm form = new MainForm();

      Assert.AreEqual(form.Service, null);
      Assert.AreEqual(form.URI, String.Empty);
      Assert.AreEqual(form.AutoDiscoverURL.Checked, true);
      Assert.AreEqual(form.ExchangeServerURL.Enabled, false);

      form.AutoDiscoverURL.Checked = false;
      Assert.AreEqual(form.AutoDiscoverURL.Checked, false);
      Assert.AreEqual(form.ExchangeServerURL.Enabled, true);

      form.AutoDiscoverURL.Checked = true;
      Assert.AreEqual(form.AutoDiscoverURL.Checked, true);
      Assert.AreEqual(form.ExchangeServerURL.Enabled, false);
    }

    //[TestMethod]
    //public void testgetExchangeVersions()
    //{
    //    MainForm form = new MainForm();

    //    form.ExchangeVersions.SelectedIndex = 0;
    //    Assert.AreEqual(form.Service.RequestedServerVersion, ExchangeVersion.Exchange2010);

    //    form.ExchangeVersions.SelectedIndex = 1;
    //    Assert.AreEqual(form.Service.RequestedServerVersion, ExchangeVersion.Exchange2010_SP1);

    //    form.ExchangeVersions.SelectedIndex = 2;
    //    Assert.AreEqual(form.Service.RequestedServerVersion, ExchangeVersion.Exchange2010_SP2);
    //}
  }
}
