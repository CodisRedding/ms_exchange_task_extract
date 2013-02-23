using CsvHelper;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Text;
using System.Windows.Forms;

namespace OutlookTaskExtract
{
  /// <summary>
  /// All string resources are supplied via a resx file.
  /// </summary>
  public partial class MainForm : Form
  {
    public ExchangeService Service = null;
    public String URI = String.Empty;
    private String csvPathToUsers = String.Empty;
    private String csvPathToSaveTasks = String.Empty;

    public MainForm()
    {
      InitializeComponent();

      ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallBack;

      // Ensure that Exchange 2010 version is selected
      ExchangeVersions.SelectedIndex = 0;

      // Disable EWS URI textbox if auto discovery is selected
      ExchangeServerURL.Enabled = !AutoDiscoverURL.Checked;
    }

    /// <summary>
    /// Kicks of the second part of the workfllow process after authentication. 
    /// </summary>
    private void execute()
    {
      // Step 1. Prompt with confirmation
      DialogResult selCsvRes = MessageBox.Show(Resource.SelectCSVFileMsg,
                                               Resource.SelectCSVFile,
                                               MessageBoxButtons.OKCancel);

      // If anything other than OK is click, bail out of workflow
      if (selCsvRes != DialogResult.OK)
      {
        return;
      }

      // Step 2. Prompt user for the CSV file that contains the users to query tasks for.
      // This file should be a CSV file with a single header column 'Email' and a list
      // of email email addresses seperated by new lines.
      csvPathToUsers = promptForUserList();

      // If a path was not selected, bail out of workflow
      if (csvPathToUsers == String.Empty)
      {
        return;
      }

      // Read CSV using open source CSV utility:
      // https://github.com/JoshClose/CsvHelper
      var csv = new CsvReader(new StreamReader(csvPathToUsers));

      // Step 3. Inform user that they need that will be selecting a dir 
      // that's used to save queried task results
      DialogResult selDirRes = MessageBox.Show(Resource.SaveCSVToDir,
                                               Resource.SelectDirectory,
                                               MessageBoxButtons.OKCancel);

      // If anything other than OK is selected, bail out of workflow
      if (selDirRes != DialogResult.OK)
      {
        return;
      }

      // Prompt user for the directory where the extracted Exchange tasks 
      // will be written
      csvPathToSaveTasks = promptForSavePath();

      // If a path is not selected, bail out of workflow
      if (csvPathToSaveTasks == String.Empty)
      {
        return;
      }

      toolStripStatusLabel1.Text = Resource.Running;

      // Iterate through users
      foreach (User user in csv.GetRecords<User>())
      {
        // Query tasks from Exchange
        List<UserTask> tasks = getUserTasks(user);

        // Write tasks CSV file
        if (!writeTasksToCSV(user, tasks))
        {
          break;
        }
      }

      toolStripStatusLabel1.Text = Resource.Complete;

      // Open directory where CSV's were saved.
      Process.Start(csvPathToSaveTasks);
    }

    /// <summary>
    /// Opens a file dialog prompting user to select CSV 
    /// list of users.
    /// </summary>
    /// <returns>File path and name</returns>
    private String promptForUserList()
    {
      Stream stream = null;
      OpenFileDialog openFileDialog = new OpenFileDialog();

      openFileDialog.InitialDirectory = "c:\\";
      openFileDialog.Filter = "CSV files (*.csv)|*.csv";
      openFileDialog.FilterIndex = 0;
      openFileDialog.RestoreDirectory = true;

      try
      {
        if (openFileDialog.ShowDialog() == DialogResult.OK)
        {
          if ((stream = openFileDialog.OpenFile()) != null)
          {
            stream.Close();
            return openFileDialog.FileName;
          }
        }
      }
      catch (Exception ex)
      {
        MessageBox.Show(String.Format(Resource.FileCouldNotReadUnknown, ex.Message));
      }

      return String.Empty;
    }

    /// <summary>
    /// Opens a dialog prompting user to select a directory where
    /// to save extracted task CSV's
    /// </summary>
    /// <returns>Directory path</returns>
    private String promptForSavePath()
    {
      FolderBrowserDialog dirPath = new FolderBrowserDialog();

      if (dirPath.ShowDialog() == DialogResult.OK)
      {
        return dirPath.SelectedPath;
      }

      return String.Empty;
    }

    /// <summary>
    /// Calls out to Exchange via Exchange Web Service querying 
    /// tasks for provided user.
    /// </summary>
    /// <param name="user">The user that tasks are queried for from Exchange</param>
    /// <returns>A list of users tasks</returns>
    private List<UserTask> getUserTasks(User user)
    {
      Service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, user.Email);

      var items = Service.FindItems(WellKnownFolderName.Tasks,
                    new ItemView(100) { PropertySet = new PropertySet(BasePropertySet.FirstClassProperties) });

      PropertySet _customPropertySet = new PropertySet(BasePropertySet.FirstClassProperties,
        AppointmentSchema.MyResponseType,
        AppointmentSchema.IsMeeting,
        AppointmentSchema.ICalUid);
      _customPropertySet.RequestedBodyType = BodyType.Text;

      List<UserTask> userTasks = new List<UserTask>();

      foreach (Task task in items.Items)
      {
        UserTask userTask = new UserTask();

        task.Load(_customPropertySet);
        userTask.Subject = task.Subject;
        userTask.StartDate = task.StartDate;
        userTask.DueDate = task.DueDate;
        userTask.ReminderOn = task.IsReminderSet;
        try
        {
          userTask.ReminderDate = task.ReminderDueBy;
        }
        catch { }
        userTask.ReminderMinutes = task.ReminderMinutesBeforeStart;
        userTask.DateCompleted = task.CompleteDate;
        userTask.PercentComplete = task.PercentComplete;
        userTask.TotalWork = task.TotalWork;
        userTask.AcutalWork = task.ActualWork;
        userTask.BillingInfo = task.BillingInformation;
        userTask.Categories = task.Categories.ToString();
        userTask.Contacts = task.Contacts.ToString();
        userTask.Mileage = task.Mileage;
        userTask.Notes = task.Body;
        userTask.Sensitivity = task.Sensitivity.ToString();
        userTask.Status = task.Status.ToString();
        userTask.Priority = task.Importance.ToString();
        userTask.IsPrivate = (userTask.Sensitivity == "Private");

        userTasks.Add(userTask);
      }

      return userTasks;
    }

    /// <summary>
    /// Writes users Exchange tasks to a CSV file.
    /// </summary>
    /// <param name="user">The users email address is used as part of the CSV file name.</param>
    /// <param name="tasks">The tasks that are written to a CSV file.</param>
    /// <returns></returns>
    private Boolean writeTasksToCSV(User user, List<UserTask> tasks)
    {
      Boolean success = true;
      String userCsv = Path.Combine(csvPathToSaveTasks, user.Email.Split('@')[0].ToLower() + ".csv");

      try
      {
        System.IO.File.WriteAllText(userCsv, "", Encoding.UTF8);

        using (var csv = new CsvWriter(new StreamWriter(userCsv)))
        {
          csv.WriteRecords(tasks);
        }
      }
      catch (System.UnauthorizedAccessException ex)
      {
        success = false;
        MessageBox.Show(Resource.FileCouldNotWriteAccessedDenied);
      }
      catch (Exception ex)
      {
        success = false;
        MessageBox.Show(String.Format(Resource.FileCouldNotWriteUnknown, ex.Message));
      }

      return success;
    }

    /// <summary>
    /// Kicks off the entire workflow. Authenticates Exchange user that is used to
    /// impersonate all Exchange users supplied throughout the workflow process. The
    /// users credentials supplied should be a user that has been giving impersonation 
    /// permissions in Exchange.
    /// </summary>
    private void extract_Click(object sender, EventArgs e)
    {
      Start.Enabled = false;
      DialogResult msgRes = MessageBox.Show(Resource.ConfirmationMsg,
                                            Resource.Confirmation,
                                            MessageBoxButtons.OKCancel);
      if (msgRes == DialogResult.OK)
      {
        if (authenticate())
        {
          execute();
        }
      }

      Start.Enabled = true;
    }

    /// <summary>
    /// Callback function used by Exchange Service to communicate with Exchange Web Service. 
    /// </summary>
    private static bool CertificateValidationCallBack(object sender,
                                            System.Security.Cryptography.X509Certificates.X509Certificate certificate,
                                            System.Security.Cryptography.X509Certificates.X509Chain chain,
                                            System.Net.Security.SslPolicyErrors sslPolicyErrors)
    {
      // If the certificate is a valid, signed certificate, return true.
      if (sslPolicyErrors == System.Net.Security.SslPolicyErrors.None)
      {
        return true;
      }

      // If there are errors in the certificate chain, look at each error to determine the cause.
      if ((sslPolicyErrors & System.Net.Security.SslPolicyErrors.RemoteCertificateChainErrors) != 0)
      {
        if (chain != null && chain.ChainStatus != null)
        {
          foreach (System.Security.Cryptography.X509Certificates.X509ChainStatus status in chain.ChainStatus)
          {
            if ((certificate.Subject == certificate.Issuer) &&
               (status.Status == System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.UntrustedRoot))
            {
              // Self-signed certificates with an untrusted root are valid. 
              continue;
            }
            else
            {
              if (status.Status != System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.NoError)
              {
                // If there are any other errors in the certificate chain, the certificate is invalid,
                // so the method returns false.
                return false;
              }
            }
          }
        }

        // When processing reaches this line, the only errors in the certificate chain are 
        // untrusted root errors for self-signed certificates. These certificates are valid
        // for default Exchange server installations, so return true.
        return true;
      }
      else
      {
        // In all other cases, return false.
        return false;
      }
    }

    /// <summary>
    /// Exchange Service validation callback.
    /// </summary>
    private static bool RedirectionUrlValidationCallback(string redirectionUrl)
    {
      // The default for the validation callback is to reject the URL.
      bool result = false;

      Uri redirectionUri = new Uri(redirectionUrl);

      // Validate the contents of the redirection URL. In this simple validation
      // callback, the redirection URL is considered valid if it is using HTTPS
      // to encrypt the authentication credentials. 
      if (redirectionUri.Scheme == "https")
      {
        result = true;
      }
      return result;
    }

    /// <summary>
    /// Gets the version of the Exchange server the user selected.
    /// </summary>
    /// <returns>Exchange Version Enum</returns>
    private ExchangeVersion getVersion()
    {
      switch (ExchangeVersions.SelectedIndex)
      {
        case 1:
          return ExchangeVersion.Exchange2010_SP1;
        case 2:
          return ExchangeVersion.Exchange2010_SP2;
        default:
          return ExchangeVersion.Exchange2010;
      }
    }

    /// <summary>
    /// Changes Exchange EWS textbox to enabled if auto discovery checkbox 
    /// is unchecked. If auto discovery is checked, then it's disabled.
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void AutoDiscoverURL_CheckedChanged(object sender, EventArgs e)
    {
      ExchangeServerURL.Enabled = !AutoDiscoverURL.Checked;
    }

    /// <summary>
    /// Authenticates the users credentials supplied. 
    /// </summary>
    /// <returns>True if the credentials are able to ping the Exchange server via the Exchange Web Service</returns>
    private Boolean authenticate()
    {
      Service = new ExchangeService(getVersion());
      Service.Credentials = new WebCredentials(Username.Text.Trim(), Password.Text.Trim(), Domain.Text.Trim());

      String status = String.Empty;
      if (AutoDiscoverURL.Checked)
      {
        toolStripStatusLabel1.Text = Resource.RunningAutoDiscovery;

        String sts = String.Empty;
        try
        {
          Service.AutodiscoverUrl(Username.Text.Trim(), RedirectionUrlValidationCallback);
        }
        catch (Microsoft.Exchange.WebServices.Data.AutodiscoverLocalException adex)
        {
          sts = Resource.AutoDiscoveryCouldNotFind;
        }
        catch (Exception ex)
        {
          sts = Resource.AutoDiscoveryFailed;
        }
        finally
        {
          if (sts == String.Empty)
          {
            sts = Resource.ExchangeServerFound;
            pingEWS();
          }

          toolStripStatusLabel1.Text = sts;
        }
      }
      else
      {
        toolStripStatusLabel1.Text = Resource.Authenticating;
        Service.Url = new Uri(ExchangeServerURL.Text.Trim());
        pingEWS();
      }

      return (toolStripStatusLabel1.Text == Resource.Authenticated);
    }

    /// <summary>
    /// Pings the Echange Web service for ues within the authentication fucntion.
    /// </summary>
    private void pingEWS()
    {
      try
      {
        var items = Service.FindItems(WellKnownFolderName.Tasks,
                    new ItemView(1) { PropertySet = new PropertySet(BasePropertySet.FirstClassProperties) });

        toolStripStatusLabel1.Text = Resource.Authenticated;
      }
      catch
      {
        toolStripStatusLabel1.Text = Resource.CouldNotAuthenticate;
      }
    }
  }
}
