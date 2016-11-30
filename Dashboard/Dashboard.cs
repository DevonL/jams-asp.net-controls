using DevExpress.DashboardWeb;
using Genghis;
using MVPSI.JAMS;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using MVPSI.JAMSPSCommon;
using System.Management.Automation;
using System.Management.Automation.Runspaces;


[assembly: WebResource("MVPSI.JAMSWeb.Controls.Dashboard.Dashboard.js", "text/javascript")]
namespace MVPSI.JAMSWeb.Controls
{
    /// <summary>
    /// The JAMS Dashboard Control can be placed on a web page to display current and history
    ///  information about a JAMS server.
    /// </summary>
    [DefaultProperty("ServerName")]
    [ToolboxData("<{0}:Dashboard runat=server></{0}:Dashboard>")]
    public class Dashboard :  ControlsCommon
    {
        private Preferences m_Preferences;
        private List<DevExpress.Data.IParameter> m_CurrentDashboardParameters;
        private ASPxDashboardViewer m_DashboardViewer;
        private Server m_Server;
       
        /// <summary>
        /// Gets or sets the name of the JAMS Server to connect to.
        /// </summary>
        [Category("Data"),
        DefaultValue(""),
        Description("Name of the JAMS server to connect to")]
        public string ServerName
        {
            get
            {
                return (m_ServerName == null) ? String.Empty : m_ServerName;
            }
            set
            {
                if (m_ServerName != value)
                {
                    m_ServerName = value;
                    m_Server = null;
                }
            }
        }
        private string m_ServerName;

        internal Server Server
        {
            get
            {
                if (m_Server == null)
                {
                    if (ServerName == null)
                    {
                        ServerName = string.Empty;
                    }

                    if (ServerName == string.Empty)
                    {
                        m_Server = JAMS.Server.GetCurrentServer();
                    }
                    else
                    {
                        m_Server = JAMS.Server.GetServer(ServerName);
                    }
                }
                return m_Server;
            }
        }

        /// <summary>
        /// Gets or sets the path to the JAMS dashboard file (.jdb) that is to be displayed.
        /// </summary>
        [Category("Data"),
        DefaultValue(""),
        Description("Path to the JAMS dashboard file (.jdb) that is to be displayed.")]
        public string DashboardFile
        {
            get
            {
                return (m_DashboardFile == null) ? String.Empty : m_DashboardFile;
            }
            set
            {
                m_DashboardFile = value;
            }
        }
        private string m_DashboardFile;

        /// <summary>
        /// Gets or sets the Dashboard Viewer's client programmatic identifier
        /// </summary>
        public string ClientInstanceName
        {
            get
            {
                return m_ClientInstanceName;
            }
            set
            {
                m_ClientInstanceName = value;
            }
        }
         private string m_ClientInstanceName;

         /// <summary>
         /// Gets or sets the Dashboard Viewer's client programmatic identifier
         /// </summary>
         public bool AllowExportDashboard
         {
             get
             {
                 return m_AllowExportDashboard;
             }
             set
             {
                 m_AllowExportDashboard = value;
             }
         }
         private bool m_AllowExportDashboard = true;

         /// <summary>
         /// Gets or Sets if FullScreenMode is enabled
         /// </summary>
         public bool FullScreenMode
         {
             get
             {
                 return m_FullScreenMode;
             }
             set
             {
                 m_FullScreenMode = value;
             }
         }
         private bool m_FullScreenMode;

        /// <summary>
        /// Raised when the dashboard is loaded
        /// </summary>
        /// <param name="e"></param>
        protected override void OnLoad(EventArgs e)
        {
            string preferencesFile = String.Empty;

            //
            // Define the Dashboard viewer
            //
            m_DashboardViewer = new ASPxDashboardViewer
            {
                ID = GetChildControlID("ASPxDashboardViewer"),
                AllowExportDashboard = m_AllowExportDashboard,
                ClientInstanceName =  m_ClientInstanceName,
                FullscreenMode = m_FullScreenMode,
                Width = this.Width,
                Height = this.Height
            };
         
            //
            // Subscribe to dashboard events for loading data
            //
            m_DashboardViewer.DataLoading += m_DashboardViewer_DataLoading;
            m_DashboardViewer.CustomParameters += m_DashboardViewer_CustomParameters;
           
            //
            // Load the dashboard preference file
            //
            preferencesFile = HttpContext.Current.Server.MapPath(Path.ChangeExtension(m_DashboardFile, ".pref"));
            if (File.Exists(preferencesFile))
            {
                m_Preferences = Preferences.GetUserNode("DataSources", preferencesFile);
            }

            //
            // Load the Dashboard
            //
            m_DashboardViewer.DashboardSource = m_DashboardFile;

            base.Controls.Add(m_DashboardViewer);
        }

        /// <summary>
        /// Gets the current parameters after a change occured
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void m_DashboardViewer_CustomParameters(object sender, CustomParametersWebEventArgs e)
        {
            m_CurrentDashboardParameters = e.Parameters;
        }

        /// <summary>
        /// Provides data to the datasources
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void m_DashboardViewer_DataLoading(object sender, DataLoadingWebEventArgs e)
        {
            string componentName = String.Empty;
            string dataSourceName = String.Empty;
            object dataCollection = null;
            Preferences subkey = null;

            //
            // Get the information about the datasource
            //
            componentName = e.DataSourceComponentName;
            dataSourceName = e.DataSourceName;

            /*
            ** Check if this is a JAMS DataSource and if so, load the data.
            */
            try
            {
                //
                // We must have a preferences object to load JAMS datasources
                //
                if (m_Preferences != null)
                {
                    // Get the Preferences subkey for this datasource
                    subkey = m_Preferences.OpenSubKey(dataSourceName);
                  
                    if (subkey != null)
                    {
                        // History datasource
                        if (componentName.StartsWith("HistoryDS"))
                        {
                            //
                            // Load each History parameter from Preferences and if the value starts with '$'
                            //  try to get the value from a matching Dashboard parameter.
                            //
                            string folderName = GetParameterValue(subkey, "QueryFolderName", @"\");
                            string jobName = GetParameterValue(subkey, "QueryJobName", "*");
                            string setupName = GetParameterValue(subkey, "QuerySetupName", "*");
                            string startDateRaw = GetParameterValue(subkey, "QueryStartDate", "Today");
                            string startTime = GetParameterValue(subkey, "QueryStartTime", "00:00:00.00");
                            string endDateRaw = GetParameterValue(subkey, "QueryEndDate", "Tomorrow");
                            string endTime = GetParameterValue(subkey, "QueryEndTime", "00:00:00.00");
                            string withinDelta = GetParameterValue(subkey, "QueryWithin", "0");
                            string includeSuccess = GetParameterValue(subkey, "QueryIncludeSuccess", "true");
                            string includeInfo = GetParameterValue(subkey, "QueryIncludeInfo", "true");
                            string includeWarning = GetParameterValue(subkey, "QueryIncludeWarning", "true");
                            string includeError = GetParameterValue(subkey, "QueryIncludeError", "true");
                            string includeFatal = GetParameterValue(subkey, "QueryIncludeFatal", "true");
                            string checkSched = GetParameterValue(subkey, "QueryCheckSched", "true");
                            string checkHold = GetParameterValue(subkey, "QueryCheckHold", "true");
                            string checkStart = GetParameterValue(subkey, "QueryCheckStart", "true");
                            string checkCompletion = GetParameterValue(subkey, "QueryCheckCompletion", "true");
                            string searchRecursivelyRaw = GetParameterValue(subkey, "QuerySearchRecursively", "true");
                            bool useDates = subkey.GetBoolean("QueryUseDates", true);

                            DateTime startDate = DateTime.MinValue;
                            DateTime endDate = DateTime.MinValue;
                            bool searchRecursively = Boolean.Parse(searchRecursivelyRaw);
                            ComparisonOperator searchFolderOperator = ComparisonOperator.MatchFolder;

                            //
                            // Determine the Start & End dates based either on the WithinTime or date range
                            //
                            if (useDates)
                            {
                                startDate = Date.Evaluate(startDateRaw,Server).Date;
                                startDate += Date.Evaluate(startTime,Server).TimeOfDay;

                                endDate = Date.Evaluate(endDateRaw,Server).Date;
                                endDate += Date.Evaluate(endTime,Server).TimeOfDay;
                            }
                            else
                            {
                                //
                                //  They are using the "Within..." controls,
                                //  the end date is right now and the start date is the current time minus the delta.
                                //
                                startDate = DateTime.Now - DeltaTime.Parse(withinDelta).TimeSpan;
                                endDate = DateTime.Now;
                            }

                            //
                            // Add History Search Criteria
                            //
                            List<HistorySelection> selectionList = new List<HistorySelection>();
                            selectionList.Add(new HistorySelection(HistorySelectionField.JobName, ComparisonOperator.Like, jobName));
                            selectionList.Add(new HistorySelection(HistorySelectionField.SetupName, ComparisonOperator.Like, setupName));

                            if (searchRecursively)
                            {
                                searchFolderOperator = ComparisonOperator.MatchFolderRecursively;
                            }
                            selectionList.Add(new HistorySelection(HistorySelectionField.FolderName, searchFolderOperator, folderName));


                            //
                            // Query History using the supplied values
                            //
                            dataCollection = JAMS.History.Find(selectionList,
                                startDate,
                                endDate,
                                Boolean.Parse(includeSuccess),
                                Boolean.Parse(includeInfo),
                                Boolean.Parse(includeWarning),
                                Boolean.Parse(includeError),
                                Boolean.Parse(includeFatal),
                                Boolean.Parse(checkSched),
                                Boolean.Parse(checkHold),
                                Boolean.Parse(checkStart),
                                Boolean.Parse(checkCompletion),
                                HistorySearchOptions.None,
                                Server);
                        }
                        else if (componentName.StartsWith("CompletionsBySeverityDS"))
                        {
                            DateTime startDate = DateTime.Now;
                            DeltaTime lookbackInterval = DeltaTime.Zero;

                            //
                            // Load the lookbackInterval from Preferences and if the value starts with '$'
                            //  try to get the value from a matching Dashboard parameter.
                            //
                            string lookbackIntervalRaw = GetParameterValue(subkey, "QueryLookBackInterval", "00:00:00.00");

                            //
                            // The startDate will be the current time minus the lookback interval
                            //
                            if (DeltaTime.TryParse(lookbackIntervalRaw, out lookbackInterval))
                            {
                                startDate = startDate.Subtract(lookbackInterval.ToTimeSpan());
                            }

                            //
                            // Retrieve the completion data that occured within 24 hours before the startDate
                            //
                            dataCollection = Statistics.GetCompletionsBySeverity(startDate, Server);
                        }
                        else if (componentName.StartsWith("QueueDS"))
                        {
                            //
                            // Load the BatchQueue mask parameter and if the value starts with '$'
                            //  try to get the value from a matching Dashboard parameter.
                            //
                            string queueMask = GetParameterValue(subkey, "QueryQueueName", "*");

                            //
                            // Retrieve the queue data
                            //
                            dataCollection = BatchQueue.Find(queueMask, Server);
                        }
                        else if (componentName.StartsWith("ResourceDS"))
                        {
                            //
                            // Load the Resource mask parameter and if the value starts with '$'
                            //  try to get the value from a matching Dashboard parameter.
                            //
                            string resourceMask = GetParameterValue(subkey, "QueryResourceName", "*");

                            //
                            // Retrieve the Resource data
                            //
                            dataCollection = Resource.Find(resourceMask, Server);
                        }
                        else if (componentName.StartsWith("AgentDS"))
                        {
                            //
                            // Load the Agent mask parameter and if the value starts with '$'
                            //  try to get the value from a matching Dashboard parameter.
                            //
                            string agentMask = GetParameterValue(subkey, "QueryAgentName", "*");

                            //
                            // Retrieve the Agent data
                            //
                            dataCollection = Agent.Find(agentMask, Server);
                        }
                        else if (componentName.StartsWith("PowerShellDS"))
                        {
                            dataCollection = GetPowerShellDataSourceData(subkey);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            if (dataCollection != null)
            {
                e.Data = dataCollection;
            }
        }


        /// <summary>
        /// Get data for a PowerShell datasource. This is broken out as a seperate method so we won't try to load the PowerShell
        /// assemblies unless we actually have a PowerShell datasource.
        /// </summary>
        /// <param name="subkey"></param>
        /// <returns></returns>
        private object GetPowerShellDataSourceData(Preferences subkey)
        {
            //
            // Get the PowerShell Source Preferences and if the value starts with '$'
            //  try to get the value from a matching Dashboard parameter.
            //
            string source = GetParameterValue(subkey, "QuerySource", String.Empty);
            string execOptions = GetParameterValue(subkey, "QueryOptions", String.Empty);
            string consoleFilePath = GetParameterValue(subkey, "QueryConsoleFilePath", String.Empty);

            //
            // Build a dictionary of the current Dashboard parameter values
            //
            Dictionary<string, object> parameterDictionary = new Dictionary<string, object>();
            foreach (DevExpress.Data.IParameter param in m_CurrentDashboardParameters)
            {
                // Get the current value
                string currentValue = GetDashboardParameterValue(param.Name, param.Value.ToString());

                // Add the element to the dictionary
                parameterDictionary.Add(param.Name, currentValue);
            }

            // Create a JAMSPSHost
            JAMSDashboardPSHost psHost = new JAMSDashboardPSHost(execOptions, consoleFilePath, parameterDictionary);

            // Define the command
            Command jobScript = new Command(source, true);

            //
            // Add Parameters from the Dashboard
            //
            if (psHost.PassParams)
            {
                foreach (KeyValuePair<string, object> param in parameterDictionary)
                {
                    jobScript.Parameters.Add(param.Key, param.Value);
                }
            }
            psHost.Pipeline.Commands.Add(jobScript);

            // Execute the Script and get the output
            ICollection<PSObject> resultList = psHost.Pipeline.Invoke();

            // Set the data collection to be the BaseObjects of the returned PSObjects.
            return resultList.Select(p => p.BaseObject);
        }

        /// <summary>
        /// Loads a parameter value from the preferences file and checks if the value
        ///  should come from a Dashboard parameter.
        /// </summary>
        /// <returns></returns>
        private string GetParameterValue(Preferences subKey, string paramName, string defaultValue)
        {
            string paramValue = String.Empty;

            try
            {
                // Load the Parameter from Preferences
                paramValue = subKey.GetString(paramName, defaultValue);

                //
                // If the value starts with '$' try to pull the value from a matching Dashboard Parameter
                //
                if (paramValue.StartsWith("$"))
                {
                    paramValue = GetDashboardParameterValue(paramName, paramValue);
                }
            }
            catch
            {
                paramValue = defaultValue;
            }

            return paramValue;
        }

        /// <summary>
        /// Checks the Dashboard for a matching parameter name and gets the value.
        /// </summary>
        /// <param name="paramName"></param>
        /// <param name="paramValue"></param>
        /// <returns></returns>
        private string GetDashboardParameterValue(string paramName, string paramValue)
        {
            bool isDashboardParam = false;

            // Strip off the '$' to get the dashboard parameter name
            if (paramValue.StartsWith("$"))
            {
                paramName = paramValue.Substring(1);
            }

            // Check if the Param exists on the Dashboard
            if (m_CurrentDashboardParameters != null &&
                m_CurrentDashboardParameters.Any(p => String.Compare(p.Name, paramName, true) == 0))
            {
                isDashboardParam = true;
            }

            if (isDashboardParam)
            {
                // Get the value from the Dashboard 
                paramValue = m_CurrentDashboardParameters.First(p => String.Compare(p.Name, paramName, true) == 0).Value.ToString();
            }

            return paramValue;
        }
    }
}
