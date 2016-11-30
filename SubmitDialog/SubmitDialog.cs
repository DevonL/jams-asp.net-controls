using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Web;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using AjaxControlToolkit;
using MVPSI.JAMS;
using System.Globalization;

[assembly: WebResource("MVPSI.JAMSWeb.Controls.SubmitMenu.JAMSSubmitStyles.css", "text/css")]
namespace MVPSI.JAMSWeb.Controls
{
    /// <summary>
    /// The JAMS SubmitDialog control can be used to submit a job from a web page.
    /// </summary>
    [Designer("MVPSI.JAMSWeb.Controls.SubmitDialogDesigner")]
    public class SubmitDialog : ControlsCommon
    {
        private ModalPopupExtender m_PopupControl;
        private TabContainer m_TabContainer;
        private List<TextBox> m_PrintQueueTextBoxes;
        private List<TextBox> m_PrintFormTextBoxes;
        private List<TextBox> m_NumCopiesTextBoxes;
        private Button m_OkButton;
        private Button m_CancelButton;
        private Panel m_InnerPanel;

        /// <summary>
        /// Default constructor
        /// </summary>
        public SubmitDialog()
        {
            DisplayModal = false;
            ServerName = string.Empty;
            JobName = string.Empty;
        }

        /// <summary>
        /// Gets the message indicating a successful or unsuccessful job submission.
        /// </summary>
        /// <value>The result text.</value>
        public string ResultText
        {
            get
            {
                object viewState = ViewState[ClientID + "ResultText"];
                if (viewState is string)
                {
                    return (string)viewState;
                }
                else
                {
                    return string.Empty;
                }
            }
            set 
            {
                ViewState[ClientID + "ResultText"] = value;
            }
        }

        private Server Server
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
        private JAMS.Server m_Server;

        /// <summary>
        /// Gets or sets the name of the job or setup to be submitted.
        /// </summary>
        [Category("Data"), 
        DefaultValue(""),
        Description("The name of the Job or Setup to be submitted.")]
        public string JobName
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the name of the JAMS Server.
        /// </summary>
        [Category("Data"),
        DefaultValue(""),
        Description("The name of the JAMS Server.")]
        public string ServerName
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the JobID or SetupID of the Job or Setup to be submitted. 
        /// </summary>
        public int SubmitID
        {
            get
            {
                object viewState = ViewState[ClientID + "SubmitID"];
                if (viewState is int)
                {
                    return (int)viewState;
                }
                else
                {
                    return 0;
                }
            }
        }

        /// <summary>
        /// Gets or sets the type of JAMS object being submitted.
        /// </summary>
        public JAMS.Submit.Type SubmitType
        {
            get
            {
                object viewState = ViewState[ClientID + "SubmitType"];
                if (viewState is JAMS.Submit.Type)
                {
                    return (JAMS.Submit.Type)viewState;
                }
                else
                {
                    return JAMS.Submit.Type.Unknown;
                }
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether or not to display the control only as a modal dialog.
        /// </summary>
        /// <value><c>true</c> if [display modal]; otherwise, <c>false</c>.</value>
        [Category("Data"),
        DefaultValue("False"),
        Description("Determines if the control will only be displayed as a modal dialog.")]
        public bool DisplayModal
        {
            get; set;
        }

        /// <summary>
        /// Raises the <see cref="E:System.Web.UI.Control.Init"/> event.
        /// </summary>
        /// <param name="e">An <see cref="T:System.EventArgs"/> object that contains the event data.</param>
        protected override void OnInit(EventArgs e)
        {
            //
            //  We have to have a ToolkitScriptManager on the page but, we can't require that the 
            //  developer add one to the page because this might be a SharePoint environment
            //  so, we added this code to Page.Init to add a ToolkitScriptManager if there isn't one.
            //
            Page.Init += delegate(object sender, EventArgs e_Init)
            {
                if (ToolkitScriptManager.GetCurrent(Page) == null)
                {
                    ToolkitScriptManager sMgr = new ToolkitScriptManager();
                    if (Page.Form == null)
                    {
                        throw new HttpException("Control must be located inside of a <form> tag.");
                    }
                    Page.Form.Controls.AddAt(0, sMgr);
                }
            };

            Width = Unit.Pixel(500);
            Height = Unit.Empty;

            base.OnInit(e);

            //
            //  Add our stylesheet
            //
            HtmlLink link = new HtmlLink();
            link.Href = Page.ClientScript.GetWebResourceUrl(this.GetType(), "MVPSI.JAMSWeb.Controls.SubmitDialog.JAMSSubmitStyles.css");
            link.Attributes.Add("type", "text/css");
            link.Attributes.Add("rel", "stylesheet");
            Page.Header.Controls.Add(link);
        }


        /// <summary>
        /// Raises the <see cref="E:System.Web.UI.Control.Load"/> event.
        /// </summary>
        /// <param name="e">The <see cref="T:System.EventArgs"/> object that contains the event data.</param>
        protected override void OnLoad(EventArgs e)
        {
            Debug.WriteLine("SubmitDialog OnLoad()");

            Debug.WriteLine("SubmitDialog CreateChildControls()");
            Server.Connect();

            //
            //  Force disconnect and reconnect because we may be impersonating a different account
            //
            //Server.Disconnect();
            //Server.Connect();

            if (!string.IsNullOrEmpty(JobName))
            {
                Submit.Info si;
                Submit.Load(out si, JobName, Server, Submit.Type.Unknown);

                if (si != null)
                {
                    ViewState[ClientID + "SubmitType"] = si.SubmitType;
                    ViewState[ClientID + "SubmitID"] = si.SubmitID;
                }
            }

            //If the click target was the cancel button, wipe out the SubmitType and SubmitID
            if (Page.IsPostBack)
            {
                var target = Page.Request["__EVENTTARGET"];
                string cancelID;

                try
                {
                    cancelID = ViewState[ClientID + "CancelID"].ToString();
                }
                    //This value has not been set yet
                catch(NullReferenceException)
                {
                    cancelID = null;
                }

                if (target != null && cancelID != null)
                {
                    if (cancelID == target)
                    {
                        ClearViewState();
                    }
                }

            }


            //
            //  Grab local copies of SubmitType and ID so we don't have to look them
            //  up in ViewState every time we reference them
            //
            JAMS.Submit.Type submitType = SubmitType;
            int submitID = SubmitID;

            if ((submitType != Submit.Type.Unknown) && (submitID != 0))
            {
                Debug.WriteLine(string.Format("Submitting {0} with ID of {1}", submitType, submitID));

                Submit.Info si = null;
                try
                {
                    if (!DesignMode)
                        Submit.Load(out si, submitID, Server, submitType);
                    ResultText = string.Empty;
                }
                catch (SecurityException ex)
                {
                    ResultText = ex.Message;
                    return;
                }

                m_InnerPanel = new Panel();
                m_InnerPanel.ID = GetChildControlID("Dialog_Panel");
                Controls.Add(m_InnerPanel);

                m_OkButton = new Button();
                m_OkButton.ID = GetChildControlID("Dialog_OK");
                m_OkButton.Text = "Submit Run Request";
                m_OkButton.UseSubmitBehavior = true;
                m_OkButton.Click += new EventHandler(okButton_Click);
                m_OkButton.Style["margin"] = "20px, 50px, 10px, 0px";

                if (DisplayModal)
                {
                    //If we don't have a modal dialog, we don't need a cancel button
                    m_CancelButton = new Button();
                    m_CancelButton.ID = GetChildControlID("Dialog_Cancel");
                    m_CancelButton.Text = "Cancel";
                    m_CancelButton.Click += new EventHandler(cancelButton_Click);
                    m_CancelButton.CausesValidation = false;
                    m_CancelButton.Style["margin"] = "20px, 0px, 10px, 0px";

                    m_PopupControl = new ModalPopupExtender{ID = GetChildControlID("PopupExtender")};
                    m_PopupControl.PopupControlID = m_InnerPanel.ClientID;
                    m_PopupControl.DropShadow = false;
                    m_PopupControl.RepositionMode = ModalPopupRepositionMode.RepositionOnWindowResize;
                    m_PopupControl.TargetControlID = GetChildControlID("TargetButton");
                    m_PopupControl.CancelControlID = ClientID + "_Dialog_Cancel";
                    ViewState[ClientID + "CancelID"] = m_PopupControl.CancelControlID;

                    m_PopupControl.OnOkScript = String.Format("__doPostBack('{0}','{1}');", m_OkButton.ClientID, "");
                    m_PopupControl.OnCancelScript = String.Format("__doPostBack('{0}','{1}');",
                                                                    ClientID + "_Dialog_Cancel",
                                                                    "");
                    m_PopupControl.BackgroundCssClass = "modalBackground";
                    m_PopupControl.PopupDragHandleControlID = GetChildControlID("TitleBar");

                    m_InnerPanel.CssClass = "modalPopup";
                    m_PopupControl.Y = 100;

                    Controls.Add(this.m_PopupControl);

                    Button button = new Button();
                    button.ID = GetChildControlID("TargetButton");
                    button.Style["display"] = "none";
                    button.Height = 0;
                    button.Width = 0;
                    Controls.AddAt(1, button);
                }         

                m_TabContainer = new TabContainer {ID = GetChildControlID("TabContainer")};
                m_TabContainer.Height = 350;

                TabPanel props = new TabPanel {ID = GetChildControlID("PropertyTab")};
                TabPanel schedule = new TabPanel {ID = GetChildControlID("ScheduleTab")};
                TabPanel reports = new TabPanel {ID = GetChildControlID("ReportTab")};

                props.HeaderText = "Parameters";
                schedule.HeaderText = "Schedule";
                reports.HeaderText = "Reports";

                m_TabContainer.Tabs.Add(props);
                m_TabContainer.Tabs.Add(reports);
                m_TabContainer.Tabs.Add(schedule);

                m_TabContainer.ActiveTabIndex = 0;
                m_TabContainer.ScrollBars = System.Web.UI.WebControls.ScrollBars.Auto;

                m_InnerPanel.Controls.Add(m_TabContainer);

                Panel buttonPanel = new Panel();
                buttonPanel.ID = GetChildControlID("Button_Panel");
                buttonPanel.Controls.Add(m_OkButton);
                if (DisplayModal)
                {
                    buttonPanel.Controls.Add(m_CancelButton);
                }
                buttonPanel.EnableViewState = false;
                buttonPanel.HorizontalAlign = System.Web.UI.WebControls.HorizontalAlign.Center;

                m_InnerPanel.Controls.Add(buttonPanel);

                if (si != null)
                {
                    JobName = si.Name;

                    Panel titleBar = new Panel();
                    titleBar.ID = GetChildControlID("TitleBar");
                    titleBar.Style["background-color"] = "#CEE5FE";
                    titleBar.Style["margin"] = "-3, -3, 6, -3";

                    Label head = new Label();
                    head.Text = si.Name + " - ";
                    titleBar.Controls.Add(head);


                    head = new Label();
                    head.Text = si.Description;
                    titleBar.Controls.Add(head);
                    m_InnerPanel.Controls.AddAt(0, titleBar);

                    head = new Label();
                    head.Text = si.Description;
                    titleBar.Controls.Add(head);
                    m_InnerPanel.Controls.AddAt(0, titleBar);

                    //Find the batch queues on this server
                    ICollection<MVPSI.JAMS.BatchQueue> queues;
                    object savedQueues = Server.GetObject("BatchQueues");
                    if (savedQueues is ICollection<MVPSI.JAMS.BatchQueue>)
                    {
                        //
                        //  There was a list of queues saved in the JAMS Server, use that
                        //
                        queues = (ICollection<MVPSI.JAMS.BatchQueue>) savedQueues;
                    }
                    else
                    {
                        //
                        //  This must be the first time they searched for batch queues
                        //
                        queues = MVPSI.JAMS.BatchQueue.Find(Server);

                        //
                        //  Save this in the Server so we'll find it the next time
                        //
                        Server.SaveObject("BatchQueues", queues);
                    }
                        
                    DisplaySchedule(si, queues);
                    DisplayReports(si);
                    DisplayParameters(si);

                    if (DisplayModal)
                    {
                        Style["Display"] = "block";
                        m_PopupControl.Show();
                    }
                }
            }
            else
            {
                //
                //  We don't have to load anything because we don't have a job to submit
                //
                if (!DisplayModal)
                {
                    Controls.Add(new Label {Text = "No Job or Setup could be loaded."});
                }
                else
                {
                    Style["Display"] = "None";
                }

                Debug.WriteLine("OnLoad doesn't have a job");
            }
        }

        /// <summary>
        /// Shows the dialog.
        /// </summary>
        /// <param name="submitType">Type of the submit.</param>
        /// <param name="submitID">The submit ID.</param>
        public void ShowDialog(JAMS.Submit.Type submitType, int submitID)
        {
            ViewState[ClientID + "SubmitType"] = submitType;
            ViewState[ClientID + "SubmitID"] = submitID;
            OnLoad(EventArgs.Empty);
        }

        internal void ClearViewState()
        {
            Console.WriteLine("ViewState cleared");
            ViewState[ClientID + "SubmitType"] = JAMS.Submit.Type.Unknown;
            ViewState[ClientID + "SubmitID"] = 0;
        }

        private void DisplaySchedule(Submit.Info si, ICollection<BatchQueue> queues)
        {
            //Add all of the Schedule information

            TextBox time = new TextBox { ID = GetChildControlID("Schedule_Time"), CssClass = "tabContents" };
            TextBox calendarTextBox = new TextBox { ID = GetChildControlID("Schedule_Date"), CssClass = "tabContents" };
            CalendarExtender calendarExtender = new CalendarExtender { TargetControlID = calendarTextBox.ClientID, ID = GetChildControlID("CalendarExtender"), Format = CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern };
            DropDownList batchQueue = new DropDownList { ID = GetChildControlID("Batch_Queues"), CssClass = "tabContents" };
            TextBox agentNode = new TextBox { ID = GetChildControlID("Agent_Node"), CssClass = "tabContents" };
            TextBox schedulingPriority = new TextBox { ID = GetChildControlID("Schedule_Priority"), CssClass = "tabContents" };
            Label scheduleWindowLabel = new Label() { CssClass = "tabContents" };

            if (si.AfterTime < DateTime.Now)
            {
                calendarExtender.SelectedDate = DateTime.Now;
                time.Text = DateTime.Now.ToString("h:mm tt");
            }
            else
            {
                //
                //  Set the After date/time
                //                               
                calendarExtender.SelectedDate = si.AfterTime;
                time.Text = si.AfterTime.ToString("h:mm tt");
            }

            if (si.ManageAccess)
            {

                //Fill DropDown
                batchQueue.Items.Add("");
                foreach (BatchQueue queue in queues)
                {
                    batchQueue.Items.Add(queue.QueueName);
                }
                batchQueue.AutoPostBack = false;
                batchQueue.Text = si.BatchQueueName;

                if (si.SubmitIsSimple)
                {
                    agentNode.Text = si.Jobs[0].AgentNode;
                }
                else
                {
                    agentNode.Enabled = false;
                }

                schedulingPriority.Text = si.SchedulingPriority.ToString();
            }
            else
            {
                //
                //  They don't have Manage access
                //
                batchQueue.Enabled = false;
                agentNode.Enabled = false;
                schedulingPriority.Enabled = false;
            }

            if ((si.ScheduleFromTime.TotalSeconds >= 0) || (si.ScheduleToTime.TotalSeconds >= 0))
            {
                //
                //  We have a submit window, display it
                //
                scheduleWindowLabel.Text = string.Format(MVPSI.JAMS.Msg.GetString("display_submit_window"),
                                                         si.ScheduleFromTime.ToString(), si.ScheduleToTime.ToString());
            }

            //Get the schedule tab and add controls to it.            
            TabPanel scheduleView = m_TabContainer.Tabs[2];

            Label startTime = new Label();
            startTime.Text = "Scheduled Start Time";
            scheduleView.Controls.Add(startTime);

            scheduleView.Controls.Add(calendarTextBox);
            scheduleView.Controls.Add(calendarExtender);

            scheduleView.Controls.Add(time);

            Label batchQueueLabel = new Label();
            batchQueueLabel.Text = "Batch Queue";
            scheduleView.Controls.Add(batchQueueLabel);

            scheduleView.Controls.Add(batchQueue);

            Label AgentNodeLabel = new Label();
            AgentNodeLabel.Text = "Agent Node";
            scheduleView.Controls.Add(AgentNodeLabel);

            scheduleView.Controls.Add(agentNode);

            Label schedulingPriorityLabel = new Label();
            schedulingPriorityLabel.Text = "Scheduling Priority";
            scheduleView.Controls.Add(schedulingPriorityLabel);

            scheduleView.Controls.Add(schedulingPriority);

            scheduleView.Controls.Add(scheduleWindowLabel);

            calendarTextBox.Text = calendarExtender.SelectedDate.Value.Date.ToShortDateString();

            AddRequiredVaild(calendarTextBox.ID, scheduleView);
            AddRegexValid(schedulingPriority.ID, @"[-+]?\b\d+\b", "Integer entered is not valid", scheduleView);
            AddRegexValid(time.ID, @"(^(?:0?[1-9]:[0-5]|1(?=[012])\d:[0-5])\d( )?(AM|am|aM|Am|PM|pm|pM|Pm))$", "Time entered is not valid", scheduleView);
            m_TabContainer.ActiveTabIndex = 2;         
            scheduleView.Visible = true;
            m_TabContainer.Visible = true;
        }

        private void DisplayParameters(Submit.Info si)
         {
            
            bool hideParamTab = true;

            //Get the parameters tab and add controls to it
            TabPanel paramView = m_TabContainer.Tabs[0];

            foreach (Submit.Parameter p in si.Parameters)
            {
                if (!p.Hide)
                {
                    hideParamTab = false;
                    // 
                    // Prompt Label
                    // 
                    if (p.DataType != DataType.Boolean)
                    {
                        Label promptLabel = new Label();
                        promptLabel.Text = p.Prompt;
                        promptLabel.CssClass = "tabContents";
                        paramView.Controls.Add(promptLabel);
                        m_TabContainer.ActiveTabIndex = 0;
                    }

                    Control baseControl = null;

                    switch (p.DataType)
                    {
                        case DataType.Date:
                        case DataType.DateTime:
                            {
                                if (p.ValidationType == ValidationType.Select)
                                {
                                    DropDownList dropDown = new DropDownList();
                                    dropDown.ID = GetChildControlID(p.ParamName + "_" + p.ParamID);
                                    baseControl = dropDown;

                                    var dateStrings = p.ValidationData.Split(new char[] { ',' },
                                                                             StringSplitOptions.RemoveEmptyEntries);
                                    foreach (var dateString in dateStrings)
                                    {
                                        try
                                        {
                                            var dateTime = Date.Evaluate(dateString, si.TodaysDate, m_Server);
                                            dropDown.Items.Add(dateTime.ToString());
                                        }
                                        catch (Exception)
                                        {
                                        }
                                    }

                                    if (dropDown.Items.FindByValue(p.ParamValue.ToString()) != null)
                                    {
                                        dropDown.SelectedValue = p.ParamValue.ToString();
                                    }
                                    else if (dropDown.Items.Count >= 1)
                                    {
                                        dropDown.SelectedIndex = 0;
                                    }

                                    //
                                    //  Add a Validation Handler
                                    //
                                    if (!p.AllowEntry)
                                    {
                                        dropDown.Enabled = false;
                                    }
                                    //
                                    //  Add help as a tooltip
                                    //
                                    if (p.HelpText.Length > 0)
                                    {
                                        dropDown.ToolTip = p.HelpText;
                                    }                                   
                                    dropDown.CssClass = "tabContents";
                                }
                                else
                                {
                                    //
                                    //  Use AJAX calendar extension with a textbox, instead of the normal calendar control
                                    //
                                    DateTime paramDT = (DateTime) p.ParamValue;
                                    TextBox calendarTextbox = new TextBox
                                                                  {
                                                                      ID = GetChildControlID(p.ParamName + "_" + p.ParamID)
                                                                  };

                                    CalendarExtender calendarExtender = new CalendarExtender
                                                                            {
                                                                                ID =
                                                                                    GetChildControlID(
                                                                                        "ParameterCalendar" + p.ParamID),
                                                                                Format = CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern
                                                                            };

                                    calendarExtender.TargetControlID = calendarTextbox.ClientID;
                                    var calendarMask = new MaskedEditExtender
                                                           {
                                                               MaskType = MaskedEditType.Date,
                                                               Mask = "99/99/9999",
                                                               TargetControlID = calendarTextbox.ClientID,
                                                               ID = GetChildControlID("CalendarMask" + p.ParamID)
                                                           };


                                    if (paramDT <= DateTime.MinValue)
                                    {
                                        //
                                        //  There is no date selected
                                        //
                                    }
                                    else
                                    {
                                        calendarExtender.SelectedDate = paramDT;
                                    }

                                    //
                                    //  Add a Validation Handler
                                    //
                                    if (!p.AllowEntry)
                                    {
                                        calendarTextbox.Enabled = false;
                                    }
                                    //
                                    //  Add help as a tooltip
                                    //
                                    if (p.HelpText.Length > 0)
                                    {
                                        calendarTextbox.ToolTip = p.HelpText;
                                    }

                                    calendarTextbox.CssClass = "tabContents";
                                    paramView.Controls.Add(calendarTextbox);
                                    paramView.Controls.Add(calendarExtender);
                                    paramView.Controls.Add(calendarMask);
                                }

                                if (baseControl != null) AddRequiredVaild(baseControl.ID, paramView);
                                break;
                            }
                        case MVPSI.JAMS.DataType.Time:
                            {
                                //
                                //  Add a Time Picker
                                //
                                MVPSI.JAMS.TimeOfDay paramTime = (TimeOfDay)p.ParamValue;

                                TextBox dt = new TextBox();
                                dt.ID = GetChildControlID(p.ParamName + "_" + p.ParamID);
                                dt.Columns = 8;

                                dt.MaxLength = p.Length;

                                dt.Enabled = p.AllowEntry;

                                if (paramTime.TotalSeconds < 0)
                                {
                                    //
                                    //  There is no time selected
                                    //
                                    if (!p.Required)
                                    {
                                        //
                                        //  There's no current value
                                        //
                                        //dt.Checked = false;
                                    }
                                }
                                else
                                {
                                    //DateTime paramDT = DateTime.Now;
                                    //paramDT -= paramDT.TimeOfDay;
                                    dt.Text = paramTime.ToString();
                                    //paramDT.AddTicks(paramTime.TotalSeconds*TimeSpan.TicksPerSecond).ToString();
                                }

                                if (!p.AllowEntry)
                                {
                                    dt.Enabled = false;
                                }
                                
                                //
                                //  Add help as a tooltip
                                //
                                if (p.HelpText.Length > 0)
                                {
                                    dt.ToolTip = p.HelpText;
                                }

                                //
                                //  Add a Validation Handler
                                //
                                if (p.Required)
                                {
                                    AddRequiredVaild(dt.ID, paramView);
                                }
                                AddRegexValid(dt.ID, @"(^(?:0?[1-9]:[0-5]|1(?=[012])\d:[0-5])\d( )?(AM|am|aM|Am|PM|pm|pM|Pm))$", "Time entered is not valid", paramView);

                                dt.CssClass = "tabContents";
                                baseControl = dt;

                                break;
                            }
                        case DataType.Boolean:
                            {
                                bool paramBool = (bool) p.ParamValue;
                                CheckBox cb = new CheckBox();
                                cb.ID = GetChildControlID(p.ParamName + "_" + p.ParamID);
                                cb.Text = p.Prompt;
                                cb.Checked = paramBool;
                                cb.Enabled = p.AllowEntry;

                                if (p.HelpText.Length > 0)
                                {
                                    cb.ToolTip = p.HelpText;
                                }

                                cb.CssClass = "tabContents checkBox";

                                baseControl = cb;
                                break;
                            }                      
                        default:
                            {                            
                                int maxControlWidth = 64;

                                switch (p.ValidationType)
                                {
                                    case ValidationType.Directory:
                                    case ValidationType.OpenFile:
                                    case ValidationType.SaveFile:

                                        var fileTextBox = new TextBox();
                                        fileTextBox.ID = GetChildControlID(p.ParamName + "_" + p.ParamID);
                                        baseControl = fileTextBox;

                                        if (!p.AllowEntry)
                                        {
                                            fileTextBox.Enabled = false;
                                        }

                                        fileTextBox.Style.Add("class", "tabContents");
                                        break;

                                    case ValidationType.MaskedEdit:
                                        var maskedExtender = new MaskedEditExtender();
                                        var textBox = new TextBox();
                                        textBox.ID = GetChildControlID(p.ParamName + "_" + p.ParamID);

                                        maskedExtender.Mask = p.ValidationData;
                                        maskedExtender.TargetControlID = GetChildControlID(p.ParamName + "_" + p.ParamID);
                                        maskedExtender.ID =
                                            GetChildControlID(p.ParamName + "_" + p.ParamID + "_" + "Extender");
                                        maskedExtender.MessageValidatorTip = true;
                                        maskedExtender.ClearMaskOnLostFocus = false;

                                        var maskedValidator = new MaskedEditValidator();
                                        maskedValidator.ControlToValidate = textBox.ClientID;
                                        maskedValidator.ControlExtender = maskedExtender.ClientID;
                                        maskedValidator.ErrorMessage = "Mask not valid.";
                                        maskedValidator.EnableClientScript = true;
                                        maskedValidator.Display = ValidatorDisplay.Dynamic;

                                        if (!p.AllowEntry)
                                        {
                                            textBox.ReadOnly = true;
                                        }

                                        paramView.Controls.Add(textBox);
                                        paramView.Controls.Add(maskedExtender);
                                        paramView.Controls.Add(maskedValidator);

                                        textBox.CssClass = "tabContents";
                                        textBox.Text = p.ParamValue.ToString();
                                        break;
                                    case ValidationType.Select:  
                                        var dropDown = new DropDownList();
                                        dropDown.ID = GetChildControlID(p.ParamName + "_" + p.ParamID);
                                        baseControl = dropDown;
                                        string[] valueStrings = p.ValidationData.Split(new char[] {','},
                                                                                       StringSplitOptions.
                                                                                           RemoveEmptyEntries);
                                        foreach (var s in valueStrings)
                                        {
                                            dropDown.Items.Add(s.Trim());
                                        }

                                        if (dropDown.Items.FindByValue(p.ParamValue.ToString()) != null)
                                        {
                                            dropDown.SelectedValue = p.ParamValue.ToString();                                          
                                        }
                                        else if (dropDown.Items.Count >= 1)
                                        {
                                            dropDown.SelectedIndex = 0;
                                        }

                                        dropDown.Enabled = p.AllowEntry;

                                        dropDown.CssClass = "tabContents";
                                        break;
                                    default:

                                        var tb = new TextBox();
                                        tb.ID = GetChildControlID(p.ParamName + "_" + p.ParamID);
                                        baseControl = tb;

                                        if (p.Uppercase)
                                        {
                                            tb.Style.Add("text-transform", "uppercase");
                                        }

                                        if (!p.AllowEntry)
                                        {
                                            tb.ReadOnly = true;
                                        }

                                        tb.Text = p.ParamValue.ToString();
                                        tb.ReadOnly = (!p.AllowEntry);

                                        if ((p.Length > maxControlWidth) || (p.Length == 0))
                                        {
                                            //
                                            //  This is larger than the dialog box
                                            //  Make it fit in the dialog but anchor each end
                                            //  so it stretches if they resize the dialog
                                            //
                                            tb.Columns = maxControlWidth;
                                        }
                                        else
                                        {
                                            tb.Columns = p.Length;
                                        }
                                        tb.MaxLength = p.Length;

                                        //
                                        //  We could have a string or an integer
                                        //
                                        if (p.ParamValue is String)
                                        {
                                            tb.Text = (string) p.ParamValue;

                                        }
                                        else if (p.ParamValue is Int32)
                                        {
                                            int paramInt = (int) p.ParamValue;
                                            tb.Text = paramInt.ToString();
                                        }
                                        else
                                        {
                                            //
                                            //  I don't know what this is!
                                            //
                                            tb.Text = string.Empty;
                                        }

                                        //
                                        //  Add help as a tooltip
                                        //
                                        if (p.HelpText.Length > 0)
                                        {
                                            tb.ToolTip = p.HelpText;
                                        }

                                        tb.CssClass = "tabContents";
                                        break;
                                }
                                break;
                            }
                    }

                    if (baseControl != null)
                    {
                        paramView.Controls.Add(baseControl);
                        SetupValidation(p, baseControl, paramView);
                    }
                }
            } 
      
            //If we don't have any parameters to show, hide the tab.
            if (hideParamTab)
                m_TabContainer.Tabs[0].Visible = false;
        }

        private void SetupValidation(Submit.Parameter p, Control ctl, TabPanel paramView)
        {
            var controlId = GetChildControlID(p.ParamName + "_" + p.ParamID);
            switch (p.DataType)
            {
                case DataType.Text:
                 
                        if (p.MustFill)
                        {
                            AddRegexValid(controlId, @".{" + p.Length + "}", "Value must be " + p.Length + " characters.", paramView);
                        }
                        else if(p.Required)
                        {
                            AddRequiredVaild(controlId, paramView);                           
                        }                   
                    if (p.ValidationType == ValidationType.Regex)
                    {
                        AddRegexValid(controlId, p.ValidationData, "Value does not match pattern.", paramView);

                        if (!string.IsNullOrEmpty(p.ValidationData))
                        {
                            AddRequiredVaild(controlId, paramView);
                        }
                    }
                    if (p.ValidationType == ValidationType.Range)
                    {
                        AddRangeValid(controlId, p, paramView);
                    }
                    break;
                case DataType.Date:
                case DataType.DateTime:
                    
                    break;
                case DataType.Time:
                    break;
                case DataType.Integer:
                    if (p.MustFill)
                    {
                        AddRegexValid(controlId, @".{" + p.Length + "}", "Value must be " + p.Length + " characters.", paramView);
                    }
                    else if (p.Required)
                    {
                        AddRequiredVaild(controlId, paramView);
                    }
                    if (p.ValidationType == ValidationType.Regex)
                    {
                        AddRegexValid(controlId, p.ValidationData, "Value does not match pattern.", paramView);
                        if (!string.IsNullOrEmpty(p.ValidationData))
                        {
                            AddRequiredVaild(controlId, paramView);
                        }
                    }
                    if (p.ValidationType == ValidationType.Range)
                    {
                        AddRangeValid(controlId, p, paramView);
                    }

                    AddRegexValid(controlId, @"[-+]?\b\d+\b", "Integer entered is not valid", paramView);

                    break;
                case DataType.Float:
                    if (p.MustFill)
                    {
                        AddRegexValid(controlId, @".{" + p.Length + "}", "Value must be " + p.Length + " characters.", paramView);
                    }
                    else if (p.Required)
                    {
                        AddRequiredVaild(controlId, paramView);
                    }
                    if (p.ValidationType == ValidationType.Regex)
                    {
                        AddRegexValid(controlId, p.ValidationData, "Value does not match pattern.", paramView);
                        if (!string.IsNullOrEmpty(p.ValidationData))
                        {
                            AddRequiredVaild(controlId, paramView);
                        }
                    }
                    if (p.ValidationType == ValidationType.Range)
                    {
                        AddRangeValid(controlId, p, paramView);
                    }

                    AddRegexValid(controlId, @"^[-+]?[0-9]*\.?[0-9]+$", "Float entered is not valid", paramView);
                    break;
                    default:
                    break;
            }
        }



        private void DisplayReports(Submit.Info si)
        {

            //Get the report tab panel to add controls to
            TabPanel reportView = m_TabContainer.Tabs[1];

            if (si.Reports.Count > 0)
            {
                m_TabContainer.ActiveTabIndex = 1;
                ListBox reportsBox = new ListBox();
                reportsBox.CssClass = "tabContents";
                reportsBox.SelectedIndex = 0;
                reportsBox.SelectedIndexChanged += new EventHandler(reportsBox_SelectedIndexChanged);
                reportsBox.SelectionMode = ListSelectionMode.Single;
                reportsBox.AutoPostBack = true;
                reportsBox.Style["width"] = "90%";

                m_PrintQueueTextBoxes = new List<TextBox>();
                m_PrintFormTextBoxes = new List<TextBox>();
                m_NumCopiesTextBoxes = new List<TextBox>();

                int i = 0;

                foreach (Submit.Report report in si.Reports)
                {
                    var newReport = new ListItem();
                    newReport.Text = report.ReportName + " - " + report.Description;
                    newReport.Value = report.ReportName + "_" + report.ReportID;
                    reportsBox.Items.Add(newReport);

                    TextBox printQueue = new TextBox { ID = "Report_" + report.ReportName + "_PrintQueue_" + report.ReportID, CssClass = "tabContents", Text = report.PrintQueue };
                    TextBox printForm = new TextBox { ID = "Report_" + report.ReportName + "_PrintForm_" + report.ReportID, CssClass = "tabContents", Text = report.PrintForm };
                    TextBox numCopies = new TextBox { ID = "Report_" + report.ReportName + "_NumberOfCopies_" + report.ReportID, CssClass = "tabContents", Text = report.Copies.ToString() };

                    printQueue.Style["Display"] = "None";
                    printForm.Style["Display"] = "None";
                    numCopies.Style["Display"] = "None";

                    m_PrintQueueTextBoxes.Add(printQueue);
                    m_PrintFormTextBoxes.Add(printForm);
                    m_NumCopiesTextBoxes.Add(numCopies);

                    AddRegexValid(numCopies.ID, @"\b\d+\b", "Copies must be a valid integer", reportView);

                    i++;
                }

                Label printQueueLabel = new Label { Text = "Print Queue" };
                Label PrintFormLabel = new Label { Text = "Print Form" };
                Label NumCopiesLabel = new Label { Text = "Number of Copies" };

                m_PrintQueueTextBoxes[0].Style.Clear();
                m_PrintFormTextBoxes[0].Style.Clear();
                m_NumCopiesTextBoxes[0].Style.Clear();

                //Add all the controls after they are setup.
                reportView.Controls.Add(reportsBox);
                reportView.Controls.Add(printQueueLabel);

                foreach (var box in m_PrintQueueTextBoxes)
                {
                    reportView.Controls.Add(box);
                }
                reportView.Controls.Add(PrintFormLabel);
                foreach (var box in m_PrintFormTextBoxes)
                {
                    reportView.Controls.Add(box);
                }
                reportView.Controls.Add(NumCopiesLabel);
                foreach (var box in m_NumCopiesTextBoxes)
                {
                    reportView.Controls.Add(box);
                }
            }
            else
            {
                //We don't have any reports, hide the reports tab
                reportView.Visible = false;
            }
        }

        void reportsBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ListBox reportBox = (ListBox)sender;

            foreach (var box in m_PrintQueueTextBoxes)
            {
                box.Style["Display"] = "None";
            }
            foreach (var box in m_PrintFormTextBoxes)
            {
                box.Style["Display"] = "None";
            }
            foreach (var box in m_NumCopiesTextBoxes)
            {
                box.Style["Display"] = "None";
            }

            m_PrintQueueTextBoxes[reportBox.SelectedIndex].Style.Clear();
            m_PrintFormTextBoxes[reportBox.SelectedIndex].Style.Clear();
            m_NumCopiesTextBoxes[reportBox.SelectedIndex].Style.Clear();
        }

        void cancelButton_Click(object sender, EventArgs e)
        {
            m_PopupControl.Hide();
            ClearViewState();
            m_TabContainer.ActiveTabIndex = 0;
        }

        void okButton_Click(object sender, EventArgs e)
        {
            Debug.WriteLine(string.Format("Submitting {0} with id of {1}", SubmitType, SubmitID));

            //
            //  Force disconnect and reconnect because we may be impersonating a different account
            //
            //Server.Disconnect();
            //Server.Connect();
            //string sst = this.SetSecuritySession(this.Page.Session.SessionID);

            Submit.Info si = new Submit.Info();
            Submit.Load(out si, SubmitID, Server, SubmitType);

            TabPanel paramView = m_TabContainer.Tabs[0];          

            foreach (Control control in paramView.Controls)
            {
                if (control is TextBox)
                {
                    //Get the last character, which is the parameter Id
                    int val = int.Parse(control.ID.Substring(control.ID.LastIndexOf('_') + 1));
                    Submit.Parameter param = si.Parameters[val-1];
                    if (param != null)
                    {
                        if (!param.Hide)
                        {
                            switch (param.DataType)
                            {                                      
                                case DataType.Date:
                                case DataType.DateTime:
                                    param.ParamValue = DateTime.Parse((control as TextBox).Text);
                                    break;
                                case DataType.Time:
                                    param.ParamValue = TimeOfDay.Parse((control as TextBox).Text);
                                    break;
                                case DataType.Integer:
                                    param.ParamValue = Int32.Parse((control as TextBox).Text);
                                    break;
                                default:
                                    if (param.Uppercase)
                                    {
                                        param.ParamValue = (control as TextBox).Text.ToUpper();
                                    }
                                    else
                                    {
                                        param.ParamValue = (control as TextBox).Text;
                                    }
                                    break;
                            }
                        }
                    }                   
                }
                else if (control is CheckBox)
                {
                    //Get the last character, which is the parameter Id
                    int val = int.Parse(control.ID.Substring(control.ID.LastIndexOf('_') + 1));
                    Submit.Parameter param = si.Parameters[val-1];
                    if ((param != null) && (param.DataType == DataType.Boolean))
                    {
                        param.ParamValue = (control as CheckBox).Checked;
                    }
                }
                else if (control is DropDownList)
                {
                    //Get the last character, which is the parameter Id
                    int val = int.Parse(control.ID.Substring(control.ID.LastIndexOf('_') + 1));
                    Submit.Parameter param = si.Parameters[val - 1];
                    if ((param != null) && (param.DataType == DataType.Text))
                    {
                        param.ParamValue = (control as DropDownList).SelectedValue;
                    }
                    else if ((param != null) && (param.DataType == DataType.Integer))
                    {
                        param.ParamValue = (control as DropDownList).SelectedValue;
                    }
                }              
            }

            TabPanel scheduleView = m_TabContainer.Tabs[2];

            var scheduledDateControl = scheduleView.FindControl(GetChildControlID("Schedule_Date")) as TextBox;
            var scheduledDate = DateTime.Parse(scheduledDateControl.Text);
            si.AfterTime = scheduledDate > DateTime.Today ? scheduledDate : DateTime.Today;

            TextBox scheduledTimeText = scheduleView.FindControl(GetChildControlID("Schedule_Time")) as TextBox;
            TimeOfDay scheduledTime = TimeOfDay.Parse(scheduledTimeText.Text);
            if (scheduledTime > TimeOfDay.Now)
            {
                si.AfterTime = si.AfterTime.Date + scheduledTime.ToTimeSpan();
            }
            else
            {
                si.AfterTime = si.AfterTime.Date + TimeOfDay.Now.ToTimeSpan();
            }

            if (si.ManageAccess)
            {
                si.BatchQueueName = (scheduleView.FindControl(GetChildControlID("Batch_Queues")) as DropDownList).Text;
                si.SchedulingPriority = int.Parse((scheduleView.FindControl(GetChildControlID("Schedule_Priority")) as TextBox).Text);
                if ((si.SubmitIsSimple) && (si.Jobs.Count > 0))
                {
                    si.Jobs[0].AgentNode = (scheduleView.FindControl(GetChildControlID("Agent_Node")) as TextBox).Text;
                    si.Jobs[0].SchedulingPriority = si.SchedulingPriority;
                }
            }

            //Get Report properties
            for (int i = 0; i < si.Reports.Count; i++)
            {
                Submit.Report report = si.Reports[i];
                report.PrintQueue = m_PrintQueueTextBoxes[i].Text;
                report.PrintForm = m_PrintFormTextBoxes[i].Text;
                report.Copies = int.Parse(m_NumCopiesTextBoxes[i].Text);
            }

            si.Submit();

            ClearViewState();

            if (DisplayModal)
            {
                this.m_PopupControl.Hide();
                this.Style["Display"] = "None";
            }

            m_TabContainer.ActiveTabIndex = 0;

            //
            //  Save the results
            //
            ViewState[ClientID + "ResultText"] = si.ResultText;
            
        }

    }
}
