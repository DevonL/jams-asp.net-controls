using System;
using System.Collections.Generic;
using System.IO;
using System.Web.UI;
using System.Web.UI.WebControls;
using AjaxControlToolkit;
using MVPSI.JAMS;

namespace MVPSI.JAMSWeb.Controls
{
    /// <summary>
    /// Displays details of the selected Job.
    /// </summary>
    internal class MonitorDialog:ControlsCommon
    {
        private TabContainer m_TabContainer;
        private Button m_OkButton;
        private const string StrNa = "N/A";
        private List<Button> m_ActionButtons;
        private Setup m_Setup;
        private Job m_Job;
        private ModalPopupExtender m_SubDialogExtender;
        private Panel m_SubDialog;

        //SubDialogControls
        private CheckBox m_AbortForceCheckbox;

        private CheckBox m_ReleaseHaltedSetup;
        private CheckBox m_ReleaseManual;
        private CheckBox m_ReleaseToRunAgain;
        private CheckBox m_ReleaseFromStepWait;
        private CheckBox m_ReleaseFromTimed;
        private CheckBox m_ReleaseFromTimeSlot;
        private CheckBox m_ReleaseIgnoreDepend;
        private CheckBox m_ReleaseIgnorePrecheck;
        private CheckBox m_ReleaseIgnoreResource;

        private DropDownList m_RescheduleBatchQueue;
        private TextBox m_RescheduleDate;
        private TextBox m_RescheduleTime;
        private CheckBox m_RescheduleHoldUntilManual;
        private TextBox m_ReschedulePriority;

        private CheckBox m_RestartHoldJob;
        private CalendarExtender m_Calendar;

        private Label m_SubLabel;
        private Label m_RescheduleBatchQueueTxt;
        private Label m_RescheduleDateTxt;
        private Label m_ReschedulePriorityTxt;
        private Label m_AbortSchedule;
        private Label m_AbortEntry;
        private Label m_AbortSubmitby;
        private Label m_AbortSubmitAt;
        private RegularExpressionValidator m_ValidPri;
        private RegularExpressionValidator m_ValidTime;
        private RequiredFieldValidator m_ValidDate;

        private MonitorAction CurrentAction
        {
            get
            {
                MonitorAction ma =
                    (MonitorAction) Enum.Parse(typeof (MonitorAction), ViewState[ClientID + "CurrentAction"].ToString());
                return ma;
            }
            set
            {
                ViewState[ClientID + "CurrentAction"] = value.ToString();
            }
        }

        private int m_CurJobRON
        {
            get
            {
                int ron = int.Parse(ViewState[ClientID + "JobRON"].ToString());
                return ron;
            }

            set
            {
                ViewState[ClientID + "JobRON"] = value;
            }
        }

        private int m_CurJobEntry { 
            get
            {
                int jamsEntry = int.Parse(ViewState[ClientID + "JAMSEntry"].ToString());
                return jamsEntry;
            }
            set
            {
                ViewState[ClientID + "JAMSEntry"] = value;
            }
        }

        private DateTime m_RescheduledDateValue
        {
            get
            {
                var date = DateTime.Parse(ViewState[ClientID + "RescheduledDateValue"].ToString());
                return date;
            }
            set
            {
                ViewState[ClientID + "RescheduledDateValue"] = value;
            }
        }

        /// <summary>
        /// Gets or sets the JAMS entry.
        /// </summary>
        /// <value>The JAMS entry.</value>
        public int JAMSEntry { get; set; }

        /// <summary>
        /// Gets or sets the extender.
        /// </summary>
        /// <value>The extender.</value>
        public ModalPopupExtender Extender { get; set; }

        /// <summary>
        /// Gets or sets the name of the server.
        /// </summary>
        /// <value>The name of the server.</value>
        public string ServerName{get;set;}

        /// <summary>
        /// Gets or sets the parent monitor.
        /// </summary>
        /// <value>The parent monitor.</value>
        public Monitor ParentMonitor { get; set; }

        public Panel Subdialog { get { return m_SubDialog; } }
        public ModalPopupExtender SubdialogExtender
        {
            get
            {
                return m_SubDialogExtender;
            }
        }

        public MonitorDialog()
        {
            m_SubDialogExtender = new ModalPopupExtender();            
            m_SubDialog = new Panel();           
        }

        internal void BuildSubDialog()
        {
            ServerName = ParentMonitor.ServerName;

            m_SubDialog.ID = GetChildControlID("SubDialog");           
            m_SubDialog.CssClass = "modalPopup";

            m_SubDialogExtender.PopupControlID = m_SubDialog.ClientID;
            m_SubDialogExtender.DropShadow = false;
            m_SubDialogExtender.RepositionMode = ModalPopupRepositionMode.RepositionOnWindowResize;
            m_SubDialogExtender.TargetControlID = GetChildControlID("SubDummyButton");
            m_SubDialogExtender.ID = GetChildControlID("SubDialogExtender");
            m_SubDialogExtender.BackgroundCssClass = "modalBackground";
            m_SubDialogExtender.CancelControlID = GetChildControlID("CancelButton");
            m_SubDialogExtender.PopupDragHandleControlID = GetChildControlID("SubDialog");
            m_SubDialogExtender.Y = 230; 

            var subDummyButton = new Button();
            subDummyButton.ID = GetChildControlID("SubDummyButton");
            subDummyButton.Style["Display"] = "none";
            subDummyButton.Height = 0;
            subDummyButton.Width = 0;
            subDummyButton.UseSubmitBehavior = false;

            var subButtonPanel = new Panel();
            var okButton = new Button { Text = "OK" };
            okButton.Click += new EventHandler(okButton_Click);
            var cancelButton = new Button { Text = "Cancel", ID = GetChildControlID("CancelButton"), UseSubmitBehavior = true };
            cancelButton.Click += new EventHandler(cancelButton_Click);
            subButtonPanel.Controls.Add(okButton);
            subButtonPanel.Controls.Add(cancelButton);

            //Add controls for each of the different actions, so that they can be hidden/displayed when needed.
            //In order for the control status to be available across postbacks, they need to be created in the Init method each time.
            //So, even though they may not be used, we need to create all of them now and add them to the subdialog.  They will all be
            //Hidden by default, and only the needed controls will be displayed, depending on what the user action is.
            m_AbortForceCheckbox = new CheckBox { Text = "Force job to abort?", Checked = false, ID = GetChildControlID("AbortForceCheckbox"), CssClass = "tabContents" };
            m_AbortForceCheckbox.Style["Display"] = "None";

            m_ReleaseHaltedSetup = new CheckBox { Text = "Release halted Setup", ID = GetChildControlID("ReleaseHaltedSetup"), CssClass = "tabContents" };
            m_ReleaseHaltedSetup.Style["Display"] = "None";
            m_ReleaseManual = new CheckBox { Text = "Release from manual hold", ID = GetChildControlID("ReleaseManual"), CssClass = "tabContents" };
            m_ReleaseManual.Style["Display"] = "None";
            m_ReleaseToRunAgain = new CheckBox { Text = "Release to run again", ID = GetChildControlID("ReleaseToRunAgain"), CssClass = "tabContents" };
            m_ReleaseToRunAgain.Style["Display"] = "None";
            m_ReleaseFromStepWait = new CheckBox { Text = "Release from Setup step wait", ID = GetChildControlID("ReleaseFromStepWait"), CssClass = "tabContents" };
            m_ReleaseFromStepWait.Style["Display"] = "None";
            m_ReleaseFromTimed = new CheckBox { Text = "Release from timed wait", ID = GetChildControlID("ReleaseFromTimed"), CssClass = "tabContents" };
            m_ReleaseFromTimed.Style["Display"] = "None";
            m_ReleaseFromTimeSlot = new CheckBox
            {
                Text = "Release from Time Slot wait",
                ID = GetChildControlID("RelaseFromTimeSlot"),
                CssClass = "tabContents"
            };
            m_ReleaseFromTimeSlot.Style["Display"] = "None";
            m_ReleaseIgnoreDepend = new CheckBox { Text = "Ignore Dependencies", ID = GetChildControlID("ReleaseIgnoreDepend"), CssClass = "tabContents" };
            m_ReleaseIgnoreDepend.Style["Display"] = "None";
            m_ReleaseIgnorePrecheck = new CheckBox
            {
                Text = "Ignore Precheck Job",
                ID = GetChildControlID("ReleaseIgnorePrecheck"),
                CssClass = "tabContents"
            };
            m_ReleaseIgnorePrecheck.Style["Display"] = "None";
            m_ReleaseIgnoreResource = new CheckBox { Text = "Ignore Resource Requirements", ID = GetChildControlID("ReleaseIgnoreResource"), CssClass = "tabContents" };
            m_ReleaseIgnoreResource.Style["Display"] = "None";

            m_RescheduleBatchQueue = new DropDownList { ID = GetChildControlID("RescheduleBatchQueue"), CssClass = "tabContents" };
            m_RescheduleBatchQueue.Style["Display"] = "None";
            m_RescheduleDate = new TextBox() { ID = GetChildControlID("RescheduleDate"), CssClass = "tabContents" };
            m_RescheduleDate.Style["Display"] = "None";
            MaskedEditExtender calendarMask = new MaskedEditExtender { MaskType = MaskedEditType.Date, Mask = "99/99/9999", TargetControlID = m_RescheduleDate.ClientID, ID = GetChildControlID("CalendarMask") };
            m_RescheduleTime = new TextBox { ID = GetChildControlID("RescheduleTime"), CssClass = "tabContents" };
            m_RescheduleTime.Style["Display"] = "None";
            MaskedEditExtender rescheduleMask = new MaskedEditExtender { AcceptAMPM = true, MaskType = MaskedEditType.Time, TargetControlID = m_RescheduleTime.ClientID, Mask = "99:99", ID = GetChildControlID("TimeMask") };
            m_RescheduleHoldUntilManual = new CheckBox
            {
                Text = "Hold until manually released.",
                ID = GetChildControlID("RescheduleHoldUntilManual"),
                CssClass = "tabContents"
            };
            m_RescheduleHoldUntilManual.Style["Display"] = "None";
            m_ReschedulePriority = new TextBox { ID = GetChildControlID("ReschedulePriority"), CssClass = "tabContents" };
            m_ReschedulePriority.Style["Display"] = "None";
            var priorityMask = new MaskedEditExtender { Mask = "9999", MaskType = MaskedEditType.Number, TargetControlID = m_ReschedulePriority.ClientID, AcceptNegative = MaskedEditShowSymbol.None };

            m_RestartHoldJob = new CheckBox { Text = "Hold Job", ID = GetChildControlID("RestartHoldJob"), Checked = false, CssClass = "tabContents" };
            m_RestartHoldJob.Style["Display"] = "None";

            //Fill dropdown with batch queues
            m_RescheduleBatchQueue.Items.Add("");

            ICollection<BatchQueue> queues;
            object savedQueues = Server.GetServer(ServerName).GetObject("BatchQueues");
            if (savedQueues is ICollection<BatchQueue>)
            {
                queues = (ICollection<BatchQueue>)savedQueues;
            }
            else
            {
                queues = BatchQueue.Find(Server.CurrentServer);

                Server.CurrentServer.SaveObject("BatchQueues", queues);
            }

            foreach (BatchQueue queue in queues)
            {
                m_RescheduleBatchQueue.Items.Add(queue.QueueName);
            }
            m_RescheduleBatchQueue.AutoPostBack = false;

            m_AbortEntry = new Label { CssClass = "tabContents" };
            m_AbortSubmitby = new Label { CssClass = "tabContents" };
            m_AbortSubmitAt = new Label {  CssClass = "tabContents" };
            m_AbortSchedule = new Label {  CssClass = "tabContents" };
           
            m_SubLabel = new Label { CssClass = "tabContents" };
            m_SubDialog.Controls.Add(m_SubLabel);
            m_SubDialog.Controls.Add(m_AbortForceCheckbox);
            m_SubDialog.Controls.Add(m_ReleaseHaltedSetup);
            m_SubDialog.Controls.Add(m_ReleaseManual);
            m_SubDialog.Controls.Add(m_ReleaseToRunAgain);
            m_SubDialog.Controls.Add(m_ReleaseFromStepWait);
            m_SubDialog.Controls.Add(m_ReleaseFromTimed);
            m_SubDialog.Controls.Add(m_ReleaseFromTimeSlot);
            m_SubDialog.Controls.Add(m_ReleaseIgnoreDepend);
            m_SubDialog.Controls.Add(m_ReleaseIgnorePrecheck);
            m_SubDialog.Controls.Add(m_ReleaseIgnoreResource);
            m_RescheduleBatchQueueTxt = new Label { Text = "Batch Queue", Visible = false, CssClass = "tabContents" };
            m_SubDialog.Controls.Add(m_RescheduleBatchQueueTxt);
            m_SubDialog.Controls.Add(m_RescheduleBatchQueue);
            m_RescheduleDateTxt =  new Label { Text = "Scheduled Time", Visible = false, CssClass = "tabContents" };
            m_SubDialog.Controls.Add(m_RescheduleDateTxt);
            m_SubDialog.Controls.Add(m_RescheduleDate);
            m_SubDialog.Controls.Add(m_RescheduleTime);
            m_SubDialog.Controls.Add(m_RescheduleHoldUntilManual);
            m_ReschedulePriorityTxt = new Label { Text = "Scheduling Priority", Visible = false, CssClass = "tabContents" };
            m_SubDialog.Controls.Add(m_ReschedulePriorityTxt);
            m_SubDialog.Controls.Add(m_ReschedulePriority);
            m_SubDialog.Controls.Add(m_RestartHoldJob);
            m_SubDialog.Controls.Add(subButtonPanel);
            m_SubDialog.Controls.Add(m_AbortEntry);
            m_SubDialog.Controls.Add(m_AbortSubmitby);
            m_SubDialog.Controls.Add(m_AbortSubmitAt);
            m_SubDialog.Controls.Add(m_AbortSchedule);
            m_SubDialog.Controls.Add(rescheduleMask);
            m_SubDialog.Controls.Add(calendarMask);
            m_SubDialog.Controls.Add(priorityMask);
            m_Calendar = new CalendarExtender();
            m_Calendar.TargetControlID = m_RescheduleDate.ID;
            
            m_SubDialog.Controls.Add(m_Calendar);
            m_ValidPri = AddRegexValid(m_ReschedulePriority.ID, @"[-+]?\b\d+\b", "Integer entered is not valid", m_SubDialog);
            m_ValidPri.Enabled = false;
            m_ValidTime = AddRegexValid(m_RescheduleTime.ID, @"(^(?:0?[1-9]:[0-5]|1(?=[012])\d:[0-5])\d( )?(AM|am|aM|Am|PM|pm|pM|Pm))$", "Time entered is not valid", m_SubDialog);
            m_ValidTime.Enabled = false;
            m_ValidDate = AddRequiredVaild(m_RescheduleDate.ID, m_SubDialog);
            m_ValidDate.Enabled = false;
            
            Controls.Add(subDummyButton);
        }

        protected override void OnInit(System.EventArgs e)
        {
            base.OnInit(e);
            
            CssClass = "modalPopup";

            m_ActionButtons = new List<Button>();

            var release = new Button { Text = "Release" };
            var hold = new Button { Text = "Hold" };
            var reschedule = new Button { Text = "Reschedule" };
            var restart = new Button { Text = "Restart" };
            var abort = new Button { Text = "Abort"};

            release.Click += new System.EventHandler(release_Click);
            hold.Click += new System.EventHandler(hold_Click);
            reschedule.Click += new System.EventHandler(reschedule_Click);
            restart.Click += new System.EventHandler(restart_Click);
            abort.Click += new System.EventHandler(abort_Click);

            m_ActionButtons.Add(release);
            m_ActionButtons.Add(hold);
            m_ActionButtons.Add(reschedule);
            m_ActionButtons.Add(restart);
            m_ActionButtons.Add(abort);

            m_OkButton = new Button();
            m_OkButton.ID = GetChildControlID("MonitorDialogOK");
            m_OkButton.Text = "OK";
            m_OkButton.UseSubmitBehavior = true;
            m_OkButton.CausesValidation = false;
            m_OkButton.Style["margin"] = "20px, 50px, 10px, 0px";

            var buttonPanel = new Panel();
            buttonPanel.ID = GetChildControlID("Monitor_Button_Panel");
            buttonPanel.Controls.Add(m_OkButton);
            foreach (var button in m_ActionButtons)
            {
                button.CausesValidation = false;
                buttonPanel.Controls.Add(button);
            }
            buttonPanel.EnableViewState = false;
            buttonPanel.HorizontalAlign = HorizontalAlign.Center;

            Controls.Add(buttonPanel);

            BuildSubDialog();

            this.ParentMonitor.SubPanel.Controls.Add(Subdialog);
            this.ParentMonitor.SubPanel.Controls.Add(SubdialogExtender);
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            if (JAMSEntry != 0)
            {
                CurJob loadedJob = new CurJob();
                try
                {
                    CurJob.Load(out loadedJob, JAMSEntry, Server.GetServer(ParentMonitor.ServerName));
                    Style["Display"] = "Inline";
                    ShowDialog(loadedJob);                    
                }
                catch(Exception)
                {
                    Style["Display"] = "None";
                    Extender.Hide();
                }

            }            
            else
            {
                Style["Display"] = "None";
                Extender.Hide();
            }

        }

        /// <summary>
        /// Shows the dialog.
        /// </summary>
        /// <param name="curJob">The cur job.</param>
        private void ShowDialog(CurJob curJob)
        {
            m_CurJobEntry = curJob.JAMSEntry;
            m_CurJobRON = curJob.RON;
         
            Extender.Show();
            Style["Display"] = "inline";

            Panel titleBar = new Panel();
            titleBar.ID = GetChildControlID("TitleBar");
            Extender.PopupDragHandleControlID = GetChildControlID("TitleBar");
            titleBar.Style["background-color"] = "#CEE5FE";
            titleBar.Style["margin"] = "-3, -3, 6, -3";

            try
            {
                JAMS.Job.Load(out m_Job, curJob.JobID, curJob.LoadedFrom);
            }
            catch (Exception ex)
            {
                m_Job = new Job();
                m_Job.Description = ex.Message;                
            }

            try
            {
                Setup.Load(out m_Setup, curJob.SetupID, curJob.LoadedFrom);
            }
            catch (Exception ex)
            {
                m_Setup = new Setup();
                m_Setup.Description = ex.Message;                
            }


            Label head = new Label();
            head.Text = "Monitor Details for Job: " + curJob.JobName + " - ";
            titleBar.Controls.Add(head);

            head = new Label();
            head.Text = "JAMS Entry: " + curJob.JAMSEntry;
            titleBar.Controls.Add(head);
            Controls.AddAt(0, titleBar);

            m_TabContainer = new TabContainer { ID = GetChildControlID("MonitorTabContainer") };

            var general = new TabPanel { ID = GetChildControlID("GeneralTab")};
            var status = new TabPanel { ID = GetChildControlID("StatusTab") };
            var statistics = new TabPanel { ID = GetChildControlID("StatisticsTab") };
            var logfile = new TabPanel { ID = GetChildControlID("LogFileTab") };
            var dependencies = new TabPanel {ID = GetChildControlID("Dependencies")};
            var parameters = new TabPanel { ID = GetChildControlID("Parameters")};

            general.HeaderText = "General";
            status.HeaderText = "Status";
            statistics.HeaderText = "Statistics";
            logfile.HeaderText = "Log File";
            dependencies.HeaderText = "Dependencies";
            parameters.HeaderText = "Parameters";

            //Todo: this will use the CustomTabStyle class, which doesn't currently display images.
            // m_TabContainer.CssClass = "CustomTabStyle";

            m_TabContainer.Tabs.Add(general);
            m_TabContainer.Tabs.Add(status);
            m_TabContainer.Tabs.Add(statistics);
            m_TabContainer.Tabs.Add(logfile);
            m_TabContainer.Tabs.Add(dependencies);
            m_TabContainer.Tabs.Add(parameters);

            m_TabContainer.ActiveTabIndex = 0;
            m_TabContainer.ScrollBars = ScrollBars.Auto;
            m_TabContainer.Height = Unit.Pixel(350);
            
            //Create the controls on the tabs, fill in data from the CurJob object.

            //General Tab
            var job = new Label { CssClass = "tabContents" };
            var system = new Label { CssClass = "tabContents" };
            var setup = new Label { CssClass = "tabContents" };
            var submitInfo = new Label { CssClass = "tabContents" };

            if (curJob.JobID != 0)
            {
                job.Text = string.Format("Job\n  {0}\n\n  {1}", curJob.JobName, m_Job.Description);
                system.Text = string.Format("System\n  {0}", m_Job.ParentFolderName);
                setup.Text = string.Format("Setup\n  {0}", StrNa);
                submitInfo.Text = string.Format("Submit Information\n  JAMS Entry {0}, Queue Entry {1}, RON {2}",
                                             curJob.JAMSEntry, curJob.BatchQueue, curJob.RON);
            }
            else if(curJob.SetupID != 0)
            {
                job.Text = string.Format("Job\n  {0}\n\n  {1}", curJob.JobName, m_Job.Description);
                system.Text = string.Format("System\n  {0}", m_Job.ParentFolderName);
                setup.Text = string.Format("Setup\n  {0}", m_Setup.SetupName);
                submitInfo.Text = string.Format("Submit Information\n  JAMS Entry {0}, Queue Entry {1}, RON {2}",
                                             curJob.JAMSEntry, curJob.BatchQueue, curJob.RON);
            }
            else
            {
                job.Text = string.Format("Job\n  {0}\n\n  {1}", m_Setup.SetupName, m_Setup.Description);
                system.Text = string.Format("System\n  {0}", m_Setup.ParentFolderName);
                setup.Text = string.Format("Setup\n  {0}", StrNa);
                submitInfo.Text = string.Format("Submit Information\n  JAMS Entry {0}, Queue Entry {1}, RON {2}",
                                             curJob.JAMSEntry, curJob.BatchQueue, curJob.RON);
            }

            general.Controls.Add(job);
            general.Controls.Add(system);
            general.Controls.Add(setup);
            general.Controls.Add(submitInfo);

            //Status Tab
            var finalStatus = new Label { Text = string.Format("Final Status\n  {0}", curJob.FinalStatus), CssClass = "tabContents" };
            var note = new Label { Text = string.Format("Note\n  {0}", curJob.Note), CssClass = "tabContents" };
            var jobStatusText = new Label { Text = string.Format("Job Status Text\n  {0}", curJob.JobStatus), CssClass = "tabContents" };

            status.Controls.Add(finalStatus);
            status.Controls.Add(note);
            status.Controls.Add(jobStatusText);

            //Stats Tab
            var times = new Label { Text = string.Format("Times\n  ") };
            var timesTable = new Table();

            for (int i = 0; i < 4; i++)
            {
                var timeRow = new TableRow();
                timesTable.Rows.Add(timeRow);
            }

            var cells = new List<TableCell>();
            cells.Add(new TableCell { Text = "Scheduled At: " });
            cells.Add(new TableCell { Text = curJob.ScheduledTime.ToString() });
            cells.Add(new TableCell { Text = "Scheduled For: " });
            cells.Add(new TableCell { Text = curJob.HoldTime.ToString() });
            cells.Add(new TableCell { Text = "Started: " });
            cells.Add(new TableCell { Text = curJob.StartTime.ToString() });
            cells.Add(new TableCell { Text = "Completed: " });
            cells.Add(new TableCell { Text = curJob.CompletionTime.ToString() });

            int j = 0;
            for (int i = 0; i < timesTable.Rows.Count; i++)
            {
                timesTable.Rows[i].Cells.Add(cells[j]);
                timesTable.Rows[i].Cells.Add(cells[j + 1]);
                j = j + 2;
            }

            var stats = new Label { Text = "Statistics" };
            var statsTable = new Table();

            for (int i = 0; i < 7; i++)
            {
                var statRow = new TableRow();
                statsTable.Rows.Add(statRow);
            }

            var statCells = new List<TableCell>();

            if(curJob.JobID != 0)
            {

                statCells.Add(new TableCell {Text = ""});
                statCells.Add(new TableCell {Text = "Current"});
                statCells.Add(new TableCell {Text = "Minimum"});
                statCells.Add(new TableCell {Text = "Maximum"});
                statCells.Add(new TableCell {Text = "Average"});
                statCells.Add(new TableCell {Text = "% of Avg."});

                statCells.Add(new TableCell {Text = "Elapsed Time: "});
                statCells.Add(new TableCell {Text = curJob.ElapsedTime.ToString()});
                statCells.Add(new TableCell {Text = m_Job.MinElapsedTime.ToString()});
                statCells.Add(new TableCell { Text = m_Job.MaxElapsedTime.ToString() });
                statCells.Add(new TableCell { Text = m_Job.AvgElapsedTime.ToString() });
                statCells.Add(new TableCell
                                  {
                                      Text = CalcPercent((int)curJob.ElapsedTime.TotalSeconds,
                                                         m_Job.AvgElapsedTime.TotalSeconds)
                                  });

                statCells.Add(new TableCell {Text = "Cpu Time: "});
                statCells.Add(new TableCell {Text = ""});//CpuTime.ToString() });
                statCells.Add(new TableCell { Text = m_Job.MinCpuTime.ToString() });
                statCells.Add(new TableCell { Text = m_Job.MaxCpuTime.ToString() });
                statCells.Add(new TableCell { Text = m_Job.AvgCpuTime.ToString() });
                statCells.Add(new TableCell
                                  {
//                                      Text = CalcPercent(curJob.CpuTime.TotalSeconds,
//                                                         curJob.Job.AvgCpuTime.TotalSeconds)
                                  });

                statCells.Add(new TableCell {Text = "Direct I/O: "});
                statCells.Add(new TableCell { Text = ""});//curJob.DirectIOCount.ToString() });
                statCells.Add(new TableCell { Text = m_Job.MinDirectIOCount.ToString() });
                statCells.Add(new TableCell { Text = m_Job.MaxDirectIOCount.ToString() });
                statCells.Add(new TableCell { Text = m_Job.AvgDirectIOCount.ToString() });
                statCells.Add(new TableCell { Text = "" });//CalcPercent(curJob.DirectIOCount, curJob.Job.AvgDirectIOCount) });

                statCells.Add(new TableCell {Text = "Buffered I/O: "});
                statCells.Add(new TableCell { Text = "" });//curJob.BufferedIOCount.ToString() });
                statCells.Add(new TableCell { Text = m_Job.MinBufferedIOCount.ToString() });
                statCells.Add(new TableCell { Text = m_Job.MaxBufferedIOCount.ToString() });
                statCells.Add(new TableCell { Text = m_Job.AvgBufferedIOCount.ToString() });
                statCells.Add(new TableCell { Text = "" });//CalcPercent(curJob.BufferedIOCount, curJob.Job.AvgBufferedIOCount) });

                statCells.Add(new TableCell {Text = "Page Faults: "});
                statCells.Add(new TableCell { Text = "" });//curJob.PageFaults.ToString() });
                statCells.Add(new TableCell { Text = m_Job.MinPageFaults.ToString() });
                statCells.Add(new TableCell { Text = m_Job.MaxPageFaults.ToString() });
                statCells.Add(new TableCell { Text = m_Job.AvgPageFaults.ToString() });
                statCells.Add(new TableCell { Text = "" });//CalcPercent(curJob.PageFaults, curJob.Job.AvgPageFaults) });

                statCells.Add(new TableCell {Text = "Peak W/S: "});
                statCells.Add(new TableCell { Text = "" });//curJob.WorkingSetPeak.ToString() });
                statCells.Add(new TableCell { Text = m_Job.MinWorkingSetPeak.ToString() });
                statCells.Add(new TableCell { Text = m_Job.MaxWorkingSetPeak.ToString() });
                statCells.Add(new TableCell { Text = m_Job.AvgWorkingSetPeak.ToString() });
                statCells.Add(new TableCell { Text = "" });//CalcPercent(curJob.WorkingSetPeak, curJob.Job.AvgWorkingSetPeak) });
            }
            else if(curJob.SetupID != 0)
            {
                statCells.Add(new TableCell {Text = ""});
                statCells.Add(new TableCell {Text = "Current"});
                statCells.Add(new TableCell {Text = "Minimum"});
                statCells.Add(new TableCell {Text = "Maximum"});
                statCells.Add(new TableCell {Text = "Average"});
                statCells.Add(new TableCell {Text = "% of Avg."});

                statCells.Add(new TableCell {Text = "Elapsed Time: "});
                statCells.Add(new TableCell { Text = curJob.ElapsedTime.ToString() });
                statCells.Add(new TableCell { Text = m_Setup.MinElapsedTime.ToString() });
                statCells.Add(new TableCell { Text = m_Setup.MaxElapsedTime.ToString() });
                statCells.Add(new TableCell { Text = m_Setup.AvgElapsedTime.ToString() });
                statCells.Add(new TableCell
                                  {
                                      Text =  ""});//CalcPercent(curJob.ElapsedTime.TotalSeconds,curJob.Setup.AvgElapsedTime.TotalSeconds)
                                  //});

                statCells.Add(new TableCell {Text = "Cpu Time: "});
                statCells.Add(new TableCell {Text = StrNa});
                statCells.Add(new TableCell {Text = StrNa});
                statCells.Add(new TableCell {Text = StrNa});
                statCells.Add(new TableCell {Text = StrNa});
                statCells.Add(new TableCell {Text = StrNa});

                statCells.Add(new TableCell {Text = "Direct I/O: "});
                statCells.Add(new TableCell {Text = StrNa});
                statCells.Add(new TableCell {Text = StrNa});
                statCells.Add(new TableCell {Text = StrNa});
                statCells.Add(new TableCell {Text = StrNa});
                statCells.Add(new TableCell {Text = StrNa});

                statCells.Add(new TableCell {Text = "Buffered I/O: "});
                statCells.Add(new TableCell {Text = StrNa});
                statCells.Add(new TableCell {Text = StrNa});
                statCells.Add(new TableCell {Text = StrNa});
                statCells.Add(new TableCell {Text = StrNa});
                statCells.Add(new TableCell {Text = StrNa});

                statCells.Add(new TableCell {Text = "Page Faults: "});
                statCells.Add(new TableCell {Text = StrNa});
                statCells.Add(new TableCell {Text = StrNa});
                statCells.Add(new TableCell {Text = StrNa});
                statCells.Add(new TableCell {Text = StrNa});
                statCells.Add(new TableCell {Text = StrNa});

                statCells.Add(new TableCell {Text = "Peak W/S: "});
                statCells.Add(new TableCell {Text = StrNa});
                statCells.Add(new TableCell {Text = StrNa});
                statCells.Add(new TableCell {Text = StrNa});
                statCells.Add(new TableCell {Text = StrNa});
                statCells.Add(new TableCell {Text = StrNa});


            }
            else
            {
                statCells.Add(new TableCell {Text = ""});
                statCells.Add(new TableCell {Text = "Current"});
                statCells.Add(new TableCell {Text = "Minimum"});
                statCells.Add(new TableCell {Text = "Maximum"});
                statCells.Add(new TableCell {Text = "Average"});
                statCells.Add(new TableCell {Text = "% of Avg."});

                statCells.Add(new TableCell {Text = "Elapsed Time: "});
                statCells.Add(new TableCell { Text = curJob.ElapsedTime.ToString() });
                statCells.Add(new TableCell {Text = StrNa});
                statCells.Add(new TableCell {Text = StrNa});
                statCells.Add(new TableCell {Text = StrNa});
                statCells.Add(new TableCell {Text = StrNa});

                statCells.Add(new TableCell {Text = "Cpu Time: "});
                statCells.Add(new TableCell { Text = "" });//curJob.CpuTime.ToString() });
                statCells.Add(new TableCell {Text = StrNa});
                statCells.Add(new TableCell {Text = StrNa});
                statCells.Add(new TableCell {Text = StrNa});
                statCells.Add(new TableCell {Text = StrNa});

                statCells.Add(new TableCell {Text = "Direct I/O: "});
                statCells.Add(new TableCell { Text = "" });//curJob.DirectIOCount.ToString() });
                statCells.Add(new TableCell {Text = StrNa});
                statCells.Add(new TableCell {Text = StrNa});
                statCells.Add(new TableCell {Text = StrNa});
                statCells.Add(new TableCell {Text = StrNa});

                statCells.Add(new TableCell {Text = "Buffered I/O: "});
                statCells.Add(new TableCell { Text = "" });//curJob.BufferedIOCount.ToString() });
                statCells.Add(new TableCell {Text = StrNa});
                statCells.Add(new TableCell {Text = StrNa});
                statCells.Add(new TableCell {Text = StrNa});
                statCells.Add(new TableCell {Text = StrNa});

                statCells.Add(new TableCell {Text = "Page Faults: "});
                statCells.Add(new TableCell { Text = "" });//curJob.PageFaults.ToString() });
                statCells.Add(new TableCell {Text = StrNa});
                statCells.Add(new TableCell {Text = StrNa});
                statCells.Add(new TableCell {Text = StrNa});
                statCells.Add(new TableCell {Text = StrNa});

                statCells.Add(new TableCell {Text = "Peak W/S: "});
                statCells.Add(new TableCell { Text = "" });//curJob.WorkingSetPeak.ToString() });
                statCells.Add(new TableCell {Text = StrNa});
                statCells.Add(new TableCell {Text = StrNa});
                statCells.Add(new TableCell {Text = StrNa});
                statCells.Add(new TableCell {Text = StrNa});
             //todo: update the current stats somehow
            }
        

            j = 0;
            for (int i = 0; i < statsTable.Rows.Count; i++)
            {
                statsTable.Rows[i].Cells.Add(statCells[j]);
                statsTable.Rows[i].Cells.Add(statCells[j + 1]);
                statsTable.Rows[i].Cells.Add(statCells[j + 2]);
                statsTable.Rows[i].Cells.Add(statCells[j + 3]);
                statsTable.Rows[i].Cells.Add(statCells[j + 4]);
                statsTable.Rows[i].Cells.Add(statCells[j + 5]);
                j = j + 6;
            }

            statistics.Controls.Add(times);
            statistics.Controls.Add(timesTable);
            statistics.Controls.Add(stats);
            statistics.Controls.Add(statsTable);

            //LogFile tab
            var logLocation = new Label { CssClass = "tabContents" };
            var logFileContent = new Label { CssClass = "tabContents" };

            try
            {
                Stream cjLog = curJob.LogFile;

                if (cjLog == null)
                {
                    logFileContent.Text = "The log file is not currently available.";
                }
                else
                {                    
                    StreamReader logSR = new StreamReader(cjLog, true);
                    logFileContent.Text = logSR.ReadToEnd();
                    cjLog.Close();

                    logLocation.Text = curJob.LogFilename;
                }
            }
            catch(Exception ex)
            {
                string exMessage = "The log file is not available";
                while (ex != null)
                {
                    exMessage = string.Format("{0}{1}{2}", exMessage, Environment.NewLine, ex.Message);
                    ex = ex.InnerException;
                }
                logFileContent.Text = exMessage;
            }

            var preTag = new System.Web.UI.HtmlControls.HtmlGenericControl();
            preTag.TagName = "pre";
            preTag.Controls.Add(logFileContent);

            logfile.Controls.Add(logLocation);
            logfile.Controls.Add(preTag);

            //Dependencies tab
            var depenTable = new Table();

            var dependencyList = JAMS.CurDependency.Find(curJob.JobID, curJob.SetupID, curJob.JAMSEntry, curJob.LoadedFrom);

            var headerRow = new TableRow();
            headerRow.Cells.Add(new TableCell { Text = "Status" });
            headerRow.Cells.Add(new TableCell { Text = "Description" });
            depenTable.Rows.Add(headerRow);

            foreach (CurDependency dependency in dependencyList)
            {
                var row = new TableRow();
                row.Cells.Add(new TableCell {Text = dependency.State.ToString()});
                row.Cells.Add(new TableCell {Text = dependency.Description});
                depenTable.Rows.Add(row);
            }

            dependencies.Controls.Add(depenTable);

            //Parameters tab

            var paramTable = new Table();

            var paramList = CurParam.Find(curJob.JAMSEntry, curJob.LoadedFrom);

            headerRow = new TableRow();
            headerRow.Cells.Add(new TableCell { Text = "Parameter Name" });
            headerRow.Cells.Add(new TableCell { Text = "Parameter Value" });
            paramTable.Rows.Add(headerRow);

            foreach (CurParam param in paramList)
            {
                var row = new TableRow();
                row.Cells.Add(new TableCell { Text = param.ParamName });
                row.Cells.Add(new TableCell { Text = param.Value.ToString() });
                paramTable.Rows.Add(row);
            }

            parameters.Controls.Add(paramTable);

            //Once all the controls have been created, add them to the dialog.
            Controls.AddAt(1, m_TabContainer);

            //Set which buttons are enabled or disabled
            SetButtonStates(curJob);
        }

        private void SetButtonStates(CurJob curJob)
        {
            foreach (var button in m_ActionButtons)
            {
                button.Enabled = false;

                switch (button.Text)
                {
                    case "ReleaseJob":
                        if (curJob.CurrentState != JobState.Executing)
                        {
                            button.Enabled = true;
                        }
                        if (curJob.MasterRON == curJob.RON)
                        {
                            if(curJob.Halted)
                            {
                                button.Enabled = true;
                            }
                        }
                        break;
                    case "Hold":
                        if(curJob.CurrentState < JobState.Executing)
                        {
                            button.Enabled = true;
                        }
                        if (curJob.MasterRON == curJob.RON)
                        {
                            if (curJob.CurrentState <= JobState.Executing)
                            {
                                button.Enabled = true;
                            }
                        }
                        break;
                    case "Reschedule":
                        if(curJob.CurrentState < JobState.Executing)
                        {
                            button.Enabled = true;
                        }
                        break;
                    default:
                        button.Enabled = true;
                        break;
                        
                }
            }

            
        }

        void cancelButton_Click(object sender, EventArgs e)
        {
            //Don't do anything, user canceled.
        }

        void okButton_Click(object sender, EventArgs e)
        {
            try
            {

                //Find what action was supposed to occur, then perform the action.
                switch (CurrentAction)
                {
                    case MonitorAction.Abort:
                        CurJob.Delete(m_CurJobEntry, m_AbortForceCheckbox.Checked, Server.CurrentServer);
                        break;
                    case MonitorAction.Hold:
                        CurJob.Hold(m_CurJobEntry);
                        break;
                    case MonitorAction.Release:

                        ReleaseType releaseType = 0;

                        if (m_ReleaseHaltedSetup.Checked)
                        {
                            releaseType |= ReleaseType.HaltedSetup;
                        }
                        if (m_ReleaseManual.Checked)
                        {
                            releaseType |= ReleaseType.ManualHold;
                        }
                        if (m_ReleaseToRunAgain.Checked)
                        {
                            releaseType |= ReleaseType.RunAgain;
                        }
                        if (m_ReleaseFromStepWait.Checked)
                        {
                            releaseType |= ReleaseType.StepWait;
                        }
                        if (m_ReleaseFromTimed.Checked)
                        {
                            releaseType |= ReleaseType.TimedWait;
                        }
                        if (m_ReleaseFromTimeSlot.Checked)
                        {
                            releaseType |= ReleaseType.TimeSlotWait;
                        }
                        if (m_ReleaseIgnoreDepend.Checked)
                        {
                            releaseType |= ReleaseType.IgnoreDependencies;
                        }
                        if (m_ReleaseIgnorePrecheck.Checked)
                        {
                            releaseType |= ReleaseType.IgnorePrecheck;
                        }
                        if (m_ReleaseIgnoreResource.Checked)
                        {
                            releaseType |= ReleaseType.IgnoreResReq;
                        }

                        CurJob.Release(m_CurJobEntry, releaseType);
                        break;
                    case MonitorAction.Reschedule:
                        CurJob curJob;
                        CurJob.Load(out curJob, m_CurJobEntry, Server.CurrentServer);

                        curJob.BatchQueueName = m_RescheduleBatchQueue.Text;

                        DateTime newDate;
                        var parse = DateTime.TryParse(m_RescheduleDate.Text, out newDate);

                        if (parse)
                        {
                            curJob.ScheduledTime = newDate += TimeOfDay.Parse(m_RescheduleTime.Text).ToTimeSpan();
                        }

                        curJob.Held = m_RescheduleHoldUntilManual.Checked;
                        curJob.SchedulingPriority = int.Parse(m_ReschedulePriority.Text);

                        curJob.Reschedule();
                        break;
                    case MonitorAction.Restart:
                        CurJob.Restart(m_CurJobEntry, m_RestartHoldJob.Checked);
                        break;
                    default:
                        break;
                }
                m_CurJobEntry = 0;
            }
            catch(Exception ex)
            {
                ParentMonitor.m_ErrorPanel.Controls.AddAt(1, new Label { Text = ex.Message, CssClass = "tabContents" });
                ParentMonitor.m_ErrorDialog.Show();
                ParentMonitor.m_ErrorPanel.Style["Display"] = "Inline";
            }
        }

        void abort_Click(object sender, System.EventArgs e)
        {
            //Show the abort dialog
            m_SubDialogExtender.Show();           

            CurJob curJobLoad;

            CurJob.Load(out curJobLoad, m_CurJobEntry, Server.GetCurrentServer());

            m_SubLabel.Text = string.Format("Are you sure you want to abort {0}?", curJobLoad.JobName);

            m_AbortEntry.Text = string.Format("Entry: {0}", curJobLoad.JAMSEntry);
            m_AbortSubmitAt.Text = string.Format("Submitted at: {0}", curJobLoad.ScheduledTime);
            m_AbortSchedule.Text = string.Format("Scheduled to start: {0}", curJobLoad.HoldTime);

            m_AbortForceCheckbox.Style.Clear();
            m_AbortForceCheckbox.Checked = false;

            ShowDialog(curJobLoad);

            CurrentAction = MonitorAction.Abort;
            m_AbortEntry.Visible = true;
            m_AbortSubmitAt.Visible = true;
            m_AbortSchedule.Visible = true;
            m_RescheduleBatchQueueTxt.Visible = false;
            m_ReschedulePriorityTxt.Visible = false;
            m_RescheduleDateTxt.Visible = false; 
        }

        void restart_Click(object sender, System.EventArgs e)
        {
            //Show the abort dialog
            m_SubDialogExtender.Show();

            CurJob curJobLoad;

            CurJob.Load(out curJobLoad, m_CurJobEntry, Server.GetCurrentServer());

            m_SubLabel.Text = string.Format("Are you sure you want to abort and restart {0}?", curJobLoad.JobName);

            ShowDialog(curJobLoad);

            CurrentAction = MonitorAction.Restart;
            m_RestartHoldJob.Style.Clear();
            m_RestartHoldJob.Checked = false;
            m_AbortEntry.Visible = false;
            m_AbortSubmitAt.Visible = false;
            m_AbortSchedule.Visible = false;
            m_RescheduleBatchQueueTxt.Visible = false;
            m_ReschedulePriorityTxt.Visible = false;
            m_RescheduleDateTxt.Visible = false;
        }

        void reschedule_Click(object sender, System.EventArgs e)
        {
            //Show the abort dialog
            m_SubDialogExtender.Show();

            CurJob curJobLoad;

            CurJob.Load(out curJobLoad, m_CurJobEntry, Server.GetCurrentServer());

            m_SubLabel.Text = string.Format("How do you want to reschedule {0}?", curJobLoad.JobName);

            m_RescheduleBatchQueue.Text = curJobLoad.BatchQueueName;
            
            //Set default times
            if (m_Calendar.SelectedDate < DateTime.Now || m_Calendar.SelectedDate == null)
            {
                if (curJobLoad.HoldTime < DateTime.Now)
                {
                    m_Calendar.SelectedDate = DateTime.Now.Date;
                    m_RescheduleTime.Text = TimeOfDay.Now.ToString("h:mm tt");
                }
                else
                {
                    m_Calendar.SelectedDate = curJobLoad.HoldTime.Date;
                    m_RescheduleTime.Text = TimeOfDay.Parse(curJobLoad.HoldTime.TimeOfDay.ToString()).ToString("h:mm tt");
                }
            }

            m_ReschedulePriority.Text = curJobLoad.SchedulingPriority.ToString();
            m_RescheduleBatchQueueTxt.Visible = true;
            m_RescheduleDateTxt.Visible = true;
            m_ReschedulePriorityTxt.Visible = true;
            //Add validation
            m_ValidDate.Enabled = true;
            m_ValidPri.Enabled = true;
            m_ValidTime.Enabled = true;
            m_AbortEntry.Visible = false;
            m_AbortSubmitAt.Visible = false;
            m_AbortSchedule.Visible = false;

            ShowDialog(curJobLoad);

            m_RescheduleHoldUntilManual.Checked = false;

            CurrentAction = MonitorAction.Reschedule;
            m_RescheduleBatchQueue.Style.Clear();
            m_RescheduleDate.Style.Clear();
            m_RescheduleHoldUntilManual.Style.Clear();
            m_ReschedulePriority.Style.Clear();
            m_RescheduleTime.Style.Clear();
        }

        void hold_Click(object sender, System.EventArgs e)
        {
            //Don't show a dialog, just put on manual hold

            CurJob.Hold(m_CurJobEntry);
        }

        void release_Click(object sender, System.EventArgs e)
        {
            //Show the abort dialog
            m_SubDialogExtender.Show();
            
            CurJob curJobLoad;
            
            CurJob.Load(out curJobLoad, m_CurJobEntry, Server.GetCurrentServer());

            m_SubLabel.Text = string.Format("How do you want to release {0}?", curJobLoad.JobName);

            ShowDialog(curJobLoad);

            CurrentAction = MonitorAction.Release;
            m_ReleaseFromStepWait.Style.Clear();
            m_ReleaseFromTimed.Style.Clear();
            m_ReleaseFromTimeSlot.Style.Clear();
            m_ReleaseHaltedSetup.Style.Clear();
            m_ReleaseIgnoreDepend.Style.Clear();
            m_ReleaseIgnorePrecheck.Style.Clear();
            m_ReleaseIgnoreResource.Style.Clear();
            m_ReleaseManual.Style.Clear();
            m_ReleaseToRunAgain.Style.Clear();
            m_AbortEntry.Visible = false;
            m_AbortSubmitAt.Visible = false;
            m_AbortSchedule.Visible = false;
            m_RescheduleBatchQueueTxt.Visible = false;
            m_ReschedulePriorityTxt.Visible = false;
            m_RescheduleDateTxt.Visible = false;
           
            m_ReleaseHaltedSetup.Checked = false;
            m_ReleaseHaltedSetup.Enabled = curJobLoad.Halted;
            
            m_ReleaseManual.Checked = false;
            m_ReleaseManual.Enabled = curJobLoad.Held;

            m_ReleaseToRunAgain.Checked = false;
            m_ReleaseToRunAgain.Enabled = curJobLoad.CurrentState == JobState.Completed;

            m_ReleaseFromTimed.Checked = false;
            m_ReleaseFromTimed.Enabled = curJobLoad.CurrentState == JobState.Timed;

            m_ReleaseFromStepWait.Checked = false;
            m_ReleaseFromStepWait.Enabled = curJobLoad.CurrentState == JobState.StepWait;

            m_ReleaseFromTimeSlot.Checked = false;
            m_ReleaseFromTimeSlot.Enabled = curJobLoad.CurrentState == JobState.Timed;

            m_ReleaseIgnoreDepend.Checked = false;
            m_ReleaseIgnoreDepend.Enabled = (!curJobLoad.DependOK);

            m_ReleaseIgnorePrecheck.Checked = false;
            m_ReleaseIgnorePrecheck.Enabled = curJobLoad.CurrentState == JobState.PrecheckWait;

            m_ReleaseIgnoreResource.Checked = false;
//            m_ReleaseIgnoreResource.Enabled = curJobLoad.CurrentState == JobState.ResourceWait;
            //Ignored in Win code.

            if (m_ReleaseHaltedSetup.Enabled)
            {
                m_ReleaseHaltedSetup.Checked = true;
            }
            else if (m_ReleaseManual.Enabled)
            {
                m_ReleaseManual.Checked = true;
            }
            else if (m_ReleaseToRunAgain.Enabled)
            {
                m_ReleaseToRunAgain.Checked = true;
            }
            else if (m_ReleaseFromStepWait.Enabled)
            {
                m_ReleaseFromStepWait.Checked = true;
            }
            else if (m_ReleaseFromTimed.Enabled)
            {
                m_ReleaseFromTimed.Checked = true;
            }
            else if (m_ReleaseFromTimeSlot.Enabled)
            {
                m_ReleaseFromTimeSlot.Checked = true;
            }
            else if (m_ReleaseIgnoreDepend.Enabled)
            {
                m_ReleaseIgnoreDepend.Checked = true;
            }
            else if (m_ReleaseIgnorePrecheck.Enabled)
            {
                m_ReleaseIgnorePrecheck.Checked = true;
            }
            else if (m_ReleaseIgnoreResource.Enabled)
            {
                m_ReleaseIgnoreResource.Checked = true;
            }
        }

        private static string CalcPercent(int top, int bottom)
        {
            if (bottom == 0)
            {
                return StrNa;
            }

            return string.Format("{0}%", ((top * 100) / bottom));
        }
    }
}
