using System;
using System.Collections.Generic;
using System.IO;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using AjaxControlToolkit;
using MVPSI.JAMS;

[assembly: WebResource("MVPSI.JAMSWeb.Controls.SubmitMenu.JAMSSubmitStyles.css", "text/css")]
namespace MVPSI.JAMSWeb.Controls
{
    /// <summary>
    /// Displays details of the selected History entry.
    /// </summary>
    internal class HistoryDialog : ControlsCommon
    {
        private TabContainer m_TabContainer;
        private HtmlButton m_OkButton;
        private bool m_CurrentlyDisplayed=false;
        private const string StrNa = "N/A";

        public ModalPopupExtender Extender { get; set; }

        protected override void OnLoad(System.EventArgs e)
        {
            base.OnLoad(e);

            CssClass = "modalPopup";
            if (!m_CurrentlyDisplayed)
            {
                Style["Display"] = "None";
                Extender.Hide();
            }
            else
            {
                Extender.Show();
            }
        }

        /// <summary>
        /// Shows the dialog.
        /// </summary>
        /// <param name="history">The history.</param>
        internal void ShowDialog(JAMS.History history)
        {
            m_CurrentlyDisplayed = true;

            Style["Display"] = "inline";

            //Create the child controls.
            Panel titleBar = new Panel();
            titleBar.ID = GetChildControlID("TitleBar");
            titleBar.Style["background-color"] = "#CEE5FE";
            titleBar.Style["margin"] = "-3, -3, 6, -3";
            Extender.PopupDragHandleControlID = GetChildControlID("TitleBar");
            Extender.OkControlID = GetChildControlID("HistoryDialogOK");

            Label head = new Label();
            head.Text = "History for Job: " + history.JobName + " - ";
            titleBar.Controls.Add(head);

            head = new Label();
            head.Text = "JAMS Entry: " + history.JAMSEntry;
            titleBar.Controls.Add(head);
            Controls.AddAt(0, titleBar);

            m_TabContainer = new TabContainer { ID = GetChildControlID("HistoryTabContainer") };
            m_TabContainer.Height = Unit.Pixel(350);

            var general = new TabPanel { ID = GetChildControlID("GeneralTab") };
            var status = new TabPanel { ID = GetChildControlID("StatusTab") };
            var statistics = new TabPanel { ID = GetChildControlID("StatisticsTab") };
            var logfile = new TabPanel { ID = GetChildControlID("LogFileTab") };

            general.HeaderText = "General";
            status.HeaderText = "Status";
            statistics.HeaderText = "Statistics";
            logfile.HeaderText = "Log File";

            m_TabContainer.Tabs.Add(general);
            m_TabContainer.Tabs.Add(status);
            m_TabContainer.Tabs.Add(statistics);
            m_TabContainer.Tabs.Add(logfile);

            m_TabContainer.ActiveTabIndex = 0;
            m_TabContainer.ScrollBars = ScrollBars.Auto;

            m_OkButton = new System.Web.UI.HtmlControls.HtmlButton();
            m_OkButton.ID = GetChildControlID("HistoryDialogOK");
            m_OkButton.InnerText = "OK";
            m_OkButton.Style["margin"] = "20px, 50px, 10px, 0px";

            var buttonPanel = new Panel();
            buttonPanel.ID = GetChildControlID("History_Button_Panel");
            buttonPanel.Controls.Add(m_OkButton);
            buttonPanel.EnableViewState = false;
            buttonPanel.HorizontalAlign = HorizontalAlign.Center;

            //Create the controls on the tabs, fill in data from the History object.

            //General Tab
            var job = new Label { Text = string.Format("Job\n  {0}\n\n  {1}", history.JobName, history.Job.Description), CssClass = "tabContents" };
            var system = new Label { Text = string.Format("System\n  {0}", history.Job.ParentFolderName), CssClass = "tabContents" };
            var setup = new Label { Text = string.Format("Setup\n  {0}", history.Setup.SetupName), CssClass = "tabContents" };
            var submitInfo = new Label
                                 {
                                     Text =
                                         string.Format(
                                         "Submit Information\n  JAMS Entry {0}, Queue Entry {1}, RON {2}",
                                         history.JAMSEntry, history.BatchQueue, history.RON),
                                     CssClass = "tabContents"
                                 };

            general.Controls.Add(job);
            general.Controls.Add(system);
            general.Controls.Add(setup);
            general.Controls.Add(submitInfo);

            //Status Tab
            var finalStatus = new Label { Text = string.Format("Final Status\n  {0}", history.FinalStatus), CssClass = "tabContents" };
            var note = new Label { Text = string.Format("Note\n  {0}", history.Note), CssClass = "tabContents" };
            var jobStatusText = new Label { Text = string.Format("Job Status Text\n  {0}", history.JobStatus), CssClass = "tabContents" };

            status.Controls.Add(finalStatus);
            status.Controls.Add(note);
            status.Controls.Add(jobStatusText);

            //Stats Tab
            var times = new Label {Text = string.Format("Times\n  ")};
            var timesTable = new Table();

            for (int i = 0; i < 4; i++)
            {
                var timeRow = new TableRow();
                timesTable.Rows.Add(timeRow);                
            }        

            var cells = new List<TableCell>();
            cells.Add(new TableCell{Text = "Scheduled At: "});
            cells.Add(new TableCell { Text = history.ScheduledTime.ToString() });
            cells.Add(new TableCell{Text = "Scheduled For: "});
            cells.Add(new TableCell{Text = history.HoldTime.ToString()});
            cells.Add(new TableCell{Text = "Started: "});
            cells.Add(new TableCell{Text = history.StartTime.ToString() });
            cells.Add(new TableCell{Text = "Completed: "});
            cells.Add(new TableCell{Text = history.CompletionTime.ToString()});

            int j = 0;
            for (int i = 0; i < timesTable.Rows.Count; i++)
            {
                timesTable.Rows[i].Cells.Add(cells[j]);
                timesTable.Rows[i].Cells.Add(cells[j + 1]);
                j = j + 2;
            }

            var stats = new Label {Text = "Statistics"};
            var statsTable = new Table();

            for (int i = 0; i < 7; i++)
            {
                var statRow = new TableRow();
                statsTable.Rows.Add(statRow);
            }

            var statCells = new List<TableCell>();

            switch (history.EntryType)
            {
                case HistoryType.Job:
                    {
                        statCells.Add(new TableCell {Text = ""});
                        statCells.Add(new TableCell {Text = "Current"});
                        statCells.Add(new TableCell {Text = "Minimum"});
                        statCells.Add(new TableCell {Text = "Maximum"});
                        statCells.Add(new TableCell {Text = "Average"});
                        statCells.Add(new TableCell {Text = "% of Avg."});

                        statCells.Add(new TableCell {Text = "Elapsed Time: "});
                        statCells.Add(new TableCell {Text = history.ElapsedTime.ToString()});
                        statCells.Add(new TableCell {Text = history.Job.MinElapsedTime.ToString()});
                        statCells.Add(new TableCell {Text = history.Job.MaxElapsedTime.ToString()});
                        statCells.Add(new TableCell {Text = history.Job.AvgElapsedTime.ToString()});
                        statCells.Add(new TableCell
                                          {
                                              Text = CalcPercent(history.ElapsedTime.TotalSeconds,
                                                                 history.Job.AvgElapsedTime.TotalSeconds)
                                          });

                        statCells.Add(new TableCell {Text = "Cpu Time: "});
                        statCells.Add(new TableCell {Text = history.CpuTime.ToString()});
                        statCells.Add(new TableCell {Text = history.Job.MinCpuTime.ToString()});
                        statCells.Add(new TableCell {Text = history.Job.MaxCpuTime.ToString()});
                        statCells.Add(new TableCell {Text = history.Job.AvgCpuTime.ToString()});
                        statCells.Add(new TableCell
                        {
                            Text = CalcPercent(history.CpuTime.TotalSeconds,
                               history.Job.AvgCpuTime.TotalSeconds)
                        });
                        
                        statCells.Add(new TableCell {Text = "Direct I/O: "});
                        statCells.Add(new TableCell {Text = history.DirectIOCount.ToString()});
                        statCells.Add(new TableCell {Text = history.Job.MinDirectIOCount.ToString()});
                        statCells.Add(new TableCell {Text = history.Job.MaxDirectIOCount.ToString()});
                        statCells.Add(new TableCell {Text = history.Job.AvgDirectIOCount.ToString()});
                        statCells.Add(new TableCell {Text = CalcPercent(history.DirectIOCount, history.Job.AvgDirectIOCount)});

                        statCells.Add(new TableCell {Text = "Buffered I/O: "});
                        statCells.Add(new TableCell {Text = history.BufferedIOCount.ToString()});
                        statCells.Add(new TableCell {Text = history.Job.MinBufferedIOCount.ToString()});
                        statCells.Add(new TableCell {Text = history.Job.MaxBufferedIOCount.ToString()});
                        statCells.Add(new TableCell {Text = history.Job.AvgBufferedIOCount.ToString()});
                        statCells.Add(new TableCell {Text = CalcPercent(history.BufferedIOCount, history.Job.AvgBufferedIOCount)});

                        statCells.Add(new TableCell {Text = "Page Faults: "});
                        statCells.Add(new TableCell {Text = history.PageFaults.ToString()});
                        statCells.Add(new TableCell {Text = history.Job.MinPageFaults.ToString()});
                        statCells.Add(new TableCell {Text = history.Job.MaxPageFaults.ToString()});
                        statCells.Add(new TableCell {Text = history.Job.AvgPageFaults.ToString()});
                        statCells.Add(new TableCell {Text = CalcPercent(history.PageFaults, history.Job.AvgPageFaults)});

                        statCells.Add(new TableCell {Text = "Peak W/S: "});
                        statCells.Add(new TableCell {Text = history.WorkingSetPeak.ToString()});
                        statCells.Add(new TableCell {Text = history.Job.MinWorkingSetPeak.ToString()});
                        statCells.Add(new TableCell {Text = history.Job.MaxWorkingSetPeak.ToString()});
                        statCells.Add(new TableCell {Text = history.Job.AvgWorkingSetPeak.ToString()});
                        statCells.Add(new TableCell {Text = CalcPercent(history.WorkingSetPeak, history.Job.AvgWorkingSetPeak)});

                        break;
                    }
                case HistoryType.Setup:
                    {
                        statCells.Add(new TableCell { Text = "" });
                        statCells.Add(new TableCell { Text = "Current" });
                        statCells.Add(new TableCell { Text = "Minimum" });
                        statCells.Add(new TableCell { Text = "Maximum" });
                        statCells.Add(new TableCell { Text = "Average" });
                        statCells.Add(new TableCell { Text = "% of Avg." });

                        statCells.Add(new TableCell { Text = "Elapsed Time: " });
                        statCells.Add(new TableCell { Text = history.ElapsedTime.ToString() });
                        statCells.Add(new TableCell { Text = history.Setup.MinElapsedTime.ToString() });
                        statCells.Add(new TableCell { Text = history.Setup.MaxElapsedTime.ToString() });
                        statCells.Add(new TableCell { Text = history.Setup.AvgElapsedTime.ToString() });
                        statCells.Add(new TableCell
                        {
                            Text = CalcPercent(history.ElapsedTime.TotalSeconds,
                                               history.Setup.AvgElapsedTime.TotalSeconds)
                        });

                        statCells.Add(new TableCell { Text = "Cpu Time: " });
                        statCells.Add(new TableCell { Text = StrNa });
                        statCells.Add(new TableCell { Text = StrNa });
                        statCells.Add(new TableCell { Text = StrNa });
                        statCells.Add(new TableCell { Text = StrNa });
                        statCells.Add(new TableCell { Text = StrNa });

                        statCells.Add(new TableCell { Text = "Direct I/O: " });
                        statCells.Add(new TableCell { Text = StrNa });
                        statCells.Add(new TableCell { Text = StrNa });
                        statCells.Add(new TableCell { Text = StrNa });
                        statCells.Add(new TableCell { Text = StrNa });
                        statCells.Add(new TableCell { Text = StrNa });

                        statCells.Add(new TableCell { Text = "Buffered I/O: " });
                        statCells.Add(new TableCell { Text = StrNa });
                        statCells.Add(new TableCell { Text = StrNa });
                        statCells.Add(new TableCell { Text = StrNa });
                        statCells.Add(new TableCell { Text = StrNa });
                        statCells.Add(new TableCell { Text = StrNa });

                        statCells.Add(new TableCell { Text = "Page Faults: " });
                        statCells.Add(new TableCell { Text = StrNa });
                        statCells.Add(new TableCell { Text = StrNa });
                        statCells.Add(new TableCell { Text = StrNa });
                        statCells.Add(new TableCell { Text = StrNa });
                        statCells.Add(new TableCell { Text = StrNa });

                        statCells.Add(new TableCell { Text = "Peak W/S: " });
                        statCells.Add(new TableCell { Text = StrNa });
                        statCells.Add(new TableCell { Text = StrNa });
                        statCells.Add(new TableCell { Text = StrNa });
                        statCells.Add(new TableCell { Text = StrNa });
                        statCells.Add(new TableCell { Text = StrNa });

                        break;
                    }
                default:
                    {
                        statCells.Add(new TableCell { Text = "" });
                        statCells.Add(new TableCell { Text = "Current" });
                        statCells.Add(new TableCell { Text = "Minimum" });
                        statCells.Add(new TableCell { Text = "Maximum" });
                        statCells.Add(new TableCell { Text = "Average" });
                        statCells.Add(new TableCell { Text = "% of Avg." });

                        statCells.Add(new TableCell { Text = "Elapsed Time: " });
                        statCells.Add(new TableCell { Text = history.ElapsedTime.ToString() });
                        statCells.Add(new TableCell { Text = StrNa });
                        statCells.Add(new TableCell { Text = StrNa });
                        statCells.Add(new TableCell { Text = StrNa });
                        statCells.Add(new TableCell { Text = StrNa });

                        statCells.Add(new TableCell { Text = "Cpu Time: " });
                        statCells.Add(new TableCell { Text = history.CpuTime.ToString() });
                        statCells.Add(new TableCell { Text = StrNa });
                        statCells.Add(new TableCell { Text = StrNa });
                        statCells.Add(new TableCell { Text = StrNa });
                        statCells.Add(new TableCell { Text = StrNa });

                        statCells.Add(new TableCell { Text = "Direct I/O: " });
                        statCells.Add(new TableCell { Text = history.DirectIOCount.ToString() });
                        statCells.Add(new TableCell { Text = StrNa });
                        statCells.Add(new TableCell { Text = StrNa });
                        statCells.Add(new TableCell { Text = StrNa });
                        statCells.Add(new TableCell { Text = StrNa });

                        statCells.Add(new TableCell { Text = "Buffered I/O: " });
                        statCells.Add(new TableCell { Text = history.BufferedIOCount.ToString() });
                        statCells.Add(new TableCell { Text = StrNa });
                        statCells.Add(new TableCell { Text = StrNa });
                        statCells.Add(new TableCell { Text = StrNa });
                        statCells.Add(new TableCell { Text = StrNa });

                        statCells.Add(new TableCell { Text = "Page Faults: " });
                        statCells.Add(new TableCell { Text = history.PageFaults.ToString() });
                        statCells.Add(new TableCell { Text = StrNa });
                        statCells.Add(new TableCell { Text = StrNa });
                        statCells.Add(new TableCell { Text = StrNa });
                        statCells.Add(new TableCell { Text = StrNa });

                        statCells.Add(new TableCell { Text = "Peak W/S: " });
                        statCells.Add(new TableCell { Text = history.WorkingSetPeak.ToString() });
                        statCells.Add(new TableCell { Text = StrNa });
                        statCells.Add(new TableCell { Text = StrNa });
                        statCells.Add(new TableCell { Text = StrNa });
                        statCells.Add(new TableCell { Text = StrNa });

                        break;
                    }
            }

            j = 0;
            for (int i = 0; i < statsTable.Rows.Count; i++)
            {                
                statsTable.Rows[i].Cells.Add(statCells[j]);
                statsTable.Rows[i].Cells.Add(statCells[j+1]);
                statsTable.Rows[i].Cells.Add(statCells[j+2]);
                statsTable.Rows[i].Cells.Add(statCells[j+3]);
                statsTable.Rows[i].Cells.Add(statCells[j+4]);
                statsTable.Rows[i].Cells.Add(statCells[j+5]);
                j = j + 6;
            }

            statistics.Controls.Add(times);
            statistics.Controls.Add(timesTable);
            statistics.Controls.Add(stats);
            statistics.Controls.Add(statsTable);

            //LogFile tab
            var logLocation = new Label { Text = history.LogFilename, CssClass = "tabContents" };

            var logLabel = new Label {CssClass = "tabContents"};

            Stream log = null;
            try
            {
                // try getting the log file. If we can't, just say the log is unavailable.
                log = history.LogFile;
            }
            catch (Exception ex)
            {
                string exMessage = "The log file is not available";
                while (ex != null)
                {
                    exMessage = string.Format("{0}{1}{2}", exMessage, Environment.NewLine, ex.Message);
                    ex = ex.InnerException;
                }
                logLabel.Text = exMessage;
            }
            if (log != null)
            {
                StreamReader logSR = new StreamReader(log, true);
                logLabel.Text = logSR.ReadToEnd();
                log.Close();
            }

            var preTag = new System.Web.UI.HtmlControls.HtmlGenericControl();
            preTag.TagName = "pre";
            preTag.Controls.Add(logLabel);

            logfile.Controls.Add(logLocation);
            logfile.Controls.Add(preTag);
            
            //Once all the controls have been created, add them to the dialog.
            Controls.Add(m_TabContainer);
            Controls.Add(buttonPanel);
        }

        private string CalcPercent(int top, int bottom)
        {
            if (bottom == 0)
            {
                return StrNa;
            }

            return string.Format("{0}%", ((top*100)/bottom));
        }

        private string CalcPercent(long top, long bottom)
        {
            if (bottom == 0)
            {
                return StrNa;
            }

            return string.Format("{0}%", ((top * 100) / bottom));
        }
    }
}
