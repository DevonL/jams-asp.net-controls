using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Web.UI;
using MVPSI.JAMS;

namespace MVPSI.JAMSWeb.Controls
{
    internal class HistoryControlDesigner:System.Web.UI.Design.ControlDesigner
    {

        /// <summary>
        /// Initializes the control designer and loads the specified component.
        /// </summary>
        /// <param name="component">The control being designed.</param>
        public override void Initialize(IComponent component)
        {
            if(component is History)
            {
                base.Initialize(component);
            }
        }

        /// <summary>
        /// Retrieves the HTML markup that is used to represent the control at design time.
        /// </summary>
        /// <returns>
        /// The HTML markup used to represent the control at design time.
        /// </returns>
        public override string GetDesignTimeHtml()
        {
            History history = this.Component as History;

            bool NoColumns = false;

            if (history != null)
            {
                history.m_GridView.Columns.Clear();
                if (history.Columns.Count == 0)
                {
                    history.Columns.Add(new HistoryColumn { Type = HistoryColumnType.JobName });
                    history.Columns.Add(new HistoryColumn { Type = HistoryColumnType.JAMSEntry });
                    history.Columns.Add(new HistoryColumn { Type = HistoryColumnType.FinalSeverity });
                    history.Columns.Add(new HistoryColumn { Type = HistoryColumnType.HoldTime });
                    history.Columns.Add(new HistoryColumn { Type = HistoryColumnType.StartTime });
                    
                    NoColumns = true;
                }
                history.m_GridView.DataSource = new List<JAMS.History>()
                                                    {
                                                        new JAMS.History
                                                            {
                                                                JobName = "HistoryEntry1",
                                                                FinalSeverity = Severity.Success,
                                                                CompletionTime = DateTime.Now,
                                                                JAMSEntry = 123,
                                                                HoldTime = DateTime.Now,
                                                                StartTime = DateTime.Now
                                                            },
                                                        new JAMS.History
                                                            {
                                                                JobName = "HistoryEntry2",
                                                                FinalSeverity = Severity.Info,
                                                                CompletionTime = DateTime.Now,
                                                                JAMSEntry = 124,
                                                                HoldTime = DateTime.Now.Subtract(new TimeSpan(1, 0, 0)),
                                                                StartTime = DateTime.Now.Subtract(new TimeSpan(1, 0, 0))
                                                            },
                                                        new JAMS.History
                                                            {
                                                                JobName = "HistoryEntry3",
                                                                FinalSeverity = Severity.Error,
                                                                CompletionTime = DateTime.Now,
                                                                JAMSEntry = 125,
                                                                HoldTime = DateTime.Now.Subtract(new TimeSpan(2, 0, 0)),
                                                                StartTime = DateTime.Now.Subtract(new TimeSpan(2, 0, 0))
                                                            }
                                                    };
                history.Refresh();
            }

            StringWriter writer = new StringWriter();
            HtmlTextWriter html = new HtmlTextWriter(writer);
            history.RenderBeginTag(html);
            history.m_GridView.RenderControl(html);
            history.RenderEndTag(html);

            string output = writer.ToString();

            if (NoColumns)
            {
                history.Columns.Clear();

            }
            history.m_GridView.DataSource = null;


            return output;
        }
    }
}