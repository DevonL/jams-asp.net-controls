using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Web.UI;
using MVPSI.JAMS;

namespace MVPSI.JAMSWeb.Controls
{
    internal class MonitorControlDesigner:System.Web.UI.Design.ControlDesigner
    {

        /// <summary>
        /// Initializes the control designer and loads the specified component.
        /// </summary>
        /// <param name="component">The control being designed.</param>
        public override void Initialize(IComponent component)
        {
            if (component is Monitor)
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
            string output = "";

            Monitor monitor = this.Component as Monitor;

            if (monitor != null)
            {
                monitor.m_GridView.Columns.Clear();
                if (monitor.Columns.Count == 0)
                {
                    monitor.Columns.Add(new MonitorColumn {Type = MonitorColumnType.JobName});
                    monitor.Columns.Add(new MonitorColumn { Type = MonitorColumnType.JAMSEntry });
                    monitor.Columns.Add(new MonitorColumn { Type = MonitorColumnType.Description });
                }

                monitor.m_GridView.DataSource = new List<JAMS.CurJob>()
                                                    {
                                                        new JAMS.CurJob()
                                                            {
                                                                JobName = "JobEntry1",
                                                                CurrentState = JobState.Executing,
                                                                StartTime = DateTime.Now,
                                                                JAMSEntry = 123,                                                                
                                                            },
                                                        new JAMS.CurJob
                                                            {
                                                                JobName = "JobEntry2",
                                                                CurrentState = JobState.Completed,
                                                                CompletionTime = DateTime.Now,
                                                                JAMSEntry = 124
                                                            },
                                                        new JAMS.CurJob
                                                            {
                                                                JobName = "JobEntry3",
                                                                CurrentState = JobState.Held,
                                                                CompletionTime = DateTime.Now,
                                                                JAMSEntry = 125
                                                            }
                                                    };

                monitor.Refresh();

            }

            StringWriter writer = new StringWriter();
            HtmlTextWriter html = new HtmlTextWriter(writer);
            monitor.RenderBeginTag(html);
            monitor.m_GridView.RenderControl(html);
            monitor.RenderEndTag(html);

            output = writer.ToString();
          
            monitor.m_GridView.DataSource = null;

            return output;
        }
    }
}