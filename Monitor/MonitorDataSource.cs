using System;
using System.Collections.Generic;
using System.Web.UI;
using System.Web.UI.WebControls;
using MVPSI.JAMS;

namespace MVPSI.JAMSWeb.Controls
{
    internal class MonitorDataSource:DataSourceControl
    {
        private const string MonitorViewName = "Monitor";
        private ParameterCollection m_Parameters;

        private ParameterCollection Parameters
        {
            get
            {
                if(m_Parameters == null)
                {
                    m_Parameters = new ParameterCollection();
                    m_Parameters.ParametersChanged += new EventHandler(m_Parameters_ParametersChanged);
                    if(IsTrackingViewState)
                    {
                        ((IStateManager) m_Parameters).TrackViewState();
                    }
                }
                return m_Parameters;
            }
        }

        private MonitorDataSourceView m_MonitorDataView;
        private MonitorDataSourceView MonitorDataView
        {
            get
            {
                if(m_MonitorDataView == null)
                {
                    m_MonitorDataView = new MonitorDataSourceView(this, MonitorViewName);
                }
                return m_MonitorDataView;
            }
        }

        /// <summary>
        /// Gets or sets the name of the job.
        /// </summary>
        /// <value>The name of the job.</value>
        public string JobName
        {
            get { return m_JobName; }
            set
            {
                if (JobName != value)
                {
                    m_JobName = value;
                    MonitorDataView.RaiseChangedEvent();
                }
            }
        }
        private string m_JobName;

        /// <summary>
        /// Gets or sets the name of the system.
        /// </summary>
        /// <value>The name of the system.</value>
        public string SystemName
        {
            get { return m_SystemName; }
            set
            {
                if (SystemName != value)
                {
                    m_SystemName = value;
                    MonitorDataView.RaiseChangedEvent();
                }
            }
        }
        private string m_SystemName;

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
                        m_Server = JAMS.Server.CurrentServer;

                        if (m_Server == null)
                        {
                            m_Server = JAMS.Server.GetServer("localhost");
                        }
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
        /// Gets or sets the name of the server.
        /// </summary>
        /// <value>The name of the server.</value>
        public string ServerName
        {
            get
            {
                return m_ServerName;
            }
            set
            {
                m_ServerName = value;
            }
        }

        /// <summary>
        /// Gets or sets the current State.
        /// </summary>
        /// <value>The State.</value>
        public StateType StateType { get; set; }

        private string m_ServerName;
        private Server m_Server;


        void m_Parameters_ParametersChanged(object sender, EventArgs e)
        {
            MonitorDataView.RaiseChangedEvent();
        }

        protected override DataSourceView GetView(string viewName)
        {
            if (string.IsNullOrEmpty(viewName) || (string.Compare(viewName, MonitorViewName, StringComparison.OrdinalIgnoreCase) == 0))            
            {
                return MonitorDataView;
            }
            throw new ArgumentOutOfRangeException("viewName");
        }

        protected override System.Collections.ICollection GetViewNames()
        {
            return new string[] {MonitorViewName};
        }

        /// <summary>
        /// Gets the monitor data.
        /// </summary>
        /// <returns></returns>
        public List<JAMS.CurJob> GetMonitorData()
        {
            return MonitorDataView.GetMonitorData();
        }

        protected override void LoadViewState(object savedState)
        {
            object baseState = null;

            if(savedState != null)
            {
                Pair p = (Pair)savedState;
                baseState = p.First;

                if(p.Second != null)
                {
                    ((IStateManager)Parameters).LoadViewState(p.Second);
                }
            }

            base.LoadViewState(savedState);
        }

        protected override void OnInit(EventArgs e)
        {
            Page.LoadComplete += new EventHandler(Page_LoadComplete);
        }

        void Page_LoadComplete(object sender, EventArgs e)
        {
            if(m_Parameters != null)
            {
                m_Parameters.UpdateValues(Context, this);
            }
        }

        protected override object SaveViewState()
        {
            object baseState = base.SaveViewState();
            object parameterState = null;

            if(m_Parameters != null)
            {
                parameterState = ((IStateManager) m_Parameters).SaveViewState();
            }

            if((baseState != null) || (parameterState != null))
            {
                return new Pair(baseState, parameterState);
            }
            return null;                        
        }

        protected override void TrackViewState()
        {
            base.TrackViewState();
            if(m_Parameters != null)
            {
                ((IStateManager) m_Parameters).TrackViewState();
            }
        }
    }
}