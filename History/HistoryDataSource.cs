using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Web.UI;
using System.Web.UI.WebControls;
using MVPSI.JAMS;

namespace MVPSI.JAMSWeb.Controls
{
    /// <summary>
    /// Provides a datasource for the history control.
    /// </summary>
    internal class HistoryDataSource:DataSourceControl
    {
        private const string HistoryViewName = "History";
        private const string QueryInfoParameterName = "QueryInfo";
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

        private HistoryDataSourceView m_HistoryDataView;
        private HistoryDataSourceView HistoryDataView
        {
            get
            {
                if(m_HistoryDataView == null)
                {
                    m_HistoryDataView = new HistoryDataSourceView(this, HistoryViewName);
                }
                return m_HistoryDataView;
            }
        }

        /// <summary>
        /// Gets or sets the query info.
        /// </summary>
        /// <value>The query info.</value>
        public QueryInfo QueryInfo
        {
            get { return m_QueryInfo; }
            set
            {
                if (QueryInfo != value)
                {
                    m_QueryInfo = value;
                    HistoryDataView.RaiseChangedEvent();
                }
            }
        }
        private QueryInfo m_QueryInfo;

        internal Server Server
        {
            get
            {
                if (m_Server == null)
                {
                    if (m_QueryInfo.ServerName == null)
                    {
                        m_QueryInfo.ServerName = string.Empty;
                    }

                    if (m_QueryInfo.ServerName == string.Empty)
                    {
                        m_Server = JAMS.Server.CurrentServer;

                        if (m_Server == null)
                        {
                            m_Server = JAMS.Server.GetServer("localhost");
                        }
                    }
                    else
                    {
                        m_Server = JAMS.Server.GetServer(m_QueryInfo.ServerName);
                    }
                }
                return m_Server;
            }
        }

        private Server m_Server;

        /// <summary>
        /// Gets the query info.
        /// </summary>
        /// <returns></returns>
        public QueryInfo GetQueryInfo()
        {
            if (m_Parameters != null)
            {
                Parameter QueryInfoParameter = m_Parameters[QueryInfoParameterName];
                if (QueryInfoParameter != null)
                {
                    IOrderedDictionary parameterValues = m_Parameters.GetValues(Context, this);
                    return (QueryInfo)parameterValues[QueryInfoParameter.Name];
                }
            }
            return QueryInfo;
        }

        void m_Parameters_ParametersChanged(object sender, EventArgs e)
        {
            HistoryDataView.RaiseChangedEvent();
        }

        protected override DataSourceView GetView(string viewName)
        {
            if (string.IsNullOrEmpty(viewName) || (string.Compare(viewName, HistoryViewName, StringComparison.OrdinalIgnoreCase) == 0))            
            {
                return HistoryDataView;
            }
            throw new ArgumentOutOfRangeException("viewName");
        }

        protected override System.Collections.ICollection GetViewNames()
        {
            return new string[] {HistoryViewName};
        }

        /// <summary>
        /// Returns the History items.
        /// </summary>
        /// <returns></returns>
        public List<JAMS.History> GetHistory()
        {
            return HistoryDataView.GetHistory();
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
