using System;
using System.Collections;
using System.Collections.Generic;
using System.Web.UI;
using MVPSI.JAMS;

namespace MVPSI.JAMSWeb.Controls
{
    /// <summary>
    /// Retrieves the History data from the specified JAMS server and returns it to the HistoryDataSource.
    /// </summary>
    internal sealed class HistoryDataSourceView : DataSourceView
    {
        private HistoryDataSource m_Owner;        

        public HistoryDataSourceView(HistoryDataSource owner, string viewName):base(owner,viewName)
        {
            m_Owner = owner;
        }

        internal List<JAMS.History> GetHistory()
        {

            List<JAMS.History> historyList = new List<JAMS.History>();

            if (m_Owner.QueryInfo == null)
                return null;
            DateTime startDate = JAMS.Date.Evaluate(m_Owner.QueryInfo.StartDate, DateTime.Today, m_Owner.Server);
            DateTime endDate = JAMS.Date.Evaluate(m_Owner.QueryInfo.EndDate, DateTime.Today, m_Owner.Server);
            startDate = startDate + DateTime.Parse(m_Owner.QueryInfo.StartTime).TimeOfDay;
            endDate = endDate + DateTime.Parse(m_Owner.QueryInfo.EndTime).TimeOfDay;

            var listSearch = new List<HistorySelection>();
            listSearch.Add(new HistorySelection(
                HistorySelectionField.JobName, 
                ComparisonOperator.Like, 
                m_Owner.QueryInfo.JobName));
            listSearch.Add(new HistorySelection(
                HistorySelectionField.SetupName, 
                ComparisonOperator.Like, 
                m_Owner.QueryInfo.SetupName));
            listSearch.Add(new HistorySelection(
                (HistorySelectionField)Enum.Parse(typeof(HistorySelectionField), m_Owner.QueryInfo.HistorySelectionField1), 
                (ComparisonOperator)Enum.Parse(typeof(ComparisonOperator), m_Owner.QueryInfo.ComparisonOperator1), 
                m_Owner.QueryInfo.ComparisonValue1));
            listSearch.Add(new HistorySelection(
                (HistorySelectionField)Enum.Parse(typeof(HistorySelectionField), m_Owner.QueryInfo.HistorySelectionField2),
                (ComparisonOperator)Enum.Parse(typeof(ComparisonOperator), m_Owner.QueryInfo.ComparisonOperator2), 
                m_Owner.QueryInfo.ComparisonValue2));

            //set the system/folder name (default = root folder)
            //var systemFolderName = (m_Owner.QueryInfo.SystemName == "*") ? "" : m_Owner.QueryInfo.SystemName;
            listSearch.Add(new HistorySelection(
                HistorySelectionField.FolderName, 
                ComparisonOperator.MatchFolderRecursively, 
                (m_Owner.QueryInfo.SystemName == "*") ? "" : m_Owner.QueryInfo.SystemName));
                        
            //listSearch.Add(new HistorySelection(HistorySelectionField.SetupFolderName, ComparisonOperator.Like, "*"));

            ICollection<JAMS.History> newList;
            try
            {
                newList = MVPSI.JAMS.History.Find(
                listSearch,
                startDate,
                endDate,
                m_Owner.QueryInfo.IncludeSuccessful,
                m_Owner.QueryInfo.IncludeInformational,
                m_Owner.QueryInfo.IncludeWarning,
                m_Owner.QueryInfo.IncludeError,
                m_Owner.QueryInfo.IncludeFatal,
                m_Owner.QueryInfo.CheckScheduledAt,
                m_Owner.QueryInfo.CheckScheduledFor,
                m_Owner.QueryInfo.CheckStarted,
                m_Owner.QueryInfo.CheckCompleted,
                HistorySearchOptions.None,
                m_Owner.Server);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            
           
            foreach (JAMS.History h in newList)
            {
                historyList.Add(h);
            }
           
            return historyList;
        }
        
        internal void RaiseChangedEvent()
        {
            OnDataSourceViewChanged(EventArgs.Empty);
        }

        protected override IEnumerable ExecuteSelect(DataSourceSelectArguments arguments)
        {
            arguments.RaiseUnsupportedCapabilitiesError(this);

            return GetHistory();
        }
    }
}