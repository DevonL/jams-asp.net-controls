using System;
using System.Collections;
using System.Collections.Generic;
using System.Web.UI;
using MVPSI.JAMS;
using System.Linq;

namespace MVPSI.JAMSWeb.Controls
{
    /// <summary>
    /// Retrieves records from the JAMS Server and returns it to the MonitorDataSource.
    /// </summary>
    internal sealed class MonitorDataSourceView : DataSourceView
    {
        private MonitorDataSource m_Owner;

        /// <summary>
        /// Initializes a new instance of the <see cref="MonitorDataSourceView"/> class.
        /// </summary>
        /// <param name="owner">The owner.</param>
        /// <param name="viewName">Name of the view.</param>
        public MonitorDataSourceView(MonitorDataSource owner, string viewName):base(owner,viewName)
        {
            m_Owner = owner;
        }

        internal List<JAMS.CurJob> GetMonitorData()
        {
            ICollection<int> folderDescendantIDList = null;
            CurJob masterEntry;
            List<JAMS.CurJob> curJobList = new List<JAMS.CurJob>();

            //
            // If a Folder was specified get its descendant IDs
            //
            if (!String.IsNullOrWhiteSpace(m_Owner.SystemName) && m_Owner.SystemName != @"\" && m_Owner.SystemName != "*")
            {
                Folder queryFolder;

                //
                // Get descendant FolderId's
                //
                JAMS.Folder.Load(out queryFolder, m_Owner.SystemName, m_Owner.Server);
                folderDescendantIDList = queryFolder.GetFolderDescendants();
            }

            //
            // Get the collection of CurJobs matching Name and State
            //
            var newList = CurJob.Find(m_Owner.JobName, m_Owner.StateType, m_Owner.Server);

            //
            // Add matching CurJobs to the result list
            //
            bool matches;
            int checkFolderID = 0;
            foreach (JAMS.CurJob h in newList)
            {
                matches = true;

                //
                // If a folder was specified check if this CurJob should be included
                //
                if (folderDescendantIDList != null)
                {
                    // By default we check the CurJob's parent Folder
                    checkFolderID = h.ParentFolderID;

                    //
                    // If this is a subjob we need to check the master entry's Folder instead
                    //
                    if (h.MasterRON != 0 && h.MasterRON != h.RON)
                    {
                        // Try to get the master entry
                        masterEntry = newList.FirstOrDefault(c => c.RON == h.MasterRON);

                        if (masterEntry != null)
                        {
                            checkFolderID = masterEntry.ParentFolderID;
                        }
                    }

                    //
                    // The parent FolderID must be a descendant to match
                    //
                    if (!folderDescendantIDList.Contains(checkFolderID))
                    {
                        matches = false;
                    }
                }

                //
                // Add the entry if it's a match
                //
                if (matches)
                {
                    curJobList.Add(h);
                }          
            }
           
            return curJobList;
        }
        
        internal void RaiseChangedEvent()
        {
            OnDataSourceViewChanged(EventArgs.Empty);
        }

        protected override IEnumerable ExecuteSelect(DataSourceSelectArguments arguments)
        {
            arguments.RaiseUnsupportedCapabilitiesError(this);

            return GetMonitorData();
        }
    }
}