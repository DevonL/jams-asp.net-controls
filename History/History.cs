using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.Drawing.Design;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using AjaxControlToolkit;

[assembly: WebResource("MVPSI.JAMSWeb.Controls.History.History.js", "text/javascript")]
[assembly: WebResource("MVPSI.JAMSWeb.Controls.Resources.JAMSLoadingStyles.css", "text/css")]
[assembly: WebResource("MVPSI.JAMSWeb.Controls.Resources.LoadIcon.gif", "image/gif")]
namespace MVPSI.JAMSWeb.Controls
{
    /// <summary>
    /// JAMS History Control
    /// </summary>    
    [ToolboxData("<{0}:History runat=server>\n<Columns>\n<{0}:HistoryColumn Type=\"JobName\" />\n<{0}:HistoryColumn Type=\"JAMSEntry\" />\n<{0}:HistoryColumn Type=\"FinalSeverity\" />\n<{0}:HistoryColumn Type=\"HoldTime\" />\n<{0}:HistoryColumn Type=\"StartTime\" />\n</Columns>\n</{0}:History>"),
    ParseChildren(true),
    PersistChildren(false),
    Designer(typeof(HistoryControlDesigner))]
    public class History: ControlsCommon, IScriptControl
    {
        private UpdatePanel m_UpdatePanel;
        private HistoryDataSource m_HistoryData;
        private Panel m_NoResultsPanel;
        private readonly HistoryDialog m_HistoryDialog;

        /// <summary>
        /// Gets or sets the Query Info, which determines the Jobs displayed in the History control.
        /// </summary>
        /// <value>The query info.</value>
        [Category("Data"),        
         Description("Query Parameters"),
        TypeConverterAttribute(typeof(QueryInfoConverter)),
        DesignerSerializationVisibility(DesignerSerializationVisibility.Visible),
        PersistenceMode(PersistenceMode.InnerProperty),
        NotifyParentProperty(true)]
        public QueryInfo QueryInfo
        {
            get
            {
                return m_QueryInfo;
            }
            set
            {
                m_QueryInfo = value;
            }
        }
        private QueryInfo m_QueryInfo;

        /// <summary>
        /// Gets the columns displayed in the History control.
        /// </summary>
        /// <value>The columns.</value>
        [Category("Data"),
         Description("The columns which will be displayed"),
        Editor(typeof(CollectionEditor),typeof(UITypeEditor)),
        DesignerSerializationVisibility(DesignerSerializationVisibility.Visible),
        PersistenceMode(PersistenceMode.InnerProperty),
        NotifyParentProperty(true)]
        public List<HistoryColumn> Columns
        {
            get
            {
                if(m_Columns == null)
                {
                    m_Columns = new List<HistoryColumn>();
                }
                return m_Columns;
            }
        }
        private List<HistoryColumn> m_Columns;

        internal GridView m_GridView;

        #region IScriptControl

        /// <summary>
        /// Gets a collection of script descriptors that represent ECMAScript (JavaScript) client components.
        /// </summary>
        /// <returns>
        /// An <see cref="T:System.Collections.IEnumerable"/> collection of <see cref="T:System.Web.UI.ScriptDescriptor"/> objects.
        /// </returns>
        public IEnumerable<ScriptDescriptor> GetScriptDescriptors()
        {
            ScriptControlDescriptor descriptor = new ScriptControlDescriptor("MVPSI.JAMSWeb.Controls.History", this.ClientID);
            //descriptor.AddProperty("doNothingProp", "val");
            return new ScriptDescriptor[] {descriptor};
        }

        /// <summary>
        /// Gets a collection of <see cref="T:System.Web.UI.ScriptReference"/> objects that define script resources that the control requires.
        /// </summary>
        /// <returns>
        /// An <see cref="T:System.Collections.IEnumerable"/> collection of <see cref="T:System.Web.UI.ScriptReference"/> objects.
        /// </returns>
        public IEnumerable<ScriptReference> GetScriptReferences()
        {
            ScriptReference reference = new ScriptReference();
            reference.Assembly = "JAMSWebControls";
            reference.Name = "MVPSI.JAMSWeb.Controls.History.History.js";

            return new ScriptReference[] { reference };
        }
        #endregion
        #region Rendering

        private ScriptManager m_ScriptManager;
        /// <summary>
        /// Raises the <see cref="E:System.Web.UI.Control.PreRender"/> event.
        /// </summary>
        /// <param name="e">An <see cref="T:System.EventArgs"/> object that contains the event data.</param>
        protected override void OnPreRender(EventArgs e)
        {
            if (!this.DesignMode)
            {
                m_ScriptManager = ToolkitScriptManager.GetCurrent(Page);

                if (m_ScriptManager == null)
                {
                    throw new HttpException("A ToolkitScriptManager control must exist on the current page.");
                }

                m_ScriptManager.RegisterScriptControl(this);

            }

            base.OnPreRender(e);
        }

        /// <summary>
        /// Renders the control to the specified HTML writer.
        /// </summary>
        /// <param name="writer">The <see cref="T:System.Web.UI.HtmlTextWriter"/> object that receives the control content.</param>
        protected override void Render(HtmlTextWriter writer)
        {
            if (!this.DesignMode)
                m_ScriptManager.RegisterScriptDescriptors(this);

         
           base.Render(writer);    
            
        }
        #endregion
        #region Implement AddParsedSubObject

        /// <summary>
        /// Notifies the server control that an element, either XML or HTML, was parsed, and adds the element to the server control's <see cref="T:System.Web.UI.ControlCollection"/> object.
        /// </summary>
        /// <param name="obj">An <see cref="T:System.Object"/> that represents the parsed element.</param>
        protected override void AddParsedSubObject(object obj)
        {
            if (obj is HistoryColumn)
            {
                this.Columns.Add((HistoryColumn)obj);
            }
        }

        #endregion

        /// <summary>
        /// Initializes a new instance of the <see cref="History"/> class.
        /// </summary>
        public History()
        {            
            m_QueryInfo = new QueryInfo();
            m_NoResultsPanel = new Panel();
            m_HistoryDialog = new HistoryDialog();
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
                                     ToolkitScriptManager sMgr = new ToolkitScriptManager{EnablePageMethods = true};
                                     if (Page.Form == null)
                                     {
                                         throw new HttpException("Control must be located inside of a <form> tag.");
                                     }

                                     Page.Form.Controls.AddAt(0, sMgr);
                                 }
                             };

//            if (!Page.IsCallback)
            {                
                m_GridView = new GridView();
                
                m_HistoryData = new HistoryDataSource();

                m_HistoryData.QueryInfo = QueryInfo;
                m_GridView.DataSource = m_HistoryData;
                m_GridView.AutoGenerateColumns = false;
            }

            if (!DesignMode)
            {
                //
                //  Add our stylesheet
                //
                HtmlLink link = new HtmlLink();
                link.Href = Page.ClientScript.GetWebResourceUrl(this.GetType(),
                                                                "MVPSI.JAMSWeb.Controls.Resources.JAMSLoadingStyles.css");
                link.Attributes.Add("type", "text/css");
                link.Attributes.Add("rel", "stylesheet");
                Page.Header.Controls.Add(link);

                //
                //  Add our stylesheet
                //
                HtmlLink link2 = new HtmlLink();
                link2.Href = Page.ClientScript.GetWebResourceUrl(this.GetType(), "MVPSI.JAMSWeb.Controls.SubmitDialog.JAMSSubmitStyles.css");
                link2.Attributes.Add("type", "text/css");
                link2.Attributes.Add("rel", "stylesheet");
                Page.Header.Controls.Add(link2);
            }

            base.OnInit(e);       
        }

        /// <summary>
        /// Raises the <see cref="E:System.Web.UI.Control.Load"/> event.
        /// </summary>
        /// <param name="e">The <see cref="T:System.EventArgs"/> object that contains the event data.</param>
        protected override void OnLoad(EventArgs e)
        {
            m_GridView.CssClass = "historyControl";
            m_UpdatePanel = new UpdatePanel {ID = GetChildControlID("UpdatePanel"), UpdateMode = UpdatePanelUpdateMode.Always};
            m_UpdatePanel.ChildrenAsTriggers = true;
            m_UpdatePanel.RenderMode = UpdatePanelRenderMode.Block;

            //Setup loading dialog
            var popup = new ModalPopupExtender();
            popup.PopupControlID = GetChildControlID("LoadingDialog");
            popup.DropShadow = false;
            popup.RepositionMode = ModalPopupRepositionMode.RepositionOnWindowResize;
            popup.TargetControlID = GetChildControlID("DummyButton");
            popup.ID = GetChildControlID("PopupExtender");
            popup.BehaviorID = GetChildControlID("PopupExtender");                       

            Button button = new Button();
            button.ID = GetChildControlID("DummyButton");
            button.Style["display"] = "none";
            button.Height = 0;
            button.Width = 0;
            button.UseSubmitBehavior = false;

            Panel loadingDialog = new Panel();
            Label loadingLabel = new Label();
            Image img = new Image();
            loadingLabel.Style["width"] = "auto";
            loadingLabel.Style["margin"] = "4%";
            loadingLabel.Text = "Loading...";

            img.ImageUrl = Page.ClientScript.GetWebResourceUrl(this.GetType(), "MVPSI.JAMSWeb.Controls.Resources.LoadIcon.gif");
            loadingDialog.Controls.Add(img);
            loadingDialog.Controls.Add(loadingLabel);
            
            loadingDialog.ID = GetChildControlID("LoadingDialog");
            loadingDialog.Style["height"] = "50px";
            loadingDialog.HorizontalAlign = HorizontalAlign.Left;
            loadingDialog.CssClass = "loadingPopup";

            Panel noResultsPanel = new Panel();
            Label noResultsLabel = new Label();
            noResultsLabel.Text = "No results found";
            noResultsLabel.Style["width"] = "auto";
            noResultsLabel.Style["margin"] = "4%";            

            noResultsPanel.ID = GetChildControlID("NoResultsPanel");
            noResultsPanel.Controls.Add(noResultsLabel);
            m_NoResultsPanel = noResultsPanel;

            m_HistoryDialog.ID = GetChildControlID("HistoryDetail");

            var dialogPopup = new ModalPopupExtender();
            dialogPopup.PopupControlID = m_HistoryDialog.ClientID;
            dialogPopup.DropShadow = false;
            dialogPopup.RepositionMode = ModalPopupRepositionMode.RepositionOnWindowResize;
            dialogPopup.TargetControlID = GetChildControlID("DummyButton2");
            dialogPopup.ID = GetChildControlID("DialogPopupExtender");
            dialogPopup.BehaviorID = GetChildControlID("DialogPopupExtender");
            dialogPopup.BackgroundCssClass = "modalBackground";
            m_HistoryDialog.Extender = dialogPopup;

            var dummyButton2 = new Button();
            dummyButton2.ID = GetChildControlID("DummyButton2");
            dummyButton2.Style["display"] = "none";
            dummyButton2.Height = 0;
            dummyButton2.Width = 0;
            dummyButton2.UseSubmitBehavior = false;

            Control controlContainer = m_UpdatePanel.ContentTemplateContainer;
            controlContainer.Controls.Add(m_GridView);
            controlContainer.Controls.Add(m_HistoryData);
            controlContainer.Controls.Add(m_HistoryDialog);
            controlContainer.Controls.Add(m_HistoryDialog.Extender);
            controlContainer.Controls.Add(dummyButton2);
            controlContainer.Controls.Add(loadingDialog);
            controlContainer.Controls.Add(button);
            controlContainer.Controls.Add(popup);
            controlContainer.Controls.Add(noResultsPanel);

            base.Controls.Add(m_UpdatePanel);
            
            Refresh();
        }

        /// <summary>
        /// Refreshes this instance.
        /// </summary>         
        public void Refresh()
        {
            if (!Page.IsPostBack)
            {

                //Build columns
                foreach (HistoryColumn column in Columns)
                {
                    m_GridView.Columns.Add(new BoundField {DataField = column.Type.ToString(), HeaderText = column.Type.ToString()});
                }

                foreach (var column in m_GridView.Columns)
                {
                    var bField = column as BoundField;
                    if(bField != null)
                        bField.HeaderText = ConvertToFriendlyText(bField.HeaderText);                    
                }
            }

            if (QueryInfo.ClearCurrentList)
            {
                //Add handler so we can setup click events on each row
                m_GridView.RowDataBound += m_GridView_RowDataBound;
                m_GridView.DataBind();
            }

            if (m_GridView.Rows.Count == 0)
            {
                m_NoResultsPanel.Visible = true;

                if (m_GridView.Columns.Count == 0)
                {
                    var label = m_NoResultsPanel.Controls[0] as Label;
                    label.Text = "No columns are currently selected";

                }
            }
            else
            {
                m_NoResultsPanel.Visible = false;
            }

            CheckDetailDialog();

        }

        private void CheckDetailDialog()
        {
            //Check if a detail row was clicked
            if (!DesignMode && Page.IsPostBack)
            {
                var RRONNumber = Page.Request["__EVENTARGUMENT"];

                if (RRONNumber != null)
                {
                    long num;
                    var isNum = long.TryParse(RRONNumber, out num);

                    if (isNum)
                    {


                        List<JAMS.History> hisotryList = m_HistoryData.GetHistory();

                        if (hisotryList.Exists(h => h.RRON == num))
                        {
                            var detailHistory = hisotryList.Find(h => h.RRON == num);

//                            if (null != detailHistory)
                            {
                                //Pass History to the dialog and show the dialog
                                m_HistoryDialog.ShowDialog(detailHistory);
                            }
                        }
                    }
                }
            }
        }

        void m_GridView_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            if (e.Row.RowType == DataControlRowType.DataRow && !DesignMode)
            {
                var eRowDataItem = e.Row.DataItem as JAMS.History;

                //When a user clicks on a row, a partial postback will occur.  The RRON number will then be passed back so that the detail
                //dialog can be displayed.

                e.Row.Attributes.Add("onclick", "__doPostBack('" + m_UpdatePanel.ClientID + "', '" + eRowDataItem.RRON + "')");
                e.Row.Attributes.Add("onmouseover", "this.style.color='red';");
                e.Row.Attributes.Add("onmouseout", "this.style.color='black';");
            }
        }

        private static string ConvertToFriendlyText(string text)
        {
            StringBuilder sb = new StringBuilder(text);
            int insertAdjustment = 0;
            
            for(int i = 0; i < text.Length; i++)
            {
                char ch = text[i];
                if(char.IsUpper(ch) && i > 0)
                {
                    if (!char.IsUpper(text[i - 1]))
                    {
                        sb.Insert(i + insertAdjustment, ' ');
                        insertAdjustment++;
                    }
                    else if (text.Length > i+1)
                    {
                        if (!char.IsUpper(text[i+1]))
                        {
                            sb.Insert(i + insertAdjustment, ' ');
                            insertAdjustment++;
                        }
                    }
                    
                }
            }
            
            return sb.ToString();
        }


    }
}
