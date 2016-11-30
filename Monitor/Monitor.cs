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
using MVPSI.JAMS;

[assembly: WebResource("MVPSI.JAMSWeb.Controls.Monitor.Monitor.js", "text/javascript")]
namespace MVPSI.JAMSWeb.Controls
{
    /// <summary>
    /// JAMS Monitor Control
    /// </summary>    
    [ToolboxData("<{0}:Monitor runat=server>\n<Columns>\n<{0}:MonitorColumn Type=\"JobName\" />\n<{0}:MonitorColumn Type=\"JAMSEntry\" />\n<{0}:MonitorColumn Type=\"Description\" />\n</Columns>\n</{0}:Monitor>"),
     ParseChildren(true),
     PersistChildren(false),
     Designer(typeof(MonitorControlDesigner))]
    public class Monitor: ControlsCommon, IScriptControl
    {
        internal UpdatePanel UpdatePanel;
        private MonitorDataSource m_MonitorData;
        private Panel m_NoResultsPanel;
        private readonly MonitorDialog m_MonitorDialog;
        internal Panel m_ErrorPanel;
        internal ModalPopupExtender m_ErrorDialog;

        /// <summary>
        /// Gets or sets the name of the JAMS Server the control should connect to.
        /// </summary>
        [Category("Data"),        
         Description("Name of the Server to connect to"),
         DesignerSerializationVisibility(DesignerSerializationVisibility.Visible),
         PersistenceMode(PersistenceMode.Attribute),
         NotifyParentProperty(true), DefaultValue("")]
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
        private string m_ServerName;

        /// <summary>
        /// Gets or sets the Job Name search string that is used to select jobs to be displayed.
        /// </summary>
        [Category("Data"),
         Description("Job Name search string"),
        DefaultValue("*")]
        public string JobName { get; set; }

        /// <summary>
        /// Gets or sets the System Name search string that is used to select jobs to be displayed.
        /// </summary>
        [Category("Data"),
         Description("System search string"),
        DefaultValue("*")]
        public string SystemName { get; set; }

        /// <summary>
        /// Gets or sets the Job States to be displayed.
        /// </summary>
        [Category("Data"),
        Description("Job States to return."),
        DefaultValue("None"),
        TypeConverter(typeof(StateTypeConverter))]
        public string StateType { get; set; }
        //StateType is a string, but is limited to the correct input values in the designer by the StateTypeConverter.
        //This prevents the user from having to reference JAMSShr dll from their web project.

        /// <summary>
        /// Gets a list of MonitorColumn objects which identify which columns to display.
        /// </summary>
        [Category("Data"),
         Description("The columns which will be displayed"),
         Editor(typeof(CollectionEditor),typeof(UITypeEditor)),
         DesignerSerializationVisibility(DesignerSerializationVisibility.Visible),
         PersistenceMode(PersistenceMode.InnerProperty),
         NotifyParentProperty(true)]
        public List<MonitorColumn> Columns
        {
            get
            {
                if(m_Columns == null)
                {
                    m_Columns = new List<MonitorColumn>();
                }
                return m_Columns;
            }
        }
        private List<MonitorColumn> m_Columns;

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
            ScriptControlDescriptor descriptor = new ScriptControlDescriptor("MVPSI.JAMSWeb.Controls.Monitor", this.ClientID);
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
            reference.Name = "MVPSI.JAMSWeb.Controls.Monitor.Monitor.js";

            return new ScriptReference[] { reference };
        }
        #endregion
        #region Rendering

        private ScriptManager m_ScriptManager;
        internal Panel SubPanel;

        /// <summary>
        /// Handles the PreRender event
        /// </summary>
        /// <param name="e"></param>
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
        /// Handles the Render event.
        /// </summary>
        /// <param name="writer"></param>
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
            Columns.Clear();

            if (obj is MonitorColumn)
            {
                this.Columns.Add((MonitorColumn)obj);
            }
        }

        #endregion

        /// <summary>
        /// Initializes a new instance of the <see cref="Monitor"/> class.
        /// </summary>
        public Monitor()
        {
            m_ServerName = "localhost";
            JobName = "*";
            SystemName = "*";
            StateType = "None";
            m_NoResultsPanel = new Panel();
            m_MonitorDialog = new MonitorDialog();
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
       
//            if (!Page.IsCallback)
            {                
                m_GridView = new GridView();
                
                m_MonitorData = new MonitorDataSource();    

                m_GridView.DataSource = m_MonitorData;
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

                HtmlLink link2 = new HtmlLink();
                link2.Href = Page.ClientScript.GetWebResourceUrl(this.GetType(), "MVPSI.JAMSWeb.Controls.SubmitDialog.JAMSSubmitStyles.css");
                link2.Attributes.Add("type", "text/css");
                link2.Attributes.Add("rel", "stylesheet");
                Page.Header.Controls.Add(link2);
            }

            m_GridView.CssClass = "monitorControl";
            UpdatePanel = new UpdatePanel { ID = GetChildControlID("UpdatePanel") };
            UpdatePanel.ChildrenAsTriggers = true;
            UpdatePanel.RenderMode = UpdatePanelRenderMode.Block;

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

            m_MonitorDialog.ID = GetChildControlID("MonitorDetail");
            m_MonitorDialog.ServerName = this.ServerName;

            var dialogPopup = new ModalPopupExtender();
            dialogPopup.PopupControlID = m_MonitorDialog.ClientID;
            dialogPopup.DropShadow = false;
            dialogPopup.RepositionMode = ModalPopupRepositionMode.RepositionOnWindowResize;
            dialogPopup.TargetControlID = GetChildControlID("DummyButton2");
            dialogPopup.ID = GetChildControlID("DialogPopupExtender");
            dialogPopup.BackgroundCssClass = "modalBackground";
            m_MonitorDialog.Extender = dialogPopup;
            m_MonitorDialog.ParentMonitor = this;

            var dummyButton2 = new Button();
            dummyButton2.ID = GetChildControlID("DummyButton2");
            dummyButton2.Style["display"] = "none";
            dummyButton2.Height = 0;
            dummyButton2.Width = 0;
            dummyButton2.UseSubmitBehavior = false;

            var errorDummyButton = new Button { ID = GetChildControlID("ErrorDummyButton"), UseSubmitBehavior = false };
            errorDummyButton.Style["Display"] = "None";

            m_ErrorPanel = new Panel { ID = GetChildControlID("ErrorDialogPanel") };
            m_ErrorPanel.Controls.Add(new Label { Text = "Error" , CssClass = "tabContents"});
            m_ErrorPanel.Controls.Add(new Button { Text = "OK", ID = GetChildControlID("ErrorOk") });
            m_ErrorPanel.CssClass = "modalPopup";
            m_ErrorPanel.Style["Display"] = "None";

            m_ErrorDialog = new ModalPopupExtender { ID = GetChildControlID("ErrorDialog") };
            m_ErrorDialog.PopupControlID = m_ErrorPanel.ClientID;
            m_ErrorDialog.DropShadow = false;
            m_ErrorDialog.RepositionMode = ModalPopupRepositionMode.RepositionOnWindowResize;
            m_ErrorDialog.TargetControlID = errorDummyButton.ClientID;
            m_ErrorDialog.BackgroundCssClass = "modalBackground";
            m_ErrorDialog.CancelControlID = GetChildControlID("ErrorOk");
            m_ErrorDialog.BehaviorID = m_ErrorDialog.ID;

            Control controlContainer = UpdatePanel.ContentTemplateContainer;
            controlContainer.Controls.Add(m_GridView);
            controlContainer.Controls.Add(m_MonitorData);
            controlContainer.Controls.Add(m_MonitorDialog);
            controlContainer.Controls.Add(m_MonitorDialog.Extender);
   
            controlContainer.Controls.Add(dummyButton2);
            controlContainer.Controls.Add(loadingDialog);
            controlContainer.Controls.Add(button);
            controlContainer.Controls.Add(popup);
            controlContainer.Controls.Add(noResultsPanel);
            controlContainer.Controls.Add(m_ErrorPanel);
            controlContainer.Controls.Add(m_ErrorDialog);
            controlContainer.Controls.Add(errorDummyButton);

            SubPanel = new Panel();
            controlContainer.Controls.Add(SubPanel);
            base.OnInit(e);
        }

        /// <summary>
        /// Raises the <see cref="E:System.Web.UI.Control.Load"/> event.
        /// </summary>
        /// <param name="e">The <see cref="T:System.EventArgs"/> object that contains the event data.</param>
        protected override void OnLoad(EventArgs e)
        {
            base.Controls.Add(UpdatePanel);

            base.OnLoad(e);

            m_MonitorDialog.Load += delegate(object sender, EventArgs e_Load)
                                        {
                                            Refresh();
                                        };
        }

        /// <summary>
        /// Refreshes this instance.
        /// </summary>         
        public void Refresh()
        {
            if (!Page.IsPostBack)
            {
                //Build columns
                foreach (MonitorColumn column in Columns)
                {
                    m_GridView.Columns.Add(new BoundField()
                                               {DataField = column.Type.ToString(), HeaderText = column.Type.ToString()});
                }
                foreach (BoundField column in m_GridView.Columns)
                {
                    column.HeaderText = ConvertToFriendlyText(column.HeaderText);
                }
            }

            m_GridView.RowDataBound += new GridViewRowEventHandler(m_GridView_RowDataBound);

            m_MonitorData.ServerName = ServerName;
            m_MonitorData.JobName = JobName;
            m_MonitorData.SystemName = SystemName;
            try
            {
                m_MonitorData.StateType = (StateType) Enum.Parse(typeof (StateType), StateType);
            }
            catch(Exception ex)
            {
                m_MonitorData.StateType = JAMS.StateType.None;

                m_ErrorPanel.Controls.AddAt(1, new Label{Text = ex.Message});
                m_ErrorDialog.Show();
                //Todo: possibly change this to a different message.
            }

            try
            {
                m_GridView.DataBind();
            }
            catch (Exception ex)
            {
                m_ErrorPanel.Controls.AddAt(1, new Label { Text = ex.Message });
                m_ErrorDialog.Show();
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
            if (!DesignMode && Page.IsPostBack)
            {
                var entryNumber = Page.Request["__EVENTARGUMENT"];

                if (entryNumber != null)
                {
                    int num;
                    var isNum = int.TryParse(entryNumber, out num);

                    if (isNum)
                    {          
                        m_MonitorDialog.JAMSEntry = num;
                    }
                }

            }
        }

        void m_GridView_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            if (e.Row.RowType == DataControlRowType.DataRow && !DesignMode)
            {
                var eRowDataItem = e.Row.DataItem as JAMS.CurJob;

                e.Row.Attributes.Add("onclick", "__doPostBack('" + UpdatePanel.ClientID + "', '" + eRowDataItem.JAMSEntry + "')");
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
