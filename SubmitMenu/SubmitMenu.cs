using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.HtmlControls;
using AjaxControlToolkit;
using MVPSI.JAMS;

[assembly: WebResource("MVPSI.JAMSWeb.Controls.Resources.System.png", "image/png")]
[assembly: WebResource("MVPSI.JAMSWeb.Controls.Resources.Job.png", "image/png")]
[assembly: WebResource("MVPSI.JAMSWeb.Controls.Resources.Setup.png", "image/png")]
[assembly: WebResource("MVPSI.JAMSWeb.Controls.SubmitDialog.JAMSSubmitStyles.css", "text/css")]
namespace MVPSI.JAMSWeb.Controls
{
    /// <summary>
    /// The JAMS Submit Menu Control can be placed on a web page to display a menu of
    /// jobs that can be submitted.
    /// </summary>
    [Designer("MVPSI.JAMSWeb.Controls.SubmitMenuDesigner"), 
    DefaultProperty("ServerName"),
    ToolboxData("<{0}:SubmitMenu runat=server></{0}:SubmitMenu>")]
    public class SubmitMenu : ControlsCommon
    {

        private UpdatePanel m_UpdatePanel;
        private SubmitTree m_SubmitTree;
        private SubmitDialog m_SubmitDialog;
        private Server m_Server;
        private Label m_ResultLabel;
        private Panel m_ResultDialog;

        /// <summary>
        /// Creates a SubmitMenu control.
        /// </summary>
        public SubmitMenu()
        {
            m_MenuType = MenuType.Folder;
            m_MenuName = string.Empty;
        }

        /// <summary>
        /// Gets or sets the name of the JAMS Server to connect to.
        /// </summary>
        [Category("Data"), 
        DefaultValue(""), 
        Description("Name of the JAMS server to connect to")]
        public string ServerName
        {
            get
            {
                return (m_ServerName == null) ? String.Empty : m_ServerName;
            }
            set
            {
                if (m_ServerName != value)
                {
                    m_ServerName = value;
                    m_Server = null;
                }
            }          
        }
        private string m_ServerName;

        /// <summary>
        /// Gets or sets the type of menu to create.  A JAMS Submit menu can be created from
        /// a specific JAMS Menu definition, a specific JAMS System or all JAMS Systems.
        /// </summary>
        [Category("Data"), 
        DefaultValue(MVPSI.JAMS.MenuType.Folder), 
        Description("The type of the root menu.")]
        public MVPSI.JAMS.MenuType MenuType
        {
            get
            {
                return m_MenuType;
            }
            set
            {
                m_MenuType = value;
            }
        }
        private MVPSI.JAMS.MenuType m_MenuType;

        /// <summary>
        /// Gets or sets the name of the menu or the SystemName (depending on the MenuType).  When the MenuType
        /// is Menu, this name is required.  When the MenyType is System, is this name is empty the generated
        /// submit menu will contain all JAMS Systems which expand into their Jobs and Setups.
        /// </summary>
        [Category("Data"), 
        DefaultValue(""), 
        Description("Name of a JAMS menu or JAMS SystemName.  An Empty string will produce the default system menu.")]
        public string MenuName
        {
            get
            {
                return (m_MenuName == null) ? String.Empty : m_MenuName;
            }
            set
            {
                m_MenuName = value;
            }
        }
        private string m_MenuName;
        private ModalPopupExtender m_ResultPopup;

        internal Server Server
        {
            get
            {
                if(m_Server == null)
                {
                    if(ServerName == null)
                    {
                        ServerName = string.Empty;
                    }

                    if(ServerName == string.Empty)
                    {
                        m_Server = JAMS.Server.GetCurrentServer();
                    }
                    else
                    {
                        m_Server = JAMS.Server.GetServer(ServerName);
                    }
                }
                return m_Server;
            }
        }

        internal SubmitDialog SubmitDialog
        {
            get
            {
                return m_SubmitDialog;
            }
        }

        /// <summary>
        /// Handles page initialization.
        /// </summary>
        /// <param name="e"></param>
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

            base.OnInit(e);

            if (string.IsNullOrEmpty(ServerName))
            {
                ServerName = "localhost";
            }

            //
            //  Add our stylesheet
            //
            HtmlLink link = new HtmlLink();
            link.Href = Page.ClientScript.GetWebResourceUrl(this.GetType(), "MVPSI.JAMSWeb.Controls.SubmitDialog.JAMSSubmitStyles.css");
            link.Attributes.Add("type", "text/css");
            link.Attributes.Add("rel", "stylesheet");
            Page.Header.Controls.Add(link);
        }

        /// <summary>
        /// Handles the Load event
        /// </summary>
        /// <param name="e"></param>
        protected override void OnLoad(EventArgs e)
        {
            //Server.Disconnect();
            //Server.Connect();

            m_UpdatePanel = new UpdatePanel
            {
                ID = GetChildControlID("UpdatePanel")
            };
            m_UpdatePanel.ChildrenAsTriggers = true;
            m_UpdatePanel.RenderMode = UpdatePanelRenderMode.Block;

            m_SubmitTree = new SubmitTree
            {
                ID = GetChildControlID("SubmitTree")
            };
            m_SubmitTree.SubmitMenu = this;

            m_SubmitDialog = new SubmitDialog
            {
                ID = GetChildControlID("SubmitDialog"),
                DisplayModal = true,
                ServerName = m_ServerName
            };

            //
            //  Create the dialog that is displayed in the m_ResultPopup when
            //  a job has been submitted
            //
            m_ResultDialog = new Panel
            {
                ID = GetChildControlID("Submit_ResultDialog")
            };

            Panel labelPanel = new Panel();
            m_ResultLabel = new Label();
            m_ResultLabel.Style["width"] = "auto";
            m_ResultLabel.Style["margin"] = "4%";
            labelPanel.Controls.Add(m_ResultLabel);
            labelPanel.Style["background-color"] = "#FFFFFF";
            labelPanel.Style["height"] = "50px";
            labelPanel.HorizontalAlign = HorizontalAlign.Left;
            m_ResultDialog.Controls.Add(labelPanel);

            Button okButton = new Button
            {
                Text = "OK",
                ID = GetChildControlID("Submit_ResultOkButton")
            };
            okButton.Style["margin"] = "2%";
            m_ResultDialog.HorizontalAlign = HorizontalAlign.Right;
            m_ResultDialog.Controls.Add(okButton);
            m_ResultDialog.Style["display"] = "None";
            m_ResultDialog.Style["Height"] = "100px";
            m_ResultDialog.CssClass = "modalPopup";

            m_ResultPopup = new ModalPopupExtender
            {
                ID = GetChildControlID("ResultExtender")
            };
            m_ResultPopup.PopupControlID = m_ResultDialog.ID;
            m_ResultPopup.TargetControlID = GetChildControlID("ResultDummyButton");
            m_ResultPopup.OkControlID = okButton.ID;
            m_ResultPopup.BackgroundCssClass = "modalBackground";
            m_ResultPopup.Controls.Add(m_ResultDialog);

            Button button2 = new Button();
            button2.ID = GetChildControlID("ResultDummyButton");
            button2.Style["display"] = "none";
            button2.Height = 0;
            button2.Width = 0;

            Control controlContainer = m_UpdatePanel.ContentTemplateContainer;
            controlContainer.Controls.Add(this.m_SubmitTree);
            controlContainer.Controls.Add(this.m_SubmitDialog);
            controlContainer.Controls.Add(m_ResultPopup);
            controlContainer.Controls.Add(m_ResultDialog);
            controlContainer.Controls.Add(button2);

            base.Controls.Add(m_UpdatePanel);
        }

        /// <summary>
        /// Handle the PreRender event
        /// </summary>
        /// <param name="e"></param>
        protected override void OnPreRender(EventArgs e)
        {
            //Check to see if we have a result value.  If so, then show it in the confirm dialog box.
            if (!string.IsNullOrEmpty(m_SubmitDialog.ResultText))
            {
                m_ResultLabel.Text = m_SubmitDialog.ResultText;
                m_SubmitDialog.ResultText = string.Empty;
                m_SubmitDialog.ClearViewState();
                m_ResultDialog.Style["Display"] = "Inline";
                m_ResultPopup.Show();
            }
            else
            {
                m_ResultDialog.Style["Display"] = "None";
                m_ResultPopup.Hide();
            }
            
            base.OnPreRender(e);
        }
    }

    /// <summary>
    /// SubmitTree which is contained within the SubmitMenu
    /// </summary>
    internal class SubmitTree : TreeView
    {       

        private SubmitMenu submitMenu;

        internal Server GetServer()
        {
            return submitMenu.Server;
        }

        /// <summary>
        /// Gets or sets the submit menu.
        /// </summary>
        /// <value>The submit menu.</value>
        public SubmitMenu SubmitMenu
        {
            get { return submitMenu; }
            set { submitMenu = value;}
        }

        protected override void OnInit(EventArgs e)
        {
            base.OnInit(e);

            if (!Page.IsCallback)
            {
                //
                //  Build the base level nodes
                //
                ExpandDepth = 0;
                PopulateMenuNode(Nodes, SubmitMenu.MenuType, SubmitMenu.MenuName);
            }

            //
            //  Enable AJAX style client side node population
            //
            //  Setting EnableClientScript to true seems like a great idea because then the tree
            //  will keep all of the expanded menu items cached and it will expand and collapse them
            //  all by itself.  But, when we did that, the tree was broken after you submitted a
            //  job.  After submitting a job, the next time the tree had to populate a branch, OnLoad
            //  would fail with an index out of range exception.
            //
            this.EnableClientScript = false;
            this.PopulateNodesFromClient = true;
            
            this.SelectedNodeChanged += new EventHandler(SubmitTree_SelectedNodeChanged);
            this.TreeNodePopulate += new TreeNodeEventHandler(SubmitTree_TreeNodePopulate);
        }

        /// <summary>
        /// Called when the user expands a menu that hasn't been expanded before.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SubmitTree_TreeNodePopulate(object sender, TreeNodeEventArgs e)
        {
            string[] nodeData = e.Node.Value.Split(',');
            MenuType menuType = (MenuType)Int32.Parse(nodeData[0]);
            int menuID = Int32.Parse(nodeData[1]);
            string menuName = nodeData[2];

            PopulateMenuNode(e.Node.ChildNodes, menuType, menuName, menuID);
        }

        private void PopulateMenuNode(TreeNodeCollection nodes, MenuType menuType, string menuName, int menuID = 0)
        {
            //
            //  Get the Image URLs
            //
            string systemImageUrl = Page.ClientScript.GetWebResourceUrl(
                this.GetType(),
                "MVPSI.JAMSWeb.Controls.Resources.System.png");
            string menuImageUrl = Page.ClientScript.GetWebResourceUrl(
                this.GetType(),
                "MVPSI.JAMSWeb.Controls.Resources.System.png");
            string jobImageUrl = Page.ClientScript.GetWebResourceUrl(
                this.GetType(),
                "MVPSI.JAMSWeb.Controls.Resources.Job.png");
            string setupImageUrl = Page.ClientScript.GetWebResourceUrl(
                this.GetType(),
                "MVPSI.JAMSWeb.Controls.Resources.Setup.png");

            //
            //  Build the menu
            //
            ICollection<MenuEntry> menu;
            if (menuID == 0)
            {
                menu = BuildMenu.Find(menuType, menuName, GetServer());
            }
            else
            {
                menu = BuildMenu.Find(menuType, menuID, GetServer());
            }

            //
            //  Add each entry in the menu as a TreeNode
            //
            foreach (MenuEntry entry in menu)
            {
                TreeNode node = new TreeNode(
                    entry.MenuText,
                    string.Format("{0},{1},{2}", (int)entry.MenuType, entry.ID, entry.MenuName));

                //
                //  Pick the correct image
                //
                switch (entry.MenuType)
                {
                    case MenuType.Job:
                        node.ImageUrl = jobImageUrl;
                        break;
                    case MenuType.Setup:
                        node.ImageUrl = setupImageUrl;
                        break;
                    case MenuType.Folder:
                        node.ImageUrl = systemImageUrl;
                        break;
                    default:
                        node.ImageUrl = menuImageUrl;
                        break;
                }

                //
                //  Should this be expandable?
                //
                if ((entry.MenuType == MenuType.Menu)
                    || (entry.MenuType == MenuType.Folder))
                {
                    node.PopulateOnDemand = true;
                    node.SelectAction = TreeNodeSelectAction.Expand;
                }

                //
                //  Add the node to the collection
                //
                nodes.Add(node);
            }
        }

        /// <summary>
        /// Clicking on a leaf node selects it, we handle that selection be submitting the
        /// job or setup that the leaf points at.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SubmitTree_SelectedNodeChanged(object sender, EventArgs e)
        {
            string[] nodeData = SelectedNode.Value.Split(',');
            MenuType menuType = (MenuType)Int32.Parse(nodeData[0]);
            int menuID = Int32.Parse(nodeData[1]);
            Debug.WriteLine(string.Format("Submitting {0} with an id of {1}", menuType, menuID));
            JAMS.Submit.Type submitType;
            if (menuType == JAMS.MenuType.Job)
            {
                submitType = Submit.Type.JobID;
            }
            else
            {
                submitType = Submit.Type.SetupID;
            }

            SubmitMenu.SubmitDialog.ShowDialog(submitType, menuID);
            SelectedNode.Selected = false;
        }
    }

    internal class SubmitMenuDesigner:System.Web.UI.Design.ControlDesigner
    {
        protected SubmitMenu submitMenu;
        public override void Initialize(IComponent component)
        {
            if (component is SubmitMenu)
            {
                base.Initialize(component);
                submitMenu = (SubmitMenu) component;

                if (string.IsNullOrEmpty(submitMenu.ServerName))
                {
                    submitMenu.ServerName = "localhost";
                }
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
            return string.Format(
                "<asp:Panel runat=\"server\" style=\"height: 267px; width: 409px\" >JAMS Submit Control<br/>Connected to Server {0}</asp:Panel>",
                submitMenu.ServerName);
        }
    }
}