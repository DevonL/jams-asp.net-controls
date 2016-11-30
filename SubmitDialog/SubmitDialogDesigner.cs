using System;
using System.ComponentModel;
using System.IO;
using System.Web.UI;
using MVPSI.JAMS;

namespace MVPSI.JAMSWeb.Controls
{
    internal class SubmitDialogDesigner : System.Web.UI.Design.ControlDesigner
    {
        private const string DesignHTML =
            @"<div style=""width:500px;height:500px;""><div style=""background-color:#CEE5FE;margin:-3, -3, 6, -3;"">
	<span>Job Name - </span><span>Description of the Job.</span></div>
    <div><span><span><a style='padding:2px;border-top:thin white inset;border-left:thin white inset; border-right:thin white inset;;' href='#'>Parameters </a></span>
    <span><a style='padding:2px;border-top:thin white inset;border-left:thin white inset; border-right:thin white inset;;' href='#'> Reports </a></span>
    <span><a style='padding:2px;border-top:thin white inset;border-left:thin white inset; border-right:thin white inset;;' href='#'> Schedule</a></span></span></div>
    <div></br>Parameters <input value=""Value""/></div></div>";

        public override void Initialize(IComponent component)
        {
            if (component is SubmitDialog)
            {
                base.Initialize(component);
                
            }
        }

        public override string GetDesignTimeHtml()
        {
            try
            {
                string output = string.Empty;

                SubmitDialog dialog = (SubmitDialog) this.Component;

                if (dialog != null)
                {
                    dialog.ShowDialog(Submit.Type.Job, 5);
                }

                StringWriter writer = new StringWriter();
                HtmlTextWriter html = new HtmlTextWriter(writer);
                dialog.RenderBeginTag(html);
                html.Write(DesignHTML);
                dialog.RenderEndTag(html);
                output = writer.ToString();

                return output;

            }
            catch (Exception ex)
            {
                return ex.ToString();
            }
        }
    }
}