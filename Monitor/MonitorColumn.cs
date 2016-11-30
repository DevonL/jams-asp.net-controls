using System.ComponentModel;

namespace MVPSI.JAMSWeb.Controls
{
    /// <summary>
    /// MonitorColumn is the class used to manage a column in a Monitor view.
    /// </summary>
    [TypeConverter(typeof(MonitorColumnConverter))]
    public class MonitorColumn
    {
        /// <summary>
        /// Gets or sets the MonitorColumnType of this column.
        /// </summary>
        [NotifyParentProperty(true)]
        public MonitorColumnType Type { get; set; }
    }
}