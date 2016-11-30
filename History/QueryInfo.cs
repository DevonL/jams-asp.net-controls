using System;
using System.ComponentModel;
using MVPSI.JAMS;

namespace MVPSI.JAMSWeb.Controls
{
    /// <summary>
    /// Holds Query information used to retrieve History records.
    /// </summary>
    public class QueryInfo
    {
        /// <summary>
        /// Gets or sets the search string of the System to query.
        /// </summary>
        /// <value>The search string of the System to query.</value>
        [DefaultValue("*"), NotifyParentProperty(true)]
        public string SystemName { get; set; }

        /// <summary>
        /// Gets or sets the search string for the name of the Job.
        /// </summary>
        /// <value>The name of the job.</value>
        [DefaultValue("*"), NotifyParentProperty(true)]
        public string JobName { get; set; }

        /// <summary>
        /// Gets or sets the search string for the name of the Setup.
        /// </summary>
        /// <value>The name of the setup.</value>
        [DefaultValue("*"), NotifyParentProperty(true)]
        public string SetupName { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether or not to check the ScheduledAt time.
        /// </summary>
        /// <value><c>true</c> if checking the ScheduledAt time; otherwise, <c>false</c>.</value>
        [DefaultValue(true), NotifyParentProperty(true)]
        public bool CheckScheduledAt { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether or not to check the ScheduledFor time.
        /// </summary>
        /// <value><c>true</c> if checking the ScheduledFor time; otherwise, <c>false</c>.</value>
        [DefaultValue(true), NotifyParentProperty(true)]
        public bool CheckScheduledFor { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether or not to check the Started time.
        /// </summary>
        /// <value><c>true</c> if checking the Started time; otherwise, <c>false</c>.</value>
        [DefaultValue(true), NotifyParentProperty(true)]
        public bool CheckStarted { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether or not to check the Completed time.
        /// </summary>
        /// <value><c>true</c> if checking the Completed time; otherwise, <c>false</c>.</value>
        [DefaultValue(true), NotifyParentProperty(true)]
        public bool CheckCompleted { get; set; }

        /// <summary>
        /// Gets or sets the start time to query.
        /// </summary>
        /// <value>The start time.</value>
        [DefaultValue("12:00:00 AM"), TypeConverter(typeof(JAMSTimeConverter)), NotifyParentProperty(true)]
        public string StartTime { get; set; }

        /// <summary>
        /// Gets or sets the end time to query.
        /// </summary>
        /// <value>The end time.</value>
        [DefaultValue("11:59:59 PM"), TypeConverter(typeof(JAMSTimeConverter)), NotifyParentProperty(true)]
        public string EndTime { get; set; }

        /// <summary>
        /// Gets or sets the start date to query.
        /// </summary>
        /// <value>The start date.</value>
        [DefaultValue("Today"), TypeConverter(typeof(JAMSDateConverter)), NotifyParentProperty(true)]
        public string StartDate { get; set; }

        /// <summary>
        /// Gets or sets the end date to query.
        /// </summary>
        /// <value>The end date.</value>
        [DefaultValue("Today"), TypeConverter(typeof(JAMSDateConverter)), NotifyParentProperty(true)]
        public string EndDate { get; set; }

        /// <summary>
        /// Gets or sets the time within the past timespan to look at.
        /// </summary>
        /// <value>The within past time.</value>
        [DefaultValue(""), NotifyParentProperty(true), TypeConverter(typeof(JAMSTimeConverter))]
        public string WithinPastTime { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether or not to include successful job completions
        /// </summary>
        /// <value><c>true</c> if successful completions are included; otherwise, <c>false</c>.</value>
        [DefaultValue(true), NotifyParentProperty(true)]
        public bool IncludeSuccessful { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether or not to include informational job completions
        /// </summary>
        /// <value><c>true</c> if informational completions are included; otherwise, <c>false</c>.</value>
        [DefaultValue(true), NotifyParentProperty(true)]
        public bool IncludeInformational { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether or not to include warning completions
        /// </summary>
        /// <value><c>true</c> if warning completions are included; otherwise, <c>false</c>.</value>
        [DefaultValue(true), NotifyParentProperty(true)]
        public bool IncludeWarning { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether or not to include error completions
        /// </summary>
        /// <value><c>true</c> if error completions are included; otherwise, <c>false</c>.</value>
        [DefaultValue(true), NotifyParentProperty(true)]
        public bool IncludeError { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether or not to include fatal completions
        /// </summary>
        /// <value><c>true</c> if fatal completions are included; otherwise, <c>false</c>.</value>
        [DefaultValue(true), NotifyParentProperty(true)]
        public bool IncludeFatal { get; set; }

        /// <summary>
        /// Gets or sets the name of the server to load history data from.
        /// </summary>
        /// <value>The name of the server.</value>
        [DefaultValue(""), NotifyParentProperty(true)]
        public string ServerName { get; set; }

        /// <summary>
        /// If true, clears the last set of results.
        /// </summary>
        [DefaultValue(true), NotifyParentProperty(true)]
        public bool ClearCurrentList { get; set; }

        /// <summary>
        /// Field to search for.
        /// </summary>
        [DefaultValue(""), NotifyParentProperty(true)]
        public string HistorySelectionField1 { get; set; }

        /// <summary>
        /// How to compare between the selected history field and the given comparison value.
        /// </summary>
        [DefaultValue(""), NotifyParentProperty(true)]
        public string ComparisonOperator1 { get; set; }

        /// <summary>
        /// Value to search for.
        /// </summary>
        [DefaultValue("*"), NotifyParentProperty(true)]
        public string ComparisonValue1 { get; set; }

        /// <summary>
        /// Field to reach for.
        /// </summary>
        [DefaultValue(""), NotifyParentProperty(true)]
        public string HistorySelectionField2 { get; set; }

        /// <summary>
        /// How to compare between the selected history field and the given comparison value.
        /// </summary>
        [DefaultValue(""), NotifyParentProperty(true)]
        public string ComparisonOperator2 { get; set; }

        /// <summary>
        /// Value to search for.
        /// </summary>
        [DefaultValue("*"), NotifyParentProperty(true)]
        public string ComparisonValue2 { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="QueryInfo"/> class.
        /// </summary>
        public QueryInfo()
        {
            SystemName = "*";
            JobName = "*";
            SetupName = "*";
            CheckScheduledAt = true;
            CheckScheduledFor = true;
            CheckStarted = true;
            CheckCompleted = true;
            StartTime = "12:00:00 AM";
            EndTime = "11:59:59 PM";
            StartDate = DateTime.Today.ToString();
            EndDate = DateTime.Today.ToString();
            IncludeSuccessful = true;
            IncludeInformational = true;
            IncludeWarning = true;
            IncludeError = true;
            IncludeFatal = true;
            ServerName = string.Empty;
            ClearCurrentList = true;
            HistorySelectionField1 = HistorySelectionField.None.ToString();
            ComparisonOperator1 = ComparisonOperator.Like.ToString();
            ComparisonValue1 = "*";
            HistorySelectionField2 = HistorySelectionField.None.ToString();
            ComparisonOperator2 = ComparisonOperator.Like.ToString();
            ComparisonValue2 = "*";
        }
      
    }


    /// <summary>
    /// Converts QueryInfo objects.
    /// </summary>
    internal class QueryInfoConverter : ExpandableObjectConverter
    {
        public override object ConvertTo(ITypeDescriptorContext context, System.Globalization.CultureInfo culture, object value, Type destinationType)
        {
            if(destinationType == typeof(string) && value is QueryInfo)
            {
                return string.Empty;
            }
            
            return base.ConvertTo(context, culture, value, destinationType);
        }
    }

    /// <summary>
    /// Converts to JAMSTime.
    /// </summary>
    internal class JAMSTimeConverter: StringConverter
    {
        public override bool IsValid(ITypeDescriptorContext context, object value)
        {
            DateTime.Parse(value.ToString());

            return true;
        }

        public override object ConvertFrom(ITypeDescriptorContext context, System.Globalization.CultureInfo culture, object value)
        {
            if(value is string)
            {
                if (context != null)
                {
                    IsValid(context, value);
                    context.OnComponentChanged();
                }
            }
           return base.ConvertFrom(context, culture, value);
        }
    }

    /// <summary>
    /// Converts JAMSDate.
    /// </summary>
    internal class JAMSDateConverter: StringConverter
    {
        public override bool IsValid(ITypeDescriptorContext context, object value)
        {
            JAMS.Date.Check(value.ToString(), Server.GetServer(((QueryInfo)context.Instance).ServerName));
            return true;
        }

        public override object ConvertFrom(ITypeDescriptorContext context, System.Globalization.CultureInfo culture, object value)
        {
            if (value is string)
            {
                if (context != null)
                {
                    IsValid(context, value);
                    context.OnComponentChanged();
                }
            }
            
            return base.ConvertFrom(context, culture, value);
        }        
    }
}
