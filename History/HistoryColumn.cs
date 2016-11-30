using System;
using System.ComponentModel;

namespace MVPSI.JAMSWeb.Controls
{    
    /// <summary>
    /// Identifies the different History columns.
    /// </summary>
    public enum HistoryColumnType
    {
        /// <summary>
        /// The Job Name.
        /// </summary>
        JobName,
        /// <summary>
        /// The JAMS entry number for this run of the Job.
        /// </summary>
        JAMSEntry,
        /// <summary>
        /// The final value of the JobStatus text.
        /// </summary>
        JobStatus,
        /// <summary>
        /// The date and time that this Job was submitted to the JAMS schedule.
        /// </summary>
        ScheduledTime,
        /// <summary>
        /// The date and time that this Job actually started.
        /// </summary>
        StartTime,
        /// <summary>
        /// The date and time that this Job completed.
        /// </summary>
        CompletionTime,
        /// <summary>
        /// The identity of the user who submitted this run.
        /// </summary>
        SubmittedBy,
        /// <summary>
        /// The Average Elapsed Time for this Job.
        /// </summary>
        AvgElapsedTime,
        /// <summary>
        /// The batch queue used by this run.
        /// </summary>
        BatchQueue,
        /// <summary>
        /// The buffered I/O count for this run.
        /// </summary>
        BufferedIOCount,
        /// <summary>
        /// The value of the checkpoint data when this run was restarted.
        /// </summary>
        CheckpointData,
        /// <summary>
        /// The Cpu time consumed by this run.
        /// </summary>
        CpuTime,
        /// <summary>
        /// True if this was a debug run. Debug runs don't satisfy dependencies or fire triggers.
        /// </summary>
        DebugMode,
        /// <summary>
        /// The direct I/O count for this run.
        /// </summary>
        DirectIOCount,
        /// <summary>
        /// The Elapsed Time for this run.
        /// </summary>
        ElapsedTime,
        /// <summary>
        /// Type of History entry. There are three types: Job, Setup, and Command Procedure.
        /// </summary>
        EntryType,
        /// <summary>
        /// The final severity for this run.
        /// </summary>
        FinalSeverity,
        /// <summary>
        /// The final status for this run.
        /// </summary>
        FinalStatus,
        /// <summary>
        /// The final status for this run.
        /// </summary>
        FinalStatusCode,
        /// <summary>
        /// The date and time that this Job was supposed to start.
        /// </summary>
        HoldTime,
        /// <summary>
        /// The Job ID.
        /// </summary>
        JobID,
        /// <summary>
        /// The JAMS Server that this run is from.
        /// </summary>
        LoadedFrom,
        /// <summary>
        /// The name of the log file for this run.
        /// </summary>
        LogFilename,
        /// <summary>
        /// The master run occurrence Number. If this was a member of a multi-job Setup, this number identifies the parent Setup.
        /// </summary>
        MasterRON,
        /// <summary>
        /// The Maximum recorded Elapsed Time for this Job.
        /// </summary>
        MaxElapsedTime,
        /// <summary>
        /// The Minimum recorded Elapsed Time for this Job.
        /// </summary>
        MinElapsedTime,
        /// <summary>
        /// The node that ran this Job.
        /// </summary>
        NodeName,
        /// <summary>
        /// The note added to this run when it was submitted.
        /// </summary>
        Note,
        /// <summary>
        /// The operating system's entry number, if any.
        /// </summary>
        OSEntry,
        /// <summary>
        /// The override name for this run.
        /// </summary>
        OverrideJobName,
        /// <summary>
        /// The number of page faults encountered by this run.
        /// </summary>
        PageFaults,
        /// <summary>
        /// The operating system's process identifier for this run.
        /// </summary>
        ProcessID,
        /// <summary>
        /// The restart count for this run. It is zero for the initial run and is incremented each time the Job is restarted.
        /// </summary>
        RestartCount,
        /// <summary>
        /// The restart sequence for this run. It is zero for the initial run and is incremented each time the Job is restarted.
        /// </summary>
        RestartSequence,
        /// <summary>
        /// The run occurrence number.
        /// </summary>
        RRON,
        /// <summary>
        /// The run occurrence number.
        /// </summary>
        RON,
        /// <summary>
        /// The Setup ID.
        /// </summary>
        SetupID,
        /// <summary>
        /// The username that this run used.
        /// </summary>
        UserName,
        /// <summary>
        /// The peak working set size for this run.
        /// </summary>
        WorkingSetPeak
    }

    /// <summary>
    /// Determines the information displayed in each Column of the History control.
    /// </summary>
    [TypeConverter(typeof(HistoryColumnConverter))]
    public class HistoryColumn
    {
        /// <summary>
        /// Gets or sets the type of the column to display.
        /// </summary>
        /// <value>The type.</value>
        [NotifyParentProperty(true)]
        public HistoryColumnType Type { get; set; }
    }

    /// <summary>
    /// Converts HistoryColumn information to strings
    /// </summary>
    public class HistoryColumnConverter : StringConverter 
    {
        /// <summary>
        /// Converts the given value object to the specified type, using the specified context and culture information.
        /// </summary>
        /// <param name="context">An <see cref="T:System.ComponentModel.ITypeDescriptorContext"/> that provides a format context.</param>
        /// <param name="culture">A <see cref="T:System.Globalization.CultureInfo"/>. If null is passed, the current culture is assumed.</param>
        /// <param name="value">The <see cref="T:System.Object"/> to convert.</param>
        /// <param name="destinationType">The <see cref="T:System.Type"/> to convert the <paramref name="value"/> parameter to.</param>
        /// <returns>
        /// An <see cref="T:System.Object"/> that represents the converted value.
        /// </returns>
        /// <exception cref="T:System.ArgumentNullException">The <paramref name="destinationType"/> parameter is null. </exception>
        /// <exception cref="T:System.NotSupportedException">The conversion cannot be performed. </exception>
        public override object ConvertTo(ITypeDescriptorContext context, System.Globalization.CultureInfo culture, object value, Type destinationType)
        {
            if (destinationType == typeof(string) && value is HistoryColumn)
            {
                return ((HistoryColumn)value).Type.ToString();
            }

            return base.ConvertTo(context, culture, value, destinationType);
        }
    }
}
