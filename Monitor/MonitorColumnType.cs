namespace MVPSI.JAMSWeb.Controls
{
    /// <summary>
    /// Identifies the different Monitor columns
    /// </summary>
    public enum MonitorColumnType
    {
        /// <summary>
        /// The name of the Job.
        /// </summary>
        JobName,
        /// <summary>
        /// The JAMS entry number.
        /// </summary>
        JAMSEntry,
        /// <summary>
        /// A Description of the entry.
        /// </summary>
        Description,
        /// <summary>
        /// The Status of the Job.
        /// </summary>
        JobStatus,
        /// <summary>
        /// The date and time that this Job was submitted to the JAMS schedule.
        /// </summary>
        ScheduledTime,
        /// <summary>
        /// The date and time that this Job completed.
        /// </summary>
        CompletionTime,
        /// <summary>
        /// The date and time that this Job actually started.
        /// </summary>
        StartTime,
        /// <summary>
        /// The current state of the Job.
        /// </summary>
        CurrentState,
        /// <summary>
        /// Indicates which step is currently executing in a setup.
        /// </summary>
        CurrentStep,
        /// <summary>
        /// The date and time that this Job was submitted to the Operating System.
        /// </summary>
        SubmitTime,
        /// <summary>
        /// The jobs current elapsed execution time. If the job has completed, the time is the total execution time of the job.
        /// </summary>
        ElapsedTime,
        /// <summary>
        /// The identity of the user who submitted this run.
        /// </summary>
        SubmittedBy,
        /// <summary>
        /// True if this Job is running on a remote system which has a JAMS Agent.
        /// </summary>
        Agent,
        /// <summary>
        /// The node the Job is to run on.
        /// </summary>
        AgentNode,
        /// <summary>
        /// The batch queue used by this run.
        /// </summary>
        BatchQueue,
        /// <summary>
        /// True if this was a debug run.
        /// </summary>
        DebugMode,
        /// <summary>
        /// Indicates if the dependencies defined for this Job have been satisfied.
        /// </summary>
        DependOK,
        /// <summary>
        /// The display name of the Job.
        /// </summary>
        DisplayName,
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
        /// True when a multi-job Setup is halted.
        /// </summary>
        Halted,
        /// <summary>
        /// True if this entry is currently Held.
        /// </summary>
        Held,
        /// <summary>
        /// The time that this entry will be released to run.
        /// </summary>
        HoldTime,
        /// <summary>
        /// True if the Job should ignore its resource requirements.
        /// </summary>
        IgnoreResReq,
        /// <summary>
        /// When true, the Job's log file is kept upon completion.
        /// </summary>
        KeepLogs,
        /// <summary>
        /// The JAMS Server that this run is from.
        /// </summary>
        LoadedFrom,
        /// <summary>
        /// Indicates if this Log filename is active for this Job.
        /// </summary>
        LogFileActive,
        /// <summary>
        /// The Log filename for this Job.
        /// </summary>
        LogFilename,
        /// <summary>
        /// The actions which should occur if this job misses it's scheduled window.
        /// </summary>
        MissedWindowAction,
        /// <summary>
        /// True if this entry is viewable in the JAMS Monitor.
        /// </summary>
        Monitor,
        /// <summary>
        /// The name for the job, the DisplayName if there is one. Otherwise, the JobName.
        /// </summary>
        Name,
        /// <summary>
        /// The node that ran this Job.
        /// </summary>
        NodeName,
        /// <summary>
        /// The Note added to this run when it was submitted.
        /// </summary>
        Note,
        /// <summary>
        /// When True, notification will be performed if this job misses it's schedule window.
        /// </summary>
        NotifyOfMissedWindow,
        /// <summary>
        /// If True, the user who manually submits the job will be added to the notification list.
        /// </summary>
        NotifyUser,
        /// <summary>
        /// This is the value of the HoldTime property when this Job was originally placed into the schedule.
        /// </summary>
        OriginalHoldTime,
        /// <summary>
        /// If the Job or Setup definition allows the precheck Job to run on an interval, this value indicates how many times the precheck Job has been released to run.
        /// </summary>
        PrecheckCount,
        /// <summary>
        /// If a precheck Job is specified for this Job, this value specifies the JAMSEntry number for that precheck Job.
        /// </summary>
        PrecheckEntry,
        /// <summary>
        /// The interval between runs of the Job's Precheck job.
        /// </summary>
        PrecheckInterval,        
        /// <summary>
        /// Indicates the next DateTime that this precheck Job will be released.
        /// </summary>
        PrecheckTime,
        /// <summary>
        /// True if this entry has been preprocessed. Jobs which are preprocessed have had their parameters evaluated and Job source created from the Job definition.
        /// </summary>
        Preprocessed,
        /// <summary>
        /// When True, the Job's log file is printed upon completion.
        /// </summary>
        PrintLogs,
        /// <summary>
        /// The operating system's process identifier for this run.
        /// </summary>
        ProcessID,
        /// <summary>
        /// True if the process which is running this Job is a proxy. A proxy process runs a Job on a system not monitored by a JAMS Agent.
        /// In this case the statistics associated with the proxy process do not represent those of the real Job.
        /// </summary>
        Proxy,
        /// <summary>
        /// The restartability of this job. A restartable job can be aborted and restarted.
        /// </summary>
        Restartable,
        /// <summary>
        /// If this Job has been restarted the value here indicates how many restarts have occurred
        /// </summary>
        RestartCount,
        /// <summary>
        /// The retention policy for this Job. The options are:
        /// Not Specified   The retention policy is taken from the Job's System definition.
        /// Always (A)      Job is always retained until specifically deleted.
        /// Never (N)       Never retain job.
        /// Error (E)       Job is retained if it completes with a severity of Warning or worse.
        /// Timed (T)       Job is retained for the specified time after completion.
        /// </summary>
        RetainOption,
        /// <summary>
        /// The length of time the Job should be retained when the retain option is set to "Timed".
        /// </summary>
        RetainTime,
        /// <summary>
        /// The run occurrence number.
        /// </summary>
        RON,
        /// <summary>
        /// This will be true if the job has exceeded one of it's execution time limits.
        /// </summary>
        Runaway,
        /// <summary>
        /// The absolute runaway Cpu time. If the job consumes more Cpu time than this it is declared a runaway job and notification will be performed.
        /// </summary>
        RunawayCpu,
        /// <summary>
        /// The absolute runaway elapsed time. If the job runs longer than this, it is declared a runaway job and notification will be performed.
        /// </summary>
        RunawayElapsed,
        /// <summary>
        /// The execution priority for this job.
        /// </summary>
        RunPriority,
        /// <summary>
        /// The beginning of this job's scheduled time window.
        /// </summary>
        ScheduleFromTime,
        /// <summary>
        /// The ending of this job's schedule time window.
        /// </summary>
        ScheduleToTime,
        /// <summary>
        /// The Job's Scheduled Window.
        /// </summary>
        ScheduleWindow,        
        /// <summary>
        /// The SchedulingPriority of the Job.
        /// </summary>
        SchedulingPriority,
        /// <summary>
        /// The SetupID.
        /// </summary>
        SetupID,
        /// <summary>
        /// True if this Job completed successfully. In a multi-job Setup if any one Job sets this property to false, then the Setup will fail.
        /// </summary>
        SetupOK,
        /// <summary>
        /// This will be true if the job has been pending longer than it's stalled time limit.
        /// </summary>
        Stalled,
        /// <summary>
        /// How much time may elapse after a job's scheduled time before the job is considered to be stalled. When the job is considered stalled, JAMS will perform notification for the Job.
        /// </summary>
        StalledTime,
        /// <summary>
        /// Each Job in a multi-job Setup can have a Setup value. If specified, Jobs run in Step order.
        /// </summary>
        Step,
        /// <summary>
        /// Indicates if this Job is waiting for previous steps to complete.
        /// </summary>
        StepWait,
        /// <summary>
        /// If set to True, Job completion notification is suppressed.
        /// </summary>
        SuppressNotify,
        /// <summary>
        /// The username that this run used.
        /// </summary>
        UserName,
        /// <summary>
        /// This is True if the Setup that contains this job should wait for the completion of this job before advancing it's step.
        /// </summary>
        WaitFor,
    }
}