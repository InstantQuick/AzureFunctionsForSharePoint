
namespace AzureFunctionsForSharePoint.Common
{
    /// <summary>
    /// Distinguishes between install and upgrade. Interpreting this is up to the client
    /// </summary>
    public enum ProvisioningAction
    {
        Install,
        Upgrade
    }

    /// <summary>
    /// Identifies the various stages of provisioning jobs
    /// </summary>
    public enum ProvisioningSteps
    {
        NotStarted,
        Upgrading,
        Features,
        GroupsAndRoles,
        Fields,
        ContentTypes,
        Lists,
        Files,
        ClassicWorkflows,
        Navigation,
        CustomActions,
        Settings,
        Events,
        DocumentTemplates,
        Complete,
        ErrorRetry
    }

    /// <summary>
    /// Service Bus queue message sent by AppLaunch the first time an app is launched to notify a client of an install event.
    /// Use this in your own jobs if you need to break up the operation or to support custom install and upgrade jobs.
    /// </summary>
    public class QueuedSharePointProvisioningEvent : QueuedSharePointEvent
    {
        /// <summary>
        /// Distinguishes between install and upgrade. Interpreting this is up to the client
        /// </summary>
        public ProvisioningAction Action { get; set; }
        /// <summary>
        /// The provisioning step to perform.
        /// </summary>
        public ProvisioningSteps ProvisioningStep { get; set; }
    }
}
