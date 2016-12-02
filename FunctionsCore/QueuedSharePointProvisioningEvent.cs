using IQAppProvisioningBaseClasses.Provisioning;

namespace FunctionsCore
{
    public enum ProvisioningAction
    {
        Install,
        Upgrade
    }
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
    public class QueuedSharePointProvisioningEvent : QueuedSharePointEvent
    {
        public ProvisioningAction Action { get; set; }
        public ProvisioningSteps ProvisioningStep { get; set; }
        public AppManifestBase AppManifest { get; set; }
    }
}
