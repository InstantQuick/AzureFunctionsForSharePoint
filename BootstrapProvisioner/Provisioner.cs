using System;
using IQAppProvisioningBaseClasses.Events;
using IQAppProvisioningBaseClasses.Provisioning;
using Microsoft.SharePoint.Client;

namespace BootstrapProvisioner
{
    public class Provisioner : ProvisioningManagerBase
    {
        protected bool IsHostWeb { get; set; }
        protected ClientContext Ctx { get; set; }
        protected Web Web { get; set; }

        private void Provisioner_Notify(object sender, IQAppProvisioningBaseClasses.Events.ProvisioningNotificationEventArgs eventArgs)
        {
            OnNotify(eventArgs.Level, eventArgs.Detail);
        }

        private bool WebHasAppinstanceId(Web web)
        {
            if (!web.IsPropertyAvailable("AppInstanceId"))
            {
                web.Context.Load(web, w => w.AppInstanceId);
                web.Context.ExecuteQueryRetry();
            }
            return web.AppInstanceId != default(Guid);
        }

        public void Provision(ClientContext ctx, Web web, AppManifestBase manifest)
        {
            IsHostWeb = !WebHasAppinstanceId(web);
            Ctx = ctx;
            Web = web;

            if (!ContextLoaded()) LoadContext();

            if (IsHostWeb)
            {
                AddFeatures(manifest);
                OnNotify(ProvisioningNotificationLevels.Verbose, "Added features");
                RemoveFeatures(manifest);
                OnNotify(ProvisioningNotificationLevels.Verbose, "Removed features");
                ProvisionGroups(manifest);
                OnNotify(ProvisioningNotificationLevels.Verbose, "Created groups");
                ProvisionRoleDefinitions(manifest);
                OnNotify(ProvisioningNotificationLevels.Verbose, "Created role definitions");
            }
            else
            {
                OnNotify(ProvisioningNotificationLevels.Verbose, "Site is an app web. Skipped features and security.");
            }
            ProvisionFields(manifest);
            OnNotify(ProvisioningNotificationLevels.Verbose, "Created fields");
            ProvisionContentTypes(manifest);
            OnNotify(ProvisioningNotificationLevels.Verbose, "Created content types");
            ProvisionLists(manifest);
            OnNotify(ProvisioningNotificationLevels.Verbose, "Created lists");
            ProvisionFiles(manifest);
            OnNotify(ProvisioningNotificationLevels.Verbose, "Created files");
            ProvisionNavigation(manifest);
            OnNotify(ProvisioningNotificationLevels.Verbose, "Set navigation");
            ProvisionClassicWorkflows(manifest);
            OnNotify(ProvisioningNotificationLevels.Verbose, "Configured 2010 style workflows");
            ProvisionCustomActions(manifest);
            OnNotify(ProvisioningNotificationLevels.Verbose, "Created custom actions");
            AttachEvents(manifest);
            OnNotify(ProvisioningNotificationLevels.Verbose, "Attached event handlers");
            ApplyDocumentTemplates(manifest);
            OnNotify(ProvisioningNotificationLevels.Verbose, "Set document templates");
            ProvisionLookAndFeel(manifest);
            OnNotify(ProvisioningNotificationLevels.Verbose, "Set look and feel");
            OnNotify(ProvisioningNotificationLevels.Verbose, "Successfully provisioned");
        }

        private void ProvisionLookAndFeel(AppManifestBase manifest)
        {
            if (manifest.LookAndFeel == null) return;

            var lfm = new LookAndFeelManager();
            lfm.Notify += Provisioner_Notify;
            lfm.ProvisionLookAndFeel(manifest, Ctx, Web);
        }

        private void ApplyDocumentTemplates(AppManifestBase manifest)
        {
            if (manifest.ListCreators != null && manifest.ListCreators.Count > 0)
            {
                foreach (var listCreator in manifest.ListCreators.Values)
                {
                    listCreator.UpdateDocumentTemplate(Ctx);
                }
            }
        }

        private void AttachEvents(AppManifestBase manifest)
        {
            if (manifest.RemoteEventRegistrationCreators == null || manifest.RemoteEventRegistrationCreators.Count == 0)
                return;
            var manager = new RemoteEventRegistrationManager();
            manager.Notify += Provisioner_Notify;
            manager.CreateEventHandlers(Ctx, Web, manifest.RemoteEventRegistrationCreators, manifest.RemoteHost);
        }

        private void AddFeatures(AppManifestBase manifest)
        {
            var featureManager = new FeatureManager { FeaturesToAdd = manifest.AddFeatures };
            featureManager.Notify += Provisioner_Notify;
            featureManager.ConfigureFeatures(Ctx, Web);
        }

        private void RemoveFeatures(AppManifestBase manifest)
        {
            var featureManager = new FeatureManager { FeaturesToRemove = manifest.RemoveFeatures };
            featureManager.Notify += Provisioner_Notify;
            featureManager.ConfigureFeatures(Ctx, Web);
        }

        private void ProvisionGroups(AppManifestBase manifest)
        {
            if (manifest.GroupCreators != null && manifest.GroupCreators.Count > 0)
            {
                var groupManager = new GroupManager { GroupCreators = manifest.GroupCreators };
                groupManager.Notify += Provisioner_Notify;
                groupManager.ProvisionGroups(Ctx, Web);
            }
        }

        private void ProvisionRoleDefinitions(AppManifestBase manifest)
        {
            if (manifest.RoleDefinitions != null && manifest.RoleDefinitions.Count > 0)
            {
                var roleDefinitionManager = new RoleDefinitionManager(Ctx, Web)
                {
                    RoleDefinitions = manifest.RoleDefinitions
                };
                roleDefinitionManager.Notify += Provisioner_Notify;
                roleDefinitionManager.Provision();
            }
        }

        private void ProvisionNavigation(AppManifestBase manifest)
        {
            if (manifest.Navigation != null)
            {
                var navigationManager = new NavigationManager(Ctx, Web)
                {
                    ClearLeftMenu = manifest.Navigation.ClearLeftMenu,
                    ClearTopMenu = manifest.Navigation.ClearTopMenu,
                    TopNavigationNodes = manifest.Navigation.TopNavigationNodes,
                    LeftNavigationNodes = manifest.Navigation.LeftNavigationNodes
                };

                navigationManager.Notify += Provisioner_Notify;
                //App webs don't have oob menus. Create menus on host web
                if (IsHostWeb)
                {
                    navigationManager.Provision();
                }
                //but create a custom action to inject the nav via JavaScript for app webs
                {
                    manifest.CustomActionCreators["IQAppWebNavigation"] =
                        navigationManager.CreateNavigationUserCustomAction(manifest.Navigation);
                }
            }
        }

        private void ProvisionCustomActions(AppManifestBase manifest)
        {
            if (manifest.CustomActionCreators != null && manifest.CustomActionCreators.Count > 0)
            {
                var actionMan = new CustomActionManager(Ctx, Web) { CustomActions = manifest.CustomActionCreators };
                actionMan.Notify += Provisioner_Notify;
                actionMan.CreateAll();
            }
        }

        private void ProvisionFiles(AppManifestBase appManifest)
        {
            var fileManager = new FileManager();
            fileManager.Notify += Provisioner_Notify;
            fileManager.ProvisionAll(Ctx, Web, appManifest);
        }

        private void ProvisionLists(AppManifestBase manifest)
        {
            if (manifest.ListCreators != null && manifest.ListCreators.Count > 0)
            {
                var lm = new ListInstanceManager(Ctx, Web, IsHostWeb) { Creators = manifest.ListCreators };
                lm.Notify += Provisioner_Notify;
                lm.CreateAll();
            }
        }

        private void ProvisionContentTypes(AppManifestBase manifest)
        {
            if (manifest.ContentTypeCreators != null && manifest.ContentTypeCreators.Count > 0)
            {
                var cm = new ContentTypeManager { Creators = manifest.ContentTypeCreators };
                cm.Notify += Provisioner_Notify;
                //ContentTypes should always be provisioned into the root or app web
                cm.CreateAll(Ctx);
            }
        }

        private void ProvisionFields(AppManifestBase manifest)
        {
            if (manifest.Fields != null && manifest.Fields.Count > 0)
            {
                var fm = new FieldManager { FieldDefinitions = manifest.Fields };
                fm.Notify += Provisioner_Notify;
                //Fields should always be provisioned into the root or app web
                fm.CreateAll(Ctx);
            }
        }

        private void ProvisionClassicWorkflows(AppManifestBase manifest)
        {
            if (manifest.ClassicWorkflowCreators == null || manifest.ClassicWorkflowCreators.Count == 0) return;

            var cm = new ClassicWorkflowManager { Creators = manifest.ClassicWorkflowCreators };
            cm.Notify += Provisioner_Notify;
            //App identities can't call the web service to register the workflow
            if (Ctx.AuthenticationMode != ClientAuthenticationMode.Anonymous)
            {
                //Normal call with credentials
                cm.CreateAll(Ctx);
            }
            else
            //So create a self destructing custom action to register via the browser
            {
                var userCustomActionTitle = "AppWorkflowAssociationCustomAction" + manifest.ManifestName;
                //manifest.CustomActionCreators = manifest.CustomActionCreators != null ? manifest.CustomActionCreators : new Dictionary<string, CustomActionCreatorBase>();
                manifest.CustomActionCreators[userCustomActionTitle] = cm.CreateAppWorkflowAssociationCustomAction(Ctx,
                    Web, manifest.ClassicWorkflowCreators, userCustomActionTitle);
            }
        }

        private bool ContextLoaded()
        {
            return !(!Ctx.Site.IsPropertyAvailable("ServerRelativeUrl") ||
                     !Ctx.Site.IsPropertyAvailable("RootWeb") ||
                     !Ctx.Site.RootWeb.IsPropertyAvailable("ServerRelativeUrl") ||
                     !Ctx.Web.IsPropertyAvailable("ServerRelativeUrl"));
        }

        private void LoadContext()
        {
            Ctx.Load(Ctx.Site);
            Ctx.Load(Ctx.Site.RootWeb);
            Ctx.Load(Ctx.Site.RootWeb.AllProperties);
            Ctx.Load(Ctx.Web);
            Ctx.Load(Ctx.Web.AllProperties);
            Ctx.ExecuteQueryRetry();
        }
    }
}
