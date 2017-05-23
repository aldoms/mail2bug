using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using log4net;
using Mail2Bug.ExceptionClasses;
using Mail2Bug.MessageProcessingStrategies;
using Microsoft.TeamFoundation.Core.WebApi;
using Microsoft.TeamFoundation.Core.WebApi.Types;
using Microsoft.TeamFoundation.Work.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;
using Microsoft.VisualStudio.Services.WebApi.Patch;
using Microsoft.VisualStudio.Services.WebApi.Patch.Json;

namespace Mail2Bug.WorkItemManagement
{
    public class VssWorkItemManager : IWorkItemManager, IDisposable
    {
        public SortedList<string, int> WorkItemsCache { get; private set; }

        public VssWorkItemManager(Config.InstanceConfig config)
        {
            ValidateConfig(config);

            _config = config;

            // Init TFS service objects
            _vssConnection = GetVssConnection();
            _projectClient = _vssConnection.GetClient<ProjectHttpClient>();
            _teamClient = _vssConnection.GetClient<TeamHttpClient>();
            _workClient = _vssConnection.GetClient<WorkHttpClient>();
            _witClient = _vssConnection.GetClient<WorkItemTrackingHttpClient>();

            Logger.InfoFormat("Geting TFS Project");
            _tfsProject = _projectClient.GetProject(config.TfsServerConfig.Project).Result;

            Logger.InfoFormat("Getting TFS WorkItem Type");
            _workItemType = _witClient.GetWorkItemTypesAsync(_tfsProject.Id).Result
                .First(wit => wit.Name.Equals(_config.TfsServerConfig.WorkItemTemplate, StringComparison.InvariantCultureIgnoreCase));
            _formalFieldNames = _workItemType.Fields.ToDictionary(f => f.Name.ToLowerInvariant(), f => f.ReferenceName);

            Logger.InfoFormat("Getting Team Config");
            var teamContext = new TeamContext(_tfsProject.Id);
            _teamSettings = _workClient.GetTeamSettingsAsync(teamContext).Result;
            if (_teamSettings.DefaultIteration != null) {
                // TODO: for some reason the path in the default iteration included in TeamSettings does not
                //       match the actual value, so get the iteration explicitly here
                _defaultIteration = _workClient.GetTeamIterationAsync(teamContext, _teamSettings.DefaultIteration.Id).Result;
             }

            Logger.InfoFormat("Initializing WorkItems Cache");
            InitWorkItemsCache();

            _nameResolver = InitNameResolver();
        }

        public void Dispose()
        {
            _projectClient?.Dispose();
            _teamClient?.Dispose();
            _witClient?.Dispose();
            _workClient?.Dispose();
        }

        ~VssWorkItemManager()
        {
            Dispose();
        }

        private VssConnection GetVssConnection()
        {
            var tfsCredentials = GetVssCredentials();

            foreach (var credentials in tfsCredentials)
            {
                try
                {
                    Logger.InfoFormat("Connecting to TFS {0} using {1} credentials", _config.TfsServerConfig.CollectionUri, credentials);
                    var tfsServer = new VssConnection(new Uri(_config.TfsServerConfig.CollectionUri), credentials);

                    Logger.InfoFormat("Successfully connected to TFS");

                    return tfsServer;
                }
                catch (Exception ex)
                {
                    Logger.WarnFormat("TFS connection attempt failed.\n Exception: {0}", ex);
                }
            }

            Logger.ErrorFormat("All TFS connection attempts failed");
            throw new Exception("Cannot connect to TFS");
        }

        private IEnumerable<VssCredentials> GetVssCredentials()
        {
            var credentials = new List<VssCredentials>();

            credentials.AddRange(GetServiceIdentityPatCredentials());

            return credentials;
        }

        private IEnumerable<VssCredentials> GetServiceIdentityPatCredentials()
        {
            if (string.IsNullOrWhiteSpace(_config.TfsServerConfig.ServiceIdentityPatFile) && _config.TfsServerConfig.ServiceIdentityPatKeyVaultSecret == null)
            {
                return new List<VssCredentials>();
            }

            var basicCred = new VssBasicCredential("", GetPatFromConfig());
            var patCred = new VssCredentials(basicCred);

            return new List<VssCredentials> { patCred };
        }

        private string GetPatFromConfig()
        {
            if (string.IsNullOrWhiteSpace(_config.TfsServerConfig.ServiceIdentityPatFile) && _config.TfsServerConfig.ServiceIdentityPatKeyVaultSecret == null)
            {
                return null;
            }

            var credentialsHelper = new Helpers.CredentialsHelper();
            return credentialsHelper.GetPassword(
                _config.TfsServerConfig.ServiceIdentityPatFile,
                _config.TfsServerConfig.EncryptionScope,
                _config.TfsServerConfig.ServiceIdentityPatKeyVaultSecret);
        }

        public void AttachFiles(int workItemId, List<string> fileList)
        {
            if (workItemId <= 0) return;

            try
            {
                fileList.ForEach(file =>
                {
                    var attachment = _witClient.CreateAttachmentAsync(file).Result;
                    var patchDocument = new JsonPatchDocument
                    {
                        new JsonPatchOperation()
                        {
                            Operation = Operation.Add,
                            Path = "/relations/-",
                            Value = new
                            {
                                rel = "AttachedFile",
                                url = attachment.Url,
                                attributes = new
                                {
                                    comment = "Mail2Bug Original Message"
                                }
                            }
                        }
                    };
                    var updatedWorkItem = _witClient.UpdateWorkItemAsync(patchDocument, workItemId).Result;
                });
            }
            catch (Exception exception)
            {
                Logger.Error(exception.ToString());
            }
        }

        private JsonPatchOperation CreateWitPatchOperation(string fieldName, string value)
        {
            return new JsonPatchOperation()
            {
                Operation = Operation.Add,
                Path = $"/fields/{GetFormalFieldName(fieldName)}",
                Value = value
            };
        }

        /// <param name="values">The list of fields and their desired values to apply to the work item</param>
        /// <returns>Work item ID of the newly created work item</returns>
        public int CreateWorkItem(Dictionary<string, string> values)
        {
            if (values == null)
            {
                throw new ArgumentNullException(nameof(values), "Must supply field values when creating new work item");
            }

            // Use patch to add fields
            JsonPatchDocument patchDocument = new JsonPatchDocument();

            foreach (var key in values.Keys)
            {
                string value = values[key];

                // Resolve current iteration
                if (_teamSettings != null && key == IterationPathFieldKey && value == _teamSettings.DefaultIterationMacro)
                {
                    if (_defaultIteration != null)
                    {
                        value = _defaultIteration.Path;
                    }
                }

                patchDocument.Add(CreateWitPatchOperation(key, value));
            }

            // Workaround for TFS issue - if you change the "Assigned To" field, and then you change the "Activated by" field, the "Assigned To" field reverts
            // to its original setting. To prevent that, we reapply the "Assigned To" field in case it's in the list of values to change.
            if (values.ContainsKey(AssignedToFieldKey))
            {
                patchDocument.Add(CreateWitPatchOperation(AssignedToFieldKey, values[AssignedToFieldKey]));
            }

            //create a work item
            var workItem = _witClient.CreateWorkItemAsync(patchDocument, _tfsProject.Id.ToString(), _workItemType.Name).Result;

            CacheWorkItem(workItem);
            return workItem.Id.Value;
        }

        /// <param name="workItemId">The ID of the work item to modify </param>
        /// <param name="comment">Comment to add to description</param>
        /// <param name="values">List of fields to change</param>
        public void ModifyWorkItem(int workItemId, string comment, Dictionary<string, string> values)
        {
            if (workItemId <= 0) return;

            var jsonPatchDocument = new JsonPatchDocument
            {
                CreateWitPatchOperation("History", comment.Replace("\n", "<br>"))
            };

            foreach (var key in values.Keys)
            {
                jsonPatchDocument.Add(CreateWitPatchOperation(key, values[key]));
            }

            var updatedWorkItem = _witClient.UpdateWorkItemAsync(jsonPatchDocument, workItemId).Result;
        }

        #region Work item caching

        public void CacheWorkItem(int workItemId)
        {
            if (WorkItemsCache.ContainsValue(workItemId)) return; // Work item already cached - nothing to do

            // It is important that we don't just get the conversation ID from the caller and update the cache with the work item
            // ID and conversation ID, because if the work item already exists, the conversation ID will be different (probably shorter
            // than the one the caller currently has)
            // That's why we get the work item from TFS and get the conversation ID from there
            var workItem = _witClient.GetWorkItemAsync(workItemId).Result;
            CacheWorkItem(workItem);
        }

        /// <returns>Sorted List of FieldValue's with ConversationIndex as the key</returns>
        private void InitWorkItemsCache()
        {
            Logger.InfoFormat("Initializing work items cache");

            WorkItemsCache = new SortedList<string, int>();

            //search TFS to get list
            var itemsToCache = _witClient.QueryByWiqlAsync(new Wiql() { Query = _config.TfsServerConfig.CacheQuery }).Result;
            Logger.InfoFormat("{0} items retrieved by TFS cache query", itemsToCache.WorkItems.Count());
            var itemIds = itemsToCache.WorkItems.Select(w => w.Id).ToArray();
            var workItems = _witClient.GetWorkItemsAsync(itemIds).Result;
            foreach (var workItem in workItems)
            {
                try
                {
                    CacheWorkItem(workItem);
                }
                catch (Exception ex)
                {
                    Logger.ErrorFormat("Exception caught while caching work item with id {0}\n{1}", workItem.Id, ex);
                }
            }
        }

        private string GetFormalFieldName(string fieldName)
        {
            fieldName = fieldName.ToLowerInvariant();
            if (_formalFieldNames.ContainsKey(fieldName) == false)
            {
                return null;
            }

            return _formalFieldNames[fieldName];
        }

        private void CacheWorkItem(Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models.WorkItem workItem)
        {
            var keyField = GetFormalFieldName(_config.WorkItemSettings.ConversationIndexFieldName);

            if (string.IsNullOrWhiteSpace(keyField))
            {
                Logger.WarnFormat("Item {0} doesn't contain the key field {1}. Not caching", workItem.Id, keyField);
                return;
            }

            var keyFieldValue = workItem.Fields[keyField].ToString().Trim();
            Logger.DebugFormat("Work item {0} conversation ID is {1}", workItem.Id, keyFieldValue);
            if (string.IsNullOrEmpty(keyFieldValue))
            {
                Logger.DebugFormat("Problem caching work item {0}. Field '{1}' is empty - using ID instead.", workItem.Id, keyField);
                WorkItemsCache[workItem.Id?.ToString(CultureInfo.InvariantCulture)] = workItem.Id.Value;
            }

            WorkItemsCache[keyFieldValue] = workItem.Id.Value;
        }

        #endregion

        public INameResolver GetNameResolver()
        {
            return _nameResolver;
        }

        private NameResolver InitNameResolver()
        {
            var fieldDef = _workItemType.Fields.Where(i => i.Name.Equals(_config.TfsServerConfig.NamesListFieldName, StringComparison.InvariantCultureIgnoreCase));
            return new NameResolver(fieldDef.Select(f => f.ReferenceName));
        }

        #region Config validation

        private static void ValidateConfig(Config.InstanceConfig config)
        {
            if (config == null) throw new ArgumentNullException(nameof(config));

            // Temp variable used for shorthand writing below
            var tfsConfig = config.TfsServerConfig;

            ValidateConfigString(tfsConfig.CollectionUri, "TfsServerConfig.CollectionUri");
            ValidateConfigString(tfsConfig.Project, "TfsServerConfig.Project");
            ValidateConfigString(tfsConfig.WorkItemTemplate, "TfsServerConfig.WorkItemTemplate");
            ValidateConfigString(tfsConfig.CacheQuery, "TfsServerConfig.CacheQuery");
            ValidateConfigString(tfsConfig.NamesListFieldName, "TfsServerConfig.NamesListFieldName");
            ValidateConfigString(tfsConfig.ServiceIdentityPatKeyVaultSecret?.KeyVaultPath, "tfsConfig.ServiceIdentityPatKeyVaultSecret.KeyVaultPath");
            ValidateConfigString(config.WorkItemSettings.ConversationIndexFieldName,
                                 "WorkItemSettings.ConversationIndexFieldName");
        }

        // ReSharper disable UnusedParameter.Local
        private static void ValidateConfigString(string value, string configValueName)
        // ReSharper restore UnusedParameter.Local
        {
            if (string.IsNullOrEmpty(value)) throw new BadConfigException(configValueName);
        }

        #endregion

        #region Consts

        private const string AssignedToFieldKey = "Assigned To";
        private const string IterationPathFieldKey = "Iteration Path";

        #endregion

        private readonly TeamProject _tfsProject;
        private readonly WorkItemType _workItemType;
        private readonly TeamSetting _teamSettings;
        private readonly TeamSettingsIteration _defaultIteration;
        private readonly NameResolver _nameResolver;
        private readonly IDictionary<string, string> _formalFieldNames;

        private readonly Config.InstanceConfig _config;

        private static readonly ILog Logger = LogManager.GetLogger(typeof(TFSWorkItemManager));
        private readonly VssConnection _vssConnection;
        private readonly TeamHttpClient _teamClient;
        private readonly ProjectHttpClient _projectClient;
        private readonly WorkHttpClient _workClient;
        private readonly WorkItemTrackingHttpClient _witClient;
    }
}
