using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;
using Microsoft.VisualStudio.Services.WebApi.Patch;
using Microsoft.VisualStudio.Services.WebApi.Patch.Json;
using System;
using System.Collections.Generic;

namespace WorkItemPublish.Class
{
    class WIOps
    {
        static string Url = "https://dev.azure.com/sagorg1";

        static WorkItemTrackingHttpClient WitClient;
        public static WorkItem CreateWorkItem(string ProjectName, string WorkItemTypeName, Dictionary<string, object> Fields)
        {
            JsonPatchDocument patchDocument = new JsonPatchDocument();

            foreach (var key in Fields.Keys)
                patchDocument.Add(new JsonPatchOperation()
                {
                    Operation = Operation.Add,
                    Path = "/fields/" + key,
                    Value = Fields[key]
                });

            return WitClient.CreateWorkItemAsync(patchDocument, ProjectName, WorkItemTypeName).Result;
        }
        public static WorkItem UpdateWorkItemLink(int parentId, int childId, string message)
        {
            JsonPatchDocument patchDocument = new JsonPatchDocument();

            patchDocument.Add(new JsonPatchOperation()
            {
                Operation = Operation.Add,
                Path = "/relations/-",
                Value = new
                {
                    rel = "System.LinkTypes.Hierarchy-Reverse",
                    url = Url + "/_apis/wit/workitems/" + parentId,
                    attributes = new
                    {
                        comment = "Linking the workitems"
                    }
                }
            });

            return WitClient.UpdateWorkItemAsync(patchDocument, childId).Result;
        }
        public static WorkItem UpdateWorkItemFields(int WIId, Dictionary<string, object> Fields)
        {
            JsonPatchDocument patchDocument = new JsonPatchDocument();

            foreach (var key in Fields.Keys)
                patchDocument.Add(new JsonPatchOperation()
                {
                    Operation = Operation.Add,
                    Path = "/fields/" + key,
                    Value = Fields[key]
                });
            if (Fields.Count != 0)
                return WitClient.UpdateWorkItemAsync(patchDocument, WIId).Result;
            else
                return null;
        }
        public static void ConnectWithPAT(string ServiceURL, string PAT)
        {
            VssConnection connection = new VssConnection(new Uri(ServiceURL), new VssBasicCredential("xx", PAT));
            InitClients(connection);
        }
        static void InitClients(VssConnection Connection)
        {
            WitClient = Connection.GetClient<WorkItemTrackingHttpClient>();
        }
    }
}
