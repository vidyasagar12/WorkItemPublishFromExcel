﻿using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
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
        static string Url;

        static WorkItemTrackingHttpClient WitClient;
        
        public static WorkItem CreateWorkItem(string ProjectName, string WorkItemTypeName, Dictionary<string, object> Fields)
        {
            
            JsonPatchDocument patchDocument = new JsonPatchDocument();
            try
            {
                foreach (var key in Fields.Keys)
                    patchDocument.Add(new JsonPatchOperation()
                    {
                        Operation = Operation.Add,
                        Path = "/fields/" + key,
                        Value = Fields[key]
                    });

                return WitClient.CreateWorkItemAsync(patchDocument, ProjectName, WorkItemTypeName).Result;
            }
            catch (Exception E)
            {
                Console.WriteLine("Error Occured While Creating Workitem"+Fields["Title"]);
                throw E;                
            }
        }
        public static WorkItem UpdateWorkItemLink(int parentId, int childId, string message)
        {
            JsonPatchDocument patchDocument = new JsonPatchDocument();
            try
            {
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
            catch (Exception E)
            {
                Console.WriteLine("Error Occured While Updating Parent Of"+childId);

                throw E;
                return null;
            }
        }
        
        public static WorkItem UpdateWorkItemFields(int WIId, Dictionary<string, object> Fields)
        {
            try
            {
                JsonPatchDocument patchDocument = new JsonPatchDocument();
                foreach (var key in Fields.Keys) {
                    JsonPatchOperation Jsonpatch = new JsonPatchOperation()
                    {
                        Operation = Operation.Replace,
                        Path = "/fields/" + key,
                        Value = Fields[key]
                    };
                    if(!patchDocument.Contains(Jsonpatch))
                    patchDocument.Add(Jsonpatch);
                }
                if (patchDocument.Count != 0)
                {
                    return WitClient.UpdateWorkItemAsync(patchDocument, WIId).Result;
                }
                else
                    return null;
            }
            catch(Exception E)
            {
                Console.WriteLine("Error Occured While Updating Fields Of "+WIId);
                throw E;
            }
        }
   
        public static void ConnectWithPAT(string ServiceURL, string PAT)
        {
            try
            {
                Url = ServiceURL;
                VssConnection connection = new VssConnection(new Uri(ServiceURL), new VssBasicCredential("xx", PAT));
                InitClients(connection);
            }
            catch (Exception E)
            {
                Console.WriteLine("Error Occured While Connecting to:"+ServiceURL);
                throw E;
            }
        }
        static void InitClients(VssConnection Connection)
        {
            try
            {
                WitClient = Connection.GetClient<WorkItemTrackingHttpClient>();
            }
            catch (Exception E)
            {
                Console.WriteLine("Error Occured To Initialize The WorkItem Client Make Sure That You Have Sufficient Permissions");
                throw E;

            }
        }
    }
}
