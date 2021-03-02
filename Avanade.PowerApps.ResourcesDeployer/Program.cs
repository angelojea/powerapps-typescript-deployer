using System.Diagnostics;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Discovery;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.Crm.Sdk.Messages;
using Newtonsoft.Json.Linq;

namespace Avanade.PowerApps.ResourcesDeployer
{
    class Program
    {
        static void Main(string[] args)
        {
            if (!File.Exists(Directory.GetCurrentDirectory() + "//settings.json")) throw new Exception("Missing \"settings.json\" file");

            var jsonContent = File.ReadAllText(Directory.GetCurrentDirectory() + "//settings.json");
            var json = JObject.Parse(jsonContent);

            var mandatoryFields = new string[] {
                "working-folder", "crm", "client-id", "client-secret", "resources",
                "prefix"
            };

            Console.WriteLine("Validating settings.json");

            foreach (var field in mandatoryFields)
            {
                if (json[field] == null) throw new Exception($"Missing property \"{field}\" on \"settings.json\" file ");
            }

            var prefix = json["prefix"].ToString() + "_";
            var workingFolder = json["working-folder"].ToString();
            var crm = json["crm"].ToString();
            var clientId = json["client-id"].ToString();
            var clientSecret = json["client-secret"].ToString();

            using (var svc = new CrmServiceClient(
                $@"AuthenticationType=ClientSecret; url={crm}; ClientId={clientId}; ClientSecret={clientSecret};"))
            {
                var existingFiles = new EntityCollection();
                var resources = new Dictionary<string, string>();
                foreach (JProperty key in json["resources"].ToList())
                    resources.Add(key.Name, key.Value.ToString());

                if (resources.Count > 0)
                {
                    Console.WriteLine("Retrieving previous records");

                    existingFiles = svc.RetrieveMultiple(new FetchExpression($@"
                    <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                        <entity name='webresource'>
                            <attribute name='name' />
                            <attribute name='webresourceid' />
                            <filter type='or'>
                                {string.Join("", resources.Keys.Select(x => $"<condition attribute='name' operator='eq' value='{prefix + resources[x]}' />").ToList())}
                            </filter>
                        </entity>
                    </fetch>"));
                }

                var reqs = getNewExecuteMultiple();

                Console.WriteLine("Reading targeted resources");

                foreach (var file in resources.Keys)
                {
                    if (!File.Exists(workingFolder + "/" + file)) continue;

                    var fragments = file.Split('.');
                    var fileName = fragments.First();
                    var extension = fragments.Last() ?? "";

                    var webResource = new Entity("webresource");
                    webResource["name"] = prefix + resources[file];
                    webResource["displayname"] = "zzzz" + resources[file];

                    // 1 - HTML
                    // 2 - CSS
                    // 3 - JS
                    switch (extension.ToLower())
                    {
                        case "css":
                            webResource["webresourcetype"] = new OptionSetValue(2);
                            break;
                        case "js":
                            webResource["webresourcetype"] = new OptionSetValue(3);
                            break;
                        case "html":
                        default:
                            webResource["webresourcetype"] = new OptionSetValue(1);
                            break;
                    }

                    if (existingFiles
                        .Entities
                        .Where(x => x["name"].ToString().ToLower() == webResource["name"].ToString().ToLower()).Count() > 0)
                    {
                        webResource.Id = existingFiles.Entities.Where(x => x["name"].ToString().ToLower() == webResource["name"].ToString().ToLower()).First().Id;
                    }

                    webResource["content"] = Convert.ToBase64String(Encoding.UTF8.GetBytes(File.ReadAllText(workingFolder + "/" + file)));
                    
                    reqs.Requests.Add(new UpsertRequest() { Target = webResource });
                }
                svc.Execute(reqs);

                existingFiles = svc.RetrieveMultiple(new FetchExpression($@"
                <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                    <entity name='webresource'>
                        <attribute name='name' />
                        <attribute name='webresourceid' />
                        <filter type='or'>
                            {string.Join("", resources.Keys.Select(x => $"<condition attribute='name' operator='eq' value='{prefix + resources[x]}' />").ToList())}
                        </filter>
                    </entity>
                </fetch>"));

                Console.WriteLine("Publishing web resources");

                if (existingFiles.Entities.Count > 0)
                {
                    svc.Execute(new PublishXmlRequest()
                    {
                        ParameterXml =
                        $@"<importexportxml><webresources>
                        {string.Join("", existingFiles.Entities.Select(x => "<webresources>" + x["webresourceid"] + "</webresources>"))}
                        </webresources></importexportxml>"
                    });
                }
            }
        }

        static ExecuteMultipleRequest getNewExecuteMultiple()
    {
        return new ExecuteMultipleRequest()
        {
            Requests = new OrganizationRequestCollection(),
            Settings = new ExecuteMultipleSettings()
            {
                ContinueOnError = true,
                ReturnResponses = true
            }
        };
    }

}
}
