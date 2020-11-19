using AzDevTcMigrator.Utilities;

using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace AzDevTcMigrator
{
    class Program
    {
        internal const string TestCase = "Test Case";

        static async Task Main(string[] args)
        {
            string azureDevOpsOrganizationUrl = string.Empty;
            string projectName = string.Empty;
            string pat = string.Empty;
            string excelFilePath = string.Empty;
            string attachmentsFolderPath = string.Empty;

            try
            {
                azureDevOpsOrganizationUrl = args[0];
                projectName = args[1];
                pat = args[2];
                excelFilePath = args[3];
                attachmentsFolderPath = args[4];
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Following exception occured. Please provide correct details and try again.\r\n");
                throw ex;
            }

            var commonMethods = new CommonMethods();

            Console.WriteLine($"Reading the data from Excel file...");

            //Read Data
            var testCasesData = await commonMethods.GetTestCasesDataFromExcel(excelFilePath).ConfigureAwait(false);

            Console.WriteLine($"Creating the connection with given Azure Devops Organization using provided PAT...");

            //Prompt user for credential
            VssConnection connection = new VssConnection(new Uri(azureDevOpsOrganizationUrl), new VssBasicCredential(string.Empty, pat));

            var workItemTrackClient = connection.GetClient<WorkItemTrackingHttpClient>();

            Console.WriteLine($"Iterating through the files in attachments folder to link to corresponding test-cases...");

            var fileNames = Directory.GetFiles(attachmentsFolderPath).Select(Path.GetFileName);
            foreach (var fileName in fileNames)
            {
                var stringSplits = fileName.Split("_");
                if (stringSplits.Count() > 1)
                {
                    if (stringSplits[0].Equals("TEST", StringComparison.OrdinalIgnoreCase))
                    {
                        var attachmentTestId = stringSplits[1];
                        var testCaseWithAttachmentTestId = testCasesData.FirstOrDefault(x => x.TestID == attachmentTestId);
                        if (testCaseWithAttachmentTestId != null)
                        {
                            try
                            {
                                var attachmentReference = await workItemTrackClient.CreateAttachmentAsync(Path.Join(attachmentsFolderPath, fileName)).ConfigureAwait(false);
                                testCaseWithAttachmentTestId.AttachmentReferences.Add(attachmentReference);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Exception with message - \"{ex.Message}\" occured while creating an attachment for TestCase with Attachment Test ID - {attachmentTestId}");
                            }
                        }
                    }
                }
            }

            Console.WriteLine($"Test-cases creation started...");

            var createdWorkItems = new List<WorkItem>();
            foreach (var testCase in testCasesData)
            {
                try
                {
                    testCase.Tags = "projectname; poc";
                    var jsonPatchDocument = commonMethods.BuildJsonPatchDocument(testCase);
                    createdWorkItems.Add(await workItemTrackClient.CreateWorkItemAsync(jsonPatchDocument, projectName, TestCase).ConfigureAwait(false));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"{ex.Message} occured for following test-case with Test ID - {testCase.TestID}");
                }
            }

            Console.WriteLine($"Test-case creation completed for given excel file !!!");
        }
    }
}
