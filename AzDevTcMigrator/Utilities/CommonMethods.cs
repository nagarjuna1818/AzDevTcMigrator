using AzDevTcMigrator.Entities;

using Ganss.Excel;

using Microsoft.TeamFoundation.TestManagement.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.WebApi.Patch.Json;

using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace AzDevTcMigrator.Utilities
{
    public class CommonMethods
    {
        public TestStep BuildTestStep(TestCase testCase)
        {
            return new TestStep()
            {
                Action = $"StepName: {testCase.StepName}\r\nDescription: {testCase.StepDescription}",
                ExpectedResult = testCase.StepExpectedResult
            };
        }

        public JsonPatchDocument BuildJsonPatchDocument(TestCase testCase)
        {
            var jsonPatchDocument = new JsonPatchDocument();
            var jsonPatchOperations = new List<JsonPatchOperation>();

            TestBaseHelper helper = new TestBaseHelper();
            ITestBase testBase = helper.Create();

            testCase.Steps.ForEach(x =>
            {
                ITestStep testStep = testBase.CreateTestStep();
                testStep.Title = x.Action ?? string.Empty;
                testStep.ExpectedResult = x.ExpectedResult ?? string.Empty;
                testStep.Description = x.Action ?? string.Empty;

                testBase.Actions.Add(testStep);
            });

            var properties = typeof(TestCase).GetProperties();
            foreach (var propertyInfo in properties.Where(x => x.Name != "Steps"))
            {
                object[] attrs = propertyInfo.GetCustomAttributes(true);
                foreach (object attr in attrs)
                {
                    AzDevFieldReferenceAttribute azDevFieldReferenceAttribute = attr as AzDevFieldReferenceAttribute;
                    if (azDevFieldReferenceAttribute != null)
                    {
                        var data = testCase.GetType().GetProperty(propertyInfo.Name).GetValue(testCase, null);
                        if (propertyInfo.Name.Contains("AttachmentReferences"))
                        {
                            var castedData = data as List<AttachmentReference>;
                            castedData.ForEach(x =>
                            {
                                jsonPatchOperations.Add(new JsonPatchOperation
                                {
                                    Operation = Microsoft.VisualStudio.Services.WebApi.Patch.Operation.Add,
                                    Path = azDevFieldReferenceAttribute.AzDevFieldReference,
                                    Value = new { rel = "AttachedFile", url = x.Url }
                                });
                            });
                        }
                        else
                        {
                            if (!string.IsNullOrWhiteSpace(data as string))
                            {
                                jsonPatchOperations.Add(new JsonPatchOperation
                                {
                                    Operation = Microsoft.VisualStudio.Services.WebApi.Patch.Operation.Add,
                                    Path = azDevFieldReferenceAttribute.AzDevFieldReference,
                                    Value = data
                                });
                            }
                        }
                    }
                }
            }

            jsonPatchDocument.AddRange(jsonPatchOperations);
            jsonPatchDocument = testBase.SaveActions(jsonPatchDocument);
            return jsonPatchDocument;
        }

        public async Task<IEnumerable<TestCase>> GetTestCasesDataFromExcel(string filePath)
        {
            var excelMapper = new ExcelMapper();
            var excelData = await excelMapper.FetchAsync<TestCase>(filePath).ConfigureAwait(false);
            var testCasesData = new List<TestCase>();
            excelData.ToList().ForEach(x =>
            {
                var testStepData = this.BuildTestStep(x);
                if (!string.IsNullOrWhiteSpace(x.TestName))
                {
                    x.Steps = new List<TestStep>
                    {
                        testStepData
                    };

                    x.AttachmentReferences = new List<AttachmentReference>();
                    testCasesData.Add(x);
                }
                else
                {
                    testCasesData.Last().Steps.Add(testStepData);
                }
            });

            return testCasesData;
        }
    }
}
