using Ganss.Excel;

using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;

using System;
using System.Collections.Generic;

namespace AzDevTcMigrator.Entities
{
    public class TestCase
    {
        [Column("Test Name")]
        [AzDevFieldReference("/fields/System.Title")]
        public string TestName { get; set; }

        [AzDevFieldReference("/fields/System.Description")]
        public string Description { get; set; }

        //[AzDevFieldReference("/fields/Microsoft.VSTS.Common.Priority")]
        public string Priority { get; set; }

        [Ignore]
        [AzDevFieldReference("/fields/Microsoft.VSTS.TCM.Steps")]
        public List<TestStep> Steps { get; set; }

        [Column("Step Name (STEP)")]
        public string StepName { get; set; }

        [Column("Expected Result (STEP)")]
        public string StepExpectedResult { get; set; }

        [Column("Description (STEP)")]
        public string StepDescription { get; set; }

        [Column("Execution Status")]
        public string ExecutionStatus { get; set; }

        [Ignore]
        [AzDevFieldReference("/fields/System.Tags")]
        public string Tags { get; set; }

        [Column("Test ID")]
        [AzDevFieldReference("/fields/Custom.Source_Test_ID")]
        public string TestID { get; set; }

        [Ignore]
        [AzDevFieldReference("/relations/-")]
        public List<AttachmentReference> AttachmentReferences { get; set; }
    }

    public class AzDevFieldReferenceAttribute : Attribute
    {
        public AzDevFieldReferenceAttribute(string azDevFieldReference)
        {
            this.AzDevFieldReference = azDevFieldReference;
        }

        public string AzDevFieldReference { get; private set; }
    }
}
