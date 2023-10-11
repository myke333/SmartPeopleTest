using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ParamAccessHelper
{
    //Class was generated
    class TestCaseJSON
    {
    }

    public class Rootobject
    {
        public int count { get; set; }
        public Value[] value { get; set; }
    }

    public class Value
    {
        //
        public int id { get; set; }
        public int rev { get; set; }
        public Fields fields { get; set; }
        public string url { get; set; }
    }

    public class Fields
    {
        public string SystemAreaPath { get; set; }
        public string SystemTeamProject { get; set; }
        public string SystemIterationPath { get; set; }
        public string SystemWorkItemType { get; set; }
        public string SystemState { get; set; }
        public string SystemReason { get; set; }
        public string SystemAssignedTo { get; set; }
        public DateTime SystemCreatedDate { get; set; }
        public string SystemCreatedBy { get; set; }
        public DateTime SystemChangedDate { get; set; }
        public string SystemChangedBy { get; set; }
        public string SystemTitle { get; set; }
        public int MicrosoftVSTSCommonPriority { get; set; }
        public DateTime MicrosoftVSTSCommonStateChangeDate { get; set; }
        public DateTime MicrosoftVSTSCommonActivatedDate { get; set; }
        public string MicrosoftVSTSCommonActivatedBy { get; set; }
        public string MicrosoftVSTSTCMAutomatedTestName { get; set; }
        public string MicrosoftVSTSTCMAutomatedTestStorage { get; set; }
        public string MicrosoftVSTSTCMAutomatedTestId { get; set; }
        public string MicrosoftVSTSTCMAutomatedTestType { get; set; }
        public string MicrosoftVSTSTCMAutomationStatus { get; set; }
        public string MicrosoftVSTSTCMSteps { get; set; }
        public string MicrosoftVSTSTCMParameters { get; set; }

        [JsonProperty("Microsoft.VSTS.TCM.LocalDataSource")]
        public string MicrosoftVSTSTCMLocalDataSource { get; set; }
    }
}
