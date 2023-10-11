using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ParamAccessHelper;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Configuration;



namespace Web_Kactus_Test
{
    class AbrirPrograma
    {
        
        //static public object[] Pdatain;
        //static public object[] Pdatain2;
        //static public object[] pdataprocesar;
        string Code;
        //string token;
        //BrowserWindow browser;
        string usuer;
        System.Diagnostics.Process p = new System.Diagnostics.Process();
        string UserDomains;
        string PassDomains;

        public AbrirPrograma(string CodigoPrograma, string usuario, string UserDomain = null, string PassDomain = null)
        {

            this.Code = CodigoPrograma;
            this.usuer = usuario;
            this.UserDomains = UserDomain;
            this.PassDomains = PassDomain;

        }
    }
    class TFSData
    {
        private string testcaseID;
        private string suites;
        private string plans;
        private string ServerURL = ConfigurationManager.AppSettings["ServerURL"];
        private string ServerPat = ConfigurationManager.AppSettings["ServerPat"];

        public TFSData( string testcaseIDs, string plan=null, string suite=null)
        {
            this.testcaseID = testcaseIDs;
            this.suites=plan;
            this.plans=suite;
        }

        public DataSet GetParams()
        { 
            DataSet ds = new DataSet();
            GetTestCaseParams p = new GetTestCaseParams();
            p.VstsURI = ServerURL;
            p.Pat = ServerPat;

            Task.Run(async () => { ds = await p.GetParams(this.testcaseID); }).GetAwaiter().GetResult();
            return ds;
        }




        public DataSet GetParamsExecutionCases()
        {
            DataSet ds = new DataSet();
            GetTestCaseParams p = new GetTestCaseParams();
            p.VstsURI = ServerURL;
            p.Pat = ServerPat;

            Task.Run(async () => { ds = await p.GetEcecutionTfsTestCase(this.plans, this.suites); }).GetAwaiter().GetResult();
            return ds;
        }


        public DataSet GetParamsBuildTest()
        {
            DataSet ds = new DataSet();
            GetTestCaseParams p = new GetTestCaseParams();
            p.VstsURI = ServerURL;
            p.Pat = ServerPat;

            Task.Run(async () => { ds = await p.GetBuildTest(); }).GetAwaiter().GetResult();
            return ds;
        }


        public List<string> GetQuery(string q)
        {
            List<string> ds = new List<string>();
            GetTestCaseParams p = new GetTestCaseParams();
            p.VstsURI = ServerURL;
            p.Pat = ServerPat;

            Task.Run(async () => { ds = await p.GetTestCasesByQuery(q); }).GetAwaiter().GetResult();
            return ds;
        }


        ~TFSData()
        {

        }

    }
}
