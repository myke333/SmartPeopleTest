using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Configuration;

using System.Text;
using System.Threading.Tasks;
using System.Net;

using System.Net.Http;
using System.Net.Http.Headers;
using System.Dynamic;

using System.Web.Script.Serialization;
using System.Collections.ObjectModel;

namespace ParamAccessHelper
{
    public class GetTestCaseParams
    {
        private string Project = ConfigurationManager.AppSettings["Project"];
        private string Collection = ConfigurationManager.AppSettings["Collection"];

        public DataSet Params { get; private set; }

        public DataSet Params2 { get; private set; }

        public DataSet Params3 { get; private set; }

        public string TestCaseJASON { get; private set; }

        public string VstsURI { get; set; }

        public string Pat { get; set; }

        public async Task<DataSet> GetEcecutionTfsTestCase(string plan, string suite)
        {


            try
            {
                string url = this.VstsURI + "/" + Project + "/" + Collection + "/_apis/test/Plans/" + plan + "/Suites/" + suite + "/points?api-version=1.0";

                using (HttpClient client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Accept.Add(
                        new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));

                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic",
                        Convert.ToBase64String(
                            System.Text.ASCIIEncoding.ASCII.GetBytes(
                                string.Format("{0}:{1}", "", this.Pat))));

                    using (HttpResponseMessage response = client.GetAsync(
                                url).Result)
                    {


                        response.EnsureSuccessStatusCode();
                        string responseBody = await response.Content.ReadAsStringAsync();

                        

                        this.TestCaseJASON = responseBody;
                        JavaScriptSerializer jss = new JavaScriptSerializer();
                        jss.RegisterConverters(new JavaScriptConverter[] { new DynamicJsonConverter() });

                        dynamic TFSResponse = jss.Deserialize(TestCaseJASON, typeof(object)) as dynamic;
                        //int con = 0;

                        dynamic b = TFSResponse.value;
                        dynamic c;
                        string Key;
                        int TestPoint = 0;
                        string state = null;
                        dynamic e;
                        string testId = null;
                        string NameTest = null;
                        string NameTest2 = null;
                        DataSet TestP = new DataSet();

                        DataTable dt = new DataTable();

                        dt.Columns.Add("testpoint");
                        dt.Columns.Add("state");
                        dt.Columns.Add("TestId");
                        dt.Columns.Add("TestName");

                        foreach (dynamic a in b)
                        {
                            c = a;

                            foreach (dynamic d in c)
                            {

                                Key = d.Key;
                                if (Key == "id")
                                {
                                    TestPoint = d.Value;

                                }
                                if (Key == "state")
                                {
                                    state = d.Value;

                                }
                                if (Key == "testCase")
                                {

                                    e = d.Value;
                                    foreach (dynamic f in e)
                                    {
                                        Key = f.Key;
                                        if (Key == "id")
                                        {
                                            testId = f.Value;
                                        }
                                    }
                                }
                                if (Key == "workItemProperties")
                                {
                                    dynamic m = d.Value;
                                    foreach (dynamic n in m)
                                    {
                                        dynamic o = n;
                                        foreach (dynamic p in o)
                                        {
                                            dynamic q = p.Value;
                                            foreach (dynamic r in q)
                                            {
                                                Key = r.Key;
                                                if (Key == "value")
                                                {
                                                    NameTest = r.Value;
                                                    if (NameTest != "Automatizada")
                                                    {
                                                        DataRow newRow = dt.NewRow();
                                                        newRow[0] = TestPoint;
                                                        newRow[1] = state;
                                                        newRow[2] = testId;
                                                        NameTest2 = r.Value;
                                                        newRow[3] = NameTest2;
                                                        dt.Rows.Add(newRow);
                                                    }
                                                }
                                            }
                                        }


                                    }

                                }
                                {
                                }

                            }
                        }

                        TestP.Tables.Add(dt);
                        this.Params2 = TestP;
                        return TestP;
                    }
                }
            }
            catch (Exception ex)
            {
                return null;
                throw ex;
            }
        }

        public async Task<DataSet> GetBuildTest()
        {
            try
            {
                DataSet TestP = new DataSet();
                string url = this.VstsURI + "/" + Project + "/" + Collection + "/_apis/build/definitions/3";

                using (HttpClient client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Accept.Add(
                        new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));

                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic",
                        Convert.ToBase64String(
                            System.Text.ASCIIEncoding.ASCII.GetBytes(
                                string.Format("{0}:{1}", "", this.Pat))));

                    using (HttpResponseMessage response = client.GetAsync(
                                url).Result)
                    {
                        DataSet dataSet = new DataSet();

                        response.EnsureSuccessStatusCode();
                        string responseBody = await response.Content.ReadAsStringAsync();

                        this.TestCaseJASON = responseBody;

                        this.TestCaseJASON = responseBody;
                        JavaScriptSerializer jss = new JavaScriptSerializer();
                        jss.RegisterConverters(new JavaScriptConverter[] { new DynamicJsonConverter() });

                        dynamic TFSResponse = jss.Deserialize(TestCaseJASON, typeof(object)) as dynamic;
                        int con = 1;
                        dynamic c;
                        string Key;
                        //int TestPoint = 0;
                        //string state = null;
                        //string testId;
                        //dynamic d;
                        dynamic f;
                        //dynamic l;
                        dynamic b = TFSResponse.build;
                        //dynamic h;
                        dynamic i;
                        string plan = null;
                        string suites;



                        DataTable dt = new DataTable();

                        dt.Columns.Add("Plan");
                        dt.Columns.Add("Suite");
                        dt.Columns.Add("ExecutionOrder");

                        foreach (dynamic a in b)
                        {
                            c = a;
                            foreach (dynamic e in c)
                            {
                                f = e;
                                Key = f.Key;
                                if (Key == "displayName")
                                {
                                    if (f.Value == "INCAP")
                                    {
                                        foreach (dynamic j in c)
                                        {
                                            Key = j.Key;
                                            if (Key == "inputs")
                                            {
                                                i = j.Value;

                                                foreach (dynamic k in i)
                                                {

                                                    Key = k.Key;
                                                    if (Key == "testPlan")
                                                    {
                                                        plan = k.Value;
                                                    }
                                                    if (Key == "testSuite")
                                                    {
                                                        suites = k.Value;
                                                        string[] suite = suites.Split(',');
                                                        foreach (string m in suite)
                                                        {
                                                            DataRow newRow = dt.NewRow();
                                                            newRow[0] = plan;
                                                            newRow[1] = m;
                                                            newRow[2] = con;
                                                            con = con + 1;
                                                            dt.Rows.Add(newRow);
                                                        }
                                                    }

                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        TestP.Tables.Add(dt);
                    }
                }
                this.Params3 = TestP;
                return TestP;
            }
            catch (Exception ex)
            {
                return null;
                throw ex;
            }
        }

        public async Task<DataSet> GetParams(string workitemID)
        {

            try
            {
                //System.Diagnostics.Debugger.Launch();
                string url = this.VstsURI + "/" + Project + "/_apis/wit/workitems?ids=" + workitemID + "&api-version=1.0";

                using (HttpClient client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Accept.Add(
                        new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));

                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic",
                        Convert.ToBase64String(
                            System.Text.ASCIIEncoding.ASCII.GetBytes(
                                string.Format("{0}:{1}", "", this.Pat))));

                    using (HttpResponseMessage response = client.GetAsync(
                                url).Result)
                    {
                        DataSet dataSet = new DataSet();

                        response.EnsureSuccessStatusCode();
                        string responseBody = await response.Content.ReadAsStringAsync();

                        this.TestCaseJASON = responseBody;

                        Rootobject ro = JsonConvert.DeserializeObject<Rootobject>(responseBody);

                        DataTable dataTable = new DataTable();

                        System.IO.StringReader xmlSR = new System.IO.StringReader(ro.value[0].fields.MicrosoftVSTSTCMLocalDataSource.ToString());
                        dataSet.ReadXml(xmlSR, XmlReadMode.Auto);

                        this.Params = dataSet;
                        return dataSet;
                    }
                }
            }
            catch (Exception ex)
            {
                return null;
                throw ex;
            }
        }

        public async Task<List<string>> GetTestCasesByQuery(string workitemQuery)
        {
            try
            {
                string url = this.VstsURI + "/" + Project + "/_apis/wit/wiql?api-version=1.0";
                using (HttpClient client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Accept.Add(
                        new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));

                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic",
                        Convert.ToBase64String(
                            System.Text.ASCIIEncoding.ASCII.GetBytes(
                                string.Format("{0}:{1}", "", this.Pat))));

                    client.BaseAddress = new Uri(url);

                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, new Uri(url));
                    request.Content = new StringContent(workitemQuery, Encoding.UTF8, "application/json");

                    var response = await client.SendAsync(request);

                    response.EnsureSuccessStatusCode();

                    string r = await response.Content.ReadAsStringAsync();

                    JObject jr = JObject.Parse(r);

                    var i = from p in jr["workItems"] select (string)p["id"];

                    List<string> ids = i.ToList();

                    return ids;
                }
            }
            catch (Exception ex)
            {
                return null;
                throw ex;
            }
        }

    }


    public class DynamicJsonObject : DynamicObject
    {
        private IDictionary<string, object> Dictionary { get; set; }

        public DynamicJsonObject(IDictionary<string, object> dictionary)
        {
            this.Dictionary = dictionary;
        }

        public override bool TryGetMember(GetMemberBinder binder, out object result)
        {
            result = this.Dictionary[binder.Name];

            if (result is IDictionary<string, object>)
            {
                result = new DynamicJsonObject(result as IDictionary<string, object>);
            }
            else if (result is ArrayList && (result as ArrayList) is IDictionary<string, object>)
            {
                result = new List<DynamicJsonObject>((result as ArrayList).ToArray().Select(x => new DynamicJsonObject(x as IDictionary<string, object>)));
            }
            else if (result is ArrayList)
            {
                result = new List<object>((result as ArrayList).ToArray());
            }

            return this.Dictionary.ContainsKey(binder.Name);
        }
    }


    public class DynamicJsonConverter : JavaScriptConverter
    {
        public override object Deserialize(IDictionary<string, object> dictionary, Type type, JavaScriptSerializer serializer)
        {
            if (dictionary == null)
                throw new ArgumentNullException("dictionary");

            if (type == typeof(object))
            {
                return new DynamicJsonObject(dictionary);
            }

            return null;
        }

        public override IDictionary<string, object> Serialize(object obj, JavaScriptSerializer serializer)
        {
            throw new NotImplementedException();
        }

        public override IEnumerable<Type> SupportedTypes
        {
            get { return new ReadOnlyCollection<Type>(new List<Type>(new Type[] { typeof(object) })); }
        }
    }


}
