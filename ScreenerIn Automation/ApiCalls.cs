using ClosedXML.Excel;
using HtmlAgilityPack;
using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace ScreenerIn_Automation
{
    public class logindata
    {
        public string csrfmiddlewaretoken { get; set; }
        public string sessionid { get; set; }
        public string csrftoken { get; set; }
    }

    public class ApiCalls
    {
        List<string> ExcelCompany = new List<string>();
        XLWorkbook wb = new XLWorkbook();
        DataTable dt = new DataTable();
        string WebUrl = "https://www.screener.in";


        public void readExcelFile(string fileName,ExcelModel model)
        {
            using (var excelWorkbook = new XLWorkbook(fileName))
            {
                var nonEmptyDataRows = excelWorkbook.Worksheet(1).RowsUsed();
                model.comapny.AddRange(nonEmptyDataRows.Select(data => Convert.ToString(data.Cell(1).Value)).ToList());
                model.url.AddRange(nonEmptyDataRows.Select(data => Convert.ToString(data.Cell(2).Value)).ToList());
            }
            AddColumns();
        }

        public List<string> readcomapnyurl(string fileName)
        {
            List<string> ExcelCompany = new List<string>();
            using (var excelWorkbook = new XLWorkbook(fileName))
            {
                var nonEmptyDataRows = excelWorkbook.Worksheet(1).RowsUsed();
                ExcelCompany.AddRange(nonEmptyDataRows.Select(data => Convert.ToString(data.Cell(2).Value)).ToList());
                ExcelCompany.AddRange(nonEmptyDataRows.Select(data => Convert.ToString(data.Cell(2).Value)).ToList());
            }
            AddColumns();
            return ExcelCompany;
        }

        public void MainExecution(string Company,logindata logindata, string Companyurlmodel)
        {
            try
            {
                string company1 = null;
                if (Companyurlmodel == "")
                {
                    CompanyModel companyModel = GetCompanyUrl(Company);
                    if (companyModel != null)
                    {
                        company1 = companyModel.name;
                        Companyurlmodel = companyModel.url;
                    }
                    else
                        Companyurlmodel = null;
                }
                else
                {
                    company1 = Company;
                }
                if (Companyurlmodel != null)
                {
                    DataRow dataRow = dt.NewRow();
                    dataRow["Requested Company"] = Company;
                    dataRow["Company"] = company1;
                    dt.Rows.Add(dataRow);
                    string Companyurl = WebUrl + Companyurlmodel;
                    string getwarehouse = GetDataWareHouseID(logindata, Companyurl);
                    Dictionary<string, string> bsense = Getbsense(logindata, Companyurl);
                    Dictionary<string, string> SectorIndus = GetSectorIndustry(logindata, Companyurl);
                    string postqueryurl = WebUrl + "/user/quick_ratios/?next=" + Companyurlmodel;
                    string Getratios = $"https://www.screener.in/api/company/{getwarehouse}/quick_ratios/";
                    List<string> QueryList = AddQueryToList();
                    foreach (DataRow dtRow in dt.Rows)
                    {
                        if (company1 == dtRow["Company"].ToString())
                        {
                            foreach (string key in bsense.Keys)
                            {
                                string key1 = null;
                                if (key.Contains("BSE"))
                                    key1 = "BSE";
                                else if (key.Contains("NSE"))
                                    key1 = "NSE";
                                dtRow[key1] = bsense[key].ToString();
                            }
                            break;
                        }
                    }
                    foreach (DataRow dtRow in dt.Rows)
                    {
                        if (company1 == dtRow["Company"].ToString())
                        {
                            foreach (string key in SectorIndus.Keys)
                            {
                                string newkey = key.Replace(":", "");
                                dtRow[newkey] = SectorIndus[key].ToString();
                            }
                            break;
                        }
                    }
                    foreach (string query in QueryList)
                    {
                        PostQuery(logindata, query, postqueryurl);
                        Dictionary<string, string> result = new Dictionary<string, string>();
                        result = Getdata(logindata, Getratios);
                        foreach (DataRow dtRow in dt.Rows)
                        {
                            if (company1 == dtRow["Company"].ToString())
                            {
                                foreach (string key in result.Keys)
                                {
                                    string newkey = key.Replace(":", "");
                                    if (string.IsNullOrEmpty(result[key].ToString()))
                                        dtRow[newkey] = null;
                                    else
                                        dtRow[newkey] = result[key].ToString();
                                }
                                break;
                            }
                        }
                    }
                }
            }
            catch
            {

            }
            
            
        }

        public logindata login(string username,string password)
        {
            logindata logindata = new logindata();
            string loginurl = WebUrl + "/login/";
            RestClient client = new RestClient(loginurl);
            RestRequest request = new RestRequest(loginurl, Method.GET);
            IRestResponse Adminresult = client.Execute(request);
            if (Adminresult.Content.Contains("csrfmiddlewaretoken") && Adminresult.Content.Contains(">"))
            {
                int Start, End;
                Start = Adminresult.Content.IndexOf("csrfmiddlewaretoken", 0) + "csrfmiddlewaretoken".Length;
                End = Adminresult.Content.IndexOf(">", Start);
                logindata.csrfmiddlewaretoken = Adminresult.Content.Substring(Start, End - Start).Split('\"')[2];
            }
            logindata.csrftoken = Adminresult.Cookies[0].Value;
            request = new RestRequest(loginurl, Method.POST);
            request.AddHeader("Content-Type", "application/x-www-form-urlencoded");
            request.AddParameter("csrfmiddlewaretoken", logindata.csrfmiddlewaretoken);
            request.AddParameter("username", username);
            request.AddParameter("password", password);
            request.AddCookie("csrftoken", logindata.csrftoken);
            request.AddHeader("Referer", loginurl);
            client.FollowRedirects = false;
            IRestResponse result = client.Execute(request);
            logindata.sessionid = result.Cookies[1].Value;
            return logindata;
        }

        public string GetDataWareHouseID(logindata logindata,string url)
        {
            string datawarehouseid = null;
            RestClient client = new RestClient(url);
            RestRequest request = new RestRequest(url, Method.GET);
            request.AddCookie("csrftoken", logindata.csrftoken);
            request.AddCookie("sessionid", logindata.sessionid);
            IRestResponse datawarehouseresult = client.Execute(request);
            if (datawarehouseresult.Content.Contains("data-warehouse-id") && datawarehouseresult.Content.Contains("data"))
            {
                int Start, End;
                Start = datawarehouseresult.Content.IndexOf("data-warehouse-id", 0) + "data-warehouse-id".Length;
                End = datawarehouseresult.Content.IndexOf("data", Start);
                datawarehouseid = datawarehouseresult.Content.Substring(Start, End - Start).Split('\"')[1];
                
            }
            return datawarehouseid;
        }

        public Dictionary<string,string> Getbsense(logindata logindata, string url)
        {
            Dictionary<string, string> result = new Dictionary<string, string>();
            RestClient client = new RestClient(url);
            RestRequest request = new RestRequest(url, Method.GET);
            request.AddCookie("csrftoken", logindata.csrftoken);
            request.AddCookie("sessionid", logindata.sessionid);
            IRestResponse response = client.Execute(request);
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(response.Content);
            if (doc.DocumentNode.SelectSingleNode("//*[@id='top']/div[2]/a[2]/span") != null)
            {
                string[] bse = doc.DocumentNode.SelectSingleNode("//*[@id='top']/div[2]/a[2]/span").InnerText.Split(':');
                result.Add(bse[0].Trim(), bse[1].Trim());
            }
            if (doc.DocumentNode.SelectSingleNode("//*[@id='top']/div[2]/a[3]/span") != null)
            {
                string[] nse = doc.DocumentNode.SelectSingleNode("//*[@id='top']/div[2]/a[3]/span").InnerText.Split(':');
                result.Add(nse[0].Trim(), nse[1].Trim());
            }
            return result;
        }

        public Dictionary<string,string> GetSectorIndustry(logindata logindata, string url)
        {
            Dictionary<string, string> result = new Dictionary<string, string>();
            RestClient client = new RestClient(url);
            RestRequest request = new RestRequest(url, Method.GET);
            request.AddCookie("csrftoken", logindata.csrftoken);
            request.AddCookie("sessionid", logindata.sessionid);
            IRestResponse response = client.Execute(request);
            string paresedvalue = null;
            if (response.Content.Contains("Peer comparison") && response.Content.Contains("</p>") && response.Content.Contains("Sector:"))
            {
                int Start, End;
                Start = response.Content.IndexOf("Peer comparison", 0) + "Peer comparison".Length;
                End = response.Content.IndexOf("</p>", Start);
                paresedvalue = response.Content.Substring(Start, End - Start);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(paresedvalue);
                StringBuilder stringbuild = new StringBuilder();
                foreach (HtmlTextNode node in doc.DocumentNode.SelectNodes("//text()"))
                {
                    if (!node.Text.Contains("\r"))
                        stringbuild.AppendLine(node.Text);
                }
                string parsedstring = stringbuild.ToString();
                string[] removespace = parsedstring.Split('\n');
                string[] finalresult = removespace.Where(x => !string.IsNullOrEmpty(x.Trim())).ToArray();
                foreach (string str in finalresult)
                {
                    if (str.Contains(":"))
                        result.Add(str.Trim(), "");
                    else if(!str.Contains("//"))
                        result[result.Keys.Last()] = result[result.Keys.Last()] + str.Trim();
                }
            }
            return result;
        }

        public void PostQuery(logindata logindata,string Query,string url)
        {
            RestClient client = new RestClient(url);
            RestRequest request = new RestRequest(url, Method.POST);
            request.AddHeader("Content-Type", "application/x-www-form-urlencoded");
            request.AddParameter("csrfmiddlewaretoken", logindata.csrfmiddlewaretoken, ParameterType.GetOrPost);
            request.AddParameter("data", Query, ParameterType.GetOrPost);
            request.AddHeader("Referer", url);
            request.AddCookie("csrftoken", logindata.csrftoken);
            request.AddCookie("sessionid", logindata.sessionid);
            client.Execute(request);
        }

        public Dictionary<string, string> Getdata(logindata logindata,string url)
        {
            char[] whitespace = new char[] { ' ', '\t' };
            Dictionary<string, string> output = new Dictionary<string, string>();
            RestClient client = new RestClient(url);
            RestRequest request = new RestRequest(url, Method.GET);
            request.AddCookie("csrftoken", logindata.csrftoken);
            request.AddCookie("sessionid", logindata.sessionid);
            IRestResponse result = client.Execute(request);
            if(!result.StatusDescription.Contains("Not Found"))
            {
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(result.Content);
                StringBuilder stringbuild = new StringBuilder();
                foreach (HtmlNode node in doc.DocumentNode.SelectNodes("//li[@class='flex flex-space-between']"))
                {
                    string keyratio = node.SelectNodes("span[@class='name']")[0].InnerText.Trim();
                    List<string> parsedvalue = node.SelectNodes("span[@class='nowrap value']")[0].InnerText.Trim().Split('\n').ToList();
                    List<string> resii1 = new List<string>();
                    string new1val = null;
                    foreach (string ele in parsedvalue)
                    {
                        if (!string.IsNullOrEmpty(ele.Trim()))
                        {
                            resii1.Add(ele.TrimStart());
                            new1val = new1val + ele.TrimStart();
                        }
                    }
                    string keyratiovalue = Mapratio(keyratio) + ":";
                    if (new1val == null)
                        new1val = "";
                    output.Add(keyratiovalue, new1val);
                }
            }
            return output;
        }

        public CompanyModel GetCompanyUrl(string Company)
        {
            HttpClient client = new HttpClient();
            client.BaseAddress = new Uri("https://www.screener.in/api/company/search/");
            HttpResponseMessage response = client.GetAsync($"?q={Company}").Result;
            var customerJsonString = response.Content.ReadAsStringAsync();
            if (customerJsonString.Result != "[]")
            {
                List<CompanyModel> deserialized = JsonConvert.DeserializeObject<List<CompanyModel>>(customerJsonString.Result);
                CompanyModel model = deserialized[0];
                return model;
            }
            else
                return null;
        }

        public void Final()
        {
            var worksheet = wb.Worksheets.Add("Result");
            worksheet.Cell(1, 1).InsertTable(dt);
            wb.SaveAs($@"C:\Users\{Environment.UserName}\Downloads\ScreenerResult{DateTime.Now.ToString("yyyyMMddHHmmss")}.xlsx");
            
        }

        public List<string> AddQueryToList()
        {
            string Query1 = "Asset Turnover Ratio,Average dividend payout 3years,Average return on capital employed 3Years,Average return on equity 3Years,Book value,Capital work in progress,Cash Equivalents,Cash from financing last year,Cash from financing preceding year,Cash from investing last year,Cash from investing preceding year,Cash from operations last year,Cash from operations preceding year,Change in promoter holding,Change in promoter holding 3Years,Current assets,Current liabilities,Current price";
            string Query2 = "Debt,Debt preceding year,Debt to equity,Depreciation,Depreciation last year,Depreciation latest quarter,Depreciation preceding year,Dividend last year,Dividend,preceding year,Dividend yield,EBIDT,EBIDT last year,EBIDT latest quarter,EBIDT preceding quarter,EBIDT preceding year quarter,EBIT,EBIT latest quarter,EBIT preceding quarter";
            string Query3 = "EBIT preceding year,EBIT preceding year quarter,Employee cost last year,Enterprise Value,EPS,EPS last year,EPS latest quarter,EPS preceding quarter,EPS preceding year,EPS preceding year quarter,Equity capital,Exports percentage,Extraordinary items latest quarter,Face value,Financial leverage,Free cash flow 3years,Free,cash flow last year,Free cash flow preceding year";
            string Query4 = "Industry PBV,Industry PE,Interest,Interest latest quarter,Interest preceding year,Inventory turnover ratio,Investing cash flow 3years,Investments,Market value of quoted investments,Net cash flow last year,Net cash flow preceding year,Net profit,Net profit 2quarters back,Net profit 3quarters back,Net Profit last year,Net Profit latest quarter,Net Profit preceding quarter,Net Profit preceding year";
            string Query5 = "Net Profit preceding year quarter,NPM last year,NPM latest quarter,NPM preceding quarter,NPM preceding year,NPM preceding year quarter,Operating cash flow 3years,Operating profit,Operating profit last year,Operating profit latest quarter,Operating profit preceding quarter,Operating profit preceding year,OPM,OPM 5Year,OPM last year,OPM latest quarter,OPM preceding quarter,OPM preceding year";
            string Query6 = "OPM preceding year quarter,Other income,Other income last year,Other income latest quarter,Pledged percentage,Price to book value,Price to Earning,Profit after tax,Profit after tax last year,Profit after tax latest quarter,Profit after tax preceding quarter,Profit after tax preceding year,Profit after tax preceding year quarter,Profit before tax last year,Profit before tax latest quarter,Profit before tax preceding year,Profit before tax preceding year quarter,Profit growth";
            string Query7 = "Promoter holding,Quick ratio,Reserves,Return on assets,Return on capital employed,Return on equity,Return on equity preceding year,Sales,Sales growth,Sales growth 3Years,Sales last year,Sales latest quarter,Sales preceding quarter,Sales preceding year,Sales preceding year quarter,Secured loan,Tax,Tax latest quarter";
            string Query8 = "Total Assets,Trade Payables,Trade receivables,Unpledged promoter holding,Unsecured loan,Working capital,Working capital preceding year,YOY Quarterly profit growth,YOY Quarterly sales growth";
            List<string> Query = new List<string>();
            Query.Add(Query1);
            Query.Add(Query2);
            Query.Add(Query3);
            Query.Add(Query4);
            Query.Add(Query5);
            Query.Add(Query6);
            Query.Add(Query7);
            Query.Add(Query8);
            return Query;
        }

        public void AddColumns()
        {
            dt.Columns.Add("Requested Company");
            dt.Columns.Add("Company");
            dt.Columns.Add("BSE");
            dt.Columns.Add("NSE");
            dt.Columns.Add("Sector");
            dt.Columns.Add("Industry");
            dt.Columns.Add("Asset Turnover Ratio");
            dt.Columns.Add("Average dividend payout 3years");
            dt.Columns.Add("Average return on capital employed 3Years");
            dt.Columns.Add("Average return on equity 3Years");
            dt.Columns.Add("Book value");
            dt.Columns.Add("Capital work in progress");
            dt.Columns.Add("Cash Equivalents");
            dt.Columns.Add("Cash from financing last year");
            dt.Columns.Add("Cash from financing preceding year");
            dt.Columns.Add("Cash from investing last year");
            dt.Columns.Add("Cash from investing preceding year");
            dt.Columns.Add("Cash from operations last year");
            dt.Columns.Add("Cash from operations preceding year");
            dt.Columns.Add("Change in promoter holding");
            dt.Columns.Add("Change in promoter holding 3Years");
            dt.Columns.Add("Current assets");
            dt.Columns.Add("Current liabilities");
            dt.Columns.Add("Current price");
            dt.Columns.Add("Debt");
            dt.Columns.Add("Debt preceding year");
            dt.Columns.Add("Debt to equity");
            dt.Columns.Add("Depreciation");
            dt.Columns.Add("Depreciation last year");
            dt.Columns.Add("Depreciation latest quarter");
            dt.Columns.Add("Depreciation preceding year");
            dt.Columns.Add("Dividend last year");
            dt.Columns.Add("Dividend preceding year");
            dt.Columns.Add("Dividend yield");
            dt.Columns.Add("EBIDT");
            dt.Columns.Add("EBIDT last year");
            dt.Columns.Add("EBIDT latest quarter");
            dt.Columns.Add("EBIDT preceding quarter");
            dt.Columns.Add("EBIDT preceding year quarter");
            dt.Columns.Add("EBIT");
            dt.Columns.Add("EBIT latest quarter");
            dt.Columns.Add("EBIT preceding quarter");
            dt.Columns.Add("EBIT preceding year");
            dt.Columns.Add("EBIT preceding year quarter");
            dt.Columns.Add("Employee cost last year");
            dt.Columns.Add("Enterprise Value");
            dt.Columns.Add("EPS");
            dt.Columns.Add("EPS last year");
            dt.Columns.Add("EPS latest quarter");
            dt.Columns.Add("EPS preceding quarter");
            dt.Columns.Add("EPS preceding year");
            dt.Columns.Add("EPS preceding year quarter");
            dt.Columns.Add("Equity capital");
            dt.Columns.Add("Exports percentage");
            dt.Columns.Add("Extraordinary items latest quarter");
            dt.Columns.Add("Face value");
            dt.Columns.Add("Financial leverage");
            dt.Columns.Add("Free cash flow 3years");
            dt.Columns.Add("Free cash flow last year");
            dt.Columns.Add("Free cash flow preceding year");
            dt.Columns.Add("Industry PBV");
            dt.Columns.Add("Industry PE");
            dt.Columns.Add("Interest");
            dt.Columns.Add("Interest latest quarter");
            dt.Columns.Add("Interest preceding year");
            dt.Columns.Add("Inventory turnover ratio");
            dt.Columns.Add("Investing cash flow 3years");
            dt.Columns.Add("Investments");
            dt.Columns.Add("Market value of quoted investments");
            dt.Columns.Add("Net cash flow last year");
            dt.Columns.Add("Net cash flow preceding year");
            dt.Columns.Add("Net profit");
            dt.Columns.Add("Net profit 2quarters back");
            dt.Columns.Add("Net profit 3quarters back");
            dt.Columns.Add("Net Profit last year");
            dt.Columns.Add("Net Profit latest quarter");
            dt.Columns.Add("Net Profit preceding quarter");
            dt.Columns.Add("Net Profit preceding year");
            dt.Columns.Add("Net Profit preceding year quarter");
            dt.Columns.Add("NPM last year");
            dt.Columns.Add("NPM latest quarter");
            dt.Columns.Add("NPM preceding quarter");
            dt.Columns.Add("NPM preceding year");
            dt.Columns.Add("NPM preceding year quarter");
            dt.Columns.Add("Operating cash flow 3years");
            dt.Columns.Add("Operating profit");
            dt.Columns.Add("Operating profit last year");
            dt.Columns.Add("Operating profit latest quarter");
            dt.Columns.Add("Operating profit preceding quarter");
            dt.Columns.Add("Operating profit preceding year");
            dt.Columns.Add("OPM");
            dt.Columns.Add("OPM 5Year");
            dt.Columns.Add("OPM last year");
            dt.Columns.Add("OPM latest quarter");
            dt.Columns.Add("OPM preceding quarter");
            dt.Columns.Add("OPM preceding year");
            dt.Columns.Add("OPM preceding year quarter");
            dt.Columns.Add("Other income");
            dt.Columns.Add("Other income last year");
            dt.Columns.Add("Other income latest quarter");
            dt.Columns.Add("Pledged percentage");
            dt.Columns.Add("Price to book value");
            dt.Columns.Add("Price to Earning");
            dt.Columns.Add("Profit after tax");
            dt.Columns.Add("Profit after tax last year");
            dt.Columns.Add("Profit after tax latest quarter");
            dt.Columns.Add("Profit after tax preceding quarter");
            dt.Columns.Add("Profit after tax preceding year");
            dt.Columns.Add("Profit after tax preceding year quarter");
            dt.Columns.Add("Profit before tax last year");
            dt.Columns.Add("Profit before tax latest quarter");
            dt.Columns.Add("Profit before tax preceding year");
            dt.Columns.Add("Profit before tax preceding year quarter");
            dt.Columns.Add("Profit growth");
            dt.Columns.Add("Promoter holding");
            dt.Columns.Add("Quick ratio");
            dt.Columns.Add("Reserves");
            dt.Columns.Add("Return on assets");
            dt.Columns.Add("Return on capital employed");
            dt.Columns.Add("Return on equity");
            dt.Columns.Add("Return on equity preceding year");
            dt.Columns.Add("Sales");
            dt.Columns.Add("Sales growth");
            dt.Columns.Add("Sales growth 3Years");
            dt.Columns.Add("Sales last year");
            dt.Columns.Add("Sales latest quarter");
            dt.Columns.Add("Sales preceding quarter");
            dt.Columns.Add("Sales preceding year");
            dt.Columns.Add("Sales preceding year quarter");
            dt.Columns.Add("Secured loan");
            dt.Columns.Add("Tax");
            dt.Columns.Add("Tax latest quarter");
            dt.Columns.Add("Total Assets");
            dt.Columns.Add("Trade Payables");
            dt.Columns.Add("Trade receivables");
            dt.Columns.Add("Unpledged promoter holding");
            dt.Columns.Add("Unsecured loan");
            dt.Columns.Add("Working capital");
            dt.Columns.Add("Working capital preceding year");
            dt.Columns.Add("YOY Quarterly profit growth");
            dt.Columns.Add("YOY Quarterly sales growth");

        }
        public string Mapratio(string keyratio)
        {
            string result = "";
            switch (keyratio)
            {
                case "Asset Turnover":
                    result = "Asset Turnover Ratio";
                    break;
                case "Avg Div Payout 3Yrs":
                    result = "Average dividend payout 3years";
                    break;
                case "ROCE 3Yr":
                    result = "Average return on capital employed 3Years";
                    break;
                case "ROE 3Yr":
                    result = "Average return on equity 3Years";
                    break;
                case "Book value":
                    result = "Book value";
                    break;
                case "CWIP":
                    result = "Capital work in progress";
                    break;
                case "Cash Equivalents":
                    result = "Cash Equivalents";
                    break;
                //case "CF Financing":
                //    result = "Average return on equity 3Years";
                //    break;
                case "CF Financing":
                    result = "Cash from financing last year";
                    break;
                //case "header":
                //    result = "Average return on equity 3Years";
                //    break;
                case "CF Financing PY":
                    result = "Cash from financing preceding year";
                    break;
                case "CF Investing":
                    result = "Cash from investing last year";
                    break;
                case "CF Investing PY":
                    result = "Cash from investing preceding year";
                    break;
                case "CF Operations":
                    result = "Cash from operations last year";
                    break;
                case "CF Operations PY":
                    result = "Cash from operations preceding year";
                    break;
                case "Change in Prom Hold":
                    result = "Change in promoter holding";
                    break;
                case "Chg in Prom Hold 3Yr":
                    result = "Change in promoter holding 3Years";
                    break;
                case "Current assets":
                    result = "Current assets";
                    break;
                case "Current liabilities":
                    result = "Current liabilities";
                    break;
                case "Current Price":
                    result = "Current price";
                    break;
                case "Debt":
                    result = "Debt";
                    break;
                case "Debt preceding year":
                    result = "Debt preceding year";
                    break;
                case "Debt to equity":
                    result = "Debt to equity";
                    break;
                case "Depreciation":
                    result = "Depreciation";
                    break;
                case "Dep Ann":
                    result = "Depreciation last year";
                    break;
                case "Dep Qtr":
                    result = "Depreciation latest quarter";
                    break;
                case "Dep Prev Ann":
                    result = "Depreciation preceding year";
                    break;
                case "Dividend last year":
                    result = "Dividend last year";
                    break;
                case "Dividend yield":
                    result = "Dividend yield";
                    break;
                //case "EBIDT last year":
                //    result = "EBIDT";
                //    break;
                case "EBIDT last year":
                    result = "EBIDT last year";
                    break;
                case "EBIDT Qtr":
                    result = "EBIDT latest quarter";
                    break;
                case "EBIDT Prev Qtr":
                    result = "EBIDT preceding quarter";
                    break;
                case "EBIDT PY Qtr":
                    result = "EBIDT preceding year quarter";
                    break;
                case "EBIT":
                    result = "EBIT";
                    break;
                case "EBIT latest quarter":
                    result = "EBIT latest quarter";
                    break;
                case "EBIT Prev Qtr":
                    result = "EBIT preceding quarter";
                    break;
                case "EBIT preceding year":
                    result = "EBIT preceding year";
                    break;
                case "EBIT PY Qtr":
                    result = "EBIT preceding year quarter";
                    break;
                case "Employee cost":
                    result = "Employee cost last year";
                    break;
                case "Enterprise Value":
                    result = "Enterprise Value";
                    break;
                case "EPS":
                    result = "EPS";
                    break;
                case "EPS last year":
                    result = "EPS last year";
                    break;
                case "EPS latest quarter":
                    result = "EPS latest quarter";
                    break;
                case "EPS Prev Qtr":
                    result = "EPS preceding quarter";
                    break;
                case "EPS preceding year":
                    result = "EPS preceding year";
                    break;
                case "EPS PY Qtr":
                    result = "EPS preceding year quarter";
                    break;
                case "Equity capital":
                    result = "Equity capital";
                    break;
                case "Exports percentage":
                    result = "Exports percentage";
                    break;
                case "Extra Ord Item Qtr":
                    result = "Extraordinary items latest quarter";
                    break;
                case "Face value":
                    result = "Face value";
                    break;
                case "Financial leverage":
                    result = "Financial leverage";
                    break;
                case "Free Cash Flow 3Yrs":
                    result = "Free cash flow 3years";
                    break;
                case "Free Cash Flow":
                    result = "Free cash flow last year";
                    break;
                case "FCF Prev Ann":
                    result = "Free cash flow preceding year";
                    break;
                case "Industry PBV":
                    result = "Industry PBV";
                    break;
                case "Industry PE":
                    result = "Industry PE";
                    break;
                case "Interest":
                    result = "Interest";
                    break;
                case "Interest Qtr":
                    result = "Interest latest quarter";
                    break;
                case "Interest Prev Ann":
                    result = "Interest preceding year";
                    break;
                case "Inven TO":
                    result = "Inventory turnover ratio";
                    break;
                case "CF Inv 3Yrs":
                    result = "Investing cash flow 3years";
                    break;
                case "Investments":
                    result = "Investments";
                    break;
                case "MV Quoted Inv":
                    result = "Market value of quoted investments";
                    break;
                case "Net CF":
                    result = "Net cash flow last year";
                    break;
                case "Net CF PY":
                    result = "Net cash flow preceding year";
                    break;
                case "Net profit":
                    result = "Net profit";
                    break;
                //case "NP 2Qtr Bk":
                //    result = "Depreciation preceding year";
                //    break;
                case "NP 2Qtr Bk":
                    result = "Net profit 2quarters back";
                    break;
                case "NP 3Qtr Bk":
                    result = "Net profit 3quarters back";
                    break;
                case "NP Ann":
                    result = "Net Profit last year";
                    break;
                case "NP Qtr":
                    result = "Net Profit latest quarter";
                    break;
                case "NP Prev Qtr":
                    result = "Net Profit preceding quarter";
                    break;
                case "NP Prev Ann":
                    result = "Net Profit preceding year";
                    break;
                case "NP PY Qtr":
                    result = "Net Profit preceding year quarter";
                    break;
                case "NPM last year":
                    result = "NPM last year";
                    break;
                case "NPM latest quarter":
                    result = "NPM latest quarter";
                    break;
                case "NPM Prev Qtr":
                    result = "NPM preceding quarter";
                    break;
                case "NPM preceding year":
                    result = "NPM preceding year";
                    break;
                case "NPM PY Qtr":
                    result = "NPM preceding year quarter";
                    break;
                case "CF Opr 3Yrs":
                    result = "Operating cash flow 3years";
                    break;
                case "Operating profit":
                    result = "Operating profit";
                    break;
                case "OP Ann":
                    result = "Operating profit last year";
                    break;
                case "OP Qtr":
                    result = "Operating profit latest quarter";
                    break;
                case "OP Prev Qtr":
                    result = "Operating profit preceding quarter";
                    break;
                case "OP Prev Ann":
                    result = "Operating profit preceding year";
                    break;
                case "OPM":
                    result = "OPM";
                    break;
                case "OPM 5Year":
                    result = "OPM 5Year";
                    break;
                case "OPM last year":
                    result = "OPM last year";
                    break;
                case "OPM latest quarter":
                    result = "OPM latest quarter";
                    break;
                case "OPM Prev Qtr":
                    result = "OPM preceding quarter";
                    break;
                case "OPM preceding year":
                    result = "OPM preceding year";
                    break;
                case "OPM PY Qtr":
                    result = "OPM preceding year quarter";
                    break;
                case "Other income":
                    result = "Other income";
                    break;
                case "Other Inc Ann":
                    result = "Other income last year";
                    break;
                case "Other Inc Qtr":
                    result = "Other income latest quarter";
                    break;
                case "Pledged percentage":
                    result = "Pledged percentage";
                    break;
                case "Price to book value":
                    result = "Price to book value";
                    break;
                case "Price to Earning":
                    result = "Price to Earning";
                    break;
                case "Profit after tax":
                    result = "Profit after tax";
                    break;
                case "PAT Ann":
                    result = "Profit after tax last year";
                    break;
                case "PAT Qtr":
                    result = "Profit after tax latest quarter";
                    break;
                case "PAT Prev Qtr":
                    result = "Profit after tax preceding quarter";
                    break;
                case "PAT Prev Ann":
                    result = "Profit after tax preceding year";
                    break;
                case "PAT PY Qtr":
                    result = "Profit after tax preceding year quarter";
                    break;
                case "PBT Ann":
                    result = "Profit before tax last year";
                    break;
                case "PBT Qtr":
                    result = "Profit before tax latest quarter";
                    break;
                case "PBT Prev Ann":
                    result = "Profit before tax preceding year";
                    break;
                case "PBT PY Qtr":
                    result = "Profit before tax preceding year quarter";
                    break;
                case "Profit growth":
                    result = "Profit growth";
                    break;
                case "Promoter holding":
                    result = "Promoter holding";
                    break;
                case "Quick ratio":
                    result = "Quick ratio";
                    break;
                case "Reserves":
                    result = "Reserves";
                    break;
                case "Return on assets":
                    result = "Return on assets";
                    break;
                case "ROCE":
                    result = "Return on capital employed";
                    break;
                case "Return on equity":
                    result = "Return on equity";
                    break;
                case "ROE Prev Ann":
                    result = "Return on equity preceding year";
                    break;
                case "Sales":
                    result = "Sales";
                    break;
                case "Sales growth":
                    result = "Sales growth";
                    break;
                case "Sales growth 3Years":
                    result = "Sales growth 3Years";
                    break;
                case "Sales last year":
                    result = "Sales last year";
                    break;
                case "Sales Qtr":
                    result = "Sales latest quarter";
                    break;
                case "Sales Prev Qtr":
                    result = "Sales preceding quarter";
                    break;
                case "Sales Prev Ann":
                    result = "Sales preceding year";
                    break;
                case "Sales PY Qtr":
                    result = "Sales preceding year quarter";
                    break;
                case "Secured loan":
                    result = "Secured loan";
                    break;
                case "Tax":
                    result = "Tax";
                    break;
                case "Tax latest quarter":
                    result = "Tax latest quarter";
                    break;
                case "Total Assets":
                    result = "Total Assets";
                    break;
                case "Trade Payables":
                    result = "Trade Payables";
                    break;
                case "Trade receivables":
                    result = "Trade receivables";
                    break;
                case "Unpledged Prom Hold":
                    result = "Unpledged promoter holding";
                    break;
                case "Unsecured loan":
                    result = "Unsecured loan";
                    break;
                case "Working capital":
                    result = "Working capital";
                    break;
                case "Work Cap PY":
                    result = "Working capital preceding year";
                    break;
                case "Qtr Profit Var":
                    result = "YOY Quarterly profit growth";
                    break;
                case "Qtr Sales Var":
                    result = "YOY Quarterly sales growth";
                    break;
            }
            return result; 
        }
    }

}
