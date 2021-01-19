using ClosedXML.Excel;
using Microsoft.Win32;
using Newtonsoft.Json;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ScreenerIn_Automation
{
    public class BusinessExecution
    {
        IWebDriver driver;
        HttpClient client = new HttpClient();
        List<string> ExcelCompany = new List<string>();
        XLWorkbook wb = new XLWorkbook();
        DataTable dt = new DataTable();
        WebDriverWait wait;
        public void Setup()
        {
            HttpClient client = new HttpClient();
            client.BaseAddress = new Uri("https://www.screener.in/api/company/search/");
            List<char> alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray().ToList<char>();
            Dictionary<string, CompanyModel> company = new Dictionary<string, CompanyModel>();

            dt.Columns.Add("Company Name");
            dt.Columns.Add("URL");

            foreach (char c in alpha)
            {
                List<CompanyModel> companyModels = Get_company(client, c.ToString());
                foreach (CompanyModel companyModel in companyModels)
                {
                    if (!company.ContainsKey(companyModel.name))
                    {
                        company.Add(companyModel.name, companyModel);
                        DataRow dat = dt.NewRow();
                        dat["Company Name"] = companyModel.name;
                        dat["URL"] = companyModel.url;
                        dt.Rows.Add(dat);
                    }

                }
                foreach (char n in alpha)
                {
                    companyModels = Get_company(client, c.ToString() + n.ToString()); 
                    foreach (CompanyModel companyModel in companyModels)
                    {
                        if (!company.ContainsKey(companyModel.name))
                        {
                            company.Add(companyModel.name, companyModel);
                            DataRow dat = dt.NewRow();
                            dat["Company Name"] = companyModel.name;
                            dat["URL"] = companyModel.url;
                            dt.Rows.Add(dat);
                        }
                    }
                    foreach (char d in alpha)
                    {
                        companyModels = Get_company(client, c.ToString() + n.ToString() + d.ToString()); ;
                        foreach (CompanyModel companyModel in companyModels)
                        {
                            if (!company.ContainsKey(companyModel.name))
                            {
                                company.Add(companyModel.name, companyModel);
                                DataRow dat = dt.NewRow();
                                dat["Company Name"] = companyModel.name;
                                dat["URL"] = companyModel.url;
                                dt.Rows.Add(dat);
                            }
                        }
                    }
                }
            }
            dt = new DataTable();
            dt.Columns.Add("Company Name");
            dt.Columns.Add("URL");
            foreach (var ele in company)
            {
                DataRow dat = dt.NewRow();
                dat["Company Name"] = ele.Key;
                dat["URL"] = ele.Value.url;
                dt.Rows.Add(dat);
            }
            XLWorkbook wb1 = new XLWorkbook();
            var worksheet = wb1.Worksheets.Add("All Company");
            worksheet.Cell(1, 1).InsertTable(dt);
            wb1.SaveAs($@"C:\Users\{Environment.UserName}\Downloads\Company{DateTime.Now.ToString("yyyyMMddHHmmss")}.xlsx");
            MessageBox.Show("Completed", "Completed");
        }

        public List<CompanyModel> Get_company(HttpClient client, string query)
        {

            HttpResponseMessage response = client.GetAsync($"?q={query}").Result;
            var customerJsonString = response.Content.ReadAsStringAsync();
            List<CompanyModel> deserialized = JsonConvert.DeserializeObject<List<CompanyModel>>(customerJsonString.Result);
            return deserialized;
        }

        public void InstalChromeDriver()
        {
            string path = (string)Registry.GetValue(@"HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe", "", null);
            string Chromeversion = FileVersionInfo.GetVersionInfo(path).FileVersion.ToString().Split('.')[0].Trim();
            string url = "https://chromedriver.storage.googleapis.com/LATEST_RELEASE_" + Chromeversion.Split('.')[0];
            WebClient webClient = new WebClient();
            string Chromever = webClient.DownloadString(url);
            string url_file = "https://chromedriver.storage.googleapis.com/" + Convert.ToString(Chromever) + "/" + "chromedriver_win32.zip";
            byte[] data = webClient.DownloadData(url_file);
            MemoryStream stream = new MemoryStream(data, true);
            ZipArchive archive = new ZipArchive(stream);
            foreach (ZipArchiveEntry entry in archive.Entries)
            {
                if (entry.FullName.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                {
                    entry.ExtractToFile($@"C:\Users\{Environment.UserName}\Desktop\" + entry.Name, true);
                }
            }
        }

        public List<string> MainExecution(string fileName)
        {
            Process.GetProcesses().Where(x => x.ProcessName.ToLower()
                         .StartsWith("chromedriver"))
                         .ToList()
                         .ForEach(x => x.Kill());
            ReadExcelFile(fileName);
            BrowerInit();
            driver.Navigate().GoToUrl("https://www.screener.in/login/");
            var js = (IJavaScriptExecutor)driver;
            js.ExecuteScript($"document.getElementById('id_username').value='{Properties.Settings.Default.UserName}'");
            js.ExecuteScript($"document.getElementById('id_password').value='{Properties.Settings.Default.Password}'");
            js.ExecuteScript("document.getElementsByClassName('button-primary')[0].click();");
            AddColumns();
            
            client.BaseAddress = new Uri("https://www.screener.in/api/company/search/");
            return ExcelCompany;
        }

        public void ExecuteQuerysteps(string company)
        {
            try
            {
                CompanyModel model = Get_company(client, company)[0];
                DataRow dataRow = dt.NewRow();
                dataRow["Requested Company"] = company;
                dataRow["Company"] = model.name;
                dt.Rows.Add(dataRow);
                GetbrowserData(model.url, model.name);
            }
            catch
            {
                DataRow dataRow = dt.NewRow();
                dataRow["Requested Company"] = company;
                dataRow["Company"] = "Company is not available in website";
                dt.Rows.Add(dataRow);
            }
        }

        public void Final()
        {
            driver.Quit();
            var worksheet = wb.Worksheets.Add("Result");
            worksheet.Cell(1, 1).InsertTable(dt);
            wb.SaveAs($@"C:\Users\{Environment.UserName}\Downloads\ScreenerResult{DateTime.Now.ToString("yyyyMMddHHmmss")}.xlsx");
            MessageBox.Show("Completed", "Completed");
        }

        public void BrowerInit()
        {
            if (!File.Exists($@"C:\Users\{Environment.UserName}\Desktop\chromedriver.exe"))
            {
                InstalChromeDriver();
            }
            var chromeOptions = new ChromeOptions();
            //chromeOptions.AddArguments(new List<string>() { "headless", "disable-gpu" });
            var chromeDriverService = ChromeDriverService.CreateDefaultService($@"C:\Users\{Environment.UserName}\Desktop");
            chromeDriverService.HideCommandPromptWindow = true;
            driver = new ChromeDriver(chromeDriverService, chromeOptions);
            driver.Manage().Window.Maximize();
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            //driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(30);
            ICapabilities capabilities = ((RemoteWebDriver)driver).Capabilities;
            string driverversion = Convert.ToString((capabilities.GetCapability("chrome") as Dictionary<string, object>)["chromedriverVersion"]).Split('(')[0].Trim();
            CheckChromeVersion(driverversion);
            

        }

        public void CheckChromeVersion(string driverversion)
        {
            driverversion = driverversion.Split('.')[0];
            string path = (string)Registry.GetValue(@"HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe", "", null);
            string Chromeversion = FileVersionInfo.GetVersionInfo(path).FileVersion.ToString().Split('.')[0].Trim();
            if (driverversion != Chromeversion)
            {
                InstalChromeDriver();
            }
        }

        public void ReadExcelFile(string fileName)
        {   
            using (var excelWorkbook = new XLWorkbook(fileName))
            {
                var nonEmptyDataRows = excelWorkbook.Worksheet(1).RowsUsed();
                ExcelCompany.AddRange(nonEmptyDataRows.Select(data => Convert.ToString(data.Cell(1).Value)).ToList());
            }
        }

        public void GetbrowserData(string URL,string company)
        {
            driver.Navigate().GoToUrl("https://www.screener.in" + URL);
            var js = (IJavaScriptExecutor)driver;
            try
            {
                string[] SectorIndus = Convert.ToString(js.ExecuteScript("return document.getElementById('peers').getElementsByClassName('sub')[0].innerText")).Split(new string[] { "//" }, StringSplitOptions.None);
                foreach (DataRow data in dt.Rows)
                {
                    if (data["Company"].ToString() == company)
                    {
                        data[SectorIndus[0].Split(':')[0].Trim()] = SectorIndus[0].Split(':')[1].Trim();
                        data[SectorIndus[1].Split(':')[0].Trim()] = SectorIndus[1].Split(':')[1].Trim();
                    }
                }
            }
            catch
            {

            }
            
            List<string> QueryList = AddQueryToList();
            foreach(string query in QueryList)
            {
                Getratiodetails(query,URL,company);
            }

        }

        public void Getratiodetails(string query,string URL,string company)
        {
            driver.Navigate().GoToUrl("https://www.screener.in/user/quick_ratios/?next=" + URL);
            var js = (IJavaScriptExecutor)driver;
            try
            {
                js.ExecuteScript($"document.getElementById('manage-list').value='{query}'");
                js.ExecuteScript("document.getElementsByClassName('button-primary')[1].click();");
                var ele = driver.Manage().Cookies;
                Thread.Sleep(1000);
                int Count = Convert.ToInt32(js.ExecuteScript("return document.getElementsByClassName('row-full-width')[2].getElementsByClassName('four columns').length"));
                for (int i = 0; i < Count; i++)
                {
                    string[] val = Convert.ToString(js.ExecuteScript($"return document.getElementsByClassName('row-full-width')[2].getElementsByClassName('four columns')[{i}].innerText")).Split(':');
                    foreach (DataRow data in dt.Rows)
                    {
                        if (data["Company"].ToString() == company)
                        {
                            data[val[0].Trim()] = val[1].Trim();
                            break;
                        }
                    }
                }
            }
            catch
            { }
            
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
            List<string> Query= new List<string>();
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
    }
}
