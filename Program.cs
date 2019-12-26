using System;
using System.Collections.Generic;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Chrome;
using System.IO;


namespace ConsoleApp4
{
    class Program
    {
        static void Main(string[] args)
        {
            string webAddress = "https://direct.dimensiondata.com/seller/direct/login.asp?";

            string sold_to_name = args[1];
            string sold_to = args[2];
            string bill_to = args[3];
            string ship_to = args[4];
            string bill_to_country = args[5];
            string ship_to_country = args[6];
            string bill_plan = args[7];
            short months = Convert.ToInt16(args[8]);
            string start_date= args[9];
            string end_date = args[10];
            string solution_type = args[11];
            string quote_number = args[12];
            string contract = args[13];
            string client_manager = args[14];
            string sales_office = args[15];
            string cust_po = args[16];
            string vend_po = args[17];
            string woid = args[18];
            string sub_id = args[19];
            string cust_po_path = args[20];
            double cost = Convert.ToDouble(args[21]);
            double sell_price = Convert.ToDouble(args[22]);

            ChromeDriver driver = Driver();
            
            driver.Navigate().GoToUrl(webAddress);

            create_direct_quote(driver,sold_to,bill_to,ship_to, bill_to_country,ship_to_country, bill_plan,months,start_date,end_date,solution_type,sub_id,cost,sell_price);
            convert_to_po(driver, sold_to_name, contract, sales_office,cust_po, vend_po, woid, cust_po_path,sub_id,client_manager, quote_number);
            driver.Quit();
        }
       

        private static void create_direct_quote(ChromeDriver driver, string sold_to, string bill_to, string ship_to, string bill_to_country, string ship_to_country, string bill_plan, short months, string start_date, string end_date,string solution_type, string sub_id,double cost, double sell_price)
        {
            select_soldto(driver, sold_to, 1);

            cancel_quote(driver);

            select_soldto(driver, sold_to, 2);

            write_csv(bill_plan, start_date, end_date, months, cost, sell_price);

            catalog(driver);

            if (bill_plan != "Prepaid")
            {
                select_bill_freq(driver, bill_plan);
            }

            select_tab(driver, "address");

            select_billto(driver, bill_to,bill_to_country);

            finish_loading(driver, "//*[@id='ajwmCnt']");

            select_shipto(driver, ship_to,ship_to_country);

            update_quote(driver, "address", 3);

            select_tab(driver, "integration");

            select_solution_type(driver, solution_type);

            update_quote(driver, "integration", 3);

            

            submit_quote(driver,sub_id);

        }

        private static void select_tab(ChromeDriver driver, string locator)
        {
            do
            {
                try
                {
                    IWebElement element = driver.FindElement(By.Id(locator));
                    element.Click();
                    break;
                }
                catch
                {
                }

            }
            while (true) ;
        }

        private static void finish_loading(ChromeDriver driver, string locator)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath(locator)));
            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.InvisibilityOfElementLocated(By.XPath(locator)));
            System.Threading.Thread.Sleep(1000);
        }

        private static void update_quote(ChromeDriver driver, string id, short tr)
        {
            string locator = "//*[@id='" + id + "_cn']/table/tbody/tr[" + tr + "]/td/input";

            do
            {
                try
                {
                    IWebElement element = driver.FindElement(By.XPath(locator));
                    element.Click();
                    break;
                }
                catch
                {
                }

            }
            while (true);
        }

        private static void cancel_quote(ChromeDriver driver)
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(15));

            string locator;
            IWebElement element;

            locator = "input[type='SUBMIT'][value='Cancel']";

            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.CssSelector(locator)));

            do
            {
                try
                {
                    element = driver.FindElement(By.CssSelector(locator));
                    element.Click();
                    break;
                }
                catch
                {
                }
            }
            while (true);

            locator = "input[type='SUBMIT'][value='OK']";

            do
            {
                try
                {
                    element = driver.FindElement(By.CssSelector(locator));
                    element.Click();
                    break;
                }
                catch
                {
                }
            }
            while (true);
        }

        private static void select_soldto(ChromeDriver driver,string sold_to, short loops)
        {
            IJavaScriptExecutor js = driver as IJavaScriptExecutor;
            string locator;
            IWebElement element;

            for (short i = 1; i <= loops; i++)
            {
                do
                {
                    try
                    {
                        element = driver.FindElement(By.CssSelector("[class='is-icon is-icon-va is-shopping-cart']"));
                        element.Click();
                        break;
                    }
                    catch
                    {
                    }
                }
                while (true);
            }

            do
            {
                try
                {
                    element = driver.FindElement(By.Name("CustCode"));
                    element.Clear();
                    element.SendKeys(sold_to);
                    if (element.GetAttribute("value") == sold_to)
                        break;
                }
                catch
                {
                }
            }
            while (true);


            //js.ExecuteScript("javascript:return ProcessSubmit();");
            do
            {
                try
                {
                    element = driver.FindElement(By.CssSelector("input[type='SUBMIT'][value='Search']"));
                    element.Click();
                    break;
                }
                catch
                {
                }
            }
            while (true);

            do
            {
                try
                {
                    element = driver.FindElement(By.XPath("//*[@id='srchDivId']/table[2]/tbody/tr/td/table/tbody/tr[3]/td[2]/a"));
                    element.Click();
                    break;
                }
                catch
                {
                }
            }
            while (true);


            try
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));

                try
                { 
                wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.CssSelector("select[name='DeliveryCountry'][onchange='ProcessInstall(this);']")));
                element = driver.FindElement(By.CssSelector("select[name='DeliveryCountry'][onchange='ProcessInstall(this);']"));
                SelectElement slct = new SelectElement(element);
                slct.SelectByValue("us");
                }
                catch
                {
                }

                element = driver.FindElement(By.CssSelector("input[type='BUTTON'][value='Continue']"));
                element.Click();

            }
            catch
            {
            }

            
            try
            {
                driver.SwitchTo().Alert().Accept();
            }
            catch
            {
            }
            

            

        }

        private static void select_billto(ChromeDriver driver,string bill_to,string bill_to_country)
        {
            string locator;
            IWebElement element;

            locator = "Bill To";

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.LinkText(locator)));

            do
            {
                try
                {
                    element = driver.FindElement(By.LinkText(locator));
                    element.Click();
                    break;
                }
                catch
                {
                }
            }
            while (true);


            do
            {
                try
                {
                    element = driver.FindElement(By.CssSelector("select[name='Country'][id='Country']"));
                    SelectElement slct = new SelectElement(element);
                    slct.SelectByValue(bill_to_country.ToLower());

                    element = driver.FindElement(By.CssSelector("input[type='TEXT'][name='AddrID']"));
                    element.Clear();
                    element.SendKeys(bill_to);
                    if (element.GetAttribute("value") == bill_to)
                        break;
                }
                catch
                {
                }
            }
            while (true);

            element = driver.FindElement(By.CssSelector("input[type='SUBMIT'][value='Search']"));
            element.Click();

            do
            {
                try
                {
                    element = driver.FindElement(By.XPath("//*[starts-with(@onclick, 'SetAddress(')]"));
                    element.Click();
                    break;
                }
                catch
                {
                }
            }
            while (true);


        }

        private static void select_shipto(ChromeDriver driver, string ship_to,string ship_to_country)
        {
            IWebElement element;
            string locator;
            IJavaScriptExecutor js = driver as IJavaScriptExecutor;

            


            js.ExecuteScript("SetCountryField(2, '" + ship_to_country + "')");
            try
            {
                driver.SwitchTo().Alert().Accept();
            }
            catch
            {
            }

            finish_loading(driver, "//*[@id='ajwmCnt']");

            js.ExecuteScript("SetCountryField(3, '" + ship_to_country + "')");
            try
            {
                driver.SwitchTo().Alert().Accept();
            }
            catch
            {
            }

            finish_loading(driver, "//*[@id='ajwmCnt']");

            locator = "Ship To";

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

            do
            {
                try
                {
                    wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.LinkText(locator)));
                    break;
                }
                catch (StaleElementReferenceException e)
                {
                    System.Threading.Thread.Sleep(500);
                }
            }
            while (true);

            

            do
            {
                try
                {
                    element = driver.FindElement(By.LinkText(locator));
                    element.Click();
                    break;
                }
                catch
                {
                }
            }


            while (true);

            do
            {
                try
                {
                    element = driver.FindElement(By.CssSelector("input[type='TEXT'][name='AddrID']"));
                    element.Clear();
                    element.SendKeys(ship_to);
                    if (element.GetAttribute("value") == ship_to)
                        break;
                }
                catch
                {
                }
            }
            while (true);


            element = driver.FindElement(By.CssSelector("input[type='SUBMIT'][value='Search']"));
            element.Click();

            do
            {
                try
                {
                    element = driver.FindElement(By.XPath("//*[starts-with(@onclick, 'SetAddress(')]"));
                    element.Click();
                    break;
                }
                catch
                {
                }
            }
            while (true);
        }

        private static void select_contract(ChromeDriver driver, string contract)
        {
            IWebElement element;

            do
            {
                try
                {
                    element = driver.FindElement(By.LinkText("Client Contract Number"));
                    element.Click();
                    break;
                }
                catch
                {
                }
            }
            while (true);

            DateTime cutoff = DateTime.Now.AddSeconds(15);

            do
            {
                System.Threading.Thread.Sleep(1000);
            }
            while (driver.FindElement(By.Id("DD:MSCONTRACTNUMbody")).Text == "Please wait...Data is loading" &&
                    DateTime.Now<cutoff);


            element = driver.FindElement(By.Name("srchDD:MSCONTRACTNUM"));
            element.SendKeys(contract);
            element = driver.FindElement(By.CssSelector("input[type='SUBMIT'][value='Search']"));
            element.Click();

            do
            {
                try
                {
                    element = driver.FindElement(By.LinkText(contract));
                    element.Click();
                    break;
                }
                catch
                {
                }
            }
            while (true);


            
        }

        private static void select_sales_office(ChromeDriver driver, string sales_office)
        {
            IWebElement element;

            do
            {
                try
                {
                    element = driver.FindElement(By.LinkText("Sales Office"));
                    element.Click();
                    break;
                }
                catch
                {
                }
            }
            while (true);

            do
            {
                try
                {
                    element = driver.FindElement(By.PartialLinkText(sales_office));
                    element.Click();
                    break;

                }
                catch
                {
                }
            }
            while (true);


        }

        private static void select_bill_freq(ChromeDriver driver, string bill_plan)
        {
            string locator;
            IWebElement element;

            locator = "MS Billing Frequency";

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.LinkText(locator)));

             do
            {
                try
                {
                    element = driver.FindElement(By.LinkText(locator));
                    element.Click();
                    break;
                }
                catch
                {
                }

            }
            while (true);

            switch (bill_plan)
            {
                case "Monthly":
                    {
                        locator = "Monthly in Advance";
                        break;
                    }

                case "Annual":
                    {
                        locator = "Annually in Advance";
                        break;
                    }
            }

            do
            {
                try
                {
                    element = driver.FindElement(By.LinkText(locator));
                    element.Click();
                    break;
                }
                catch
                {
                }

            }
            while (true);
        }

        private static void select_contract(ChromeDriver driver)
        {

        }

        private static void select_solution_type(ChromeDriver driver, string solution)
        {
            IWebElement element;
            do
            {
                try
                {
                    element = driver.FindElement(By.LinkText("Solution Type"));
                    element.Click();
                    break;
                }
                catch
                {
                }

            }
            while (true);

            do
            {
                try
                {
                    element = driver.FindElement(By.PartialLinkText(solution + " -"));
                    element.Click();
                    break;
                }
                catch
                {
                }

            }
            while (true);
        }

        private static void select_delivery_date(ChromeDriver driver)
        {
            string today = DateTime.Today.ToShortDateString();

            do
            {
                try
                {
                    IWebElement element = driver.FindElement(By.Name("REGION_REQDELDATE"));
                    element.Clear();

                    element.SendKeys(today);
                    if (element.GetAttribute("value") == today);
                        break;
                }
                catch
                {
                }

            }
            while (true);
        }

        private static void catalog(ChromeDriver driver)
        {
            IWebElement element;

            do
            {
                try
                {
                    element = driver.FindElement(By.XPath("//*[@id='BottomRow']/a"));
                    break;
                }
                catch
                {
                }

            }
            while (true);


            string catalogURL = "https://direct.dimensiondata.com/seller/direct/" + element.GetAttribute("onclick").Split('\'')[1];
            driver.Navigate().GoToUrl(catalogURL);



            do
            {
                try
                {
                    element = driver.FindElement(By.LinkText("BOM Upload"));
                    element.Click();
                    break;
                }
                catch
                {
                }

            }
            while (true);

            do
            {
                try
                {
                    element = driver.FindElement(By.CssSelector("[class='is-search ng-star-inserted']"));
                    element.SendKeys("AM_Service Codes Backouts");
                    break;
                }
                catch
                {
                }

            }
            while (true);

            do
            {
                try
                {
                    element = driver.FindElement(By.LinkText("AM_Service Codes Backouts"));
                    element.Click();
                    break;
                }
                catch
                {
                }

            }
            while (true);

            do
            {
                try
                {
                    element = driver.FindElement(By.CssSelector("[class='is-button is-primary'][type='submit']"));
                    element.Click();
                    break;
                }
                catch
                {
                }

            }
            while (true);

            do
            {
                try
                {
                    element = driver.FindElement(By.Name("FileName"));
                    element.SendKeys(Directory.GetCurrentDirectory() + "\\bomupload.csv");
                    System.Threading.Thread.Sleep(500);
                    break;
                }
                catch
                {
                }

            }
            while (true);

            

            element = driver.FindElement(By.CssSelector("input[type='SUBMIT'][name='UPDATE'][value='Upload'][onclick='return CheckFile();']"));
            element.Click();

        }

        private static void deselect_quote_ccs(ChromeDriver driver, string locator)
        {
            IWebElement element = driver.FindElement(By.XPath(locator));
            IReadOnlyCollection<IWebElement> elements = element.FindElements(By.TagName("option"));

            foreach (IWebElement ele in elements)
            {
                if (ele.Text == "(Not Selected)")
                {
                    if (ele.Selected == false)
                        ele.Click();
                else
                    if (ele.Selected == true)
                        ele.Click();
                }
            }
        }

        private static void submit_quote(ChromeDriver driver, string sub_id)
        {
            string locator;
            IWebElement element;
            string quote_number;


            locator = "input[type='SUBMIT'][value='Submit']";


            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.CssSelector(locator)));

            do
            {
                try
                {
                    element = driver.FindElement(By.CssSelector(locator));
                    OpenQA.Selenium.Interactions.Actions actions = new OpenQA.Selenium.Interactions.Actions(driver);
                    actions.MoveToElement(element);
                    actions.Perform();
                    element.Click();
                    break;
                }
                catch
                {
                }
            }
            while (true);


            do
            {
                try
                {
                    element = driver.FindElement(By.CssSelector("input[type='SUBMIT'][value='Apply']"));
                    element.Click();
                    break;
                }
                catch
                {
                    try
                    {
                        element = driver.FindElement(By.CssSelector("input[type='SUBMIT'][value='Continue']"));
                        element.Click();
                    }
                    catch
                    {
                    }
                }
            }
            while (true);


            do
            {
                try
                {
                    element = driver.FindElement(By.Name("txtQNAME"));
                    element.Clear();
                    
                    
                    element.SendKeys("none");
                    if (element.GetAttribute("value") == "none")
                        break;
                }
                catch
                {
                }
            }
            while (true);


            try
            {
                deselect_quote_ccs(driver, "//*[contains(@name,'cbCC')]");
            }
            catch
            {
            }

            try
            {
                deselect_quote_ccs(driver, "//*[@name='CCCustomerContact']");
            }
            catch
            {
            }

            try
            {
                deselect_quote_ccs(driver, "//*[@name='CCAssignedSales']");
            }
            catch
            {
            }

            element = driver.FindElement(By.CssSelector("input[type='SUBMIT'][value='Submit'][id='SUBMIT1'][name='SUBMIT1']"));
            element.Click();

            do
            {
                try
                {
                    locator = "DoNotSendEmail";
                    element = driver.FindElement(By.Name(locator));
                    break;
                }
                catch
                {
                }
            }
            while (true);

            switch (element.Selected)
            {
                case true:
                    {
                        break;
                    }

                case false:
                    {
                        element.Click();
                        break;
                    }
            }


            element = driver.FindElement(By.CssSelector("input[type='SUBMIT'][value='Finalize']"));
            element.Click();

            do
            {
                try
                {
                    element = driver.FindElement(By.XPath("/html/body/div[1]/main/table/tbody/tr/td/p/table/tbody/tr/td/table/tbody/tr[2]/td/b/a"));
                    quote_number = element.Text;
                    element.Click();
                    break;
                }
                catch
                {
                }
            }
            while (true);

            Console.WriteLine("Direct Quote " + quote_number + " Successfully Created.");
        }

        private static void convert_to_po(ChromeDriver driver, string sold_to_name, string contract, string sales_office, string cust_po, string vend_po, string woid, string cust_po_path, string sub_id, string client_manager="", string quote_number= "")
        {
            string locator;
            IWebElement element;
            IJavaScriptExecutor js = driver as IJavaScriptExecutor;

            do
            {
                try
                {
                    js.ExecuteScript("CheckWfWarning('POCONVERT');return false;");
                    break;
                }
                catch
                {
                }
            }
            while (true);

            do
            {
                try
                {
                    js.ExecuteScript("ValidateRowIDs('DoRefreshAction()');return false;");
                    break;
                }
                catch
                {
                    try
                    {
                        element = driver.FindElement(By.Name("PONumber"));
                        break;
                    }
                    catch
                    {
                    }
                }
            }
            while (true);

            do
            {
                try
                {
                    element = driver.FindElement(By.Name("PONumber"));
                    element.Clear();
                    element.SendKeys(cust_po);
                    if (element.GetAttribute("value") == cust_po)
                        break;
                }
                catch
                {
                }
            }
            while (true);



            string ContactPerson = "Other";

            element = driver.FindElement(By.Name("selBILLCONTACTS"));
            SelectElement slct_Bill = new SelectElement(element);
            slct_Bill.SelectByText(ContactPerson);

            element = driver.FindElement(By.Name("selSHIPCONTACTS"));
            SelectElement slct_Ship = new SelectElement(element);
            slct_Ship.SelectByText(ContactPerson);

            do
            {
                try
                {
                    string name = "Other";
                    string phone = "1234567890";
                    IWebElement BillOtherCName = driver.FindElement(By.Name("BillOtherCName"));
                    IWebElement BillOtherPhone = driver.FindElement(By.Name("BillOtherPhone"));
                    IWebElement ShipOtherCName = driver.FindElement(By.Name("ShipOtherCName"));
                    IWebElement ShipOtherPhone = driver.FindElement(By.Name("ShipOtherPhone"));

                    BillOtherCName.Clear();
                    BillOtherPhone.Clear();
                    ShipOtherCName.Clear();
                    ShipOtherPhone.Clear();

                    BillOtherCName.SendKeys(name);
                    BillOtherPhone.SendKeys(phone);
                    ShipOtherCName.SendKeys(name);
                    ShipOtherPhone.SendKeys(phone);

                    if (BillOtherCName.GetAttribute("value") == name && 
                        BillOtherPhone.GetAttribute("value") == phone &&
                        ShipOtherCName.GetAttribute("value") == name &&
                        ShipOtherPhone.GetAttribute("value") == phone)
                        break;
                }
                catch
                {
                }
            }
            while (true);

            js.ExecuteScript("AddAttachments();return false;");

            driver.SwitchTo().Window(driver.WindowHandles[1]);
            element = driver.FindElement(By.Name("Attachments"));
            element.SendKeys(cust_po_path);
            element = driver.FindElement(By.Name("UPDATE"));
            element.Click();
            driver.SwitchTo().Window(driver.WindowHandles[0]);

            js.ExecuteScript("SetValue(); return CheckDuplicatePO();");

            try
            {
                js.ExecuteScript("HidePopup();ContinueKeepPO();");
            }
            catch
            {
            }
            element = driver.FindElement(By.CssSelector("input[type='SUBMIT'][value='Finalize - View PO']"));
            element.Click();

            select_tab(driver, "additional");

            select_contract(driver, contract);

            select_tab(driver, "integration");

            select_sales_office(driver, sales_office);

            select_delivery_date(driver);

            update_quote(driver, "integration", 3);

            select_tab(driver, "comments");

            string comments="";


            if (sub_id != "empty")
                comments = comments+ sub_id + " Renewal.";

            if (vend_po != "empty")
                comments = comments + "Please add to PO " + vend_po+".";

            if (woid != "empty")
                comments = comments + "WOID is " + woid + ".";

            do
            {
                try
                {
                    element = driver.FindElement(By.Name("InternalComments"));
                    element.Clear();
                    element.SendKeys(comments);
                    if (element.GetAttribute("value")==comments)
                        break;
                }
                catch
                {
                }
            }
            while (true);


            update_quote(driver, "comments", 5);

            element=driver.FindElement(By.CssSelector("input[type='BUTTON'][name='ERPPost'][id='ERPPost']"));
            element.Click();

            do
            {
                try
                {
                    element = driver.FindElement(By.CssSelector("input[type='SUBMIT'][name='FINISH']"));
                    element.Click();
                    break;
                }
                catch
                {
                }
            }
            while (true);

            do
            {
                try
                {
                    element = driver.FindElement(By.LinkText("ERP SUBMISSION AUDIT TRAIL"));
                    element.Click();
                    break;
                }
                catch
                {
                }
            }
            while (true);

            bool finished = false;


            do
            {

                IReadOnlyCollection <IWebElement> elements = driver.FindElements(By.ClassName("ListAltRow1"));

                foreach (IWebElement ele in elements)
                {
                    if (ele.Text.Contains("SAP Quote Successfully Created."))
                    {

                        long conf_number = Convert.ToInt64(ele.Text.Split('#')[1]);

                        string email_comments = "Conf#: " + conf_number;
                        if (vend_po != "empty")
                            email_comments = email_comments+ " <br /> PO: " + vend_po;

                        if (woid != "empty")
                            email_comments = email_comments+ " <br /> WOID: " + woid;

                        string email_title = sold_to_name+" ";

                        if (sub_id != "empty")
                        { 
                            email_title = email_title+sub_id + " Change/Renewal SaaS Order";
                        }
                        else
                            email_title = email_title+" New SaaS Order";

                        MailMessage("", "", cust_po_path, email_title, "Hi,<br /><br /> "+email_comments);
                        finished=true;
                    }

                    Console.WriteLine(ele.Text);
                }
                
                while(finished==false)
                {
                    try
                    {
                        element = driver.FindElement(By.CssSelector("input[type='SUBMIT'][value='Cancel']"));
                        element.Click();
                        element = driver.FindElement(By.LinkText("ERP SUBMISSION AUDIT TRAIL"));
                        element.Click();
                        System.Threading.Thread.Sleep(2000);
                        break;
                    }
                    catch
                    {
                    }
                }
            }
            while (finished==false);
        }

        static private ChromeDriver Driver()
        {
            ChromeDriverService driverService = ChromeDriverService.CreateDefaultService();
            driverService.HideCommandPromptWindow = true;

            ChromeOptions driverOptions = new ChromeOptions();
            driverOptions.AddArgument("--window-size=1280,800");
            driverOptions.AddArgument("--headless");
            driverOptions.AddArgument("--disable-gpu");
            driverOptions.AddArgument("--hide-scrollbars");
            driverOptions.AddArgument("--allow-insecure-localhost");
            driverOptions.AddArgument("--disable-notifications");
            driverOptions.AddAdditionalCapability("acceptInsecureCerts", true, true);
            driverOptions.AddUserProfilePreference("download.prompt_for_download", false);
            driverOptions.AddUserProfilePreference("plugins.always_open_pdf_externally", true);

            Dictionary<string, object> driver_params = new Dictionary<string, object>();

            driver_params.Add("behavior", "allow");
            driver_params.Add("downloadPath", System.Reflection.Assembly.GetExecutingAssembly().Location);


            ChromeDriver driver = new ChromeDriver(driverService, driverOptions);

            driver.ExecuteChromeCommandWithResult("Page.setDownloadBehavior", driver_params);

            driver.Manage().Window.Maximize();

            return driver;
        }

        private static void write_csv(string bill_plan,string start_date, string end_date,short months,double cost, double sell_price)
        {
            string csv_path = Directory.GetCurrentDirectory() + "\\bomupload.csv";

            string csv_headers = Properties.Resources.bomupload;

            long qty = months;
            double total_cost = cost;
            double total_sell_price = sell_price;

            switch (bill_plan)
            {
                case "Prepaid":
                    {
                        qty = 1;
                        total_cost = cost * months;
                        total_sell_price = sell_price * months;
                        break;
                    }
                case "Annual":
                    {
                        qty = months/12;
                        total_cost = cost * months/12;
                        total_sell_price = sell_price * months / 12;
                        break;
                    }
            }

            string csv_text = "10" + ",," + "Dimension Data" + ",," + qty + "," + "DDSP-VBR-VendorBrandedResell" + "," + "MS VBR Base Services" + "," + "7" + ",,," + "1" + "," + "EA" + "," + "USD" + "," + "Fixed" + "," + "0" + "," + total_sell_price + "," + total_cost + ",,,,,,," + start_date + "," + end_date;

            File.WriteAllText(csv_path, csv_headers);
            File.AppendAllText(csv_path, csv_text);

        }

        private static void MailMessage(string recipient, string CC, string attachment, string subject, string body)
        {
            Microsoft.Office.Interop.Outlook.Application AppOutlook = new Microsoft.Office.Interop.Outlook.Application();
            Microsoft.Office.Interop.Outlook.MailItem OutlookMessage = AppOutlook.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);

            try
            {
                {

                    Microsoft.Office.Interop.Outlook.Recipients Recipents = OutlookMessage.Recipients;
                    Recipents.Add(recipient);
                    OutlookMessage.CC = CC;
                    OutlookMessage.Attachments.Add(attachment);
                    OutlookMessage.Subject = subject;
                    OutlookMessage.BodyFormat = Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatHTML;
                    OutlookMessage.Display();
                    OutlookMessage.HTMLBody = body + OutlookMessage.HTMLBody;
                    OutlookMessage.Send();
                    Console.WriteLine("Email successfully sent to " + recipient);
                }
            }
            // OutlookMessage.Send()
            catch
            {
                Console.WriteLine("Email Send Failed");
            }
            // MessageBox.Show("Mail could not be sent") 'if you dont want this message, simply delete this line 
            finally
            {
                OutlookMessage = null/* TODO Change to default(_) if this is not a reference type */;
                AppOutlook = null/* TODO Change to default(_) if this is not a reference type */;
            }
        }

    }
}
