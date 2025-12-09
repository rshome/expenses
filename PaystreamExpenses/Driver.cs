using OpenQA.Selenium;
using System;
using System.Threading;
using OpenQA.Selenium.Chrome;
using excel = Microsoft.Office.Interop.Excel;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;
using SeleniumExtras.WaitHelpers;

namespace PaystreamExpenses
{
    public class Driver
    {
        IWebDriver driver = new ChromeDriver();
        string url = "https://portal.paystream.co.uk/";

        excel.Application xlApp = new excel.Application();
        Helper _helper = new Helper();

        public void Login()
        {
            string username = "";
            string password = "";

            excel.Application xlApp = new excel.Application();

            excel.Workbook xlWorkbook = xlApp.Workbooks.Open("C:\\Passwords\\Passwords.xlsx");
            excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            excel.Range xlRange = xlWorksheet.UsedRange;

            username = xlRange.Cells[1][1].Value;
            password = xlRange.Cells[1][2].Value;            

            driver.Navigate().GoToUrl(url);
            driver.Manage().Cookies.DeleteAllCookies();
            driver.Manage().Window.Maximize();
            driver.FindElement(By.XPath("//*[@id='login-form']/form/button")).Click();

            //driver.FindElement(By.Id("Username")).SendKeys(username);
            _helper.EnterData(driver,"Username", username);            
            _helper.EnterData(driver,"Password", password);
            
            //driver.FindElement(By.Id("Password")).SendKeys(password);
            driver.FindElement(By.XPath("/html/body/div/div[2]/div/form/div[3]/button[1]")).Click();
            
            //driver.FindElement(By.XPath(".//*[@id='login']/ul/li[4]/button")).Click();
            Thread.Sleep(1000);
        }

        public void SelectAccountandExpenses()
        {
            Thread.Sleep(1000);

            driver.FindElement(By.Name("choose-user")).Click();
            //driver.FindElement(By.XPath(".//*[@id='content']/div[2]/div[2]/form/button")).Click();

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20));
            // Replace all occurrences of ExpectedConditions with the correct namespace
            wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.LinkText("Expenses")));
            Thread.Sleep(1000);
            driver.FindElement(By.LinkText("Expenses")).Click();

            Thread.Sleep(1000);
        }

        public void Broadband()
        {
            string bBand;

            excel.Workbook xlWorkbook = xlApp.Workbooks.Open("C:\\Passwords\\ExpensesDemo.xlsx");
            excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            excel.Range xlRange = xlWorksheet.UsedRange;

            Thread.Sleep(1000);
                      
                bBand = xlRange.Cells[1][2].Value2.ToString();

                driver.FindElement(By.XPath(".//*[@id='add-item-links']/div/div/div[2]/div[2]/div/button")).Click();

                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20));
                wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.Id("ExpenseParentCategoryId")));

            //IWebElement week = driver.FindElement(By.Id("WeekEndingDateDisplay"));
            //week.SendKeys(Keys.ArrowDown);

            IWebElement internet = driver.FindElement(By.Id("ExpenseParentCategoryId"));
                internet.SendKeys(Keys.ArrowDown);

                //type
                IWebElement type = driver.FindElement(By.Id("ExpenseCategoryId"));
                type.SendKeys(Keys.ArrowDown);
                type.SendKeys(Keys.ArrowDown);
                type.SendKeys(Keys.ArrowDown);

                // description
                driver.FindElement(By.Id("Description")).SendKeys("Broadband");

                //amount
                driver.FindElement(By.Id("GrossAmount")).Clear();

                driver.FindElement(By.Id("GrossAmount")).SendKeys(bBand);
                Thread.Sleep(1000);

                driver.FindElement(By.XPath("(//button[@type='button'])[2]")).Click();

                Thread.Sleep(1000);
                
            xlWorkbook.Close();            
        }

        public void Hotel()
        {
            string hotel;

            excel.Workbook xlWorkbook = xlApp.Workbooks.Open("C:\\Passwords\\ExpensesDemo.xlsx");
            excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            excel.Range xlRange = xlWorksheet.UsedRange;

            Thread.Sleep(1000);

            hotel = xlRange.Cells[10][2].Value2.ToString();

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20));
            Thread.Sleep(1000);

            wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id='add-receipted-item']")));
            driver.FindElement(By.XPath("//*[@id='add-receipted-item']")).Click();

            wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.Id("ExpenseParentCategoryId")));

            IWebElement week = driver.FindElement(By.Id("WeekEndingDateDisplay"));
            //week.SendKeys(Keys.ArrowDown);

            IWebElement internet = driver.FindElement(By.Id("ExpenseParentCategoryId"));
            internet.SendKeys(Keys.ArrowDown);

            //type
            IWebElement type = driver.FindElement(By.Id("ExpenseCategoryId"));
            type.SendKeys(Keys.ArrowDown);

            // description
            driver.FindElement(By.Id("Description")).SendKeys("Hotel");

            //amount
            driver.FindElement(By.Id("GrossAmount")).Clear();

            driver.FindElement(By.Id("GrossAmount")).SendKeys(hotel);
            Thread.Sleep(1000);
            driver.FindElement(By.Id("HaveVatReceipt")).Click();

            driver.FindElement(By.XPath("(//button[@type='button'])[2]")).Click();

            Thread.Sleep(1000);
            driver.FindElement(By.Id("Nights")).Clear();
            driver.FindElement(By.Id("Nights")).SendKeys("1");

            driver.FindElement(By.XPath("/html/body/div[4]/div[3]/div/button[1]")).Click();
            Thread.Sleep(1000);

            xlWorkbook.Close();
        }

        public void MonthlyTrainPass()
        {
            string tPass;

            excel.Workbook xlWorkbook = xlApp.Workbooks.Open("C:\\Passwords\\ExpensesDemo.xlsx");
            excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            excel.Range xlRange = xlWorksheet.UsedRange;

            Thread.Sleep(1000);

            if (xlRange.Cells[9][2].Value2 != null)
            {
                tPass = xlRange.Cells[9][2].Value.ToString();
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20));

                Thread.Sleep(1500);                
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//button[@value='AddReceiptedItem']")));
                driver.FindElement(By.XPath("//button[@value='AddReceiptedItem']")).Click();

                wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.Id("ExpenseParentCategoryId")));

                //IWebElement week = driver.FindElement(By.Id("WeekEndingDateDisplay"));
                //week.SendKeys(Keys.ArrowUp);
                //week.SendKeys(Keys.ArrowDown);

                IWebElement internet = driver.FindElement(By.Id("ExpenseParentCategoryId"));
                internet.SendKeys(Keys.ArrowDown);
                internet.SendKeys(Keys.ArrowDown);
                internet.SendKeys(Keys.ArrowDown);

                //type
                IWebElement type = driver.FindElement(By.Id("ExpenseCategoryId"));
                type.SendKeys(Keys.ArrowDown);

                // description
                driver.FindElement(By.Id("Description")).SendKeys("Train Pass to London");

                //amount
                driver.FindElement(By.Id("GrossAmount")).Clear();
                driver.FindElement(By.Id("GrossAmount")).SendKeys(tPass);
                Thread.Sleep(1000);

                driver.FindElement(By.XPath("(//button[@type='button'])[2]")).Click();                

                Thread.Sleep(1000);
            }
            xlWorkbook.Close();
        }

        public void DeclareExpensesLunch()
        {
            string lunch;  //using figures from Expenses spreadsheet
            excel.Workbook xlWorkbook = xlApp.Workbooks.Open("C:\\Passwords\\ExpensesDemo.xlsx");
            excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            excel.Range xlRange = xlWorksheet.UsedRange;

            Thread.Sleep(1000);
       
            for (int i = 2; i < 7; i++)
            {                
                    lunch = xlRange.Cells[3][i].Value2.ToString();
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20));

                Thread.Sleep(1000);
                wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.XPath("/html/body/div[3]")));
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id='add-receipted-item']")));
                driver.FindElement(By.XPath("//*[@id='add-receipted-item']")).Click();

                wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.Id("ExpenseParentCategoryId")));

                //IWebElement week = driver.FindElement(By.Id("WeekEndingDateDisplay"));
                //week.SendKeys(Keys.ArrowDown);

                //lunch
                IWebElement meals = driver.FindElement(By.Id("ExpenseParentCategoryId"));
                meals.SendKeys(Keys.ArrowDown);
                meals.SendKeys(Keys.ArrowDown);

                //breakfast
                IWebElement type = driver.FindElement(By.Id("ExpenseCategoryId"));
                type.SendKeys(Keys.ArrowDown);
                type.SendKeys(Keys.ArrowDown);

                // description
                driver.FindElement(By.Id("Description")).SendKeys("Lunch");

                    //amount
                    driver.FindElement(By.Id("GrossAmount")).Clear();
                    driver.FindElement(By.Id("GrossAmount")).SendKeys(lunch);
                    Thread.Sleep(1000);

                    driver.FindElement(By.XPath("(//button[@type='button'])[2]")).Click();
                   
                    Thread.Sleep(1000);
            }
            xlWorkbook.Close();
        }

        public void DeclareExpensesBreakfast()
        {
            string breakFast;  //using figures from Expenses spreadsheet
            excel.Workbook xlWorkbook = xlApp.Workbooks.Open("C:\\Passwords\\ExpensesDemo.xlsx");
            excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            excel.Range xlRange = xlWorksheet.UsedRange;

            Thread.Sleep(1500);

            for (int i = 2; i < 7; i++)
            {
                    breakFast = xlRange.Cells[2][i].Value.ToString();
                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20));

                    Thread.Sleep(1000);
                wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.XPath("/html/body/div[3]")));
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id='add-receipted-item']")));
                driver.FindElement(By.XPath("//*[@id='add-receipted-item']")).Click();

                wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.Id("ExpenseParentCategoryId")));

                //IWebElement week = driver.FindElement(By.Id("WeekEndingDateDisplay"));
                //week.SendKeys(Keys.ArrowDown);

                //breakfast
                IWebElement meals = driver.FindElement(By.Id("ExpenseParentCategoryId"));
                    meals.SendKeys(Keys.ArrowDown);
                    meals.SendKeys(Keys.ArrowDown);

                    //breakfast
                    IWebElement type = driver.FindElement(By.Id("ExpenseCategoryId"));
                    type.SendKeys(Keys.ArrowDown);

                    // description
                    driver.FindElement(By.Id("Description")).SendKeys("Breakfast");

                    //amount
                    driver.FindElement(By.Id("GrossAmount")).Clear();

                    driver.FindElement(By.Id("GrossAmount")).SendKeys(breakFast.ToString());
                    Thread.Sleep(1000);

                    driver.FindElement(By.XPath("(//button[@type='button'])[2]")).Click();
                    
                Thread.Sleep(1000);                
            }
            xlWorkbook.Close();
        }

        public void DeclareExpensesCoffee()
        {
            string coffee;  //using figures from Expenses spreadsheet
            excel.Workbook xlWorkbook = xlApp.Workbooks.Open("C:\\Passwords\\ExpensesDemo.xlsx");
            excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            excel.Range xlRange = xlWorksheet.UsedRange;

            Thread.Sleep(1000);

            for (int i = 2; i < 7; i++)
            {
                    coffee = xlRange.Cells[4][i].Value2.ToString();
                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20));                
                Thread.Sleep(1000);
                wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.XPath("/html/body/div[3]")));
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id='add-receipted-item']")));
                driver.FindElement(By.XPath("//*[@id='add-receipted-item']")).Click();

                wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.Id("ExpenseParentCategoryId")));

                //IWebElement week = driver.FindElement(By.Id("WeekEndingDateDisplay"));
                //week.SendKeys(Keys.ArrowDown);

                //coffee
                IWebElement meals = driver.FindElement(By.Id("ExpenseParentCategoryId"));
                    meals.SendKeys(Keys.ArrowDown);
                    meals.SendKeys(Keys.ArrowDown);

                    //breakfast
                    IWebElement type = driver.FindElement(By.Id("ExpenseCategoryId"));
                    type.SendKeys(Keys.ArrowDown);

                    // description
                    driver.FindElement(By.Id("Description")).SendKeys("Coffees");

                    //amount
                    driver.FindElement(By.Id("GrossAmount")).Clear();
                    driver.FindElement(By.Id("GrossAmount")).SendKeys(coffee.ToString());
                    Thread.Sleep(1000);

                    driver.FindElement(By.XPath("(//button[@type='button'])[2]")).Click();                
                Thread.Sleep(1000);
                //}
            }
            xlWorkbook.Close();
        }

        public void DeclareExpensesParking()
        {
            string park;  //using figures from Expenses spreadsheet
            excel.Workbook xlWorkbook = xlApp.Workbooks.Open("C:\\Passwords\\ExpensesDemo.xlsx");
            excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            excel.Range xlRange = xlWorksheet.UsedRange;

            Thread.Sleep(1000);

            for (int i = 2; i < 7; i++)
            {
                    park = xlRange.Cells[5][i].Value2.ToString();
                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20));                
                Thread.Sleep(1500);
                wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.XPath("/html/body/div[3]")));
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id='add-receipted-item']")));
                driver.FindElement(By.XPath("//*[@id='add-receipted-item']")).Click();

                //Thread.Sleep(1000);
                wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.Id("ExpenseParentCategoryId")));

                //IWebElement week = driver.FindElement(By.Id("WeekEndingDateDisplay"));
                //week.SendKeys(Keys.ArrowDown);

                IWebElement parking = driver.FindElement(By.Id("ExpenseParentCategoryId"));
                    parking.SendKeys(Keys.ArrowDown);
                    parking.SendKeys(Keys.ArrowDown);
                    parking.SendKeys(Keys.ArrowDown);

                    //parking
                    IWebElement type = driver.FindElement(By.Id("ExpenseCategoryId"));
                    type.SendKeys(Keys.ArrowDown);
                    type.SendKeys(Keys.ArrowDown);
                    type.SendKeys(Keys.ArrowDown);
                    type.SendKeys(Keys.ArrowDown);
                    type.SendKeys(Keys.ArrowDown);
                    type.SendKeys(Keys.ArrowDown);
                    type.SendKeys(Keys.ArrowDown);
                    type.SendKeys(Keys.ArrowDown);
                    type.SendKeys(Keys.ArrowDown);
                    type.SendKeys(Keys.ArrowDown);

                    // description
                    driver.FindElement(By.Id("Description")).SendKeys("Parking");

                    //amount
                    driver.FindElement(By.Id("GrossAmount")).Clear();
                    driver.FindElement(By.Id("GrossAmount")).SendKeys(park.ToString());

                    driver.FindElement(By.XPath("(//button[@type='button'])[2]")).Click();                

                Thread.Sleep(1000);                
            }
            xlWorkbook.Close();
        }

        public void DeclarePhoneCalls()
        {
            string phone;  //using figures from Expenses spreadsheet
            excel.Workbook xlWorkbook = xlApp.Workbooks.Open("C:\\Passwords\\ExpensesDemo.xlsx");
            excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            excel.Range xlRange = xlWorksheet.UsedRange;

            Thread.Sleep(1000); //test

            if (xlRange.Cells[6][2].Value != null)
            {
                phone = xlRange.Cells[6][2].Value.ToString();
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20));
                Thread.Sleep(1000);
                wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.XPath("/html/body/div[3]")));
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id='add-receipted-item']")));
                driver.FindElement(By.XPath("//*[@id='add-receipted-item']")).Click();

                wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.Id("ExpenseParentCategoryId")));

                //IWebElement week = driver.FindElement(By.Id("WeekEndingDateDisplay"));
                //week.SendKeys(Keys.ArrowDown);

                //category
                IWebElement cat = driver.FindElement(By.Id("ExpenseParentCategoryId"));

                for (int i = 0; i < 4; i++)
                {
                    cat.SendKeys(Keys.ArrowDown);
                }

                //type
                IWebElement type = driver.FindElement(By.Id("ExpenseCategoryId"));

                for (int i = 0; i < 2; i++)
                {
                    type.SendKeys(Keys.ArrowDown);
                }

                driver.FindElement(By.Id("Description")).SendKeys("Work calls");

                //amount
                driver.FindElement(By.Id("GrossAmount")).Clear();
                driver.FindElement(By.Id("GrossAmount")).SendKeys(phone.ToString());
                Thread.Sleep(1000);

                driver.FindElement(By.XPath("(//button[@type='button'])[2]")).Click();                
                Thread.Sleep(1000);
            }
            xlWorkbook.Close();
        }

        public void DeclareExpensesDriving()
        {
            //string miles;  //using figures from Expenses spreadsheet
            string eSize;
            int rowEnd;

            excel.Workbook xlWorkbook = xlApp.Workbooks.Open("C:\\Passwords\\ExpensesDemo.xlsx");
            excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            excel.Range xlRange = xlWorksheet.UsedRange;
            rowEnd = xlRange.Rows.Count;

            Thread.Sleep(1000);

                eSize = xlRange.Cells[8][2].Value;

                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(40));
                Thread.Sleep(1000);

                wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.XPath("//*[@id='add-mileage-item']")));
                wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath("//*[@id='add-mileage-item']")));
                driver.FindElement(By.XPath("//*[@id='add-mileage-item']")).Click();

                Thread.Sleep(1500);
                wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.Id("VehicleType")));

               //IWebElement week = driver.FindElement(By.Id("WeekEndingDateDisplay"));
               // week.SendKeys(Keys.ArrowDown);

            Thread.Sleep(1000);
                IWebElement vehicle = driver.FindElement(By.Id("VehicleType"));
                vehicle.SendKeys(eSize);

                // description
                driver.FindElement(By.Id("Description")).SendKeys("Drive Commute");

                //total number of miles for the week
                driver.FindElement(By.Id("toClaim")).SendKeys("130");

                //clain WAT for petrol
                var element = driver.FindElement(By.Id("ReclaimVat"));
                Actions actions = new Actions(driver);
                actions.MoveToElement(element);
                element.Click();
                Thread.Sleep(1000);

                driver.FindElement(By.XPath("(//button[@type='button'])[2]")).Click();
                Thread.Sleep(1000);

            xlWorkbook.Close();
        }       
    }
    
    
}
