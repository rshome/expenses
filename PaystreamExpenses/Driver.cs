using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using System;
using System.Threading;
using OpenQA.Selenium.Chrome;
using Microsoft.VisualBasic;
using Excel;

namespace PaystreamExpenses
{
    public class Driver
    {        
            IWebDriver driver;
            string url = "https://portal.paystream.co.uk/";
            

            //change firefox version to lower(currently version 43.0 works, latest is 50)

            public void Login()
            {
                var username = "ricky_shome@yahoo.co.uk";
                var password = "mem2as";

                driver = new FirefoxDriver();                

                driver.Navigate().GoToUrl(url);
                driver.Manage().Cookies.DeleteAllCookies();
                driver.Manage().Window.Maximize();

                driver.FindElement(By.Id("EmailAddress")).SendKeys(username);
                driver.FindElement(By.Id("Password")).SendKeys(password);

                driver.FindElement(By.XPath(".//*[@id='login']/ul/li[4]/button")).Click();
            }

            public void SelectAccountandExpenses()
            {
                Thread.Sleep(500);
                driver.FindElement(By.XPath(".//*[@id='content']/ul/li[2]/form/button")).Click();

                Thread.Sleep(500);
                driver.FindElement(By.LinkText("Expenses")).Click();

                Thread.Sleep(1000);
                
            }

            public void Broadband()
            {
                Thread.Sleep(1000);
                driver.FindElement(By.XPath(".//*[@id='add-item-links']/div/div/div[2]/div[2]/div/button")).Click();

                //broadband
                Thread.Sleep(2000);
                
                IWebElement week = driver.FindElement(By.Id("WeekEndingDateDisplay"));          

            IWebElement internet = driver.FindElement(By.Id("ExpenseParentCategoryId"));                
                internet.SendKeys(Keys.ArrowDown);
                internet.SendKeys(Keys.ArrowDown);
                internet.SendKeys(Keys.ArrowDown);
                internet.SendKeys(Keys.ArrowDown);
                
                //type
                IWebElement type = driver.FindElement(By.Id("ExpenseCategoryId"));                
                type.SendKeys(Keys.ArrowDown);
                type.SendKeys(Keys.ArrowDown);
                type.SendKeys(Keys.ArrowDown);
                type.SendKeys(Keys.ArrowDown);

                // description
                driver.FindElement(By.Id("Description")).SendKeys("Broadband");

                //amount
                driver.FindElement(By.Id("GrossAmount")).Clear();
                driver.FindElement(By.Id("GrossAmount")).SendKeys("4.00");
                Thread.Sleep(500);

                driver.FindElement(By.XPath("(//button[@type='button'])[2]")).Click();

                Thread.Sleep(1000);                
        }

            public void WeekTrainPass()
            {
            Thread.Sleep(2000);            
            driver.FindElement(By.XPath("//*[@id='add-receipted-item']")).Click();        

            //meals
            driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(10));
            //wait = new WebDriverWait(driver, new TimeSpan(0, 0, 5));
            //wait.Until(drv => drv.FindElement(By.Id("ExpenseParentCategoryId")));
            IWebElement week = driver.FindElement(By.Id("WeekEndingDateDisplay"));

            IWebElement internet = driver.FindElement(By.Id("ExpenseParentCategoryId"));
            internet.SendKeys(Keys.ArrowDown);
            internet.SendKeys(Keys.ArrowDown);
            internet.SendKeys(Keys.ArrowDown);        

            //type
            IWebElement type = driver.FindElement(By.Id("ExpenseCategoryId"));
            type.SendKeys(Keys.ArrowDown);

            // description
            driver.FindElement(By.Id("Description")).SendKeys("Weekly Pass to London");

            //amount
            driver.FindElement(By.Id("GrossAmount")).Clear();
            driver.FindElement(By.Id("GrossAmount")).SendKeys("176.00");
            Thread.Sleep(2000);

            driver.FindElement(By.XPath("(//button[@type='button'])[2]")).Click();

            Thread.Sleep(1000);
        }

            public void DeclareExpensesLunch()
            {
                Thread.Sleep(2000);
            driver.FindElement(By.XPath("//*[@id='add-receipted-item']")).Click();

            //meals
            driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(10));
                //wait = new WebDriverWait(driver, new TimeSpan(0, 0, 5));
                //wait.Until(drv => drv.FindElement(By.Id("ExpenseParentCategoryId")));
                IWebElement week = driver.FindElement(By.Id("WeekEndingDateDisplay"));

            IWebElement meals = driver.FindElement(By.Id("ExpenseParentCategoryId"));
                meals.SendKeys(Keys.ArrowDown);
                meals.SendKeys(Keys.ArrowDown);

                //type
                IWebElement type = driver.FindElement(By.Id("ExpenseCategoryId"));
                type.SendKeys(Keys.ArrowDown);
                type.SendKeys(Keys.ArrowDown);


                // description
                driver.FindElement(By.Id("Description")).SendKeys("Leon");

                //amount
                driver.FindElement(By.Id("GrossAmount")).Clear();
                driver.FindElement(By.Id("GrossAmount")).SendKeys("10.95");
                Thread.Sleep(500);

                driver.FindElement(By.XPath("(//button[@type='button'])[2]")).Click();

                Thread.Sleep(1000);

            }

            public void DeclareExpensesBreakfast()
            {           
                Thread.Sleep(2000);
            driver.FindElement(By.XPath("//*[@id='add-receipted-item']")).Click();

            //meals
            driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(10));
                
                IWebElement week = driver.FindElement(By.Id("WeekEndingDateDisplay"));

            IWebElement meals = driver.FindElement(By.Id("ExpenseParentCategoryId"));
                meals.SendKeys(Keys.ArrowDown);
                meals.SendKeys(Keys.ArrowDown);

                //breakfast
                IWebElement type = driver.FindElement(By.Id("ExpenseCategoryId"));
                type.SendKeys(Keys.ArrowDown);

                // description
                driver.FindElement(By.Id("Description")).SendKeys("Drink and porridge");

                //amount
                driver.FindElement(By.Id("GrossAmount")).Clear();

                driver.FindElement(By.Id("GrossAmount")).SendKeys("4.45");
                Thread.Sleep(500);

                driver.FindElement(By.XPath("(//button[@type='button'])[2]")).Click();

                Thread.Sleep(1000);

            }

            public void DeclareExpensesCoffee()
            {
                Thread.Sleep(2000);
            driver.FindElement(By.XPath("//*[@id='add-receipted-item']")).Click();

            //meals
            driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(10));
                //wait = new WebDriverWait(driver, new TimeSpan(0, 0, 5));
                //wait.Until(drv => drv.FindElement(By.Id("ExpenseParentCategoryId")));
                IWebElement week = driver.FindElement(By.Id("WeekEndingDateDisplay"));

            IWebElement meals = driver.FindElement(By.Id("ExpenseParentCategoryId"));
                meals.SendKeys(Keys.ArrowDown);
                meals.SendKeys(Keys.ArrowDown);

                //breakfast
                IWebElement type = driver.FindElement(By.Id("ExpenseCategoryId"));
                type.SendKeys(Keys.ArrowDown);

                // description
                driver.FindElement(By.Id("Description")).SendKeys("Coffee x 3");

                //amount
                driver.FindElement(By.Id("GrossAmount")).Clear();
                driver.FindElement(By.Id("GrossAmount")).SendKeys("9.75");
                Thread.Sleep(500);

                driver.FindElement(By.XPath("(//button[@type='button'])[2]")).Click();                
            }

            public void DeclareExpensesParking()
            {
                Thread.Sleep(2000);
            driver.FindElement(By.XPath("//*[@id='add-receipted-item']")).Click();

            driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(10));
                //wait = new WebDriverWait(driver, new TimeSpan(0, 0, 5));
                //wait.Until(drv => drv.FindElement(By.Id("ExpenseParentCategoryId")));
                IWebElement week = driver.FindElement(By.Id("WeekEndingDateDisplay"));

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
                driver.FindElement(By.Id("Description")).SendKeys("Parking at Hemel station");

                //amount
                driver.FindElement(By.Id("GrossAmount")).Clear();
                driver.FindElement(By.Id("GrossAmount")).SendKeys("6.70");
                Thread.Sleep(2000);

                driver.FindElement(By.XPath("(//button[@type='button'])[2]")).Click();

                Thread.Sleep(1000);
            }        

            public void DeclareExpensesDriveHemel()
            {
                Thread.Sleep(3000);
                driver.FindElement(By.XPath(".//*[@id='add-favourite-item']")).Click();                            

                IWebElement drive = driver.FindElement(By.Id("SelectedFavouriteId"));
                drive.SendKeys(Keys.ArrowDown);
                drive.SendKeys(Keys.ArrowDown);

                driver.FindElement(By.XPath("(//button[@type='button'])[2]")).Click();
                Thread.Sleep(3000);

                IWebElement week = driver.FindElement(By.Id("WeekEndingDateDisplay"));
                week.SendKeys(Keys.ArrowDown);

            //add mileage                
            driver.FindElement(By.XPath("(//button[@type='button'])[2]")).Click();

                Thread.Sleep(2000);
            }

            public void DeclarePhoneCalls()
            {
            Thread.Sleep(2000); //test
            driver.FindElement(By.XPath("//*[@id='add-receipted-item']")).Click();
            //phone calls
            driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(10));
            //wait = new WebDriverWait(driver, new TimeSpan(0, 0, 5));
            //wait.Until(drv => drv.FindElement(By.Id("ExpenseParentCategoryId")));
            IWebElement week = driver.FindElement(By.Id("WeekEndingDateDisplay"));

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

            driver.FindElement(By.Id("Description")).SendKeys("Calls to Offshore team");

            String phoneAmount = Interaction.InputBox("Enter Amount in 0.00 format", "Phone Calls Expenses", "0.00", 0, 0);

            //amount
            driver.FindElement(By.Id("GrossAmount")).Clear();
            driver.FindElement(By.Id("GrossAmount")).SendKeys(phoneAmount);
            Thread.Sleep(500);

            driver.FindElement(By.XPath("(//button[@type='button'])[2]")).Click();
        }

    }
}
