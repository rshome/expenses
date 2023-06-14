using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

namespace PaystreamExpenses
{
    public class Helper
    {
        IWebDriver driver = new ChromeDriver();
        public IWebElement EnterData(IWebDriver driver, string id, string data)
        {            
            var element = driver.FindElement(By.Id(id));
            element.SendKeys(data);
            return element;
        }
    }
}
