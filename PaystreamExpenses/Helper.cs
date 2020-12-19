using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;

namespace PaystreamExpenses
{
    public class Helper
    {
        public IWebElement EnterData(IWebDriver driver, string id, string data)
        {            
            var element = driver.FindElement(By.Id(id));
            element.SendKeys(data);
            return element;
        }
    }
}
