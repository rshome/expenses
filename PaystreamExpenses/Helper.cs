using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;


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
