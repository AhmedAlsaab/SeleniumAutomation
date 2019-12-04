using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using AutomationExcel;
using SE = SeleniumExtras;
using System.Threading;
using NUnit.Framework;
using System.Collections.Generic;
using System.Text;
using OpenQA.Selenium.Chrome;
using System.IO;

namespace AutomationTest.TestSuite
{
   public class DriverLogic
    {
        public string chromeDriverPath = @"C:\Automation";
        private static string URL_Form1 = "URL to start Automation";
        private static string URL_Form2 = "Other URL for Automation";
        public static IWebDriver chrome;
        readonly int secondsToWait = 5;
        ExcelFileReader excelFileReader = new ExcelFileReader();
        


        public void SetupAndPrepareChromeDriver(int sheetNum)
        {
            var chromeOptions = new ChromeOptions();
            
            // Chrome Driver Options
            // Comment or uncomment to add or remove
            chromeOptions.AddArguments(new List<string>() {
           "--window-size=1920,1080",
           "--start-maximized",
            //  "--proxy-server='direct://'",
           "--disable-extensions",
            //  "--proxy-bypass-list=*",
            //   "--disable-gpu",
            //     "no-sandboxgit st
            //   "headless"
            });

            try
            {
                if (sheetNum == 2)
                {
                    System.Diagnostics.Debug.WriteLine("Starting Form 1 Automation");
                    chrome = new ChromeDriver(chromeDriverPath, chromeOptions)
                    {

                        Url = URL_Form1
                    };
                }
                else if (sheetNum == 1 || sheetNum == 3)
                {
                    System.Diagnostics.Debug.WriteLine("Starting Form 2 Automation");
                    chrome = new ChromeDriver(chromeDriverPath, chromeOptions)
                    {

                        Url = URL_Form2
                    };
                }
            }
            catch (WebDriverException)
            {
                System.Diagnostics.Debug.WriteLine("Caught WebDriver Related Failure ");
            }
            
            
        }

        // Method to find via Xpath
        // Stand alone method so that it can be reused
        public IWebElement FindByXpath(string elementLocation)
        {
            var waitForElementToLoad = new WebDriverWait(chrome, TimeSpan.FromSeconds(secondsToWait));
            IWebElement elementXpath = waitForElementToLoad.Until(SE.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.XPath(elementLocation)));
            return elementXpath;
        }

        // Method to find via ID
        public IWebElement FindById(string elementLocation)
        {
            var waitForElementToLoad = new WebDriverWait(chrome, TimeSpan.FromSeconds(secondsToWait));
            IWebElement elementByID = waitForElementToLoad.Until(SE.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.Id(elementLocation)));
            return elementByID;
        }

        // Logic handler based on switch case number given
        // 'whatToSend' parameter is optional
        public void WebdriverOperations(string elementLocation, int methodToUse, string whatToSend = "")
        {

            try
            {
                switch (methodToUse)
                {
                    // Find By Xpath and Click
                    case 1:
                        FindByXpath(elementLocation).Click();
                        Assert.IsTrue(FindByXpath(elementLocation).Displayed);
                        break;

                    // Find By Xpath and Send Keys (Type into field)
                    case 2:
                        FindByXpath(elementLocation).SendKeys(whatToSend);
                        Assert.AreEqual(FindByXpath(elementLocation).GetAttribute("value"), whatToSend);

                        System.Diagnostics.Debug.WriteLine("Assessing");
                        System.Diagnostics.Debug.WriteLine("Value", FindByXpath(elementLocation).GetAttribute("value"));
                        System.Diagnostics.Debug.WriteLine("What To Send", whatToSend);
                        break;

                    // Find By Xpath, Scroll into view and Click
                    case 3:
                        ((IJavaScriptExecutor)chrome).ExecuteScript("arguments[0].scrollIntoView(true);", FindByXpath(elementLocation));
                        ((IJavaScriptExecutor)chrome).ExecuteScript("arguments[0].click();", FindByXpath(elementLocation));
                        break;

                    // Find By ID and Send Keys
                    case 4:
                        FindById(elementLocation).SendKeys(whatToSend);
                        Assert.AreEqual(FindById(elementLocation).GetAttribute("value"), whatToSend);
                        System.Diagnostics.Debug.WriteLine("The data has matched the expected input\n", whatToSend);
                        break;

                    // Find by ID and Click
                    case 5:
                        FindById(elementLocation).Click();
                        Assert.IsTrue(FindById(elementLocation).Displayed);
                        break;

                    // Find By ID and Pick Date
                    case 6:
                        ((IJavaScriptExecutor)chrome).ExecuteScript("arguments[0].scrollIntoView(true);", FindById(elementLocation));
                        ((IJavaScriptExecutor)chrome).ExecuteScript("document.getElementById('" + elementLocation + "').setAttribute('value', '" + whatToSend + "')");
                        break;
                }
            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.Write("ELEMENT NOT FOUND\n", elementLocation);
                ErrorLogging(e);
            }
        }

        // Logging error when exceptions are thrown
        public static void ErrorLogging(Exception ex)
        {
            string errorLogPath = @"C:\YourDirectory\ErrorLog.txt";

            if (!File.Exists(errorLogPath))
            {
                File.Create(errorLogPath).Dispose();
            }
            using (StreamWriter sw = File.AppendText(errorLogPath))
            {
                sw.WriteLine("=============Error Logging ===========");
                sw.WriteLine("===========Start============= " + DateTime.Now);
                sw.WriteLine("Error Message: " + ex.Message);
                sw.WriteLine("Stack Trace: " + ex.StackTrace);
                sw.WriteLine("===========End============= " + DateTime.Now);

            }
        }

        // Page Waiter
        // Can be used as support for conditional additional waits
            public void WaitForPageLoad()
        {
            try
            {
                new WebDriverWait(chrome, TimeSpan.FromMinutes(1)).Until(
                 d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));
            }
             catch (ArgumentNullException)
            {
                System.Diagnostics.Debug.WriteLine("Caught Receive Failure (Automation stopped?)");
            }
           
        }


        public void SaveAsNoQuote(int row, int sheetNum)
        {
            excelFileReader.SaveToExcel(row, 80, "No Quote", sheetNum);
            excelFileReader.SaveToExcel(row, 81, "No Quote", sheetNum);
            excelFileReader.SaveToExcel(row, 82, "No Quote", sheetNum);
        }



        // saving results based on URL given at the result screen
        // results written to Excel
        public void SaveQuoteResults(int row, int column, string elementLocation, int methodToUse, int sheetNum)
        {
            string noQuotePL = "No Results URL - Web Form 1";
            string noQuotePI = "No Results URL - Web Form 2";
            string noQuoteSM = "No Results URL - Web Form 3";
            string ErrorQuote = "http://quote.test.constructaquote.com/Error/Index";
            string CurrentURL = chrome.Url;
            var waitForQuoteResult = new WebDriverWait(chrome, TimeSpan.FromSeconds(30));
            var listThreadIds = Thread.CurrentThread.ManagedThreadId.ToString();
            StringBuilder companyIdBuilder = new StringBuilder();
            StringBuilder quotePriceBuilder = new StringBuilder();


            if (CurrentURL == noQuotePL)
            {
                System.Diagnostics.Debug.WriteLine("No Quote: Web Form 1 - Check input data");
                SaveAsNoQuote(row, sheetNum);
            }

            if (CurrentURL == noQuotePI)
            {
                System.Diagnostics.Debug.WriteLine("No Quote: Web Form 2 - Change input data");
                SaveAsNoQuote(row, sheetNum);
            }

            if (CurrentURL == noQuoteSM)
            {
                System.Diagnostics.Debug.WriteLine("No Quote: Web Form 3 - Change input data");
                SaveAsNoQuote(row, sheetNum);
            }

            if (CurrentURL == ErrorQuote)
            {
                System.Diagnostics.Debug.WriteLine("Something went wrong!");
            }
            switch (methodToUse)
            {
                case 1:
                    if (CurrentURL != ErrorQuote && CurrentURL != noQuotePI && CurrentURL !=noQuotePL)
                    {
                        try
                        {
                            WaitForPageLoad();

                            IList<IWebElement> foundElements = chrome.FindElements(By.XPath(elementLocation));
                            foreach (IWebElement element in foundElements)
                            {
                                if (element.Displayed && element.Enabled)
                                {
                                    var result = element.Text;
                                    quotePriceBuilder.Append(result + "\n");

                                    // Add each result into string builder
                                    System.Diagnostics.Debug.WriteLine("Quote Price OR Quote Reference: ", result + "\n");
                                    System.Diagnostics.Debug.WriteLine("Saved Quote Results\n");
                                }
                            }
                        }
                        catch (Exception e)
                        {
                            // Do something
                            System.Diagnostics.Debug.Write("Could not save quote for data set running on thread: ", listThreadIds);

                        }
                        finally
                        {
                            // Once all loops are done
                            // Save stringbuilder to excel
                            excelFileReader.SaveToExcel(row, column, quotePriceBuilder.ToString(), sheetNum);
                            System.Diagnostics.Debug.Write("Finished attempting to save\n");
                        }
                    }
                    break;

                case 2:
                    if (CurrentURL != ErrorQuote && CurrentURL != noQuotePL && CurrentURL != noQuotePI)
                    {
                        try
                        {
                            WaitForPageLoad();

                            IList<IWebElement> foundElements = chrome.FindElements(By.XPath(elementLocation));
                            foreach (IWebElement element in foundElements)
                            {
                                if (element.Displayed && element.Enabled)
                                {
                                    if(sheetNum == 1 || sheetNum == 3)
                                    {
                                        string companyListForPL = element.GetAttribute("data-id");
                                        companyIdBuilder.Append(companyListForPL + "\n");
                                        System.Diagnostics.Debug.WriteLine("Company Results ", companyIdBuilder.ToString() + "\n");
                                        System.Diagnostics.Debug.WriteLine("Saved Company Results\n");
                                    }
                                    else if (sheetNum == 2)
                                    {
                                        string companyListForPI = element.Text;
                                        companyIdBuilder.Append(companyListForPI + "\n");
                                        System.Diagnostics.Debug.WriteLine("Company Results: ", companyIdBuilder.ToString() + "\n");
                                        System.Diagnostics.Debug.WriteLine("Saved Company Results\n");
                                    }
                                }
                            }
                        }
                        catch (Exception e)
                        {
                            // Do something
                            System.Diagnostics.Debug.Write("Could not save company for data set running on thread: ", listThreadIds);

                        }
                        finally
                        {
                            // Once all loops are done
                            // Save stringbuilder to excel
                            excelFileReader.SaveToExcel(row, column, companyIdBuilder.ToString(), sheetNum);
                            System.Diagnostics.Debug.Write("Finished attempting to save\n");
                        }
                    }
                    break;
            }
        }
    }
}



