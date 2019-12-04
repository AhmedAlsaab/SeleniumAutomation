using NUnit.Framework;
using System;

namespace AutomationTest.TestSuite.TestCases
{

    // Automation of Web Form 1,2 and 3

    [TestFixture]
    public class AutomationOrdered : DriverLogic
    {

        // sheetNumber represents a dynamic value that changes based on paramater
        // paramater value inserted by Automation GUI - Drop Down Selection
        


        // rownum specify which record to automate

        [OneTimeSetUp]
        public void SetUp(int sheetNumber)
        {
            SetupAndPrepareChromeDriver(sheetNumber);
        }

        [Test, Order(1)]
        public void ShortForm(int rowNum, int sheetNumber)
        {

            UserDetails userDetails = new UserDetails();
            userDetails.UserDetailsInput(rowNum, sheetNumber);
        }

        // .. Other tests omitted

        [OneTimeTearDown]
        public void TearDown()
        {
            try
            {
                System.Diagnostics.Debug.Write("Tearing down chrome");
                chrome.Close();
                chrome.Quit();
            }
            catch(NullReferenceException)
            {
                System.Diagnostics.Debug.Write("Null Referenece Caught when Tearing Down");
            }
           
        }
    }
}