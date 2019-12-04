using NUnit.Framework;

namespace AutomationTest.TestSuite.TestCases
{

    // Automation in batch of web form 1,2 and 3

    [TestFixture]
    public class AutomationLooped : DriverLogic
    {
        

        UserDetails userDetails = new UserDetails();
        // ..redacted

      
       
        [Test]
        public void LoopedAutomation(int rowNum, int endRow, int sheetNum)
        {
            for (int i = rowNum; i < endRow; i++)
            {
                Setup(sheetNum);
                userDetails.UserDetailsInput(i, sheetNum);
                // ... Other sections to automate omitted 
                TearDown();
              
            }
        }

        public void Setup(int sheetNum)
        {
            SetupAndPrepareChromeDriver(sheetNum);
        }
        public void TearDown()
        {
            System.Diagnostics.Debug.Write("Tearing down chrome");
            chrome.Close();
            chrome.Quit();
        }
    }
}
