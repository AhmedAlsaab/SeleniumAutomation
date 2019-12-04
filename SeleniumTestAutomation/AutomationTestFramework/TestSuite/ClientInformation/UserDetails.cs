using AutomationExcel;


namespace AutomationTest.TestSuite
{

    public class UserDetails : DriverLogic
    {
        
        ExcelFileReader excelFileReader = new ExcelFileReader();

        // User Details 
        
        public void UserDetailsInput(int row, int sheetNum)
        {
            // Redacted and adjusted for public repo
            
            System.Diagnostics.Debug.WriteLine("INSERTING USER DETAILS\n");

            WaitForPageLoad();

                // Contact Name
                string checkForContactName = excelFileReader.ExcelLookup(2, row, sheetNum);
                string contactName = "ContactName";
                WebdriverOperations(contactName, 4, checkForContactName);

                // Email
                string checkForEmail = excelFileReader.ExcelLookup(3, row, sheetNum);
                string email = "Email";
                WebdriverOperations(email, 4, checkForEmail);
                        
                // Company Name
                string checkForCompanyName = excelFileReader.ExcelLookup(4, row, sheetNum);
                string companyName = "CompanyName";
                WebdriverOperations(companyName, 4, checkForCompanyName);

                // Mobile Number
                string checkForMobileNumber = excelFileReader.ExcelLookup(5, row, sheetNum);
                string mobileNumber = "Mobile";
                WebdriverOperations(mobileNumber, 4, checkForMobileNumber);

              

                // Submit 1st Form 
                string submitBtn = "//*[@id='quoteForm']/div[6]/div[2]/button";
                WebdriverOperations(submitBtn, 1);
          

                WaitForPageLoad();


                //// Corresponde Address (POSTCODE)
                string checkPostcode = excelFileReader.ExcelLookup(6, row, sheetNum);
                string ps = "Postcode";
                WebdriverOperations(ps, 4, checkPostcode);

          
        }
    }
}
