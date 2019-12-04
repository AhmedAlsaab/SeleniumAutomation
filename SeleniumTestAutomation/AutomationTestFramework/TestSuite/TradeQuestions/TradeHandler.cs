using AutomationExcel;


namespace AutomationTest.TestSuite
{
    //... Adjusted and redacted, this class has some content left in for examples

    public class TradeHandler : DriverLogic
    {
        readonly string yes = "1";
        ExcelFileReader excelFileReader = new ExcelFileReader();

        // Handles Trades in Block 1
        public void Block1(int row, int sheetNum)
        {
            System.Diagnostics.Debug.WriteLine("CHECKING FOR TRADE TYPE QUESTIONS, DATA SET: BLOCK 1");
            
            // primary question for this trade block
            string data_A = excelFileReader.ExcelLookup(24, row, sheetNum);
            string element_A = "//input[contains(@class, 'element_a_classname') and contains(@value, '" + data_A + "')]/parent::label";
            WebdriverOperations(element_A, 1, data_A);
            if (data_A == yes)
            {
                // How many x peform y? 
                string data_b = excelFileReader.ExcelLookup(25, row, sheetNum);
                string element_b = "//select[contains(@class, 'element_b_classname')]/option[contains(text(), '" + data_b + "')]";
                WebdriverOperations(element_b, 1, data_b);

                // how many xa perform yx?
                string data_c = excelFileReader.ExcelLookup(26, row, sheetNum);
                string element_c = "//input[contains(@class, 'element_c_classname') and contains(@value, '" + data_c + "')]/parent::label";
                WebdriverOperations(element_c, 1, data_c);

                if (data_c == yes)
                {
                    // sub question
                    string data_d = excelFileReader.ExcelLookup(27, row, sheetNum);
                    string element_d = "//input[contains(@class, 'element_d_classname')]";
                    WebdriverOperations(element_d, 2, data_d);
                }
            }
        }

       // Handles Trades in Block 2

        public void Block2(int row, int sheetNum)
        {
           // ... omitted
           // ... Block 2 to 20 redacted
        }

        // omitted
    }
}
