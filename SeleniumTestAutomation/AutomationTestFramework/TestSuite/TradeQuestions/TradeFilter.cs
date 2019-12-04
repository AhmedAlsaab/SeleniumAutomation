using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AutomationExcel;


namespace AutomationTest.TestSuite
{
    public class TradeFilter : TradeHandler
    {
        ExcelFileReader excelFileReader = new ExcelFileReader();

        // Storing list of trades into new lists that will be operated on
        // Then checking whether clients main or secondary trade matches a certain trade found inside the strings (checkTradeBlock)
        // If a match is found, then the input/methods corresponding to the found trade type are run (Block)


        public void FilterTradeQuestions(int row, int sheetNum)
        {
         // Redacted & Adjusted for Public Repo

            if (sheetNum ==1)
            {
                string ClientMainTrade = excelFileReader.ExcelLookup(10, row, sheetNum);
                string ClientSecondTrade = excelFileReader.ExcelLookup(12, row, sheetNum);

                List<String> checkTradeBlock1 = ListOfTrades.TradeBlock1;

                if (checkTradeBlock1.Where(x => checkTradeBlock1.Contains(ClientMainTrade) || checkTradeBlock1.Contains(ClientSecondTrade)).Any())
                {
                    Block1(row, sheetNum);
                }

                List<String> checkTradeBlock2 = ListOfTrades.TradeBlock2;

                if (checkTradeBlock2.Where(x => checkTradeBlock2.Contains(ClientMainTrade) || checkTradeBlock2.Contains(ClientSecondTrade)).Any())
                {
                    Block2(row, sheetNum);
                }

                List<String> checkTradeBlock3 = ListOfTrades.TradeBlock3;

                if (checkTradeBlock3.Where(x => checkTradeBlock3.Contains(ClientMainTrade) || checkTradeBlock3.Contains(ClientSecondTrade)).Any())
                {
                    Block3(row, sheetNum);
                }

                List<String> checkTradeBlock4 = ListOfTrades.TradeBlock4;

                if (checkTradeBlock4.Where(x => checkTradeBlock4.Contains(ClientMainTrade) || checkTradeBlock4.Contains(ClientSecondTrade)).Any())
                {
                    Block4(row, sheetNum);
                }

                List<String> checkTradeBlock5 = ListOfTrades.TradeBlock5;

                if (checkTradeBlock5.Where(x => checkTradeBlock5.Contains(ClientMainTrade) || checkTradeBlock5.Contains(ClientSecondTrade)).Any())
                {
                    Block5(row, sheetNum);
                }

            }
            
            if (sheetNum == 2 || sheetNum == 3)
            {

                //Web Form 2
                string clientMainTrade = excelFileReader.ExcelLookup(10, row, sheetNum);

               

                List<string> checkTradeBlock6 = ListOfTrades.TradeBlock6;

                if (checkTradeBlock6.Where(x => checkTradeBlock6.Contains(clientMainTrade)).Any())
                {
                    Block6(row, sheetNum);
                }

                List<string> checkTradeBlock7 = ListOfTrades.TradeBlock7;

                if (checkTradeBlock7.Where(x => checkTradeBlock7.Contains(clientMainTrade)).Any())
                {
                    Block7(row, sheetNum);
                }

                List<string> checkTradeBlock8 = ListOfTrades.TradeBlock8;

                if (checkTradeBlock8.Where(x => checkTradeBlock8.Contains(clientMainTrade)).Any())
                {
                    Block8(row, sheetNum);
                }
                List<string> checkTradeBlock9 = ListOfTrades.TradeBlock9;

                if (checkTradeBlock9.Where(x => checkTradeBlock9.Contains(clientMainTrade)).Any())
                {
                    Block9(row, sheetNum);
                }

                string checkTradeBlock10 = ListOfTrades.TradeBlock10;

                if(checkTradeBlock10 == clientMainTrade)
                {
                    Block10(row, sheetNum);
                }
                string checkTradeBlock11 = ListOfTrades.TradeBlock11;

                if(checkTradeBlock11 == clientMainTrade)
                {
                    Block11(row, sheetNum);
                }

                List<string> checkTradeBlock12 = ListOfTrades.TradeBlock12;

                if(checkTradeBlock12.Where(x => checkTradeBlock12.Contains(clientMainTrade)).Any())
                {
                    Block12(row, sheetNum);
                }

                List<string> checkTradeBlock13 = ListOfTrades.TradeBlock13;

                if(checkTradeBlock13.Where(x=> checkTradeBlock13.Contains(clientMainTrade)).Any())
                {
                     Block13(row, sheetNum);
                }

                List<string> checkTradeBlock14 = ListOfTrades.TradeBlock14;

                if(checkTradeBlock14.Where(x => checkTradeBlock14.Contains(clientMainTrade)).Any())
                {
                    Block14(row, sheetNum);
                }

                string checkTradeBlock15 = ListOfTrades.TradeBlock15;

                if(checkTradeBlock15 == clientMainTrade)
                {
                    Block15(row, sheetNum);
                }

                string checkTradeBlock16 = ListOfTrades.TradeBlock16;
                if(checkTradeBlock16 == clientMainTrade)
                {
                    Block16(row, sheetNum);
                }

                List<string> checkTradeBlock17 = ListOfTrades.TradeBlock17;
                if(checkTradeBlock17.Where(x => checkTradeBlock17.Contains(clientMainTrade)).Any())
                {
                    Block17(row, sheetNum);
                }
            }
        } 
    }
}
