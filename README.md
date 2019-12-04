# Test Automation with Selenium

The automation of data entry into  web-forms to assess quotes and element defects in batches. 

*Note: This is a lightweight repo that acts as a demo and showcase, most of the content and classes have been removed as it contained sensitive company related data. Alternatively, for a lightweight GUI-less version that fully works with different automation strategies (ordered, multithreaded and looped) view my DataDrivenTesting Repo!*


## Table of contents

* Introduction
* Build status
* Requirements
* Examples
* Installation
* How to run
* Video Guides
* FAQ
* Contributors

## Introduction

![GUI Preview](https://i.imgur.com/ecyYy4G.png)

Instead of going through the web-forms manually, one form at a time, all you would have to do is prepare data sets into the provide Excel WorkBook, boot up the Desk App and specify which data sets to run!

Whilst technically you would be doing the same on the web-form (entering data manually to assess quotes), the benefit of this approach:
* Do not have to wait for all the dynamic elements to load 
* No load screens if you wanted, or no web form at all with headless!
* Large quantities of datasets can be created rapidly
* Stored datasets can be used for regression testing 
* Provides a much faster way to compare, extract and analyse the different datasets and their results
* You can automate in large batches without needing to supervise it
* Enables stakeholders to manage dis-functional/broken elements to be fixed prior to deployment!
* Exceptions and logs are printed in a dedicated log file with time stamps

## Build Status

```Version 1 released on 19/11/2019```

## Examples

#### UserDetails.cs
``` c#
// Each class represents a section on the targetted web-form, this class represents the short-form
public void UserDetailsInput(int row, int sheetNum)
        {
            
            System.Diagnostics.Debug.WriteLine("Starting Quote, Automating Short Form");

// Storing exctracted Excel values into specified variables
// Finding element locations 
// Telling the logic handler how to automate the data set into the specified element
    
            WaitForPageLoad();

                // Contact Name
                string checkForContactName = excelFileReader.ExcelLookup(2, row, sheetNum);
                string contactName = "ContactName";
                WebdriverOperations(contactName, 4, checkForContactName);

                // Email
                string checkForEmail = excelFileReader.ExcelLookup(3, row, sheetNum);
                string email = "Email";
                WebdriverOperations(email, 4, checkForEmail);
        }
                ... omitted
```

#### DriverLogic.cs

```c#
// This is where the logic is handled
 public void WebdriverOperations(string elementLocation, int methodToUse, string whatToSend = "")
        {

            try
            {
                switch (methodToUse)
                {
                    // Find By Xpath and Click
                    case 1:
                       ... omitted

                    // Find By Xpath and Send Keys (Type into field)
                    case 2:
                       ... omitted
                        break;

                    // Find By Xpath, Scroll into view and Click
                    case 3:
                        ... omitted
                        break;

                    // Find By ID and Send Keys
                    case 4:
                        FindById(elementLocation).SendKeys(whatToSend);
                        Assert.AreEqual(FindById(elementLocation).GetAttribute("value"), whatToSend);
                        System.Diagnostics.Debug.WriteLine("The data has matched the expected input\n", whatToSend);
                        break;
                }
            ... omitted
```

#### AutomationOrdered.cs

```C#
// Test Fixture; where the data and the logic is invoked and put together to automate

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
```
#### AutomationControls.cs
``` C#
// Desk App Logic

// This particular section highlights how the row number and name of each data set is captured to pass in as arguements into the TestFixure (AutomationOrdered) later
   public void ExcelOperation(int sheetNum)
        {
            // Clear dict with each invocation (used as users change which data/form otherwise it stacks up)
            if(myDict != null)
            {
                myDict.Clear();
            }

            excel.Application Xapp = new excel.Application();
            excel.Workbook xWorkbook = Xapp.Workbooks.Open(ExcelFileLocationHome);

            // Get the row number for each data set 
            // Look in the A Column of each specified work sheet (this is where data set names are given by the user)
            try
            {
                int dataSetCount = 0;
                excel.Worksheet xWorksheet = xWorkbook.Sheets[sheetNum];
                int rows = xWorksheet.Cells.SpecialCells(excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
                excel.Range range = xWorksheet.get_Range("A1:A" + rows);
                foreach(excel.Range dataSet in range.Cells)
                {
                    if(dataSet.Value2 != null)
                    {
                        dataSetCount += 1;
                    }
                    System.Diagnostics.Debug.WriteLine(dataSetCount);
                }

                List<string> newList = new List<string>();

                // add the rownumber and given data set name in a dictionairy 
                // Data set names have row numbers. These row numbers are used as the rowNum arguement/parameter
                for (int i = 3; i <= dataSetCount; i++)
                {
                    excel.Range xRange = xWorksheet.UsedRange;
                    string extractedData = xRange.Cells[1][i].Value2.ToString();
                    int extractedRow = xRange.Cells[1][i].Row;

                    myDict.Add(extractedRow, extractedData);
                    newList.Add(extractedData);
                }
                DataSetSelector.DataSource = newList;
            } catch(Exception e)
            {
               // Exception
            }

            finally
            {
                System.Diagnostics.Debug.WriteLine("Closing Excel: Automation Controls");
                xWorkbook.Close(true);
                Xapp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(Xapp);
            }
        }
```
## Requirements

```.NET 4.5+```

``` Excel ```

## Installation

* Fork the repository and clone it on your own environment. Alternatively, download the ZIP and extract the files.
* Double click the 'automation-desk-app' Application to install and run the desk app

## How to run

1. ~Launch the TestData WorkBook found inside the Automation directory that you cloned or download~ Edit: Not available in public repo, create your own data sets and read through the code to see how data is looked up in Excel!
2. ~Select a Sheet you want to automate and start filling the rows out. Each row from start to the specified finish column represents one data set. The WorkBook is dynamic, has drop-downs and behaves the same way the website does, following the same rule set. Make sure you fill out all the white cells for each row and give each data set a name!~  Edit: Not available in public repo
3. ~After have one or multiple data sets ready, save and exit the WorkBook~
4. ~Launch the automation-desk-app~
5. ~Select which form you want to automate from the first dropdown~
6. ~The data sets you created and named should automatically load after you select the correct sheet from the desk-app~
7. ~Select which data set you want to automate~
8. ~Click start to start automating!~
9. ~After the progress bar reaches 100, click 'Finish' and open up the TestData WorkBook once again to see the quote results~
10. ~Once done, run the terminate-driver.bat file to clean up~

*Note: How to run is not applicablabe for this repo; but left in to highlight how to run this project when reproduced. Alternatively, for a lightweight GUI-less version that fully works with different automation strategies (ordered, multithreaded and looped) view my DataDrivenTesting Repo!*

## Video Guides & Tutorials for this project

Selenium & Excel
> Redacted

Windows Forms Desk-App
> Redacted


## FAQ

> I cannot get the desk application to run
* Most of the functionality and data in this project has been removed; the ideas and code is free to use under the MIT licence 

> Where is the rest of the FAQ?
* Hosted and viewable in full detail on the companies private Git



## Contributors

Ahmed Alsaab


## License

[MIT License]()

