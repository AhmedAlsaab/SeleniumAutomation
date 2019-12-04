# Test Automation with Selenium

The automation of data entry into insurance web-forms to assess quotes and element defects in batches.


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

Instead of going through the web-forms manually, one form at a time, all you would have to do is prepare data sets into the provide Excel WorkBook, boot up the Desk App and specify which data sets to run!

Whilst technically you would be doing the same on the web-form (entering data manually to assess quotes), the benefit of doing it this way is that you 
* Do not have to wait for all the dynamic elements to load 
* No load screens if you wanted, or no web form at all!
* Large quantities of datasets can be created rapidly
* Stored datasets can be used for regression testing 
* Provides a much faster way to compare the different datasets vs quote results; allowing stakeholders to quickly 
understand what sort of quotes potential customers would be getting. 
* You can automate in large batches without needing to supervise it
* Enables IT to manage any 
dis-functional/broken elements to be fixed prior to deployment!

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

1. Launch the TestData WorkBook found inside the Automation directory that you cloned or download
2. Select a Sheet you want to automate and start filling the rows out. Each row from start to the specified finish represents one data set. The WorkBook is dynamic, has drop-downs and behaves the same way the website does, following the same rule set. Make sure you fill out all the white cells for each row and give each data set a name!
3. After have one or multiple data sets ready, save and exit the WorkBook
4. Launch the automation-desk-app
5. Select which form you want to automate from the first dropdown
6. The data sets you created and named should automatically load after you select the correct sheet from the desk-app
7. Select which data set you want to automate
8. Click start to start automating!
9. After the progress bar reaches 100, click 'Finish' and open up the TestData WorkBook once again to see the quote results
10. Once done, run the terminate-driver.bat file to clean up

## Video Guides

Selenium & Excel
> https://youtu.be/VivIxP6gPA4

Windows Forms Desk-App
>https://youtu.be/7Fjx9RfkYnw


## FAQ

>I cannot get the desk application to run
* Make sure you have the correct .NET version installed

>I don't know how to use the WorkBook
* Each section within each sheet represents an input section on the corresponding web-form
* Based on what you enter, other shells should go either gray or white 
* Gray means do not enter data in this cell for this row, white implies the opposite
* Enter data for the whole row, from Column A to the specified end
* Watch the first part of the Selenium and Excel video guide 

> I cannot find my data set inside the Automation App
* Make sure you created it in the correct sheet that you are selected inside the app
* Make sure you gave the data set a name!

> How do I load a new data set into the web app ?
* After creating your new data-set, save and exit Excel
* Re-boot the desk application
* Select the form (sheet) where the data set is located in from the app's dropdown

> The Desk App is failing halway, at a certain stage or not working at all!
* This can mean a number of things
* Did you enter the correct data inside the WorkBook? 
* Something might have changed with the page structure of the form your trying to automate

> I keep getting NoQuotes!
* This means that the data set you prepared and automated is not eligible for online quotes
* Either change the data set or keep it for further analysis

> I am getting or keep getting Error Quotes!
* Is the website offline?
* Have you tried alteast twice?
* Check the other FAQ and alternatively contact IT

> My System is behaving slow/sluggish after automation!
* Run the termindate-driver.bat file by double clicking it
* This should clean background/underlying procedures up

> I accidentally clicked on the close button or want to click the close button whilst automating
* Run terminate-driver.bat

>Some of the rules inside the WorkBook are incorrect or don't follow the rules on the website!
* Contact IT

> I am having persisted issues with the application or it still doesn't work!
* Contact IT

> Suggestions? Improvements?
* Contact IT




## Contributors

Ahmed Alsaab


## License

[MIT License]()

