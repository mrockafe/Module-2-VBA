# Module-2-VBA (Read in Code View)
'Module 2 VS Codes and Result Images
'In this repository you will find 4 VS Code Scripts and 3 screenshots for result visuals
'The first level of code is "Sub_Ticker_Group().vb" 
  'This code focused on the similar tickers and grouping them together 
    'Secondly the last close date and first open date are located and saved as a variable
    'Third these values are used to find the Yearly change
    'Fourth the code sums the volume per created ticker group 
    'Fifth the code calculates the percent change via the saved open and closed values with percent formatting 
    'Lastly the results table is created and the data is populated
'The next level of code is "Sub Overall Results().vb"
  'This code focused on making the overall results table 
    'The code creates variables for max and min percent changes & max volume
    'The table is created via cell position and the data is created from the created variable and inserted text
'The next level of code is "Sub All_Sheets().vb"
   'This code creates a second sub that goes through all sheets, sets them to active, then runs the main Sub on each sheet
'The last level of code is "Sub Module_2_Final().vb"
  'This code level combines all previous levels to run the full sub and overview per sheet.
    'There is added code to re-run the numbers for the established Ticker groups to verify the data

'Lastly I have included in this repository a scrren shot for each of the following years - 2018,2019,2020 showing the full results for each sheet

'Note the full sheet data is to large to add to a github repository or the canvas website, that is why the screenshots are shared
'If testing the code is needed the file "alphabetical_testing.xlsm" has been added to t6his repository with the final code loaded in
