# VBA-Challenge
 **Brief Summary**
 In this project, VBA scripting is used to analyze historical stock price data across multiple years (2018,2019 and 2020). The results are presented in each of the corresponding worksheets. Finally, screenshots of the results for each of the worksheets is also posted on a separate document. 
 **Source Code Organization**
 The VBA script for this project is organized across two logical sections as explained below:
    1. Worksheet Definition and Variable Declarations: The initial section of the script declares the key variables to be used in the script along with essential elements of the header row, including related formating of the relevant fields. For performance optimization, screen updating and enabling events are turned OFF at the beginning of the code construct and set to ON only at the end of the code. 
    2. Looping Construct: The core part of the script is a looping construct that executes the required steps for each of the worksheets in order to publish the desired summary statistics for each year. 
**References**
1. "Top Ten Tips To Speed Up Your VBA Code", https://eident.co.uk/2016/03/top-ten-tips-to-speed-up-your-vba-code/, Eident Training, accessed Jan 2024
2. Excel VBA ColorIndex, https://analysistabs.com/excel-vba/colorindex/, accessed Dec 2023
3. Automate Excel.com, https://www.automateexcel.com/vba/autofit-columns-rows/, Steve Rynearson and Editorial Team, accessed Dec 2023
4. Visual Basic for Applications (VBA) Developer Documentation - Worksheet Object, https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet, accessed Dec 2023