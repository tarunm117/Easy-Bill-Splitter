# Easy-Bill-Splitter
This Google Sheets app integration will easily split the cost of a bill amongst your friends, you just have to input the numbers in the bill and fill a matrix that records who had how much stuff.

## Steps to execute:
1. Create a Google Sheets spreadsheet with the following cells
<img width="743" height="148" alt="image" src="https://github.com/user-attachments/assets/b2d1fe7d-e3ab-4899-b62f-f16a2c4ffbb8" />

2. Go to Extensions > Apps Script

3. Replace the code in the Code Editor with the bill-splitter.js code

4. Create a button in the spreadsheet using Insert > Drawing

5. Click on the button once, click on the three dots option, click on assign script, enter the string "buildBillTableFull"

6. Fill the empty cells with relevant details contained in the bill and press the button

7. Once the script is executed, fill in details for the consumption matrix as shown below:
<img width="1191" height="119" alt="image" src="https://github.com/user-attachments/assets/8074cfaa-5b07-44d1-b4e8-e71dd420840b" />

8. Fill in values for the corresponding "Menu Price"

## That's it, you have your final split to be sent on your group chat
