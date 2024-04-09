
This VBA macro automates the process of extracting data from the Customer Service Interface (CSI) diagnostic tool. 

**Browser Interaction:**
Starts a Chrome browser session and navigates to the Amazon CSI diagnostic tool webpage (https://csi.*******.com/diag/Knowhere_Deals?resourcePath=Knowhere_Deals).
Waits for specific elements to become present on the page before proceeding.

**Data Retrieval:**
Iterates over each row of data in the "Sheet1" worksheet of the Excel workbook.
Scrolls to specific sections of the webpage and clicks on elements to access deal details.
Enters deal information such as deal ID and marketplace into input fields and submits the form.
Waits for the data to load and be displayed on the page.

**Data Extraction:**
Finds and extracts deal-related information from the webpage, such as deal status, escalation count, and last update date.
Populates the worksheet with the extracted data.
