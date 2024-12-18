## **Automated Withdrawal Form System**

Designed and implemented an automated withdrawal system using **Excel**, **VBA**, and **Python** to streamline the process of inventory withdrawals and approvals. The system integrates several powerful features, ensuring efficiency and accuracy in tracking and managing withdrawal requests.

### **Key Features & Workflow:**

•	**Initiation & Data Entry:**

The process begins by entering a name and date in **Sheet1** (_Withdrawal_Control_Number_), along with the corresponding control number. This control number is automatically transferred to **Sheet2** (_Withdrawal_Form_). Based on this number, the system automatically generates the _requestor's name_ and _date_ using pre-defined formulas in Excel.

•	**Automated Item Details:**

For **Sheet2** (_Withdrawal_Form_), the user only needs to input the _item code_ and other relevant details (_such as quantity to order and reason for withdrawal_). Excel formulas (e.g., **VLOOKUP**) automatically pull additional item information from an external spare parts database, populating fields such as item description, part number, available balance, and other necessary details.

•	**Request Submission:**

After entering the required data, the user can submit the withdrawal request by clicking a button. This action triggers a **VBA** function that copies **Sheet2** (_Withdrawal_Form_) to a **Pending Withdrawal** folder as a new Excel file for approval by both the manager and the spare parts officer. The submission process is limited to a maximum of 10 items per withdrawal.

•	**Approval & Finalization:**

Once the request is in the **Pending Withdrawal** folder, both authorized personnel (manager and spare parts officer) review the details. If both approve the request, the **Finalize** button could activate by spare parts officer. Upon finalization:

o	**Available Balance Update:** Python code, executed via **VBA**, automatically calculates and updates the available balance in the spare parts database by deducting the order quantity from the existing balance.

o	If the request is not approved by any authorized personnel, the withdrawal is denied, and no changes are made to the spare parts database.

•	**Post-Approval Actions:**

Regardless of approval, the system performs several critical post-processing actions:

o	**PDF Generation:** The withdrawal request form is automatically converted to a PDF and stored in the **Withdrawals** folder.

o	**Tracking & Logging:** Key details from the approved or denied request are extracted and logged into a **Tracking Excel file**, including the approval status and other relevant metadata.

### **Technologies Used:**

•	**Excel Formulas:** Utilized advanced formulas (e.g., **VLOOKUP, IF, INDEX,** and **error handling**) to automate data retrieval and calculations. Also used **Tables, Conditional Formatting,** and **Input Validation** for better visualization and user experience.

•	**VBA (Visual Basic for Applications):** Custom **VBA** functions handle the **Submit** and **Finalize** buttons, including state changes (e.g., enabling/disabling buttons based on approval). VBA also manages the copying of files and generates PDFs from the copied files. It performs formatting manipulations to handle complex formats and prevent errors or incorrect data entry.

•	**Python (with Openpyxl):**

o	Used **Openpyxl** to interact with Excel files for data manipulation.

o	Python scripts manage the logic for updating balances and generating/updating the tracking file.

o	Python libraries used include **openpyxl, os, sys, datetime,** and **openpyxl.styles** (for formatting).
