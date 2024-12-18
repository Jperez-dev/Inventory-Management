
Sub GenerateWithdrawalForm()
    Dim withdrawalControlNumber As String
    Dim requestorName As String
    Dim withdrawalDate As String
    Dim fileName As String
    Dim filePath As String
    Dim newWorkbook As Workbook
    Dim sourceSheet As Worksheet
    Dim newSheet As Worksheet

    ' Set reference to the source sheet (Withdrawal Form)
    Set sourceSheet = ThisWorkbook.Sheets("Withdrawal_Form")

    ' Retrieve values from the form (Control Number, Requestor's Name, Date, etc.)
    withdrawalControlNumber = sourceSheet.Range("D4").Value
    withdrawalDate = sourceSheet.Range("Q4").Value
    requestorName = sourceSheet.Range("D20").Value

    ' Create the file name for the new Excel form
    fileName = withdrawalControlNumber & "_" & requestorName & ".xlsm"
    filePath = "C:\Users\admin\git-directories\Data_Management\Inventory_Management\Pending_Withdrawals\" & fileName ' Update with correct folder path

    ' Create a copy of the withdrawal form (Sheet2) as a new workbook
    sourceSheet.Copy
    Set newWorkbook = ActiveWorkbook
    Set newSheet = newWorkbook.Sheets(1) ' This will be the copied sheet
    
    ' Rename the sheet to avoid conflicts with the original workbook
    newSheet.Name = "Finalized_Withdrawal_Form"

    ' Clear formulas and leave only values
    Dim cell As Range
    For Each cell In newSheet.UsedRange
        If cell.HasFormula Then
            cell.Value = cell.Value ' Replace formulas with their values
        End If
    Next cell

    ' Paste formatting (fonts, colors, borders, etc.)
    newSheet.Cells.Copy
    newSheet.Cells.PasteSpecial Paste:=xlPasteFormats

    ' Copy column widths
    newSheet.Cells.Copy
    newSheet.Cells.PasteSpecial Paste:=xlPasteColumnWidths

    ' Copy data validation
    newSheet.Cells.Copy
    newSheet.Cells.PasteSpecial Paste:=xlPasteValidation

    ' Re-apply merged cells with values and formatting only
    For Each cell In sourceSheet.UsedRange
        If cell.MergeCells Then
            ' Copy values only for merged cells
            cell.MergeArea.Copy
            newSheet.Range(cell.Address).PasteSpecial Paste:=xlPasteValues
            
            ' Copy formatting for merged cells (if needed)
            newSheet.Range(cell.Address).PasteSpecial Paste:=xlPasteFormats
        End If
    Next cell

    ' Re-apply wrap text
    For Each cell In newSheet.UsedRange
        If cell.WrapText Then
            cell.WrapText = True
        End If
    Next cell

    ' Change the button functionality and label in the new workbook (not the source workbook)
    With newSheet
        Set SubmitButton = .Shapes("Button 1")
        SubmitButton.TextFrame.Characters.Text = "Finalize Withdrawal"
        SubmitButton.OnAction = "FinalizeWithdrawalForm" ' Point to the second functionality
    End With
    
    ' Save the new workbook as an .xlsm file (Macro-Enabled)
    newWorkbook.SaveAs filePath, FileFormat:=52 ' 52 is the constant for .xlsm (macro-enabled workbooks)

    ' Reset the form for new input
    sourceSheet.Range("C10:C19").ClearContents ' Clear Item Codes
    sourceSheet.Range("G10:O19").ClearContents
    
    ' Close the new workbook
    newWorkbook.Close SaveChanges:=True

    MsgBox "Withdrawal form has been generated and saved as " & fileName
    
    'Close this workbook
    ThisWorkbook.Close SaveChanges:=True
End Sub

Sub FinalizeWithdrawalForm()
    Dim withdrawalControlNumber As String
    Dim requestorName As String
    Dim withdrawalDate As String
    Dim pdfFileName As String
    Dim pdfFilePath As String
    Dim pdfFolder As String
    Dim approvalManager As String
    Dim approvalSpareParts As String
    Dim result As Integer
    Dim pythonScript As String
    Dim pythonArguments As String
    Dim tempFilePath As String
    Dim sourceSheet As Worksheet
    Dim openWb As Workbook
    Dim historyFile As String
    Dim trackingFolder As String
    Dim fileName As String
    Dim filePath As String

    ' Set reference to the source sheet (Withdrawal Form)
    Set sourceSheet = ActiveWorkbook.Sheets("Finalized_Withdrawal_Form")

    ' Retrieve values from the form (Control Number, Requestor's Name, Date, Approvals)
    withdrawalControlNumber = sourceSheet.Range("D4").Value
    withdrawalDate = sourceSheet.Range("Q4").Value
    requestorName = sourceSheet.Range("D20").Value
    approvalManager = sourceSheet.Range("G21").Value
    approvalSpareParts = sourceSheet.Range("O21").Value

    ' Ensure both approvals are not blank
    If approvalManager <> "Approved" And approvalManager <> "Denied" Or approvalSpareParts <> "Approved" And approvalSpareParts <> "Denied" Then
        MsgBox "Both approvals must be either 'Approved' or 'Denied'."
        Exit Sub
    End If

    ' Define tracking folder and history file name (with year-month format)
    trackingFolder = "C:\Users\admin\git-directories\Data_Management\Inventory_Management\Withdrawal_Tracking\"
    historyFile = trackingFolder & "Withdrawal_History_" & Format(Date, "yyyy-mm") & ".xlsx"

    ' Loop through all open workbooks and check if one matches the history file name
    On Error Resume Next ' Ignore errors in case the file is not open
    For Each openWb In Application.Workbooks
        If openWb.FullName = historyFile Then
            openWb.Close SaveChanges:=True ' Close the workbook if it matches
            MsgBox "The tracking history file was open and has been closed."
            Exit Sub ' Exit after closing the file
        End If
        
    Next openWb
    On Error GoTo 0 ' Reset error handling to default
    
    fileName = withdrawalControlNumber & "_" & requestorName & ".xlsm"
    filePath = "C:\Users\admin\git-directories\Data_Management\Inventory_Management\Pending_Withdrawals\" & fileName ' Update with correct folder path

    ' Prepare the Python script and arguments
    pythonScript = "C:\Users\admin\git-directories\Data_Management\Inventory_Management\Code\update_spare_parts_database.py"
    pythonArguments = filePath

    ' Run Python script with arguments (call from VBA)
    If approvalManager = "Approved" And approvalSpareParts = "Approved" Then
        result = Shell("python """ & pythonScript & """ " & pythonArguments, vbNormalFocus)
    End If

    ' Generate PDF file name
    pdfFolder = "C:\Users\admin\git-directories\Data_Management\Inventory_Management\Withdrawals\"
    pdfFileName = withdrawalControlNumber & "_" & requestorName & ".pdf"
    pdfFilePath = pdfFolder & pdfFileName
    
    Set exportRange = sourceSheet.Range("B3:S26")
    
    ' Ensure the content fits to a single page before exporting
    With sourceSheet.PageSetup
        .Orientation = xlPortrait ' Change to xlLandscape if required
        .Zoom = False ' Disable automatic zoom
        .FitToPagesWide = 1 ' Fit to 1 page wide
        .FitToPagesTall = 1 ' Fit to 1 page tall
        .LeftMargin = Application.InchesToPoints(0.31) ' Adjust left margin if necessary
        .RightMargin = Application.InchesToPoints(0.29) ' Adjust right margin if necessary
        .TopMargin = Application.InchesToPoints(0.3) ' Adjust top margin if necessary
        .BottomMargin = Application.InchesToPoints(0.3) ' Adjust bottom margin if necessary
    End With

    ' Save the form as PDF
    exportRange.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfFilePath

    ' Prepare the Python script and arguments
    pythonScript2 = "C:\Users\admin\git-directories\Data_Management\Inventory_Management\Code\generate_tracking_table.py"
    pythonArguments2 = withdrawalControlNumber & " " & _
                      requestorName & " " & _
                      withdrawalDate & " " & _
                      pdfFileName & " " & _
                      approvalManager & " " & _
                      approvalSpareParts

    ' Run Python script with arguments (call from VBA)
    result = Shell("python """ & pythonScript2 & """ " & pythonArguments2, vbNormalFocus)

    ' After successful execution, delete the temporary Excel file in Pending_Withdrawal folder
    tempFilePath = "C:\Users\admin\git-directories\Data_Management\Inventory_Management\Pending_Withdrawals\" & withdrawalControlNumber & "_" & requestorName & ".xlsm"
    
    ' Close the workbook if it's still open
    On Error Resume Next ' Ignore errors (in case it's already closed)
    Set newWorkbook = Workbooks(withdrawalControlNumber & "_" & requestorName & ".xlsm")
    If Not newWorkbook Is Nothing Then
        newWorkbook.Close SaveChanges:=True ' Close the workbook without saving changes
    End If
    On Error GoTo 0 ' Reset error handling
    
    If Dir(tempFilePath) <> "" Then
        Kill tempFilePath ' Delete the file
    End If

    MsgBox "Withdrawal form has been approved and history updated successfully."
End Sub

