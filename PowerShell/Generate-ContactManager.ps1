<#
.SYNOPSIS
  Generates a macro-enabled Excel contact manager workbook with two sheets, a table, form-control buttons, and injected VBA.

.PARAMETER OutputPath
  Full path (including filename) where the .xlsm will be saved. Defaults to .\ContactManager.xlsm.

.EXAMPLE
  .\Generate-ContactManager.ps1 -OutputPath "C:\Temp\ContactManager.xlsm"
#>

param(
    # Define the output path for the Excel file, defaulting to the current folder
    [string]$OutputPath = "$(Get-Location)\ContactManager.xlsm"
)

# Create an Excel COM object (hidden instance of Excel)
$excel = New-Object -ComObject Excel.Application
$excel.Visible       = $false      # Run Excel in background
$excel.DisplayAlerts = $false      # Suppress Excel alert dialogs

try {
    # 1) Create a new blank workbook
    $wb = $excel.Workbooks.Add()

    # 2) Add two custom sheets: "Contacts" (data storage) and "Form" (data entry)
    $wsContacts = $wb.Worksheets.Add()
    $wsContacts.Name = "Contacts"
    $wsForm     = $wb.Worksheets.Add()
    $wsForm.Name = "Form"

    # 3) Delete any extra default sheets (e.g., "Sheet1", "Sheet2")
    foreach ($sheet in @($wb.Worksheets)) {
        if ($sheet.Name -notin @("Contacts","Form")) {
            $sheet.Delete()
        }
    }

    # 4) Create a table in the "Contacts" sheet
    #    Add column headers and turn them into a ListObject (Excel table)
    $wsContacts.Range("A1:E1").Value2 = @("ID","FirstName","LastName","Email","Phone")
    $lo = $wsContacts.ListObjects.Add(1, $wsContacts.Range("A1:E1"), $null, 1)
    $lo.Name = "ContactDB"   # Name the table for easy reference in VBA

    # 5) Set up labels and named ranges on the "Form" sheet for data entry
    #    Each named range corresponds to a field in the "Contacts" table
    $wsForm.Range("B2").Value2 = "ID:"     ; $wsForm.Range("C2").Name = "Form_ID"
    $wsForm.Range("B3").Value2 = "First:"  ; $wsForm.Range("C3").Name = "Form_First"
    $wsForm.Range("B4").Value2 = "Last:"   ; $wsForm.Range("C4").Name = "Form_Last"
    $wsForm.Range("B5").Value2 = "Email:"  ; $wsForm.Range("C5").Name = "Form_Email"
    $wsForm.Range("B6").Value2 = "Phone:"  ; $wsForm.Range("C6").Name = "Form_Phone"

    # 6) Add form-control buttons on the "Form" sheet
    #    Each button is tied to a VBA macro for managing contacts
    $buttons = @(
        @{Caption="Initialize"; Macro="InitApp"},
        @{Caption="New";        Macro="AddNewContact"},
        @{Caption="Save";       Macro="SaveContact"},
        @{Caption="Load";       Macro="LoadContact"},
        @{Caption="Delete";     Macro="DeleteContact"},
        @{Caption="Clear";      Macro="ClearForm"}
    )

    # Define layout for button placement
    $left   = 250
    $top    = 40
    $width  = 80
    $height = 24

    # Loop through each button definition and add it to the sheet
    for ($i = 0; $i -lt $buttons.Count; $i++) {
        # Add button (using Buttons.Add rather than Shapes.AddFormControl)
        $btn = $wsForm.Buttons().Add(
            $left,
            $top + ($i * ($height + 4)),  # Stack vertically with spacing
            $width,
            $height
        )
        $btn.Caption  = $buttons[$i].Caption  # Button text
        $btn.OnAction = $buttons[$i].Macro    # Macro it triggers when clicked
    }

    # 7) Store VBA code as a string (to inject into workbook later)
    #    This code defines macros for adding, saving, loading, deleting contacts, etc.
    $vba = @'
Option Explicit
' VBA code omitted (per your request)
'@

    # 8) Inject the VBA code into a new module in the workbook
    #    vbext_ct_StdModule = 1 means we’re adding a standard code module
    $vbModule = $wb.VBProject.VBComponents.Add(1)
    $vbModule.Name = "ContactManager"
    $vbModule.CodeModule.AddFromString($vba)

    # 9) Save the workbook as macro-enabled (.xlsm)
    #    FileFormat 52 = xlOpenXMLWorkbookMacroEnabled
    $wb.SaveAs($OutputPath, 52)

} finally {
    # Cleanup: close workbook, quit Excel, and release COM objects to prevent memory leaks
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
    [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# Confirmation message
Write-Host "Workbook generated and saved to $OutputPath" -ForegroundColor Green
