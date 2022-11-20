Sub Table2_a()

'// Creating the worksheets

    Sheets.Add after:=Sheets(Sheets.Count)
    Sheets(ActiveSheet.Name).Name = "T2"
    Sheets.Add after:=Sheets(Sheets.Count)
    Sheets(ActiveSheet.Name).Name = "T3"
    Sheets.Add after:=Sheets(Sheets.Count)
    Sheets(ActiveSheet.Name).Name = "T4"
    Sheets.Add after:=Sheets(Sheets.Count)
    Sheets(ActiveSheet.Name).Name = "T5"
    Sheets.Add after:=Sheets(Sheets.Count)
    Sheets(ActiveSheet.Name).Name = "T6"


Dim wb As Workbook
Dim wsSource As Worksheet, wsTarget As Worksheet
Dim LastRow As Long, LastColumn As Long
Dim SourceDataRange As Range
Dim PTCache As PivotCache
Dim PT As PivotTable
Dim pf_vlookup As PivotField


On Error GoTo errHandler
Set wb = ActiveWorkbook
Set wsTarget = wb.Worksheets("T2")
wsTarget.Select
wsTarget.Cells.Clear


'// Step 1. Define data soruce
Set wsSource = wb.Worksheets("Sales")
With wsSource
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
    LastColumn = .Cells(1, Columns.Count).End(xlToLeft).Column
    

'// Define source data object
Set SourceDataRange = .Range(.Cells(1, 1), .Cells(LastRow, LastColumn))
End With

'// 2. Set PT Cache
Set PTCache = wb.PivotCaches.Create(xlDatabase, SourceDataRange)

'// 3. Create Pivot Table
Set PT = PTCache.CreatePivotTable(wsTarget.Range("B5"), "T2A")
With PT
'// Show GrandTotals
    .ColumnGrand = True
    .RowGrand = True
    
    .RowAxisLayout xlTabularRow
    
    .TableStyle2 = "PivotStyleDark16"
    
    
    
    '// Add Pivot Fields
    
    '// Filters
    
    'If (ActiveSheet.PivotTables("T2A").PivotFields("vlookup").CurrentPage = "#N/A") Then
    
'    On Error Resume Next
'  pf_vlookup = ActiveSheet.PivotTable("T2A").PivotFields("vlookup").PivotItems("#N/A")
'  If Err = 0 Then ActiveSheet.PivotTables("T2A").PivotFields("vlookup").CurrentPage = "#N/A"
'  Else: ActiveSheet.PivotTables(T2A).PivotFields("vlookup").CurrentPage = "blank"
'    Err.Clear
'
'    End If
    
    With .PivotFields("vlookup")
    .Orientation = xlPageField
    .EnableMultiplePageItems = True
    End With

    Set pf_vlookup = PT.PivotFields("vlookup")

    pf_vlookup.ClearAllFilters

    '// Enable multiple filters selection
    pf_vlookup.EnableMultiplePageItems = True
    
    
    pf_vlookup.PivotItems("#N/A").Visible = False
    
    'Else
    
    
    'End If
    

    
    
    '// Rows Section
    
    With .PivotFields("Group")
    .Orientation = xlRowField
    .Subtotals(1) = False
    End With
    
    '// Columns
    'With .PivotFields("Group")
    '.Orientation = xlColumnField
    'End With
    
    '// Values
     With .PivotFields("Sales Units")
    .Orientation = xlDataField
    .Function = xlSum
    .NumberFormat = "#,###"
    End With




End With

CleanUp:
    Set PT = Nothing
    Set PTCache = Nothing
    Set SourceDataRange = Nothing
    Set wsSource = Nothing
    Set wsTarget = Nothing
    Set wb = Nothing
    Set pf_vlookup = Nothing
    
Exit Sub

errHandler:
    MsgBox "Error: " & Err.Description, vbExclamation
    GoTo CleanUp
End Sub

Sub test2_b()

Dim wb As Workbook
Dim wsSource As Worksheet, wsTarget As Worksheet
Dim LastRow As Long, LastColumn As Long
Dim SourceDataRange As Range
Dim PTCache As PivotCache
Dim PT As PivotTable
Dim pf_vlookup As PivotField
Dim pf_duplicate As PivotField
Dim vlookupArray(1) As String
Dim numberOfElements As Integer
Dim i As Integer
Dim j As Integer

vlookupArray(1) = "#N/A"


On Error GoTo errHandler
Set wb = ActiveWorkbook
Set wsTarget = wb.Worksheets("T2")
wsTarget.Select
'wsTarget.Cells.Clear


'// Step 1. Define data soruce
Set wsSource = wb.Worksheets("Complaints")
With wsSource
    LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
    LastColumn = .Cells(1, Columns.Count).End(xlToLeft).Column
    

'// Define source data object
Set SourceDataRange = .Range(.Cells(1, 1), .Cells(LastRow, LastColumn))
End With

'// 2. Set PT Cache
Set PTCache = wb.PivotCaches.Create(xlDatabase, SourceDataRange)

'// 3. Create Pivot Table
Set PT = PTCache.CreatePivotTable(wsTarget.Range("G5"), "T2B")
With PT
'// Show GrandTotals
    .ColumnGrand = True
    .RowGrand = True
    
    .RowAxisLayout xlTabularRow
    
    .TableStyle2 = "PivotStyleDark16"
    
    
    
    '// Add Pivot Fields
    
    '// Filters
    With .PivotFields("vlookup")
    .Orientation = xlPageField
'    .EnableMultiplePageItems = True
    End With
    
    With .PivotFields("Dup?")
    .Orientation = xlPageField
    .EnableMultiplePageItems = True
    End With
    
    Set pf_vlookup = PT.PivotFields("vlookup")
    Set pf_duplicate = PT.PivotFields("Dup?")
    
'    pf_vlookup.ClearAllFilters
    pf_duplicate.ClearAllFilters
    
    '// Enable multiple filters selection
    pf_vlookup.EnableMultiplePageItems = True
    pf_duplicate.EnableMultiplePageItems = True
        
'    pf_vlookup.PivotItems("#N/A").Visible = False
    pf_duplicate.CurrentPage = ""
        
           '// only apply filter if present in data
   
    numberOfElements = UBound(vlookupArray) - LBound(vlookupArray) + 1

If numberOfElements > 0 Then
    With pf_vlookup
        For i = 1 To pf_vlookup.PivotItems.Count
        j = 0
'            MsgBox pf_vlookup.PivotItems.Count
'            MsgBox pf_vlookup.PivotItems(i).Name
'            MsgBox vlookupArray(j)
            
        Do While j < numberOfElements
            
            
            If pf_vlookup.PivotItems(i).Name = vlookupArray(j) Then
            
                pf_vlookup.PivotItems(pf_vlookup.PivotItems(i).Name).Visible = False
                Exit Do
            Else
                pf_vlookup.PivotItems(pf_vlookup.PivotItems(i).Name).Visible = True
            End If
            j = j + 1
        Loop
        Next i
    End With
End If
    
    
    '// Rows Section
    
    With .PivotFields("Group")
    .Orientation = xlRowField
    .Subtotals(1) = False
    End With
    
    '// Columns
    'With .PivotFields("Group")
    '.Orientation = xlColumnField
    'End With
    
    '// Values
     With .PivotFields("Complaint ID")
    .Orientation = xlDataField
    .Function = xlCount
    .NumberFormat = "#,###"
    End With




End With

CleanUp:
    Set PT = Nothing
    Set PTCache = Nothing
    Set SourceDataRange = Nothing
    Set wsSource = Nothing
    Set wsTarget = Nothing
    Set wb = Nothing
    Set pf_vlookup = Nothing
    Set pf_duplicate = Nothing
    
Exit Sub

errHandler:
    MsgBox "Error: " & Err.Description, vbExclamation
    GoTo CleanUp
End Sub

Sub test2_c()

Dim wb As Workbook
Dim wsSource As Worksheet, wsTarget As Worksheet
Dim LastRow As Long, LastColumn As Long
Dim SourceDataRange As Range
Dim PTCache As PivotCache
Dim PT As PivotTable
Dim pf_vlookup As PivotField
Dim pf_arCode As PivotField
Dim pf_ARC As PivotField
Dim vlookupArray(1) As String
Dim arcArray(0 To 10) As String
Dim numberOfElements As Integer
Dim numberOfElementsTwo As Integer
Dim i As Integer
Dim j As Integer
Dim c As Integer
Dim d As Integer


vlookupArray(1) = "#N/A"

arcArray(0) = "No Consequences or Impact to Patient"
arcArray(1) = "No Known Impact Or Consequence To Patient"
arcArray(2) = ""
arcArray(3) = "Device No Known Device Problem"
arcArray(4) = "Device No Reported Allegation"
arcArray(5) = "Insufficient Information"
arcArray(6) = "No Clinical Signs, Symptoms or Conditions"
arcArray(7) = "No Code Available"
arcArray(8) = "No Health Consequences or Impact"
arcArray(9) = "No Information"
arcArray(10) = "No Patient Involvement"


On Error GoTo errHandler
Set wb = ActiveWorkbook
Set wsTarget = wb.Worksheets("T2")
wsTarget.Select
'wsTarget.Cells.Clear


'// Step 1. Define data soruce
Set wsSource = wb.Worksheets("Complaints")
With wsSource
LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
LastColumn = .Cells(1, Columns.Count).End(xlToLeft).Column


'// Define source data object
Set SourceDataRange = .Range(.Cells(1, 1), .Cells(LastRow, LastColumn))
End With

'// 2. Set PT Cache
Set PTCache = wb.PivotCaches.Create(xlDatabase, SourceDataRange)

'// 3. Create Pivot Table
Set PT = PTCache.CreatePivotTable(wsTarget.Range("K5"), "T2C")
With PT
'// Show GrandTotals
.ColumnGrand = True
.RowGrand = True

.RowAxisLayout xlTabularRow

.TableStyle2 = "PivotStyleDark16"



'// Add Pivot Fields

'// Filters
With .PivotFields("vlookup")
.Orientation = xlPageField
'    .EnableMultiplePageItems = True
End With

'    With .PivotFields("AR Code Description (GCMS)")
'    .Orientation = xlPageField
'    .EnableMultiplePageItems = True
'    End With

With .PivotFields("AR Code Description (GCMS)2")
.Orientation = xlPageField
'    .EnableMultiplePageItems = True
End With

    '// Rows Section

With .PivotFields("Group")
.Orientation = xlRowField
.Subtotals(1) = False
End With

'// Columns
'With .PivotFields("Group")
'.Orientation = xlColumnField
'End With

'// Values
 With .PivotFields("Complaint ID")
.Orientation = xlDataField
.Function = xlCount
.NumberFormat = "#,###"
End With

    '// ARC filter
Set pf_ARC = PT.PivotFields("AR Code Description (GCMS)2")


Set pf_vlookup = PT.PivotFields("vlookup")
'    Set pf_arCode = PT.PivotFields("AR Code Description (GCMS)")


pf_vlookup.ClearAllFilters
'    pf_arCode.ClearAllFilters

'// Enable multiple filters selection

pf_vlookup.EnableMultiplePageItems = True
'    pf_arCode.EnableMultiplePageItems = True

   pf_ARC.EnableMultiplePageItems = True



'// only apply filter if present in data

numberOfElementsTwo = UBound(arcArray) - LBound(arcArray) + 1
'    MsgBox numberOfElementsTwo

If numberOfElementsTwo > 0 Then
With pf_ARC
    For c = 1 To pf_ARC.PivotItems.Count

    d = 0
'            MsgBox pf_ARC.PivotItems.Count
'            MsgBox "Arc Array" + pf_ARC.PivotItems(c).Name
'            MsgBox d
        
    Do While d < numberOfElementsTwo
        
        

        If pf_ARC.PivotItems(c).Name = arcArray(d) Then
        
            pf_ARC.PivotItems(pf_ARC.PivotItems(c).Name).Visible = False
            Exit Do
        Else
            pf_ARC.PivotItems(pf_ARC.PivotItems(c).Name).Visible = True
        End If
        d = d + 1
    Loop
    Next c
End With
End If

           '// only apply filter if present in data

numberOfElements = UBound(vlookupArray) - LBound(vlookupArray) + 1

If numberOfElements > 0 Then
With pf_vlookup
    For i = 1 To pf_vlookup.PivotItems.Count
    j = 0
'            MsgBox pf_vlookup.PivotItems.Count
'            MsgBox pf_vlookup.PivotItems(i).Name
'            MsgBox vlookupArray(j)
        
    Do While j < numberOfElements
        
        
        If pf_vlookup.PivotItems(i).Name = vlookupArray(j) Then
        
            pf_vlookup.PivotItems(pf_vlookup.PivotItems(i).Name).Visible = False
            Exit Do
        Else
            pf_vlookup.PivotItems(pf_vlookup.PivotItems(i).Name).Visible = True
        End If
        j = j + 1
    Loop
    Next i
End With
End If



'    pf_vlookup.PivotItems("#N/A").Visible = False
'    pf_arCode.PivotItems("").Visible = False
'    pf_arCode.PivotItems("No Patient Involvement").Visible = False
'    pf_arCode.PivotItems("No Health Consequences or Impact").Visible = False
'    pf_arCode.PivotItems("Insufficient Information").Visible = False
'    pf_arCode.PivotItems("No Known Impact Or Consequence To Patient").Visible = False
'    pf_arCode.PivotItems("No Consequences or Impact to Patient").Visible = False
'    pf_arCode.PivotItems("No Clinical Signs, Symptoms or Conditions").Visible = False
'    pf_arCode.PivotItems("Device No Reported Allegation").Visible = False
'    pf_arCode.PivotItems("Device No Known Device Problem").Visible = False









End With

CleanUp:
Set PT = Nothing
Set PTCache = Nothing
Set SourceDataRange = Nothing
Set wsSource = Nothing
Set wsTarget = Nothing
Set wb = Nothing
Set pf_vlookup = Nothing
Set pf_arCode = Nothing

Exit Sub

errHandler:
MsgBox "Error: " & Err.Description, vbExclamation
GoTo CleanUp
End Sub

Sub test2_d()

Dim wb As Workbook
Dim wsSource As Worksheet, wsTarget As Worksheet
Dim LastRow As Long, LastColumn As Long
Dim SourceDataRange As Range
Dim PTCache As PivotCache
Dim PT As PivotTable
Dim pf_globalRei As PivotField
Dim pf_vlookup As PivotField
Dim pf_duplicate As PivotField
Dim vlookupArray(1) As String
Dim numberOfElements As Integer
Dim i As Integer
Dim j As Integer

vlookupArray(1) = "#N/A"


On Error GoTo errHandler
Set wb = ActiveWorkbook
Set wsTarget = wb.Worksheets("T2")
wsTarget.Select
'wsTarget.Cells.Clear


'// Step 1. Define data soruce
Set wsSource = wb.Worksheets("Complaints")
With wsSource
LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
LastColumn = .Cells(1, Columns.Count).End(xlToLeft).Column


'// Define source data object
Set SourceDataRange = .Range(.Cells(1, 1), .Cells(LastRow, LastColumn))
End With

'// 2. Set PT Cache
Set PTCache = wb.PivotCaches.Create(xlDatabase, SourceDataRange)

'// 3. Create Pivot Table
Set PT = PTCache.CreatePivotTable(wsTarget.Range("B35"), "T2D")
With PT
'// Show GrandTotals
.ColumnGrand = True
.RowGrand = True

.RowAxisLayout xlTabularRow

.TableStyle2 = "PivotStyleDark16"



'// Add Pivot Fields

'// Filters
With .PivotFields("vlookup")
.Orientation = xlPageField
'    .EnableMultiplePageItems = True
End With

With .PivotFields("Dup?")
.Orientation = xlPageField
.EnableMultiplePageItems = True
End With

With .PivotFields("Global REI")
.Orientation = xlPageField
.EnableMultiplePageItems = True
End With


Set pf_globalRei = PT.PivotFields("Global REI")
Set pf_vlookup = PT.PivotFields("vlookup")
Set pf_duplicate = PT.PivotFields("Dup?")

pf_globalRei.ClearAllFilters
pf_vlookup.ClearAllFilters
pf_duplicate.ClearAllFilters

'// Enable multiple filters selection
pf_globalRei.EnableMultiplePageItems = True
pf_vlookup.EnableMultiplePageItems = True
pf_duplicate.EnableMultiplePageItems = True


pf_globalRei.CurrentPage = "Yes"
'    pf_vlookup.PivotItems("#N/A").Visible = False
pf_duplicate.CurrentPage = ""

           '// only apply filter if present in data

numberOfElements = UBound(vlookupArray) - LBound(vlookupArray) + 1

If numberOfElements > 0 Then
With pf_vlookup
    For i = 1 To pf_vlookup.PivotItems.Count
    j = 0
'            MsgBox pf_vlookup.PivotItems.Count
'            MsgBox pf_vlookup.PivotItems(i).Name
'            MsgBox vlookupArray(j)
        
    Do While j < numberOfElements
        
        
        If pf_vlookup.PivotItems(i).Name = vlookupArray(j) Then
        
            pf_vlookup.PivotItems(pf_vlookup.PivotItems(i).Name).Visible = False
            Exit Do
        Else
            pf_vlookup.PivotItems(pf_vlookup.PivotItems(i).Name).Visible = True
        End If
        j = j + 1
    Loop
    Next i
End With
End If



'// Rows Section

With .PivotFields("Group")
.Orientation = xlRowField
.Subtotals(1) = False
End With

'// Columns
'With .PivotFields("Group")
'.Orientation = xlColumnField
'End With

'// Values
 With .PivotFields("Complaint ID")
.Orientation = xlDataField
.Function = xlCount
.NumberFormat = "#,###"
End With




End With

CleanUp:
Set PT = Nothing
Set PTCache = Nothing
Set SourceDataRange = Nothing
Set wsSource = Nothing
Set wsTarget = Nothing
Set wb = Nothing
Set pf_globalRei = Nothing
Set pf_vlookup = Nothing
Set pf_duplicate = Nothing

Exit Sub

errHandler:
MsgBox "Error: " & Err.Description, vbExclamation
GoTo CleanUp
End Sub


Sub test3_a1()

Dim wb As Workbook
Dim wsSource As Worksheet, wsTarget As Worksheet
Dim LastRow As Long, LastColumn As Long
Dim SourceDataRange As Range
Dim PTCache As PivotCache
Dim PT As PivotTable
Dim pf_duplicate As PivotField
Dim pf_vlookup As PivotField
Dim vlookupArray(1) As String
Dim numberOfElements As Integer
Dim i As Integer
Dim j As Integer

vlookupArray(1) = "#N/A"


On Error GoTo errHandler
Set wb = ActiveWorkbook
Set wsTarget = wb.Worksheets("T3")
wsTarget.Select
wsTarget.Cells.Clear


'// Step 1. Define data soruce
Set wsSource = wb.Worksheets("Complaints")
With wsSource
LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
LastColumn = .Cells(1, Columns.Count).End(xlToLeft).Column


'// Define source data object
Set SourceDataRange = .Range(.Cells(1, 1), .Cells(LastRow, LastColumn))
End With

'// 2. Set PT Cache
Set PTCache = wb.PivotCaches.Create(xlDatabase, SourceDataRange)

'// 3. Create Pivot Table
Set PT = PTCache.CreatePivotTable(wsTarget.Range("B5"), "T3A1")
With PT
'// Show GrandTotals
.ColumnGrand = True
.RowGrand = True

.RowAxisLayout xlTabularRow

.TableStyle2 = "PivotStyleDark16"



'// Add Pivot Fields

'// Filters
With .PivotFields("vlookup")
.Orientation = xlPageField
.EnableMultiplePageItems = True
End With

With .PivotFields("Dup?")
.Orientation = xlPageField
.EnableMultiplePageItems = True
End With

Set pf_duplicate = PT.PivotFields("Dup?")
Set pf_vlookup = PT.PivotFields("vlookup")

pf_duplicate.ClearAllFilters
'    pf_vlookup.ClearAllFilters

'// Enable multiple filters selection
pf_duplicate.EnableMultiplePageItems = True
pf_vlookup.EnableMultiplePageItems = True


pf_duplicate.CurrentPage = ""
'    pf_vlookup.PivotItems("#N/A").Visible = False

           '// only apply filter if present in data

numberOfElements = UBound(vlookupArray) - LBound(vlookupArray) + 1

If numberOfElements > 0 Then
    With pf_vlookup
        For i = 1 To pf_vlookup.PivotItems.Count
        j = 0
'            MsgBox pf_vlookup.PivotItems.Count
'            MsgBox pf_vlookup.PivotItems(i).Name
'            MsgBox vlookupArray(j)
            
        Do While j < numberOfElements
            
            
            If pf_vlookup.PivotItems(i).Name = vlookupArray(j) Then
            
                pf_vlookup.PivotItems(pf_vlookup.PivotItems(i).Name).Visible = False
                Exit Do
            Else
                pf_vlookup.PivotItems(pf_vlookup.PivotItems(i).Name).Visible = True
            End If
            j = j + 1
        Loop
        Next i
    End With
End If



'// Rows Section

With .PivotFields("Group")
.Orientation = xlRowField
.Subtotals(1) = False
End With

'// Columns
With .PivotFields("Event Region")
.Orientation = xlColumnField
End With

'// Values
 With .PivotFields("Complaint ID")
.Orientation = xlDataField
.Function = xlCount
.NumberFormat = "#,###"
End With




End With

CleanUp:
Set PT = Nothing
Set PTCache = Nothing
Set SourceDataRange = Nothing
Set wsSource = Nothing
Set wsTarget = Nothing
Set wb = Nothing
Set pf_duplicate = Nothing
Set pf_vlookup = Nothing

Exit Sub

errHandler:
MsgBox "Error: " & Err.Description, vbExclamation
GoTo CleanUp
End Sub


Sub Table3_a2()

Dim wb As Workbook
Dim wsSource As Worksheet, wsTarget As Worksheet
Dim LastRow As Long, LastColumn As Long
Dim SourceDataRange As Range
Dim PTCache As PivotCache
Dim PT As PivotTable
Dim pf_vlookup As PivotField


On Error GoTo errHandler
Set wb = ActiveWorkbook
Set wsTarget = wb.Worksheets("T3")
wsTarget.Select
'wsTarget.Cells.Clear


'// Step 1. Define data soruce
Set wsSource = wb.Worksheets("Sales")
With wsSource
LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
LastColumn = .Cells(1, Columns.Count).End(xlToLeft).Column


'// Define source data object
Set SourceDataRange = .Range(.Cells(1, 1), .Cells(LastRow, LastColumn))
End With

'// 2. Set PT Cache
Set PTCache = wb.PivotCaches.Create(xlDatabase, SourceDataRange)

'// 3. Create Pivot Table
Set PT = PTCache.CreatePivotTable(wsTarget.Range("B25"), "T3A2")
With PT
'// Show GrandTotals
.ColumnGrand = True
.RowGrand = True

.RowAxisLayout xlTabularRow

.TableStyle2 = "PivotStyleDark16"



'// Add Pivot Fields

'// Filters
With .PivotFields("vlookup")
.Orientation = xlPageField
.EnableMultiplePageItems = True
End With


Set pf_vlookup = PT.PivotFields("vlookup")
pf_vlookup.ClearAllFilters

'// Enable multiple filters selection
pf_vlookup.EnableMultiplePageItems = True

pf_vlookup.PivotItems("#N/A").Visible = False


'// Rows Section

With .PivotFields("Group")
.Orientation = xlRowField
.Subtotals(1) = False
End With

'// Columns
With .PivotFields("Event Region")
.Orientation = xlColumnField
End With

'// Values
 With .PivotFields("Sales Units")
.Orientation = xlDataField
.Function = xlSum
.NumberFormat = "#,###"
End With




End With

CleanUp:
Set PT = Nothing
Set PTCache = Nothing
Set SourceDataRange = Nothing
Set wsSource = Nothing
Set wsTarget = Nothing
Set wb = Nothing
Set pf_vlookup = Nothing

Exit Sub

errHandler:
MsgBox "Error: " & Err.Description, vbExclamation
GoTo CleanUp
End Sub


Sub test3_b1()

Dim wb As Workbook
Dim wsSource As Worksheet, wsTarget As Worksheet
Dim LastRow As Long, LastColumn As Long
Dim SourceDataRange As Range
Dim PTCache As PivotCache
Dim PT As PivotTable
Dim pf_duplicate As PivotField
Dim pf_vlookup As PivotField
Dim vlookupArray(1) As String
Dim numberOfElements As Integer
Dim i As Integer
Dim j As Integer

vlookupArray(1) = "#N/A"


On Error GoTo errHandler
Set wb = ActiveWorkbook
Set wsTarget = wb.Worksheets("T3")
wsTarget.Select
'wsTarget.Cells.Clear


'// Step 1. Define data soruce
Set wsSource = wb.Worksheets("Complaints")
With wsSource
LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
LastColumn = .Cells(1, Columns.Count).End(xlToLeft).Column


'// Define source data object
Set SourceDataRange = .Range(.Cells(1, 1), .Cells(LastRow, LastColumn))
End With

'// 2. Set PT Cache
Set PTCache = wb.PivotCaches.Create(xlDatabase, SourceDataRange)

'// 3. Create Pivot Table
Set PT = PTCache.CreatePivotTable(wsTarget.Range("J5"), "T3B1")
With PT
'// Show GrandTotals
.ColumnGrand = True
.RowGrand = True

.RowAxisLayout xlTabularRow

.TableStyle2 = "PivotStyleDark16"



'// Add Pivot Fields

'// Filters
With .PivotFields("vlookup")
.Orientation = xlPageField
.EnableMultiplePageItems = True
End With

With .PivotFields("Dup?")
.Orientation = xlPageField
.EnableMultiplePageItems = True
End With


Set pf_duplicate = PT.PivotFields("Dup?")
Set pf_vlookup = PT.PivotFields("vlookup")

pf_duplicate.ClearAllFilters
'    pf_vlookup.ClearAllFilters

'// Enable multiple filters selection
pf_duplicate.EnableMultiplePageItems = True
pf_vlookup.EnableMultiplePageItems = True


pf_duplicate.CurrentPage = ""
'    pf_vlookup.PivotItems("#N/A").Visible = False

              '// only apply filter if present in data

numberOfElements = UBound(vlookupArray) - LBound(vlookupArray) + 1

If numberOfElements > 0 Then
    With pf_vlookup
        For i = 1 To pf_vlookup.PivotItems.Count
        j = 0
'            MsgBox pf_vlookup.PivotItems.Count
'            MsgBox pf_vlookup.PivotItems(i).Name
'            MsgBox vlookupArray(j)
            
        Do While j < numberOfElements
            
            
            If pf_vlookup.PivotItems(i).Name = vlookupArray(j) Then
            
                pf_vlookup.PivotItems(pf_vlookup.PivotItems(i).Name).Visible = False
                Exit Do
            Else
                pf_vlookup.PivotItems(pf_vlookup.PivotItems(i).Name).Visible = True
            End If
            j = j + 1
        Loop
        Next i
    End With
End If



'// Rows Section

With .PivotFields("Group")
.Orientation = xlRowField
.Subtotals(1) = False
End With

'// Columns
With .PivotFields("Year")
.Orientation = xlColumnField
End With

'// Values
 With .PivotFields("Complaint ID")
.Orientation = xlDataField
.Function = xlCount
.NumberFormat = "#,###"
End With




End With

CleanUp:
Set PT = Nothing
Set PTCache = Nothing
Set SourceDataRange = Nothing
Set wsSource = Nothing
Set wsTarget = Nothing
Set wb = Nothing
Set pf_duplicate = Nothing
Set pf_vlookup = Nothing

Exit Sub

errHandler:
MsgBox "Error: " & Err.Description, vbExclamation
GoTo CleanUp
End Sub


Sub Table3_b2()

Dim wb As Workbook
Dim wsSource As Worksheet, wsTarget As Worksheet
Dim LastRow As Long, LastColumn As Long
Dim SourceDataRange As Range
Dim PTCache As PivotCache
Dim PT As PivotTable
Dim pf_vlookup As PivotField


On Error GoTo errHandler
Set wb = ActiveWorkbook
Set wsTarget = wb.Worksheets("T3")
wsTarget.Select
'wsTarget.Cells.Clear


'// Step 1. Define data soruce
Set wsSource = wb.Worksheets("Sales")
With wsSource
LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
LastColumn = .Cells(1, Columns.Count).End(xlToLeft).Column


'// Define source data object
Set SourceDataRange = .Range(.Cells(1, 1), .Cells(LastRow, LastColumn))
End With

'// 2. Set PT Cache
Set PTCache = wb.PivotCaches.Create(xlDatabase, SourceDataRange)

'// 3. Create Pivot Table
Set PT = PTCache.CreatePivotTable(wsTarget.Range("J25"), "T3B2")
With PT
'// Show GrandTotals
.ColumnGrand = True
.RowGrand = True

.RowAxisLayout xlTabularRow

.TableStyle2 = "PivotStyleDark16"



'// Add Pivot Fields

'// Filters
With .PivotFields("vlookup")
.Orientation = xlPageField
.EnableMultiplePageItems = True
End With


Set pf_vlookup = PT.PivotFields("vlookup")
pf_vlookup.ClearAllFilters

'// Enable multiple filters selection
pf_vlookup.EnableMultiplePageItems = True

pf_vlookup.PivotItems("#N/A").Visible = False


'// Rows Section

With .PivotFields("Group")
.Orientation = xlRowField
.Subtotals(1) = False
End With

'// Columns
With .PivotFields("Year")
.Orientation = xlColumnField
End With

'// Values
 With .PivotFields("Sales Units")
.Orientation = xlDataField
.Function = xlSum
.NumberFormat = "#,###"
End With




End With

CleanUp:
Set PT = Nothing
Set PTCache = Nothing
Set SourceDataRange = Nothing
Set wsSource = Nothing
Set wsTarget = Nothing
Set wb = Nothing
Set pf_vlookup = Nothing

Exit Sub

errHandler:
MsgBox "Error: " & Err.Description, vbExclamation
GoTo CleanUp
End Sub


Sub test_4A()

Dim wb As Workbook
Dim wsSource As Worksheet, wsTarget As Worksheet
Dim LastRow As Long, LastColumn As Long
Dim SourceDataRange As Range
Dim PTCache As PivotCache
Dim PT As PivotTable
Dim pf_vlookup As PivotField
Dim pf_ARC As PivotField
Dim vlookupArray(1) As String
Dim arcArray(0 To 10) As String
Dim numberOfElements As Integer
Dim numberOfElementsTwo As Integer
Dim i As Integer
Dim j As Integer
Dim c As Integer
Dim d As Integer


vlookupArray(1) = "#N/A"

arcArray(0) = "No Consequences or Impact to Patient"
arcArray(1) = "No Known Impact Or Consequence To Patient"
arcArray(2) = ""
arcArray(3) = "Device No Known Device Problem"
arcArray(4) = "Device No Reported Allegation"
arcArray(5) = "Insufficient Information"
arcArray(6) = "No Clinical Signs, Symptoms or Conditions"
arcArray(7) = "No Code Available"
arcArray(8) = "No Health Consequences or Impact"
arcArray(9) = "No Information"
arcArray(10) = "No Patient Involvement"



On Error GoTo errHandler
Set wb = ActiveWorkbook
Set wsTarget = wb.Worksheets("T4")
wsTarget.Select
wsTarget.Cells.Clear


'// Step 1. Define data soruce
Set wsSource = wb.Worksheets("Complaints")
With wsSource
LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
LastColumn = .Cells(1, Columns.Count).End(xlToLeft).Column


'// Define source data object
Set SourceDataRange = .Range(.Cells(1, 1), .Cells(LastRow, LastColumn))
End With

'// 2. Set PT Cache
Set PTCache = wb.PivotCaches.Create(xlDatabase, SourceDataRange)

'// 3. Create Pivot Table
Set PT = PTCache.CreatePivotTable(wsTarget.Range("B5"), "T4A")
With PT
'// Show GrandTotals
.ColumnGrand = True
.RowGrand = True

.RowAxisLayout xlTabularRow

.TableStyle2 = "PivotStyleDark16"



'// Add Pivot Fields

'// Filters
With .PivotFields("vlookup")
.Orientation = xlPageField
'    .EnableMultiplePageItems = True
End With

 Set pf_vlookup = PT.PivotFields("vlookup")

'// Enable multiple filters selection
pf_vlookup.EnableMultiplePageItems = True

'   pf_vlookup.PivotItems("#N/A").Visible = False

With .PivotFields("AR Code Description (GCMS)2")
.Orientation = xlPageField
'    .EnableMultiplePageItems = True
End With

'// only apply filter if present in data

numberOfElements = UBound(vlookupArray) - LBound(vlookupArray) + 1

If numberOfElements > 0 Then
With pf_vlookup
    For i = 1 To pf_vlookup.PivotItems.Count
    j = 0
'            MsgBox pf_vlookup.PivotItems.Count
'            MsgBox pf_vlookup.PivotItems(i).Name
'            MsgBox vlookupArray(j)
        
    Do While j < numberOfElements
        
        
        If pf_vlookup.PivotItems(i).Name = vlookupArray(j) Then
        
            pf_vlookup.PivotItems(pf_vlookup.PivotItems(i).Name).Visible = False
            Exit Do
        Else
            pf_vlookup.PivotItems(pf_vlookup.PivotItems(i).Name).Visible = True
        End If
        j = j + 1
    Loop
    Next i
End With
End If

'// Rows Section

With .PivotFields("Group")
.Orientation = xlRowField
.Subtotals(1) = False
End With

With .PivotFields("AR Code Description (GCMS)")
.Orientation = xlRowField
.Subtotals(1) = False
End With





'// Columns
With .PivotFields("Year")
.Orientation = xlColumnField
End With

'// Values
 With .PivotFields("Complaint ID")
.Orientation = xlDataField
.Function = xlCount
.NumberFormat = "#,###"
End With

    '// ARC filter
     Set pf_ARC = PT.PivotFields("AR Code Description (GCMS)2")

'    MsgBox pf_ARC
'// Enable multiple filters selection
pf_ARC.EnableMultiplePageItems = True



'// only apply filter if present in data

numberOfElementsTwo = UBound(arcArray) - LBound(arcArray) + 1
'    MsgBox numberOfElementsTwo

If numberOfElementsTwo > 0 Then
With pf_ARC
    For c = 1 To pf_ARC.PivotItems.Count

    d = 0
'            MsgBox pf_ARC.PivotItems.Count
'            MsgBox "Arc Array" + pf_ARC.PivotItems(c).Name
'            MsgBox d
        
    Do While d < numberOfElementsTwo
        
        

        If pf_ARC.PivotItems(c).Name = arcArray(d) Then
        
            pf_ARC.PivotItems(pf_ARC.PivotItems(c).Name).Visible = False
            Exit Do
        Else
            pf_ARC.PivotItems(pf_ARC.PivotItems(c).Name).Visible = True
        End If
        d = d + 1
    Loop
    Next c
End With
End If


ActiveSheet.PivotTables(1).NullString = "0"

End With

CleanUp:
Set PT = Nothing
Set PTCache = Nothing
Set SourceDataRange = Nothing
Set wsSource = Nothing
Set wsTarget = Nothing
Set wb = Nothing
Set pf_vlookup = Nothing

Exit Sub

errHandler:
MsgBox "Error: " & Err.Description, vbExclamation
GoTo CleanUp
End Sub

Sub test_4B()

Dim wb As Workbook
Dim wsSource As Worksheet, wsTarget As Worksheet
Dim LastRow As Long, LastColumn As Long
Dim SourceDataRange As Range
Dim PTCache As PivotCache
Dim PT As PivotTable
Dim pf_vlookup As PivotField
Dim pf_ARC As PivotField
Dim vlookupArray(1) As String
Dim arcArray(0 To 10) As String
Dim numberOfElements As Integer
Dim numberOfElementsTwo As Integer
Dim i As Integer
Dim j As Integer
Dim c As Integer
Dim d As Integer


vlookupArray(1) = "#N/A"

arcArray(0) = "No Consequences or Impact to Patient"
arcArray(1) = "No Known Impact Or Consequence To Patient"
arcArray(2) = ""
arcArray(3) = "Device No Known Device Problem"
arcArray(4) = "Device No Reported Allegation"
arcArray(5) = "Insufficient Information"
arcArray(6) = "No Clinical Signs, Symptoms or Conditions"
arcArray(7) = "No Code Available"
arcArray(8) = "No Health Consequences or Impact"
arcArray(9) = "No Information"
arcArray(10) = "No Patient Involvement"


On Error GoTo errHandler
Set wb = ActiveWorkbook
Set wsTarget = wb.Worksheets("T4")
wsTarget.Select
'wsTarget.Cells.Clear


'// Step 1. Define data soruce
Set wsSource = wb.Worksheets("Complaints")
With wsSource
LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
LastColumn = .Cells(1, Columns.Count).End(xlToLeft).Column


'// Define source data object
Set SourceDataRange = .Range(.Cells(1, 1), .Cells(LastRow, LastColumn))
End With

'// 2. Set PT Cache
Set PTCache = wb.PivotCaches.Create(xlDatabase, SourceDataRange)

'// 3. Create Pivot Table
Set PT = PTCache.CreatePivotTable(wsTarget.Range("L5"), "T4B")
With PT
'// Show GrandTotals
.ColumnGrand = True
.RowGrand = True

.RowAxisLayout xlTabularRow

.TableStyle2 = "PivotStyleDark16"



'// Add Pivot Fields

'// Filters
With .PivotFields("vlookup")
.Orientation = xlPageField
'    .EnableMultiplePageItems = True
End With

 Set pf_vlookup = PT.PivotFields("vlookup")
 
'     pf_vlookup.ClearAllFilters
 
 '// Enable multiple filters selection
pf_vlookup.EnableMultiplePageItems = True

'    pf_vlookup.PivotItems("#N/A").Visible = False

   With .PivotFields("AR Code Description (GCMS)2")
.Orientation = xlPageField
'    .EnableMultiplePageItems = True
End With

'// only apply filter if present in data

numberOfElements = UBound(vlookupArray) - LBound(vlookupArray) + 1

If numberOfElements > 0 Then
With pf_vlookup
    For i = 1 To pf_vlookup.PivotItems.Count
    j = 0
'            MsgBox pf_vlookup.PivotItems.Count
'            MsgBox pf_vlookup.PivotItems(i).Name
'            MsgBox vlookupArray(j)
        
    Do While j < numberOfElements
        
        
        If pf_vlookup.PivotItems(i).Name = vlookupArray(j) Then
        
            pf_vlookup.PivotItems(pf_vlookup.PivotItems(i).Name).Visible = False
            Exit Do
        Else
            pf_vlookup.PivotItems(pf_vlookup.PivotItems(i).Name).Visible = True
        End If
        j = j + 1
    Loop
    Next i
End With
End If

'// Rows Section

With .PivotFields("Group")
.Orientation = xlRowField
.Subtotals(1) = False
End With

With .PivotFields("AR Code Description (GCMS)")
.Orientation = xlRowField
.Subtotals(1) = False
End With

'// Columns
With .PivotFields("Event Region")
.Orientation = xlColumnField
End With

'// Values
 With .PivotFields("Complaint ID")
.Orientation = xlDataField
.Function = xlCount
.NumberFormat = "#,###"
End With

        '// ARC filter
     Set pf_ARC = PT.PivotFields("AR Code Description (GCMS)2")

'    MsgBox pf_ARC
'// Enable multiple filters selection
pf_ARC.EnableMultiplePageItems = True



'// only apply filter if present in data

numberOfElementsTwo = UBound(arcArray) - LBound(arcArray) + 1
'    MsgBox numberOfElementsTwo

If numberOfElementsTwo > 0 Then
With pf_ARC
    For c = 1 To pf_ARC.PivotItems.Count

    d = 0
'            MsgBox pf_ARC.PivotItems.Count
'            MsgBox "Arc Array" + pf_ARC.PivotItems(c).Name
'            MsgBox d
        
    Do While d < numberOfElementsTwo
        
        

        If pf_ARC.PivotItems(c).Name = arcArray(d) Then
        
            pf_ARC.PivotItems(pf_ARC.PivotItems(c).Name).Visible = False
            Exit Do
        Else
            pf_ARC.PivotItems(pf_ARC.PivotItems(c).Name).Visible = True
        End If
        d = d + 1
    Loop
    Next c
End With
End If


ActiveSheet.PivotTables(1).NullString = "0"

End With

CleanUp:
Set PT = Nothing
Set PTCache = Nothing
Set SourceDataRange = Nothing
Set wsSource = Nothing
Set wsTarget = Nothing
Set wb = Nothing
Set pf_vlookup = Nothing

Exit Sub

errHandler:
MsgBox "Error: " & Err.Description, vbExclamation
GoTo CleanUp
End Sub


Sub test5_a1()

Dim wb As Workbook
Dim wsSource As Worksheet, wsTarget As Worksheet
Dim LastRow As Long, LastColumn As Long
Dim SourceDataRange As Range
Dim PTCache As PivotCache
Dim PT As PivotTable
Dim pf_globalRei As PivotField
Dim pf_vlookup As PivotField
Dim pf_duplicate As PivotField
Dim vlookupArray(1) As String
Dim numberOfElements As Integer
Dim i As Integer
Dim j As Integer

vlookupArray(1) = "#N/A"


On Error GoTo errHandler
Set wb = ActiveWorkbook
Set wsTarget = wb.Worksheets("T5")
wsTarget.Select
wsTarget.Cells.Clear


'// Step 1. Define data soruce
Set wsSource = wb.Worksheets("Complaints")
With wsSource
LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
LastColumn = .Cells(1, Columns.Count).End(xlToLeft).Column


'// Define source data object
Set SourceDataRange = .Range(.Cells(1, 1), .Cells(LastRow, LastColumn))
End With

'// 2. Set PT Cache
Set PTCache = wb.PivotCaches.Create(xlDatabase, SourceDataRange)

'// 3. Create Pivot Table
Set PT = PTCache.CreatePivotTable(wsTarget.Range("B5"), "T5A1")
With PT
'// Show GrandTotals
.ColumnGrand = True
.RowGrand = True

.RowAxisLayout xlTabularRow

.TableStyle2 = "PivotStyleDark16"



'// Add Pivot Fields

'// Filters
With .PivotFields("vlookup")
.Orientation = xlPageField
.EnableMultiplePageItems = True
End With

With .PivotFields("Dup?")
.Orientation = xlPageField
.EnableMultiplePageItems = True
End With

With .PivotFields("Global REI")
.Orientation = xlPageField
.EnableMultiplePageItems = True
End With


Set pf_globalRei = PT.PivotFields("Global REI")
Set pf_vlookup = PT.PivotFields("vlookup")
Set pf_duplicate = PT.PivotFields("Dup?")

pf_globalRei.ClearAllFilters
pf_vlookup.ClearAllFilters
pf_duplicate.ClearAllFilters


'// Enable multiple filters selection
pf_duplicate.EnableMultiplePageItems = True
pf_vlookup.EnableMultiplePageItems = True
pf_globalRei.EnableMultiplePageItems = True


pf_duplicate.CurrentPage = ""
'    pf_vlookup.PivotItems("#N/A").Visible = False
pf_globalRei.CurrentPage = "Yes"


  '// only apply filter if present in data

numberOfElements = UBound(vlookupArray) - LBound(vlookupArray) + 1

If numberOfElements > 0 Then
    With pf_vlookup
        For i = 1 To pf_vlookup.PivotItems.Count
        j = 0
'            MsgBox pf_vlookup.PivotItems.Count
'            MsgBox pf_vlookup.PivotItems(i).Name
'            MsgBox vlookupArray(j)
            
        Do While j < numberOfElements
            
            
            If pf_vlookup.PivotItems(i).Name = vlookupArray(j) Then
            
                pf_vlookup.PivotItems(pf_vlookup.PivotItems(i).Name).Visible = False
                Exit Do
            Else
                pf_vlookup.PivotItems(pf_vlookup.PivotItems(i).Name).Visible = True
            End If
            j = j + 1
        Loop
        Next i
    End With
End If

'// Rows Section

With .PivotFields("Group")
.Orientation = xlRowField
.Subtotals(1) = False
End With

'// Columns
'With .PivotFields("Reg Report Priority")
'.Orientation = xlColumnField
'End With

'// Values
 With .PivotFields("Complaint ID")
.Orientation = xlDataField
.Function = xlCount
.NumberFormat = "#,###"
End With




End With

CleanUp:
Set PT = Nothing
Set PTCache = Nothing
Set SourceDataRange = Nothing
Set wsSource = Nothing
Set wsTarget = Nothing
Set wb = Nothing
Set pf_duplicate = Nothing
Set pf_vlookup = Nothing
Set pf_globalRei = Nothing

Exit Sub

errHandler:
MsgBox "Error: " & Err.Description, vbExclamation
GoTo CleanUp
End Sub


Sub test5_a2()

Dim wb As Workbook
Dim wsSource As Worksheet, wsTarget As Worksheet
Dim LastRow As Long, LastColumn As Long
Dim SourceDataRange As Range
Dim PTCache As PivotCache
Dim PT As PivotTable
Dim pf_duplicate As PivotField
Dim pf_vlookup As PivotField
Dim pf_MDR As PivotField
Dim vlookupArray(1) As String
Dim numberOfElements As Integer
Dim i As Integer
Dim j As Integer

vlookupArray(1) = "#N/A"


On Error GoTo errHandler
Set wb = ActiveWorkbook
Set wsTarget = wb.Worksheets("T5")
wsTarget.Select
'wsTarget.Cells.Clear


'// Step 1. Define data soruce
Set wsSource = wb.Worksheets("Complaints")
With wsSource
LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
LastColumn = .Cells(1, Columns.Count).End(xlToLeft).Column


'// Define source data object
Set SourceDataRange = .Range(.Cells(1, 1), .Cells(LastRow, LastColumn))
End With

'// 2. Set PT Cache
Set PTCache = wb.PivotCaches.Create(xlDatabase, SourceDataRange)

'// 3. Create Pivot Table
Set PT = PTCache.CreatePivotTable(wsTarget.Range("B20"), "T5A2")
With PT
'// Show GrandTotals
.ColumnGrand = True
.RowGrand = True

.RowAxisLayout xlTabularRow

.TableStyle2 = "PivotStyleDark16"



'// Add Pivot Fields

'// Filters
With .PivotFields("vlookup")
.Orientation = xlPageField
.EnableMultiplePageItems = True
End With

With .PivotFields("Dup?")
.Orientation = xlPageField
.EnableMultiplePageItems = True
End With

With .PivotFields("MDR REI")
.Orientation = xlPageField
.EnableMultiplePageItems = True
End With


Set pf_duplicate = PT.PivotFields("Dup?")
Set pf_vlookup = PT.PivotFields("vlookup")
Set pf_MDR = PT.PivotFields("MDR REI")

pf_duplicate.ClearAllFilters
pf_vlookup.ClearAllFilters
pf_MDR.ClearAllFilters

'// Enable multiple filters selection
pf_duplicate.EnableMultiplePageItems = True
pf_vlookup.EnableMultiplePageItems = True
pf_MDR.EnableMultiplePageItems = True


pf_duplicate.CurrentPage = ""
'    pf_vlookup.PivotItems("#N/A").Visible = False
pf_MDR.CurrentPage = "Yes"


    '// only apply filter if present in data

numberOfElements = UBound(vlookupArray) - LBound(vlookupArray) + 1

If numberOfElements > 0 Then
    With pf_vlookup
        For i = 1 To pf_vlookup.PivotItems.Count
        j = 0
'            MsgBox pf_vlookup.PivotItems.Count
'            MsgBox pf_vlookup.PivotItems(i).Name
'            MsgBox vlookupArray(j)
            
        Do While j < numberOfElements
            
            
            If pf_vlookup.PivotItems(i).Name = vlookupArray(j) Then
            
                pf_vlookup.PivotItems(pf_vlookup.PivotItems(i).Name).Visible = False
                Exit Do
            Else
                pf_vlookup.PivotItems(pf_vlookup.PivotItems(i).Name).Visible = True
            End If
            j = j + 1
        Loop
        Next i
    End With
End If

'// Rows Section

With .PivotFields("Group")
.Orientation = xlRowField
.Subtotals(1) = False
End With

'// Columns
'With .PivotFields("Reg Report Priority")
'.Orientation = xlColumnField
'End With

'// Values
 With .PivotFields("Complaint ID")
.Orientation = xlDataField
.Function = xlCount
.NumberFormat = "#,###"
End With




End With

CleanUp:
Set PT = Nothing
Set PTCache = Nothing
Set SourceDataRange = Nothing
Set wsSource = Nothing
Set wsTarget = Nothing
Set wb = Nothing
Set pf_duplicate = Nothing
Set pf_vlookup = Nothing
Set pf_MDR = Nothing

Exit Sub

errHandler:
MsgBox "Error: " & Err.Description, vbExclamation
GoTo CleanUp
End Sub

Sub test5_a3()

Dim wb As Workbook
Dim wsSource As Worksheet, wsTarget As Worksheet
Dim LastRow As Long, LastColumn As Long
Dim SourceDataRange As Range
Dim PTCache As PivotCache
Dim PT As PivotTable
Dim pf_duplicate As PivotField
Dim pf_vlookup As PivotField
Dim pf_MDV As PivotField
Dim vlookupArray(1) As String
Dim numberOfElements As Integer
Dim i As Integer
Dim j As Integer

vlookupArray(1) = "#N/A"


On Error GoTo errHandler
Set wb = ActiveWorkbook
Set wsTarget = wb.Worksheets("T5")
wsTarget.Select
'wsTarget.Cells.Clear


'// Step 1. Define data soruce
Set wsSource = wb.Worksheets("Complaints")
With wsSource
LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
LastColumn = .Cells(1, Columns.Count).End(xlToLeft).Column


'// Define source data object
Set SourceDataRange = .Range(.Cells(1, 1), .Cells(LastRow, LastColumn))
End With

'// 2. Set PT Cache
Set PTCache = wb.PivotCaches.Create(xlDatabase, SourceDataRange)

'// 3. Create Pivot Table
Set PT = PTCache.CreatePivotTable(wsTarget.Range("E20"), "T5A3")
With PT
'// Show GrandTotals
.ColumnGrand = True
.RowGrand = True

.RowAxisLayout xlTabularRow

.TableStyle2 = "PivotStyleDark16"



'// Add Pivot Fields

'// Filters
With .PivotFields("vlookup")
.Orientation = xlPageField
.EnableMultiplePageItems = True
End With

With .PivotFields("Dup?")
.Orientation = xlPageField
.EnableMultiplePageItems = True
End With

With .PivotFields("MDV REI")
.Orientation = xlPageField
.EnableMultiplePageItems = True
End With

Set pf_duplicate = PT.PivotFields("Dup?")
Set pf_vlookup = PT.PivotFields("vlookup")
Set pf_MDV = PT.PivotFields("MDV REI")

pf_duplicate.ClearAllFilters
pf_vlookup.ClearAllFilters
pf_MDV.ClearAllFilters

'// Enable multiple filters selection
pf_duplicate.EnableMultiplePageItems = True
pf_vlookup.EnableMultiplePageItems = True
pf_MDV.EnableMultiplePageItems = True


pf_duplicate.CurrentPage = ""
'    pf_vlookup.PivotItems("#N/A").Visible = False
pf_MDV.CurrentPage = "Yes"

      '// only apply filter if present in data

numberOfElements = UBound(vlookupArray) - LBound(vlookupArray) + 1

If numberOfElements > 0 Then
    With pf_vlookup
        For i = 1 To pf_vlookup.PivotItems.Count
        j = 0
'            MsgBox pf_vlookup.PivotItems.Count
'            MsgBox pf_vlookup.PivotItems(i).Name
'            MsgBox vlookupArray(j)
            
        Do While j < numberOfElements
            
            
            If pf_vlookup.PivotItems(i).Name = vlookupArray(j) Then
            
                pf_vlookup.PivotItems(pf_vlookup.PivotItems(i).Name).Visible = False
                Exit Do
            Else
                pf_vlookup.PivotItems(pf_vlookup.PivotItems(i).Name).Visible = True
            End If
            j = j + 1
        Loop
        Next i
    End With
End If


'// Rows Section

With .PivotFields("Group")
.Orientation = xlRowField
.Subtotals(1) = False
End With

'// Columns
'With .PivotFields("Reg Report Priority")
'.Orientation = xlColumnField
'End With

'// Values
 With .PivotFields("Complaint ID")
.Orientation = xlDataField
.Function = xlCount
.NumberFormat = "#,###"
End With




End With

CleanUp:
Set PT = Nothing
Set PTCache = Nothing
Set SourceDataRange = Nothing
Set wsSource = Nothing
Set wsTarget = Nothing
Set wb = Nothing
Set pf_duplicate = Nothing
Set pf_vlookup = Nothing
Set pf_MDV = Nothing
Exit Sub

errHandler:
MsgBox "Error: " & Err.Description, vbExclamation
GoTo CleanUp
End Sub


Sub test5_b()

Dim wb As Workbook
Dim wsSource As Worksheet, wsTarget As Worksheet
Dim LastRow As Long, LastColumn As Long
Dim SourceDataRange As Range
Dim PTCache As PivotCache
Dim PT As PivotTable
Dim pf_globalRei As PivotField
Dim pf_vlookup As PivotField
Dim pf_duplicate As PivotField
Dim vlookupArray(1) As String
Dim numberOfElements As Integer
Dim i As Integer
Dim j As Integer

vlookupArray(1) = "#N/A"



On Error GoTo errHandler
Set wb = ActiveWorkbook
Set wsTarget = wb.Worksheets("T5")
wsTarget.Select
'wsTarget.Cells.Clear


'// Step 1. Define data soruce
Set wsSource = wb.Worksheets("Complaints")
With wsSource
LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
LastColumn = .Cells(1, Columns.Count).End(xlToLeft).Column


'// Define source data object
Set SourceDataRange = .Range(.Cells(1, 1), .Cells(LastRow, LastColumn))
End With

'// 2. Set PT Cache
Set PTCache = wb.PivotCaches.Create(xlDatabase, SourceDataRange)

'// 3. Create Pivot Table
Set PT = PTCache.CreatePivotTable(wsTarget.Range("B35"), "T5B")
With PT
'// Show GrandTotals
.ColumnGrand = True
.RowGrand = True

.RowAxisLayout xlTabularRow

.TableStyle2 = "PivotStyleDark16"



'// Add Pivot Fields

'// Filters
With .PivotFields("vlookup")
.Orientation = xlPageField
.EnableMultiplePageItems = True
End With

With .PivotFields("Dup?")
.Orientation = xlPageField
.EnableMultiplePageItems = True
End With

With .PivotFields("Global REI")
.Orientation = xlPageField
.EnableMultiplePageItems = True
End With


Set pf_globalRei = PT.PivotFields("Global REI")
Set pf_vlookup = PT.PivotFields("vlookup")
Set pf_duplicate = PT.PivotFields("Dup?")

pf_globalRei.ClearAllFilters
pf_vlookup.ClearAllFilters
pf_duplicate.ClearAllFilters

'// Enable multiple filters selection
pf_globalRei.EnableMultiplePageItems = True
pf_vlookup.EnableMultiplePageItems = True
pf_duplicate.EnableMultiplePageItems = True


pf_globalRei.CurrentPage = "Yes"
'    pf_vlookup.PivotItems("#N/A").Visible = False
pf_duplicate.CurrentPage = ""


      '// only apply filter if present in data

numberOfElements = UBound(vlookupArray) - LBound(vlookupArray) + 1

If numberOfElements > 0 Then
    With pf_vlookup
        For i = 1 To pf_vlookup.PivotItems.Count
        j = 0
'            MsgBox pf_vlookup.PivotItems.Count
'            MsgBox pf_vlookup.PivotItems(i).Name
'            MsgBox vlookupArray(j)
            
        Do While j < numberOfElements
            
            
            If pf_vlookup.PivotItems(i).Name = vlookupArray(j) Then
            
                pf_vlookup.PivotItems(pf_vlookup.PivotItems(i).Name).Visible = False
                Exit Do
            Else
                pf_vlookup.PivotItems(pf_vlookup.PivotItems(i).Name).Visible = True
            End If
            j = j + 1
        Loop
        Next i
    End With
End If

'// Rows Section

With .PivotFields("Group")
.Orientation = xlRowField
.Subtotals(1) = False
End With

With .PivotFields("Reg Report Priority")
.Orientation = xlRowField
.Subtotals(1) = False
End With

'// Columns
With .PivotFields("Year")
.Orientation = xlColumnField
End With

'// Values
 With .PivotFields("Complaint ID")
.Orientation = xlDataField
.Function = xlCount
.NumberFormat = "#,###"
End With


ActiveSheet.PivotTables(1).NullString = "0"

End With

CleanUp:
Set PT = Nothing
Set PTCache = Nothing
Set SourceDataRange = Nothing
Set wsSource = Nothing
Set wsTarget = Nothing
Set wb = Nothing
Set pf_globalRei = Nothing
Set pf_vlookup = Nothing
Set pf_duplicate = Nothing

Exit Sub

errHandler:
MsgBox "Error: " & Err.Description, vbExclamation
GoTo CleanUp
End Sub

Sub test6_a()

Dim wb As Workbook
Dim wsSource As Worksheet, wsTarget As Worksheet
Dim LastRow As Long, LastColumn As Long
Dim SourceDataRange As Range
Dim PTCache As PivotCache
Dim PT As PivotTable
Dim pf_globalRei As PivotField
Dim pf_vlookup As PivotField
Dim vlookupArray(1) As String
Dim numberOfElements As Integer
Dim pf_ARC As PivotField
Dim arcArray(0 To 10) As String
Dim numberOfElementsTwo As Integer
Dim i As Integer
Dim j As Integer
Dim c As Integer
Dim d As Integer


vlookupArray(1) = "#N/A"

arcArray(0) = "No Consequences or Impact to Patient"
arcArray(1) = "No Known Impact Or Consequence To Patient"
arcArray(2) = ""
arcArray(3) = "Device No Known Device Problem"
arcArray(4) = "Device No Reported Allegation"
arcArray(5) = "Insufficient Information"
arcArray(6) = "No Clinical Signs, Symptoms or Conditions"
arcArray(7) = "No Code Available"
arcArray(8) = "No Health Consequences or Impact"
arcArray(9) = "No Information"
arcArray(10) = "No Patient Involvement"


On Error GoTo errHandler
Set wb = ActiveWorkbook
Set wsTarget = wb.Worksheets("T6")
wsTarget.Select
wsTarget.Cells.Clear


'// Step 1. Define data soruce
Set wsSource = wb.Worksheets("Complaints")
With wsSource
LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
LastColumn = .Cells(1, Columns.Count).End(xlToLeft).Column


'// Define source data object
Set SourceDataRange = .Range(.Cells(1, 1), .Cells(LastRow, LastColumn))
End With

'// 2. Set PT Cache
Set PTCache = wb.PivotCaches.Create(xlDatabase, SourceDataRange)

'// 3. Create Pivot Table
Set PT = PTCache.CreatePivotTable(wsTarget.Range("B5"), "T6A")
With PT
'// Show GrandTotals
.ColumnGrand = True
.RowGrand = True

.RowAxisLayout xlTabularRow

.TableStyle2 = "PivotStyleDark16"



'// Add Pivot Fields

'// Filters
With .PivotFields("vlookup")
.Orientation = xlPageField
.EnableMultiplePageItems = True
End With


With .PivotFields("Global REI")
.Orientation = xlPageField
.EnableMultiplePageItems = True
End With

With .PivotFields("AR Code Description (GCMS)2")
.Orientation = xlPageField
'    .EnableMultiplePageItems = True
End With

Set pf_globalRei = PT.PivotFields("Global REI")
Set pf_vlookup = PT.PivotFields("vlookup")
Set pf_ARC = PT.PivotFields("AR Code Description (GCMS)2")

pf_globalRei.ClearAllFilters
pf_vlookup.ClearAllFilters

'// Enable multiple filters selection
pf_globalRei.EnableMultiplePageItems = True
pf_vlookup.EnableMultiplePageItems = True
pf_ARC.EnableMultiplePageItems = True


pf_globalRei.CurrentPage = "Yes"
'    pf_vlookup.PivotItems("#N/A").Visible = False


   '// only apply filter if present in data

numberOfElements = UBound(vlookupArray) - LBound(vlookupArray) + 1

If numberOfElements > 0 Then
    With pf_vlookup
        For i = 1 To pf_vlookup.PivotItems.Count
        j = 0
'            MsgBox pf_vlookup.PivotItems.Count
'            MsgBox pf_vlookup.PivotItems(i).Name
'            MsgBox vlookupArray(j)
            
        Do While j < numberOfElements
            
            
            If pf_vlookup.PivotItems(i).Name = vlookupArray(j) Then
            
                pf_vlookup.PivotItems(pf_vlookup.PivotItems(i).Name).Visible = False
                Exit Do
            Else
                pf_vlookup.PivotItems(pf_vlookup.PivotItems(i).Name).Visible = True
            End If
            j = j + 1
        Loop
        Next i
    End With
End If

'// Rows Section

With .PivotFields("Group")
.Orientation = xlRowField
.Subtotals(1) = False
End With

With .PivotFields("AR Code Description (GCMS)")
.Orientation = xlRowField
.Subtotals(1) = False
End With

'// Columns
With .PivotFields("Reg Report Priority")
.Orientation = xlColumnField
End With

'// Values
 With .PivotFields("Complaint ID")
.Orientation = xlDataField
.Function = xlCount
.NumberFormat = "#,###"
End With

 '// only apply filter if present in data

    numberOfElementsTwo = UBound(arcArray) - LBound(arcArray) + 1
'    MsgBox numberOfElementsTwo

If numberOfElementsTwo > 0 Then
    With pf_ARC
        For c = 1 To pf_ARC.PivotItems.Count

        d = 0
'            MsgBox pf_ARC.PivotItems.Count
'            MsgBox "Arc Array" + pf_ARC.PivotItems(c).Name
'            MsgBox d
            
        Do While d < numberOfElementsTwo
            
            

            If pf_ARC.PivotItems(c).Name = arcArray(d) Then
            
                pf_ARC.PivotItems(pf_ARC.PivotItems(c).Name).Visible = False
                Exit Do
            Else
                pf_ARC.PivotItems(pf_ARC.PivotItems(c).Name).Visible = True
            End If
            d = d + 1
        Loop
        Next c
    End With
End If


ActiveSheet.PivotTables(1).NullString = "0"

End With

CleanUp:
Set PT = Nothing
Set PTCache = Nothing
Set SourceDataRange = Nothing
Set wsSource = Nothing
Set wsTarget = Nothing
Set wb = Nothing
Set pf_globalRei = Nothing
Set pf_vlookup = Nothing

Exit Sub

errHandler:
MsgBox "Error: " & Err.Description, vbExclamation
GoTo CleanUp
End Sub


Sub test6_b()

Dim wb As Workbook
Dim wsSource As Worksheet, wsTarget As Worksheet
Dim LastRow As Long, LastColumn As Long
Dim SourceDataRange As Range
Dim PTCache As PivotCache
Dim PT As PivotTable
Dim pf_globalRei As PivotField
Dim pf_vlookup As PivotField
Dim vlookupArray(1) As String
Dim numberOfElements As Integer
Dim pf_ARC As PivotField
Dim arcArray(0 To 10) As String
Dim numberOfElementsTwo As Integer
Dim i As Integer
Dim j As Integer
Dim c As Integer
Dim d As Integer


vlookupArray(1) = "#N/A"

arcArray(0) = "No Consequences or Impact to Patient"
arcArray(1) = "No Known Impact Or Consequence To Patient"
arcArray(2) = ""
arcArray(3) = "Device No Known Device Problem"
arcArray(4) = "Device No Reported Allegation"
arcArray(5) = "Insufficient Information"
arcArray(6) = "No Clinical Signs, Symptoms or Conditions"
arcArray(7) = "No Code Available"
arcArray(8) = "No Health Consequences or Impact"
arcArray(9) = "No Information"
arcArray(10) = "No Patient Involvement"



On Error GoTo errHandler
Set wb = ActiveWorkbook
Set wsTarget = wb.Worksheets("T6")
wsTarget.Select
'wsTarget.Cells.Clear


'// Step 1. Define data soruce
Set wsSource = wb.Worksheets("Complaints")
With wsSource
LastRow = .Cells(Rows.Count, "A").End(xlUp).Row
LastColumn = .Cells(1, Columns.Count).End(xlToLeft).Column


'// Define source data object
Set SourceDataRange = .Range(.Cells(1, 1), .Cells(LastRow, LastColumn))
End With

'// 2. Set PT Cache
Set PTCache = wb.PivotCaches.Create(xlDatabase, SourceDataRange)

'// 3. Create Pivot Table
Set PT = PTCache.CreatePivotTable(wsTarget.Range("J5"), "T6B")
With PT
'// Show GrandTotals
.ColumnGrand = True
.RowGrand = True

.RowAxisLayout xlTabularRow

.TableStyle2 = "PivotStyleDark16"

'Layout and Format
'.MergeLabels = True

.DisplayNullString = True
.NullString = 0

'// Add Pivot Fields

'// Filters
With .PivotFields("vlookup")
.Orientation = xlPageField
.EnableMultiplePageItems = True
End With


With .PivotFields("Global REI")
.Orientation = xlPageField
.EnableMultiplePageItems = True
End With

With .PivotFields("AR Code Description (GCMS)2")
.Orientation = xlPageField
'    .EnableMultiplePageItems = True
End With

Set pf_globalRei = PT.PivotFields("Global REI")
Set pf_vlookup = PT.PivotFields("vlookup")
Set pf_ARC = PT.PivotFields("AR Code Description (GCMS)2")

pf_globalRei.ClearAllFilters
 pf_vlookup.ClearAllFilters

'// Enable multiple filters selection
pf_globalRei.EnableMultiplePageItems = True
pf_vlookup.EnableMultiplePageItems = True
pf_ARC.EnableMultiplePageItems = True


pf_globalRei.CurrentPage = "Yes"
'    pf_vlookup.PivotItems("#N/A").Visible = False


      '// only apply filter if present in data

numberOfElements = UBound(vlookupArray) - LBound(vlookupArray) + 1

If numberOfElements > 0 Then
    With pf_vlookup
        For i = 1 To pf_vlookup.PivotItems.Count
        j = 0
'            MsgBox pf_vlookup.PivotItems.Count
'            MsgBox pf_vlookup.PivotItems(i).Name
'            MsgBox vlookupArray(j)
            
        Do While j < numberOfElements
            
            
            If pf_vlookup.PivotItems(i).Name = vlookupArray(j) Then
            
                pf_vlookup.PivotItems(pf_vlookup.PivotItems(i).Name).Visible = False
                Exit Do
            Else
                pf_vlookup.PivotItems(pf_vlookup.PivotItems(i).Name).Visible = True
            End If
            j = j + 1
        Loop
        Next i
    End With
End If

'// Rows Section

With .PivotFields("Group")
.Orientation = xlRowField
.Subtotals(1) = False
End With

With .PivotFields("AR Code Description (GCMS)")
.Orientation = xlRowField
.Subtotals(1) = False
End With

'// Columns
With .PivotFields("Year")
.Orientation = xlColumnField
End With

'// Values
 With .PivotFields("Complaint ID")
.Orientation = xlDataField
.Function = xlCount
.NumberFormat = "#,###"
End With


 '// only apply filter if present in data

    numberOfElementsTwo = UBound(arcArray) - LBound(arcArray) + 1
'    MsgBox numberOfElementsTwo

If numberOfElementsTwo > 0 Then
    With pf_ARC
        For c = 1 To pf_ARC.PivotItems.Count

        d = 0
'            MsgBox pf_ARC.PivotItems.Count
'            MsgBox "Arc Array" + pf_ARC.PivotItems(c).Name
'            MsgBox d
            
        Do While d < numberOfElementsTwo
            
            

            If pf_ARC.PivotItems(c).Name = arcArray(d) Then
            
                pf_ARC.PivotItems(pf_ARC.PivotItems(c).Name).Visible = False
                Exit Do
            Else
                pf_ARC.PivotItems(pf_ARC.PivotItems(c).Name).Visible = True
            End If
            d = d + 1
        Loop
        Next c
    End With
End If


End With

CleanUp:
Set PT = Nothing
Set PTCache = Nothing
Set SourceDataRange = Nothing
Set wsSource = Nothing
Set wsTarget = Nothing
Set wb = Nothing
Set pf_globalRei = Nothing
Set pf_vlookup = Nothing

Exit Sub

errHandler:
MsgBox "Error: " & Err.Description, vbExclamation
GoTo CleanUp
End Sub