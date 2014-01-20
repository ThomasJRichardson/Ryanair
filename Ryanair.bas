Attribute VB_Name = "Module1"
Option Explicit

Sub Upload_Ryanair()
Dim objPerson As PSPerson
Dim objAddress As PSAddress
Dim objEmployee As PSEmployee
Dim objSecondEmployee As PSEmployee
Dim objEmployer As PSEmployer
Dim objJobClass As PSHistory
Dim objSecondJobClass As PSHistory
Dim objEmploymentEvents As PSHistory
Dim objSecondEmploymentEvents As PSHistory
Dim objScheme As PSScheme
Dim objSchemeMember As PSSchemeMember
Dim objSecondSchemeMember As PSSchemeMember
Dim objSalaryHist As PSHistory
Dim objSecondSalaryHist As PSHistory
Dim objMarStatus As PSHistory

Dim objEnv As PSEnvironment
Dim varCatID As Variant
Dim v2 As Variant
Dim EmployeeRef As Long
Dim blnValidated As Boolean
Dim blnInCommitment As Boolean
Dim blnIn2ndCommitment As Boolean
Dim blnSecondEmployee As Boolean
Dim blnHasAddress As Boolean

Dim xlWb As Excel.Workbook
Dim xlWs As Excel.Worksheet
Dim xlRow As Long

On Error GoTo ErrorHandler

'Initialise
blnInCommitment = False
blnIn2ndCommitment = False

Set xlWb = Excel.Application.ActiveWorkbook
Set xlWs = xlWb.Worksheets("Full benefits")

xlRow = 2

EmployeeRef = 0
Set objEnv = New PSEnvironment

    'Init Employer
If Not objEmployer Is Nothing Then asControl(objEmployer).Clear
Set objEmployer = New PSEmployer
asControl(objEmployer).CurrentEnvironment = objEnv
asControl(objEmployer).Load "EmployerCode = 'Ryanair'", "", PSReadWrite
    
    'Init Scheme
If Not objScheme Is Nothing Then asControl(objScheme).Clear
Set objScheme = New PSScheme
asControl(objScheme).CurrentEnvironment = objEnv
asControl(objScheme).Load "Name = 'Ryanair Retirement and Death Benefit Plan'", "", PSReadWrite

'Process workdata
While xlRow <= 440
    xlWs.Cells(xlRow, 1).Select
    blnSecondEmployee = False
    
    'Person
    If Not objPerson Is Nothing Then asControl(objPerson).Clear
    Set objPerson = New PSPerson
    
    asControl(objPerson).CurrentEnvironment = objEnv
    asControl(objPerson).Load "1=0", "", PSReadWrite
    
    asBase(objPerson).MakeNew
    
    With objPerson
        .NationalIDNumber = xlWs.Cells(xlRow, 1).Value  'Username
        .Reference = .NationalIDNumber
        .Salutation = xlWs.Cells(xlRow, 3).Value        'Actual PPSN
        .NationalIDValidType = "OTH"
        .Surname = xlWs.Cells(xlRow, 4).Value
	.Initials = xlWs.Cells(xlRow, 5).Value
        .Forename = xlWs.Cells(xlRow, 6).Value
        .Sex = Left(UCase(xlWs.Cells(xlRow, 14).Value), 1)
        .DateOfBirth = Format(xlWs.Cells(xlRow, 12).Value, "dd/mm/yyyy")
        .PrevSurname = "NotUpdated"
    End With
    
    'Address
    If Not objAddress Is Nothing Then asControl(objAddress).Clear
    
    If xlWs.Cells(xlRow, 7).Formula = "" Then
        blnHasAddress = False
    Else
        blnHasAddress = True
    End If
    
    If blnHasAddress Then
        Set objAddress = asBase(objPerson).Addresses
        asChild(objAddress).MakeNewOfType "HOMEADD"
        
        With objAddress
            .Line1 = xlWs.Cells(xlRow, 7).Value

            If xlWs.Cells(xlRow, 8).Formula <> "" then
		.Line2 = xlWs.Cells(xlRow, 8).Value

            	If xlWs.Cells(xlRow, 9).Formula <> "" then
			.Line3 = xlWs.Cells(xlRow, 9).Value

	    		If xlWs.Cells(xlRow, 11).Formula <> "" then
	    			If xlWs.Cells(xlRow, 11).Formula <> "" then
            				.Line4 = xlWs.Cells(xlRow, 10).Value + ", " + xlWs.Cells(xlRow, 11).Value
	    			Else
            				.Line4 = xlWs.Cells(xlRow, 10).Value
	    			End If
	    		End If
		End If
	    End If

            .EffDate = Format(xlWs.Cells(xlRow, 18).Value, "dd/mm/yyyy")
        End With
    End If 'If blnHasAddress Then
    
    'Employee
    If Not objEmployee Is Nothing Then asControl(objEmployee).Clear
    Set objEmployee = New PSEmployee
    
    asControl(objEmployee).CurrentEnvironment = objEnv
    asControl(objEmployee).Load "1=0", "", PSReadWrite
    
    asBase(objEmployee).MakeNew
    
    With objEmployee
        .SetPerson objPerson
        .SetEmployer objEmployer
        .DateFirstEmployed = Format(xlWs.Cells(xlRow, 18).Value, "dd/mm/yyyy")
        If xlWs.Cells(xlRow, 23).Value = 0 Then
            .PayrollNumber = "NO tsfr"
        Else
            .PayrollNumber = "YES tsfr"
        End If
    End With

    blnSecondEmployee = False
    
    'Employee events
    If Not objEmploymentEvents Is Nothing Then asControl(objEmploymentEvents).Clear
    Set objEmploymentEvents = New PSHistory
    asControl(objEmploymentEvents).CurrentEnvironment = objEnv
    
    Set varCatID = Nothing
    varCatID = asControl(objEmployee).GetCatids("EMPEVHIST")
    'For Each v2 In varCatID
    '    MsgBox CStr(v2)
    'Next v2
    objEmploymentEvents.[_loadHistoryForBatch] varCatID, "1=0", "", "IntegerHistory", PSReadWrite
    
    asChild(objEmploymentEvents).MakeNewOfType "EMPEVHIST"
    asChild(objEmploymentEvents).ParentUID = asBase(objEmployee).Uid
    
    objEmploymentEvents.Value = 4113  '(EMPLOYED)
    objEmploymentEvents.Date = objEmployee.DateFirstEmployed
    
    If xlWs.Cells(xlRow, 22).Formula <> "" Then
        asChild(objEmploymentEvents).MakeNewOfType "EMPEVHIST"
        asChild(objEmploymentEvents).ParentUID = asBase(objEmployee).Uid
    
        objEmploymentEvents.Value = 4119  '(LEAVES)
        objEmploymentEvents.Date = Format(xlWs.Cells(xlRow, 22).Value, "dd/mm/yyyy")
    End If
    
    'Job class
    If Not objJobClass Is Nothing Then asControl(objJobClass).Clear
    Set objJobClass = New PSHistory
    asControl(objJobClass).CurrentEnvironment = objEnv
    
    Set varCatID = Nothing
    varCatID = asControl(objEmployee).GetCatids("EEEJCGRP")
    'For Each v2 In varCatID
    '    MsgBox CStr(v2)
    'Next v2
    objJobClass.[_loadHistoryForBatch] varCatID, "1=0", "", "IntegerHistory", PSReadWrite
    asChild(objJobClass).MakeNewOfType "EEEJOBCL"
    asChild(objJobClass).ParentUID = asBase(objEmployee).Uid

    objJobClass.Date = objEmployee.DateFirstEmployed
    objJobClass.Value = DeriveJobClass(xlWs.Cells(xlRow, 16).Value)
        
    'Scheme Member
    If Not objSchemeMember Is Nothing Then asControl(objSchemeMember).Clear
    Set objSchemeMember = New PSSchemeMember
    asControl(objSchemeMember).CurrentEnvironment = objEnv
    
    asControl(objSchemeMember).Load "1=0", "", PSReadWrite
    asBase(objSchemeMember).MakeNew
    
    objSchemeMember.SetEmployee objEmployee
    objSchemeMember.SetScheme objScheme
        'DJS
    If xlWs.Cells(xlRow, 19).Formula <> "" Then
        If IsDate(xlWs.Cells(xlRow, 19).Value) Then
            objSchemeMember.DateJoinedScheme = Format(xlWs.Cells(xlRow, 19).Value, "dd/mm/yyyy")
        End If
    Else
        If xlWs.Cells(xlRow, 18).Formula <> "" Then
            If IsDate(xlWs.Cells(xlRow, 18).Value) Then
                objSchemeMember.DateJoinedScheme = Format(xlWs.Cells(xlRow, 18).Value, "dd/mm/yyyy")
            End If
        End If
    End If

        'Normal retirement date
        If xlWs.Cells(xlRow, 13).Formula <> "" Then
            If IsDate(xlWs.Cells(xlRow, 13).Value) Then
                objSchemeMember.SchemeRetirementDate = Format(xlWs.Cells(xlRow, 13).Value, "dd/mm/yyyy")
            End If
        End If

        'Scheme Reference increment
    objSchemeMember.MemberReference = EmployeeRef
        'increment employee ref ??
    EmployeeRef = EmployeeRef + 1


    objSchemeMember.RetainedBenefitRulesApply = False
    objSchemeMember.AVCPayer = False


    objSchemeMember.NominationReceived = False
        
    'Marital Status
    If Not objMarStatus Is Nothing Then asControl(objMarStatus).Clear
    Set objMarStatus = New PSHistory
    asControl(objMarStatus).CurrentEnvironment = objEnv

    Set varCatID = Nothing
    varCatID = asControl(objPerson).GetCatids("MARSTATUS")
    'For Each v2 In varCatID
    '    MsgBox CStr(v2)
    'Next v2
    objMarStatus.[_loadHistoryForBatch] varCatID, "1=0", "", "StringHistory", PSReadWrite
    
    asChild(objMarStatus).MakeNewOfType "MARSTATUS"
    asChild(objMarStatus).ParentUID = asBase(objPerson).Uid
    objMarStatus.Date = objSchemeMember.DateJoinedScheme
    objMarStatus.Value = DeriveMarStatus(xlWs.Cells(xlRow, 15).Value)
    
    'Salary
    If Not objSalaryHist Is Nothing Then asControl(objSalaryHist).Clear
    
    Set objSalaryHist = New PSHistory
    asControl(objSalaryHist).CurrentEnvironment = objEnv

    Set varCatID = Nothing
    varCatID = asControl(objEmployee).GetCatids("SALGRP")
    'For Each v2 In varCatID
    '    MsgBox CStr(v2)
    'Next v2
    objSalaryHist.[_loadHistoryForBatch] varCatID, "1=0", "", "CurrencyHistory", PSReadWrite

    asChild(objSalaryHist).MakeNewOfType "BASSAL"
    asChild(objSalaryHist).ParentUID = asBase(objEmployee).Uid
    
    If xlWs.Cells(xlRow, 26).Formula = "" Then
        objSalaryHist.Value = 0
    ElseIf IsNumeric(xlWs.Cells(xlRow, 26).Value) Then
        objSalaryHist.Value = Format(xlWs.Cells(xlRow, 26).Value, "###,###.00")
    Else
        objSalaryHist.Value = 0
    End If
    
    If xlWs.Cells(xlRow, 22).Formula = "" Then
        objSalaryHist.Date = Format("01-jan-2014", "dd/mm/yyyy")
    Else
        objSalaryHist.Date = Format(xlWs.Cells(xlRow, 22).Value, "dd/mm/yyyy")
    End If

    asChild(objSalaryHist).MakeNewOfType "PENSAL"
    asChild(objSalaryHist).ParentUID = asBase(objEmployee).Uid
    
    If xlWs.Cells(xlRow, 27).Formula = "" Then
        objSalaryHist.Value = 0
    ElseIf IsNumeric(xlWs.Cells(xlRow, 27).Value) Then
        objSalaryHist.Value = Format(xlWs.Cells(xlRow, 27).Value, "###,###.00")
    Else
        objSalaryHist.Value = 0
    End If
    
    If xlWs.Cells(xlRow, 22).Formula = "" Then
        objSalaryHist.Date = Format("01-jan-2014", "dd/mm/yyyy")
    Else
        objSalaryHist.Date = Format(xlWs.Cells(xlRow, 22).Value, "dd/mm/yyyy")
    End If

    asChild(objSalaryHist).MakeNewOfType "SCHEMSALRY"
    asChild(objSalaryHist).ParentUID = asBase(objEmployee).Uid
    
    If xlWs.Cells(xlRow, 30).Formula = "" Then
        objSalaryHist.Value = 0
    ElseIf IsNumeric(xlWs.Cells(xlRow, 30).Value) Then
        objSalaryHist.Value = Format(xlWs.Cells(xlRow, 30).Value, "###,###.00")
    Else
        objSalaryHist.Value = 0
    End If
    
    If xlWs.Cells(xlRow, 22).Formula = "" Then
        objSalaryHist.Date = Format("01-jan-2014", "dd/mm/yyyy")
    Else
        objSalaryHist.Date = Format(xlWs.Cells(xlRow, 22).Value, "dd/mm/yyyy")
    End If
        
    'End of Object creation, now validate - any issues set blnValidated to False
    blnValidated = True
    Call Validate(objPerson, "Person", blnValidated)
    If blnHasAddress Then Call Validate(objAddress, "Address", blnValidated)
    Call Validate(objEmployee, "Employee", blnValidated)
    Call Validate(objEmploymentEvents, "Employment Events", blnValidated)
    Call Validate(objJobClass, "Job Class", blnValidated)
    Call Validate(objSchemeMember, "Scheme Member", blnValidated)
    Call Validate(objMarStatus, "Marital Status", blnValidated)
    Call Validate(objSalaryHist, "Salary", blnValidated)
    
    With Excel.Application.Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5296274
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

    'Save if blnValidated still true
    If blnValidated Then
        objEnv.StartTx
        blnInCommitment = True
        
        asControl(objPerson).Commit
        If blnHasAddress Then asControl(objAddress).Commit
        asControl(objEmployee).Commit
        asControl(objEmploymentEvents).Commit
        asControl(objJobClass).Commit
        asControl(objSchemeMember).Commit
        asControl(objMarStatus).Commit
        asControl(objSalaryHist).Commit
        
        objEnv.CommitTx
        blnInCommitment = False
    
    Else
        Debug.Print xlRow & " => Not Committed"
        
        With Excel.Application.Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 255
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    End If 'If blnValidated Then
    
NextXlRow:
    xlRow = xlRow + 1
    DoEvents
Wend
    
Exit Sub

ErrorHandler:
    Debug.Print xlRow & " => Failed => " & Err.Description
    
    If blnInCommitment Then
        Call Validate(objPerson, "Person", blnValidated)
        If blnHasAddress Then Call Validate(objAddress, "Address", blnValidated)
        Call Validate(objEmployee, "Employee", blnValidated)
        Call Validate(objEmploymentEvents, "Employment Events", blnValidated)
        Call Validate(objJobClass, "Job Class", blnValidated)
        Call Validate(objSchemeMember, "Scheme Member", blnValidated)
        Call Validate(objMarStatus, "Marital Status", blnValidated)
        Call Validate(objSalaryHist, "Salary", blnValidated)
        objEnv.AbortTx
    End If
    
    With Excel.Application.Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Resume NextXlRow
Resume
End Sub

Public Sub Validate(myObj As Object, strType As String, ByRef blnValidated As Boolean)
Dim objError As PSErrors
Dim enumVal As Long
    
    If Not asControl(myObj).IsValid Then
        blnValidated = False
        
        Set objError = asControl(myObj).ErrorList
        
        For enumVal = 0 To objError.Count - 1
            Debug.Print Excel.Application.ActiveCell.Row & " => Validation Error => " & strType & " => " & objError.Item(enumVal)
        Next enumVal
    End If
    
End Sub


Public Function DeriveJobClass(JobClass As String) As Long
Select Case JobClass
    Case "1- Staff DB"
        DeriveJobClass = 88662509
    Case "2- Pilots DB"
        DeriveJobClass = 88662510
    Case Else
End Select
End Function

Public Function DeriveMarStatus(MarStatus As String) As String
Select Case UCase(MarStatus)
    Case "SINGLE", "S", "SIN"
        DeriveMarStatus = "SIN"
    Case "MARRIED", "M", "MAR"
        DeriveMarStatus = "MAR"
    Case "DIVORCED", "D", "DIV"
        DeriveMarStatus = "DIV"
    Case "SEPARATED", "SEPERATED", "A", "APART", "LEGALLY SEPARATED", "LEGALLY SEPERATED", "APA"
        DeriveMarStatus = "APA"
    Case "WIDOWER", "W", "WID"
        DeriveMarStatus = "WID"
    Case Else
        DeriveMarStatus = "UNK"
End Select
End Function


