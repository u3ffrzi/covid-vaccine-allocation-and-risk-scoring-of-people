VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   15420
   ClientLeft      =   300
   ClientTop       =   1152
   ClientWidth     =   19752
   OleObjectBlob   =   "vaccine.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim idx As Integer
Dim val As Integer
Private Sub cbProv_Change()
If Not IsEmpty(cbProv) Then
        cbDist.Clear
       For Each c In Range(cbProv)
          cbDist.AddItem c.Value
       Next c
       End If
End Sub
'==================================================Supply Side============================================================='
'---------------This routine makes levels of planning selectable based on national, provincial or district--------------------
Private Sub ComboBox1_Change()  
    If ComboBox1.Value = "Provincial" Or ComboBox1.Value = "District" Then
        ComboBox2.Visible = True
    Else
        ComboBox2.Visible = False
        ComboBox3.Visible = False
        ComboBox3.Clear
        
    End If
    If ComboBox1.Value = "District" Then
        ComboBox3.Visible = True
        If ComboBox2.Value <> "" Then
        For Each c In Range(ComboBox2.Value)
        ComboBox3.AddItem c.Value
       Next c
       End If
    End If
    If ComboBox1.Value = "Provincial" Then
        ComboBox3.Visible = False
    End If
End Sub

'------------- This routine updates or removes district dropdown list --------------
Private Sub ComboBox2_Change()
    ComboBox3.Clear
    If ComboBox1.Value = "District" Then
        ComboBox3.Visible = True
        For Each c In Range(ComboBox2.Value)
            ComboBox3.AddItem c.Value
            Next c
        Else
            ComboBox3.Visible = False
        End If
    vacNo.Value = Application.WorksheetFunction.VLookup(ComboBox2.Value, Sheets("controls").Range("o:P"), 2, 0)
     ComboBox1.Value = "Provincial"
  Call ComboBox1_Change
    End Sub





Private Sub ComboBox3_Change()
On Error Resume Next
If Not IsEmpty(ComboBox3.Value) Then
    vacNo.Value = Application.WorksheetFunction.VLookup(ComboBox3.Value, Sheets("controls").Range("q:r"), 2, 0)
End If
End Sub

'============ Demand Side =================
Private Sub CommandButton1_Click()

'i = 1
'For Each ctl In UserForm1.Controls
'Cells(2 + i, 1) = ctl.Name
'i = i + 1
'Next


'-----------------risk score calculation
' The beta from regression analysis
Dim risk, risk2 As Variant
risk = Array(-5.150746, 0.2367982, 0.0418373, 1.025089, 0.4038391, 0.5943882, 0.3562976, 0.2469078, 0.1880846)
risk2 = risk


If userSex.Value <> Range("d2").Value Then
risk2(1) = 0
End If
If malig.Value = False Then
risk2(3) = 0
End If
If idd.Value = False Then
risk2(4) = 0
End If
If liver.Value = False Then
risk2(5) = 0
End If
If kidney.Value = False Then
risk2(6) = 0
End If
If dm.Value = False Then
risk2(7) = 0
End If
If cardio.Value = False Then
risk2(8) = 0
End If

'Calculate risk of the person based on their selections so far
score = Exp(risk2(0) + risk2(1) + risk2(2) * userAge.Value + risk2(3) + risk2(4) + risk2(5) + risk2(6) + risk2(7) + risk2(8))
riskScore = score / (1 + score)
riskResult.Caption = "Your risk score is : " + Format(riskScore, "###0.00") + Chr(10)

' leveling Risk of people to 5 groups
If riskScore <= 0.02352454 Then
gp = "EI"
gi = 5
ElseIf riskScore <= 0.09409816 Then
gp = "DI"
gi = 4
ElseIf riskScore <= 0.1881963 Then
gp = "CI"
gi = 3
ElseIf riskScore <= 0.2822945 Then
gp = "BI"
gi = 2
ElseIf riskScore > 0.2822945 Then
gp = "AI"
gi = 1
End If

'Create and Update the output  text
riskResult.Caption = riskResult.Caption + " " + Sheets("controls").Range("j1").Value + Application.WorksheetFunction.VLookup(gp, Sheets("controls").Range("I:J"), 2, 0) + Sheets("controls").Range("i1").Value + Chr(10)
'pos = Application.Match("0", classValue, False)

If val > gi Then
res = Sheets("controls").Range("k1").Value
ElseIf val = gi Then
res = Sheets("controls").Range("k2").Value

Else
res = Sheets("controls").Range("k3").Value
End If
riskResult.Caption = riskResult.Caption + res

' Check national id format
If Not (IsNumeric(userId.Value) And Len(userId.Value) = 10) Then
        MsgBox "only 10 numbers allowed in NationalID"
Else

' Save the user data to database 
Sheets("USERS").Activate
nRow = Range("a2").CurrentRegion.Rows.Count
nRow = nRow + 1
Cells(nRow, 1) = userName.Value
Cells(nRow, 2) = userId.Value
Cells(nRow, 3) = userAge.Value
Cells(nRow, 4) = userSex.Value
Cells(nRow, 5) = userOcc.Value
Cells(nRow, 6) = userPreg.Value

MsgBox "Done"
End If
End Sub



Public Sub lbDrillDown_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

If ComboBox1.Value = "National" Then
ComboBox1.Value = "Provincial"
Call ComboBox1_Change
Call ComboBox2_Change
ComboBox2.Value = lbDrillDown.Value
vacNo.Value = lbDrillDown2.List(lbDrillDown.ListIndex)
idx = lbDrillDown.ListIndex
'ElseIf ComboBox1.Value = "Provincial" Then
Else
vacNo.Value = lbDrillDown2.List(lbDrillDown.ListIndex)


ComboBox1.Value = "District"
Call ComboBox1_Change
Call ComboBox2_Change
ComboBox3.Value = lbDrillDown.Value

idx = lbDrillDown.ListIndex
End If
Call startAlloc_Click
'MsgBox lbDrillDown.Value
'MsgBox lbDrillDown.ListIndex

End Sub

======== Supply Side
-------------- vaccine allocation to provinces and districts based on 5 risk levels and also public servant categories
    Public Sub startAlloc_Click()
    'Main routine to start the allocation system 
    
   
    'Initializing variables and sheets
        Sheets(2).ListObjects("vacNeed").Range.AutoFilter
        lbResult.Caption = Sheets("controls").Range("m1").Value
        Dim classess(), classNames(), classNeed() As Variant
        classNeed = Array(0, 0, 0, 0, 0, 0, 0, 0, 0)
        classess = Array("AP", "AI", "BP", "BI", "CP", "CI", "DI", "S", "EI")
       classNames = Array("High Risk - Health Professionals", "Very High Risk - General Population", "Medium Risk - Health Professionals", "High Risk - General Population", "Low Risk - Health Professionals", "Medium Risk - General Population", "Low Risk - General Population", "Residentials", "Very Low Risk - General Population")
        'classNames = Sheets("controls").Range("m2:m10").Value
        
        'Total need for vaccine in different levels of planning are calculated for different risk groups and printed on UI
        For i = 0 To UBound(classess)
            If ComboBox1.Value = "National" Then
            Sheets(4).Range("a1").Formula = "=GETPIVOTDATA(""VaccineNeed"" ,pivot!$A$3,""Class"",""" + classess(i) + """)"

            ElseIf ComboBox3.Visible = True Then
                
                Sheets(4).Range("a1").Formula = "=GETPIVOTDATA(""VaccineNeed"" ,pivot!$A$3,""Province"",""" + ComboBox2.Value + """,""District"",""" + ComboBox3.Value + """,""Class"",""" + classess(i) + """)"
            ElseIf ComboBox3.Visible = False Then
                Sheets(4).Range("a1").Formula = "=GETPIVOTDATA(""VaccineNeed"" ,pivot!$A$3,""Province"",""" + ComboBox2.Value + """,""Class"",""" + classess(i) + """)"
            End If
            lb = Sheets(4).Range("a1").Value
            classNeed(i) = lb
            lbResult.Caption = lbResult.Caption & Chr(10) & classNames(i) & ": " & CStr(lb)
        Next
        
        ' Initialize allocated vaccine numbers to risk groups
        classValue = Array(0, 0, 0, 0, 0, 0, 0, 0, 0)
        classValue(0) = Application.WorksheetFunction.Min(vacNo, classNeed(0))
        
        '------------- Allocate vaccine to different risk groups based on entered value by user
        '------------- based on flowchart of prioritization
        If Int(vacNo) >= classNeed(0) + classNeed(1) + classNeed(2) + classNeed(3) + classNeed(4) + classNeed(5) + classNeed(6) + classNeed(7) + classNeed(8) Then
            classValue(0) = classNeed(0)
            classValue(1) = classNeed(1)
            classValue(2) = classNeed(2)
            classValue(3) = classNeed(3)
            classValue(4) = classNeed(4)
            classValue(5) = classNeed(5)
            classValue(6) = classNeed(6)
            classValue(7) = classNeed(7)
            classValue(8) = classNeed(8)
          val = 5
        ElseIf Int(vacNo) > classNeed(0) + classNeed(1) + classNeed(2) + classNeed(3) + classNeed(4) + classNeed(5) + classNeed(6) + classNeed(7) Then
            classValue(0) = classNeed(0)
            classValue(1) = classNeed(1)
            classValue(2) = classNeed(2)
            classValue(3) = classNeed(3)
            classValue(4) = classNeed(4)
            classValue(5) = classNeed(5)
            classValue(6) = classNeed(6)
            usedVac = classValue(0) + classValue(1) + classValue(2) + classValue(3) + classValue(4) + classValue(5) + classValue(6)
            classValue(7) = Application.WorksheetFunction.Min(classNeed(7), 0.8 * (vacNo - usedVac))
            usedVac = usedVac + classValue(7)
            classValue(8) = Application.WorksheetFunction.Min(classNeed(8), vacNo - usedVac)
            val = 5
        ElseIf Int(vacNo) > classNeed(0) + classNeed(1) + classNeed(2) + classNeed(3) + classNeed(4) + classNeed(5) + classNeed(6) Then
            classValue(0) = classNeed(0)
            classValue(1) = classNeed(1)
            classValue(2) = classNeed(2)
            classValue(3) = classNeed(3)
            classValue(4) = classNeed(4)
            classValue(5) = classNeed(5)
            usedVac = classValue(0) + classValue(1) + classValue(2) + classValue(3) + classValue(4) + classValue(5)
            classValue(6) = Application.WorksheetFunction.Min(classNeed(6), 0.8 * (vacNo - usedVac))
            usedVac = usedVac + classValue(6)
            classValue(7) = Application.WorksheetFunction.Min(classNeed(7), vacNo - usedVac)
            val = 4
        ElseIf Int(vacNo) > classNeed(0) + classNeed(1) + classNeed(2) + classNeed(3) + classNeed(4) + classNeed(5) Then
            classValue(0) = classNeed(0)
            classValue(1) = classNeed(1)
            classValue(2) = classNeed(2)
            classValue(3) = classNeed(3)
            classValue(4) = classNeed(4)
            usedVac = classValue(0) + classValue(1) + classValue(2) + classValue(3) + classValue(4)
            classValue(5) = Application.WorksheetFunction.Min(classNeed(5), 0.8 * (vacNo - usedVac))
            usedVac = usedVac + classValue(5)
            classValue(6) = Application.WorksheetFunction.Min(classNeed(6), vacNo - usedVac)
        val = 4
        ElseIf Int(vacNo) > classNeed(0) + classNeed(1) + classNeed(2) + classNeed(3) + classNeed(4) Then
            classValue(0) = classNeed(0)
            classValue(1) = classNeed(1)
            classValue(2) = classNeed(2)
            classValue(3) = classNeed(3)
            usedVac = classValue(0) + classValue(1) + classValue(2) + classValue(3)
            classValue(4) = Application.WorksheetFunction.Min(classNeed(4), 0.8 * (vacNo - usedVac))
            usedVac = usedVac + classValue(4)
            classValue(5) = Application.WorksheetFunction.Min(classNeed(5), vacNo - usedVac)
            val = 3
        ElseIf Int(vacNo) > classNeed(0) + classNeed(1) + classNeed(2) + classNeed(3) Then
            classValue(0) = classNeed(0)
            classValue(1) = classNeed(1)
            classValue(2) = classNeed(2)
            usedVac = classValue(0) + classValue(1) + classValue(2)
            classValue(3) = Application.WorksheetFunction.Min(classNeed(3), 0.8 * (vacNo - usedVac))
            usedVac = usedVac + classValue(3)
            classValue(4) = Application.WorksheetFunction.Min(classNeed(4), vacNo - usedVac)
           val = 2
        ElseIf Int(vacNo) > classNeed(0) + classNeed(1) + classNeed(2) Then
            classValue(0) = classNeed(0)
            classValue(1) = classNeed(1)
            usedVac = classValue(0) + classValue(1)
            classValue(2) = Application.WorksheetFunction.Min(classNeed(2), 0.8 * (vacNo - usedVac))
            usedVac = usedVac + classValue(2)
            classValue(3) = vacNo - usedVac
        val = 2
        ElseIf Int(vacNo) > classNeed(0) + classNeed(1) Then
            classValue(0) = classNeed(0)

            usedVac = classValue(0)
            classValue(1) = Application.WorksheetFunction.Min(classNeed(1), 0.8 * (vacNo - usedVac))
            usedVac = usedVac + classValue(1)
            classValue(2) = Application.WorksheetFunction.Min(classNeed(2), vacNo - usedVac)
           val = 1
        Else
            usedVac = 0
            classValue(0) = Application.WorksheetFunction.Min(classNeed(0), 0.8 * (vacNo - usedVac))
            usedVac = usedVac + classValue(0)
            classValue(1) = Application.WorksheetFunction.Min(classNeed(1), vacNo - usedVac)
        val = 1
        End If
        
        ' Show the results in UI
        lbAllocated.Caption = Sheets("controls").Range("n1").Value
        For j = 0 To UBound(classess)
            lbAllocated.Caption = lbAllocated.Caption & Chr(10) & classNames(j) & ": " & CStr(Round(classValue(j), 0))
        Next
        
        
    'fill the details window
    

    lbDrillDown.TextAlign = fmTextAlignLeft
    lbDrillDown2.TextAlign = fmTextAlignCenter
   
     If ComboBox1.Value = "National" Then
             lbDrillDown.Clear
    lbDrillDown2.Clear
    Range("o:p").Clear
  
    For i = 1 To 31

    lbDrillDown.AddItem (Sheets("provPivot").Range("a" + CStr(4 + i)).Value)
    Sheets("controls").Range("o" + CStr(i)).Value = (Sheets("provPivot").Range("a" + CStr(4 + i)).Value)
  ' ,": ", Format(Sheets("provPivot").Range("k" + CStr(4 + i)).Value, "0.00%"))
    lbDrillDown2.AddItem (Format(Round(Sheets("provPivot").Range("k" + CStr(4 + i)).Value * vacNo, 0), "#,##0"))
     Sheets("controls").Range("p" + CStr(i)).Value = (Format(Round(Sheets("provPivot").Range("k" + CStr(4 + i)).Value * vacNo, 0), "#,##0"))
    Next
    Else
    Sheets("distPivot").Activate
     ActiveSheet.PivotTables("districtPivot").PivotFields("Province").ClearAllFilters
    For i = 1 To ActiveSheet.PivotTables("districtPivot").PivotFields("Province").PivotItems.Count
    If ActiveSheet.PivotTables("districtPivot").PivotFields("Province").PivotItems(i) = ComboBox2.Value Then
         ActiveSheet.PivotTables("districtPivot").PivotFields("Province").PivotItems(i).Visible = True
     Else
        ActiveSheet.PivotTables("districtPivot").PivotFields("Province").PivotItems(i).Visible = False
      End If
      Next
        Range("q:r").Clear
  If ComboBox3.Visible = False Then
     nRow = Range("b6").End(xlDown).Row - 1
     lbDrillDown.Clear
    lbDrillDown2.Clear
        For i = 6 To nRow
    lbDrillDown.AddItem (Sheets("distPivot").Range("a" + CStr(i)).Value)
     Sheets("controls").Range("q" + CStr(i)).Value = (Sheets("distPivot").Range("a" + CStr(i)).Value)
  '  ,": ", Format(Sheets("provPivot").Range("k" + CStr(4 + i)).Value, "0.00%"))
    lbDrillDown2.AddItem (Format(Round(Sheets("distPivot").Range("k" + CStr(i)).Value * vacNo, 0), "#,##0"))
     Sheets("controls").Range("r" + CStr(i)).Value = (Format(Round(Sheets("distPivot").Range("k" + CStr(i)).Value * vacNo, 0), "#,##0"))
    Next
End If
    
    End If
    
    

    
    
    
    End Sub
    

'============== UI/UX  controls
    Private Sub UserForm_Initialize()
      
    'populate radioButtons
    Sheets("controls").Activate
    
    For i = 1 To Sheets(1).Range("b1").End(xlDown).Row
            cbProv.AddItem Sheets(1).Range("b" + CStr(i)).Value
        Next
        
    Dim ctl As Control
    
  For Each ctl In UserForm1.Controls
   If Not IsEmpty(Application.WorksheetFunction.VLookup(ctl.Name, Range("A:C"), 3, 0)) Then
   ctl.Caption = Application.WorksheetFunction.VLookup(ctl.Name, Range("A:C"), 3, 0)

  End If
  Next
    For i = 4 To Sheets("controls").Range("d1").End(xlDown).Row
            userSex.AddItem Sheets("controls").Range("d" + CStr(i)).Value
   Next
       For i = 6 To Sheets("controls").Range("e1").End(xlDown).Row
            userOcc.AddItem Sheets("controls").Range("e" + CStr(i)).Value
   Next

 
       For i = 4 To Sheets("controls").Range("f1").End(xlDown).Row
            userPreg.AddItem Sheets("controls").Range("f" + CStr(i)).Value
   Next

    
    'Populate dropdowns at start of application
        ComboBox2.Visible = False
        ComboBox3.Visible = False
        
        For i = 1 To Sheets(1).Range("a1").End(xlDown).Row
            ComboBox1.AddItem Sheets(1).Range("a" + CStr(i)).Value
        Next
        For i = 1 To Sheets(1).Range("b1").End(xlDown).Row
            ComboBox2.AddItem Sheets(1).Range("b" + CStr(i)).Value
        Next
        
        'Set Color and other Visual effects of UI
   
        
        
        

        For Each ctl In Me.Controls
            If TypeName(ctl) = "Label" Or TypeName(ctl) = "TextBox" Or TypeName(ctl) = "ComboBox" Then
               ctl.Font.Size = 13
                ctl.BorderColor = &HA9A9A9
                ctl.BackColor = &HFFFFFF
                ctl.ForeColor = &H464646
                 ctl.Font.Name = "Calibri"
            End If
                        If TypeName(ctl) = "TextBox" Or TypeName(ctl) = "ComboBox" Then
               ctl.SpecialEffect = 0
            ctl.BorderStyle = fmBorderStyleSingle
                ctl.BorderColor = &HA9A9A9
                                 ctl.Font.Name = "Calibri"

            End If
            If TypeName(ctl) = "CommandButton" Then
         ctl.Font.Size = 16
           
                ctl.BackColor = RGB(97, 81, 146)
                 ctl.Font.Name = "Calibri"
ctl.ForeColor = RGB(170, 151, 57)
            End If
            If TypeName(ctl) = "Label" Then
            
             ctl.BackColor = &H8000000F
            End If
            Next ctl
            lbTitleDemand.Font.Size = 25
            lbResult.BorderStyle = fmBorderStyleSingle
            lbAllocated.BorderStyle = fmBorderStyleSingle
            lbDrillDown.BorderStyle = fmBorderStyleSingle
            lbDrillDown.SpecialEffect = fmSpecialEffectBump
            lbDrillDown.Font.Name = "Calibri"
            lbDrillDown.Font.Size = 10
            lbDrillDown2.BorderStyle = fmBorderStyleSingle
            lbDrillDown2.SpecialEffect = fmSpecialEffectBump
            lbDrillDown2.Font.Name = "Calibri"
            lbDrillDown2.Font.Size = 10
      
'            lbResult.BackColor = RGB(10, 28, 2)
'            lbResult.ForeColor = RGB(100, 200, 2)
'            lbAllocated.BackColor = RGB(10, 28, 2)
'            lbAllocated.ForeColor = RGB(100, 200, 2)
'            startAlloc.BackColor = RGB(175, 38, 53)
'            startAlloc.ForeColor = RGB(10, 28, 2)
        End Sub
        
        
        
'---------- update UI for female or male specific risk factors
Private Sub userSex_Change()
If userSex.Value = Range("D3").Value Then
lbPregnancy.Visible = True
userPreg.Visible = True
Else
lbPregnancy.Visible = False
userPreg.Visible = False


End If

End Sub


