Public tdc As Object



Sub login()
    
    Dim qcURL As String
    Dim qcID As String
    Dim qcPWD As String
    Dim qcDomain As String
    Dim qcProject As String
    
    On Error GoTo err
    qcURL = Sheets("login").TextBox1.Text
    qcID = Sheets("login").TextBox2.Text
    qcPWD = Sheets("login").TextBox3.Text
    qcDomain = Sheets("login").TextBox4.Text
    qcProject = Sheets("login").TextBox5.Text
'Display a message in Status bar
    Application.StatusBar = "Connecting to Quality Center.. Wait..."
' Create a Connection object to connect to Quality Center
    Set tdc = CreateObject("TDApiOle80.TDConnection")
'Initialise the Quality center connection
    tdc.InitConnectionEx qcURL
'Authenticating with username and password
    tdc.login qcID, qcPWD
'connecting to the domain and project
    tdc.Connect qcDomain, qcProject
'On successfull login display message in Status bar
    Application.StatusBar = "........QC Connection is done Successfully"
    MsgBox "QC Connection Successfully"
    
err:
    MsgBox err.Description
    'Display the error message in Status bar
    Application.StatusBar = err.Description
    
    
End Sub


Sub clear()
    Sheets("login").TextBox2.Value = ""
    Sheets("login").TextBox3.Value = ""
End Sub

Sub import()
    
    If Sheets("login").TextBox6.Value = "" And Sheets("login").TextBox7.Value = "" Then
        MsgBox "  Enter the last row & column number until the update is needed  "
'End Sub
    Else
        Dim rc As Integer
        Dim cc As Integer
        rc = Sheets("login").TextBox6.Value
        cc = Sheets("login").TextBox7.Value
        Dim q As Integer
        q = 0
        
        For i = 2 To rc
            
            For j = 2 To cc
                testPlanID = Sheets("data").Cells(i, 1).Value
                Set TestList = tdc.TestFactory
                Set TestPlanFilter = TestList.Filter
                TestPlanFilter.Filter("TS_TEST_ID") = testPlanID
                Set TestPlanList = TestList.NewList(TestPlanFilter.Text)
                Set myTestPlan = TestPlanList.Item(1)
                
                'posting data
                myTestPlan.Field(Sheets("reference").Cells(2, j).Value) = Sheets("data").Cells(i, j).Value
                myTestPlan.Post
                
                Set TestPlanFilter = Nothing
                Set myTestPlan = Nothing
                Set TestList = Nothing
                Set TestPlanFilter = Nothing
                q = q + 1
                Next j
                Next i
                MsgBox "  Data importeded to " & q & " test cases "
                
            End If
            
        End Sub
        
        
        Sub export()
            
            If Sheets("login").TextBox6 = "" Or Sheets("login").TextBox7 = "" Then
                MsgBox "  Enter the last row & column number until the update is needed  "
            Else
                Dim rc As Integer
                Dim cc As Integer
                rc = Sheets("login").TextBox6.Value
                cc = Sheets("login").TextBox7.Value
                Dim q As Integer
                q = 0
                
                For i = 2 To rc
                    
                    For j = 2 To cc
                        testPlanID = Sheets("data").Cells(i, 1).Value
                        
                        Set TestList = tdc.TestFactory
                        Set TestPlanFilter = TestList.Filter
                        TestPlanFilter.Filter("TS_TEST_ID") = testPlanID
                        Set TestPlanList = TestList.NewList(TestPlanFilter.Text)
                        Set myTestPlan = TestPlanList.Item(1)
                        
                        'getting data
                        Dim dt As String
                        dt = myTestPlan.Field(Sheets("reference").Cells(2, j).Value)
                        Sheets("data").Cells(i, j).Value = dt
                        
                        Set TestPlanFilter = Nothing
                        Set myTestPlan = Nothing
                        Set TestList = Nothing
                        Set TestPlanFilter = Nothing
                        q = q + 1
                        Next j
                        Next i
                        MsgBox "  Data exported for " & q & " test cases "
                        
                    End If
                End Sub
                

Sub UnhideAllSheets()

    'Unhide all sheets in workbook.
    Dim ws As Worksheet
    
    For Each ws In ActiveWorkbook.Worksheets
    ws.Visible = xlSheetVisible
    Next ws

End Sub
