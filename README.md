- 👋 Hi, I’m @DimitrisAmpatzidis
- This is a project I was enabled, web scrapping data from a Maritime related site
- With different credentials to log-in and different vessels to navigate thorugh the website
- Using VBA and Excel macros to build the data-entry operation
- If you have any question please contact me E: d.ampatzidis@hotmail.gr, and LIN: https://www.linkedin.com/in/dimitriosampatzidis/

Start the procedure

      Sub WebScrapping()

Declare variables

    Dim k As Integer
    Dim P_INSP_Start As String
    Dim P_INSP_End As String
    Dim counter As Long
    Dim loginCounter As Integer
    Dim count As Integer, cc As Integer
    Dim Username, Password As String
    Dim idoc As MSHTML.HTMLDocument
    Dim IEobj As Object
    Dim Button As MSHTML.IHTMLElement
    Dim Buttons As MSHTML.IHTMLElementCollection
    Dim PressVessel As String, PV1 As String, PV2 As String, PV3 As String
    Dim string1 As String, string2 As String, string3 As String
    Dim ws As Worksheet
    Dim hTable As Object
    Dim td As Object, tr As Object, th As Object, r As Long, c As Long
    Dim i As Integer
    Dim j As Integer
    Dim IMO As String
    Dim Result As Integer
    Dim Result0 As Integer
    Dim CountNoInsp As Integer
    Dim refresher As Integer
    Dim CounterNoInsp As Integer

Place questions on the end-user making the procedure more friendly

    Result = 999

    Sheets("Database").Visible = True
    Sheets("Database").Select

    Result0 = MsgBox("Do you want to extract the 'Non Found IDs'?", vbQuestion + vbYesNo)
    If Result0 = vbYes Then
      Call NonFoundIDs
      Exit Sub
    End If

    Result0 = MsgBox("The procedure will start now; Please do not use any Excel file in the meantime. The total execution time estimated at: " & Sheets("User Sheet").Range("B4").Value & ". Press YES to continue..", vbQuestion + vbYesNo)
    If Result0 = vbNo Then
      Sheets("User Sheet").Select
      Exit Sub
    End If

    Result0 = MsgBox("Do you want to see the Excel updating?", vbQuestion + vbYesNo)
    If Result0 = vbNo Then
      Application.ScreenUpdating = False 'Not to update the Excel while the function is running..
    Else
      Application.ScreenUpdating = True
    End If

    Sheets("User Sheet").Select
    Sheets("Sheet1").Visible = True
    Sheets("Emails").Visible = True
    Sheets("Non Found IDs").Visible = True

Sorting emails in a random way so that we will log in each time with a different account and not trigger the website's security. Just FYI I used 80 different emails.

    ActiveWorkbook.Worksheets("Emails").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Emails").AutoFilter.Sort.SortFields.Add2 Key:=Range("C1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("Emails").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

Find how many unique searches we are gonna perform to have a live status-bar as the Excel is operating

    counter = 0
    P_INSP_Start = Sheets("User Sheet").Range("B1").Value
    P_INSP_End = Sheets("User Sheet").Range("B2").Value
    loginCounter = 0
    count = (P_INSP_End - P_INSP_Start) 'Count how many loops we will have
    cc = 0

Ask the end-user if he/she wants to see the IE while the Excel is extracting

    If Result = 999 Then
      Result = MsgBox("Do you want to see the Internet Explorer browser while web-scrapping?", vbQuestion + vbYesNo)
    End If
    
     If refresher = 0 Then
    'Create Internet Explorer with Question handling
            If Result = vbYes Then
                Set IE = New InternetExplorer
                IE.Visible = True 'Present Internet Explorer
            Else
                Set IE = New InternetExplorer
                IE.Visible = False 'Hide Internet Explorer
            End If
    End If

Open status bar (see at the end the code for this) & place a 'hot-point' to get back if it is needed

    OpenStatusBar 'Open the Status Bar for loading..

    Start:

    counter = counter + 1

    refresher = 0

Navigate to the site

    With IE
        .navigate ("website URL is typen here")
        While IE.ReadyState <> 4
            DoEvents
        Wend
    End With
    Application.Wait (Now + TimeValue("00:00:01"))

Read the email and password

    Sheets("Emails").Select
    Username = Sheets("Emails").Range("A" & loginCounter + 2).Value
    Password = Sheets("Emails").Range("B" & loginCounter + 2).Value
    Sheets("User Sheet").Select

Log in

    Application.Wait (Now + TimeValue("00:00:01"))
    Err.Clear
    On Error Resume Next
    Set IEobj = IE.document.getElementById("home-login")
    IEobj.Value = Username
    Application.Wait (Now + TimeValue("00:00:01"))
    Set IEobj = IE.document.getElementById("home-password")
    IEobj.Value = Password
    Application.Wait (Now + TimeValue("00:00:01"))
    Set idoc = IE.document
    Set Buttons = idoc.getElementsByClassName("pull-right btn btn-lg gris-bleu-copyright")
    For Each Button In Buttons
        If Button.className = "pull-right btn btn-lg gris-bleu-copyright" Then
            Button.Click
            Exit For
        End If
    Next Button
    Application.Wait (Now + TimeValue("00:00:01"))
    With IE
        While IE.ReadyState <> 4
            DoEvents
        Wend
    End With
    Application.Wait (Now + TimeValue("00:00:01"))

    Sheets("User Sheet").Range("K4").Value = "Error99" 'Error handling
    

Search for a randomly choses vessel just to trigger the systems shows us the right page

    Application.Wait (Now + TimeValue("00:00:01"))
    Set idoc = IE.document
    Set IEobj = idoc.getElementById("P_ENTREE_HOME")
    Application.Wait (Now + TimeValue("00:00:01"))
    IMO = Sheets("Database").Range("C10").Value
    IEobj.Value = IMO
    Application.Wait (Now + TimeValue("00:00:01"))
    Set idoc = IE.document
    Set Buttons = idoc.getElementsByClassName("btn btn-default")
    Application.Wait (Now + TimeValue("00:00:01"))
    For Each Button In Buttons
        If Button.className = "btn btn-default" Then
            Button.Click
            Exit For
        End If
    Next Button
    Application.Wait (Now + TimeValue("00:00:01"))
    With IE
        While IE.ReadyState <> 4
            DoEvents
        Wend
    End With
    Application.Wait (Now + TimeValue("00:00:02"))

Start the scrapping procedure

    Do Until P_INSP_Start > P_INSP_End
    Sheets("Database").Range("A2").Value = P_INSP_Start
    Application.Wait (Now + TimeValue("00:00:02"))
    
    Err.Clear
    On Error Resume Next

Press vessel-enable onclick function

    IMO = Sheets("Database").Range("C10").Value
    Set idoc = IE.document
    PV1 = ("document.formShip.P_IMO.value='")
    PV2 = PV1 & IMO
    PV3 = ("';document.formShip.submit();")
    PressVessel = PV2 & PV3
    Call idoc.parentWindow.execScript(PressVessel, "JavaScript")
    Application.Wait (Now + TimeValue("00:00:01"))
    With IE
      While IE.ReadyState <> 4
      DoEvents
      Wend
    End With
    Application.Wait (Now + TimeValue("00:00:02"))

Press Inspections-enable onclick

    Set idoc = IE.document
    Call idoc.parentWindow.execScript("document.formOngletShip.action ='ShipInspection?fs=ShipInfo';document.formOngletShip.submit();", "JavaScript")

    Application.Wait (Now + TimeValue("00:00:02"))
    Err.Clear
    On Error Resume Next


Access specific page based on the P_INSP.value we want to extract

    Set idoc = IE.document
    string1 = ("document.formShipInspection.P_INSP.value='")
    string2 = string1 & P_INSP_Start
    string3 = string2 & ("';document.formShipInspection.action='DetailsPSC?fs=ShipInspection';document.formShipInspection.submit();")
    Call idoc.parentWindow.execScript(string3, "JavaScript")
    Application.Wait (Now + TimeValue("00:00:01"))
    With IE
      While IE.ReadyState <> 4
      DoEvents
      Wend
    End With

    Application.Wait (Now + TimeValue("00:00:02"))
    Err.Clear
    On Error Resume Next

Extract Basic vessel's info by ClassName

    Set idoc = IE.document
    Application.Wait (Now + TimeValue("00:00:01"))
    Sheets("Sheet1").Range("A1").Value = idoc.getElementsByClassName("color-gris-bleu-copyright")(0).innerText 'VesselName and IMO
    Sheets("Sheet1").Range("A2").Value = idoc.getElementsByClassName("col-lg-4 col-md-4 col-sm-6 col-xs-6")(2).innerText 'Flag
    Sheets("Sheet1").Range("A3").Value = idoc.getElementsByClassName("col-lg-4 col-md-4 col-sm-6 col-xs-6")(8).innerText 'GRT
    Sheets("Sheet1").Range("A4").Value = idoc.getElementsByClassName("col-lg-4 col-md-4 col-sm-6 col-xs-6")(11).innerText 'DWT
    Sheets("Sheet1").Range("A5").Value = idoc.getElementsByClassName("col-lg-4 col-md-4 col-sm-6 col-xs-6")(13).innerText 'TypeOfShip
    Sheets("Sheet1").Range("A6").Value = idoc.getElementsByClassName("col-lg-4 col-md-4 col-sm-6 col-xs-6")(16).innerText 'YearOfBuild
    Sheets("Sheet1").Range("A7").Value = idoc.getElementsByClassName("col-lg-6 col-md-6 col-sm-6 col-xs-6")(3).innerText 'PSC Organization
    Sheets("Sheet1").Range("A8").Value = idoc.getElementsByClassName("col-lg-6 col-md-6 col-sm-6 col-xs-6")(5).innerText 'Authority
    Sheets("Sheet1").Range("A9").Value = idoc.getElementsByClassName("col-lg-6 col-md-6 col-sm-6 col-xs-6")(7).innerText 'Port
    Sheets("Sheet1").Range("A10").Value = idoc.getElementsByClassName("col-lg-6 col-md-6 col-sm-6 col-xs-6")(9).innerText 'TypeOfInsp
    Sheets("Sheet1").Range("A11").Value = idoc.getElementsByClassName("col-lg-6 col-md-6 col-sm-6 col-xs-6")(11).innerText 'Date
    Sheets("Sheet1").Range("A12").Value = idoc.getElementsByClassName("col-lg-6 col-md-6 col-sm-6 col-xs-6")(13).innerText 'Detention
    Sheets("Sheet1").Range("A13").Value = idoc.getElementsByClassName("col-lg-6 col-md-6 col-sm-6 col-xs-6")(15).innerText 'NumberOfDeficiencies
        
    If Sheets("Sheet1").Range("A1").Value = "" Then
      Sheets("Non Found IDs").Select
      Rows("2:2").Select
      Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
      Application.Wait (Now + TimeValue("00:00:01"))
      Sheets("Non Found IDs").Range("A2").Value = P_INSP_Start
      GoTo NoInsp
    End If

Extract info coming from a Table using TagName

    Set ws = ThisWorkbook.Worksheets("Sheet1")
    Application.Wait (Now + TimeValue("00:00:01"))
    Err.Clear
    On Error Resume Next

    Set hTable = idoc.getElementsByClassName("tableLS table table-striped table-responsive")(0)
    
    r = 0
    For Each tr In hTable.getElementsByTagName("tr")
      r = r + 1: c = 5
      For Each th In tr.getElementsByTagName("th")
        ws.Cells(r, c) = th.innerText
        c = c + 1
      Next
      For Each td In tr.getElementsByTagName("td")
        ws.Cells(r, c) = td.innerText
        c = c + 1
      Next
    Next
    Application.Wait (Now + TimeValue("00:00:01"))

    Sheets("User Sheet").Range("K4").Value = "Error1" 'Error handling
    
    Err.Clear
    On Error Resume Next

    Set hTable = idoc.getElementsByClassName("tableLS table table-striped table-responsive")(1)
    
    r = 0
    For Each tr In hTable.getElementsByTagName("tr")
      r = r + 1: c = 10
      For Each th In tr.getElementsByTagName("th")
        ws.Cells(r, c) = th.innerText
        c = c + 1
      Next
      For Each td In tr.getElementsByTagName("td")
         ws.Cells(r, c) = td.innerText
         c = c + 1
       Next
     Next
     Application.Wait (Now + TimeValue("00:00:01"))

    Sheets("User Sheet").Range("K4").Value = "Error2" 'Error handling

    If Sheets("Sheet1").Range("A13").Value = 0 Then
        GoTo NoDef
    End If

    Err.Clear
    On Error Resume Next
    Set hTable = idoc.getElementsByClassName("tableLS table table-striped table-responsive")(2)
        
    r = 0
    For Each tr In hTable.getElementsByTagName("tr")
      r = r + 1: c = 15
      For Each th In tr.getElementsByTagName("th")
        ws.Cells(r, c) = th.innerText
        c = c + 1
      Next
      For Each td In tr.getElementsByTagName("td")
         ws.Cells(r, c) = td.innerText
         c = c + 1
      Next
    Next
    Application.Wait (Now + TimeValue("00:00:01"))

    Sheets("User Sheet").Range("K4").Value = "Error3" 'Error handling
    
    Err.Clear
    On Error Resume Next
    If Sheets("Sheet1").Range("A12").Value Like "*Yes*" Then

      Set hTable = idoc.getElementsByClassName("tableLS table table-striped table-responsive")(3)
    
      r = 0
      For Each tr In hTable.getElementsByTagName("tr")
        r = r + 1: c = 20
        For Each th In tr.getElementsByTagName("th")
          ws.Cells(r, c) = th.innerText
          c = c + 1
        Next
        For Each td In tr.getElementsByTagName("td")
          ws.Cells(r, c) = td.innerText
          c = c + 1
        Next
      Next
    
    End If
    Application.Wait (Now + TimeValue("00:00:01"))

    Sheets("User Sheet").Range("K4").Value = "Error8" 'Error handling
    NoDef:

    Err.Clear
    On Error Resume Next
    Set idoc = IE.document
    Application.Wait (Now + TimeValue("00:00:01"))
    Call idoc.parentWindow.execScript("document.formOngletShip.action ='ShipHistory?fs=ShipInspection';document.formOngletShip.submit();", "JavaScript")
    Application.Wait (Now + TimeValue("00:00:02"))
    With IE
      While IE.ReadyState <> 4
        DoEvents
      Wend
    End With
    Application.Wait (Now + TimeValue("00:00:02"))
    Set idoc = IE.document
    
    Err.Clear
    On Error Resume Next
    Set hTable = idoc.getElementsByClassName("tableLS table table-striped table-responsive")(0)
    
    r = 0
    For Each tr In hTable.getElementsByTagName("tr")
      r = r + 1: c = 26
      For Each th In tr.getElementsByTagName("th")
        ws.Cells(r, c) = th.innerText
        c = c + 1
      Next
      For Each td In tr.getElementsByTagName("td")
        ws.Cells(r, c) = td.innerText
        c = c + 1
      Next
    Next
    Application.Wait (Now + TimeValue("00:00:01"))

    Sheets("User Sheet").Range("K4").Value = "Error4" 'Error handling

    'On Error GoTo Er5
     Err.Clear
     On Error Resume Next
     Set hTable = idoc.getElementsByClassName("tableLS table table-striped table-responsive")(1)
     Application.Wait (Now + TimeValue("00:00:01"))
    
     r = 0
    For Each tr In hTable.getElementsByTagName("tr")
      r = r + 1: c = 31
      For Each th In tr.getElementsByTagName("th")
        ws.Cells(r, c) = th.innerText
        c = c + 1
      Next
      For Each td In tr.getElementsByTagName("td")
         ws.Cells(r, c) = td.innerText
         c = c + 1
      Next
    Next
    Application.Wait (Now + TimeValue("00:00:01"))

    Sheets("User Sheet").Range("K4").Value = "Error5" 'Error handling
    
    'On Error GoTo Er6
    Err.Clear
    On Error Resume Next
    Set hTable = idoc.getElementsByClassName("tableLS table table-striped table-responsive")(2)
    Application.Wait (Now + TimeValue("00:00:01"))
    
    r = 0
    For Each tr In hTable.getElementsByTagName("tr")
      r = r + 1: c = 36
      For Each th In tr.getElementsByTagName("th")
        ws.Cells(r, c) = th.innerText
        c = c + 1
      Next
      For Each td In tr.getElementsByTagName("td")
        ws.Cells(r, c) = td.innerText
        c = c + 1
      Next
    Next
    Application.Wait (Now + TimeValue("00:00:01"))

    Sheets("User Sheet").Range("K4").Value = "Error6" 'Error handling

    'On Error GoTo Er7
    Err.Clear
    On Error Resume Next
    Set hTable = idoc.getElementsByClassName("tableLS table table-striped table-responsive")(3)
    Application.Wait (Now + TimeValue("00:00:01"))
    
    r = 0
    For Each tr In hTable.getElementsByTagName("tr")
      r = r + 1: c = 41
      For Each th In tr.getElementsByTagName("th")
        ws.Cells(r, c) = th.innerText
        c = c + 1
      Next
      For Each td In tr.getElementsByTagName("td")
        ws.Cells(r, c) = td.innerText
        c = c + 1
      Next
    Next
    Application.Wait (Now + TimeValue("00:00:01"))

    Sheets("User Sheet").Range("K4").Value = "Error7" 'Error handling

Data entry from Sheet1 to Database on Excel Sheets

    'Find DOC Issuer & SMC Issuer
     Application.Wait (Now + TimeValue("00:00:01"))
     i = 2
     Sheets("Database").Range("O2").Value = ""
     Sheets("Database").Range("P2").Value = ""
     Sheets("Database").Range("Q2").Value = ""
     Sheets("Database").Range("R2").Value = ""
     Do
      If Sheets("Sheet1").Range("E" & i).Value Like "*DoC*" Then
        Sheets("Database").Range("O2").Value = Sheets("Sheet1").Range("F" & i).Value
        Sheets("Database").Range("P2").Value = Sheets("Sheet1").Range("G" & i).Value
        Exit Do
      End If
      If Sheets("Sheet1").Range("E" & i).Value = "" Then
        Exit Do
      End If
      i = i + 1
    Loop
    i = 2
    Do
      If Sheets("Sheet1").Range("E" & i).Value Like "*SMC*" Then
        Sheets("Database").Range("Q2").Value = Sheets("Sheet1").Range("F" & i).Value
        Sheets("Database").Range("R2").Value = Sheets("Sheet1").Range("G" & i).Value
        Exit Do
      End If
      If Sheets("Sheet1").Range("E" & i).Value = "" Then
        Exit Do
      End If
      i = i + 1
    Loop
        
    'Find ISM Manager
    Application.Wait (Now + TimeValue("00:00:01"))
    Sheets("Sheet1").Select
    Call Extract_Date
    Call Extract_PSC_Date
    Call Check_Dates
    i = 2
    Sheets("Database").Range("U2").Value = ""
    Do
      If Sheets("Sheet1").Range("AP" & i).Value Like "*ISM*" And Sheets("Sheet1").Range("AT" & i).Value = True Then
        Sheets("Database").Range("U2").Value = Sheets("Sheet1").Range("AO" & i).Value
        Exit Do
      End If
      If Sheets("Sheet1").Range("AP" & i).Value = "" Then
        Exit Do
      End If
      i = i + 1
    Loop

    'Deficiencies Entry
    i = 2
    j = 0
    Sheets("Sheet1").Select
    
    If Sheets("Sheet1").Range("A7").Value = "US Coast Guard " Then
      Do
        If Sheets("Sheet1").Range("J" & i).Value <> "" Then
          j = j + 1
          i = i + 1
        End If
        If Sheets("Sheet1").Range("J" & i).Value = "" Then
          Exit Do
        End If
      Loop
            
      i = 2
      k = 0
      Do
        If Sheets("Sheet1").Range("O" & i).Value <> "" Then
          k = k + 1
          i = i + 1
        End If
        If Sheets("Sheet1").Range("O" & i).Value = "" Then
          Exit Do
        End If
      Loop
    End If
        
    If Sheets("Sheet1").Range("A7").Value = "Vina Del Mar MoU " Then
      Do
        If Sheets("Sheet1").Range("J" & i).Value <> "" Then
          j = j + 1
          i = i + 1
        End If
        If Sheets("Sheet1").Range("J" & i).Value = "" Then
          Exit Do
        End If
      Loop
            
      i = 2
      k = 0
      Do
        If Sheets("Sheet1").Range("O" & i).Value <> "" Then
          k = k + 1
          i = i + 1
        End If
        If Sheets("Sheet1").Range("O" & i).Value = "" Then
          Exit Do
        End If
      Loop
    End If
        
    If Sheets("Sheet1").Range("A7").Value = "Black Sea MoU " Then
      Do
        If Sheets("Sheet1").Range("E" & i).Value <> "" Then
          j = j + 1
          i = i + 1
        End If
        If Sheets("Sheet1").Range("E" & i).Value = "" Then
          Exit Do
        End If
      Loop
            
      i = 2
      k = 0
      Do
        If Sheets("Sheet1").Range("J" & i).Value <> "" Then
          k = k + 1
          i = i + 1
        End If
        If Sheets("Sheet1").Range("J" & i).Value = "" Then
          Exit Do
        End If
      Loop
    End If
        
    If Sheets("Sheet1").Range("A7").Value = "Abuja MoU " Then
      Do
        If Sheets("Sheet1").Range("E" & i).Value <> "" Then
          j = j + 1
          i = i + 1
        End If
        If Sheets("Sheet1").Range("E" & i).Value = "" Then
          Exit Do
        End If
      Loop
            
      i = 2
      k = 0
      Do
        If Sheets("Sheet1").Range("J" & i).Value <> "" Then
          k = k + 1
          i = i + 1
        End If
        If Sheets("Sheet1").Range("J" & i).Value = "" Then
          Exit Do
        End If
      Loop
    End If
        
    If Sheets("Sheet1").Range("A7").Value = "Tokyo MoU " Then
      Do
        If Sheets("Sheet1").Range("E" & i).Value <> "" Then
          j = j + 1
          i = i + 1
        End If
        If Sheets("Sheet1").Range("E" & i).Value = "" Then
          Exit Do
        End If
      Loop
            
      i = 2
      k = 0
      Do
        If Sheets("Sheet1").Range("J" & i).Value <> "" Then
          k = k + 1
          i = i + 1
        End If
        If Sheets("Sheet1").Range("J" & i).Value = "" Then
          Exit Do
        End If
      Loop
    End If
        
    If Sheets("Sheet1").Range("A7").Value = "Indian Ocean MoU " Then
      Do
        If Sheets("Sheet1").Range("J" & i).Value <> "" Then
          j = j + 1
          i = i + 1
        End If
        If Sheets("Sheet1").Range("J" & i).Value = "" Then
          Exit Do
        End If
      Loop
            
      i = 2
      k = 0
      Do
        If Sheets("Sheet1").Range("O" & i).Value <> "" Then
          k = k + 1
          i = i + 1
        End If
        If Sheets("Sheet1").Range("O" & i).Value = "" Then
          Exit Do
        End If
      Loop
    End If
        
    If Sheets("Sheet1").Range("A7").Value = "Paris MoU " Then
      Do
        If Sheets("Sheet1").Range("O" & i).Value <> "" Then
          j = j + 1
          i = i + 1
        End If
        If Sheets("Sheet1").Range("O" & i).Value = "" Then
          Exit Do
        End If
      Loop
    
      i = 2
      k = 0
      Do
        If Sheets("Sheet1").Range("T" & i).Value <> "" Then
          k = k + 1
          i = i + 1
        End If
        If Sheets("Sheet1").Range("T" & i).Value = "" Then
          Exit Do
        End If
      Loop
    End If
        
    If Sheets("Sheet1").Range("A7").Value = "Caribbean MoU " Then
      Do
        If Sheets("Sheet1").Range("O" & i).Value <> "" Then
          j = j + 1
          i = i + 1
        End If
        If Sheets("Sheet1").Range("O" & i).Value = "" Then
          Exit Do
        End If
      Loop
    
      i = 2
      k = 0
      Do
        If Sheets("Sheet1").Range("T" & i).Value <> "" Then
          k = k + 1
          i = i + 1
        End If
        If Sheets("Sheet1").Range("T" & i).Value = "" Then
          Exit Do
        End If
      Loop
    End If
        
    Sheets("Database").Select
    If j > 0 Then
      i = 1
        For i = 1 To j
          Sheets("Database").Select
          Rows("5:5").Select
          Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
          Application.Wait (Now + TimeValue("00:00:01"))
            
          Sheets("Database").Select
          Sheets("Database").Range("A2:AB2").Select
          Selection.Copy
          Sheets("Database").Range("A5:AB5").Select
          Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
              :=False, Transpose:=False
            
          Sheets("Database").Range("V5").Value = Sheets("Database").Range("A5").Value & "_" & i
                
          If Sheets("Sheet1").Range("A7").Value = "US Coast Guard " Then
            Sheets("Database").Range("W5").Value = Sheets("Sheet1").Range("J" & i + 1).Value
            Sheets("Database").Range("X5").Value = Sheets("Sheet1").Range("K" & i + 1).Value
            Sheets("Database").Range("Y5").Value = Sheets("Sheet1").Range("L" & i + 1).Value
             Sheets("Database").Range("Z5").Value = "No"
             Sheets("Database").Range("AA5").Value = "No"
          End If
                
          If Sheets("Sheet1").Range("A7").Value = "Vina Del Mar MoU " Then
            Sheets("Database").Range("W5").Value = Sheets("Sheet1").Range("J" & i + 1).Value
            Sheets("Database").Range("X5").Value = Sheets("Sheet1").Range("K" & i + 1).Value
            Sheets("Database").Range("Y5").Value = Sheets("Sheet1").Range("L" & i + 1).Value
            Sheets("Database").Range("Z5").Value = "No"
             Sheets("Database").Range("AA5").Value = "No"
          End If
                
          If Sheets("Sheet1").Range("A7").Value = "Black Sea MoU " Then
            Sheets("Database").Range("W5").Value = Sheets("Sheet1").Range("E" & i + 1).Value
            Sheets("Database").Range("X5").Value = Sheets("Sheet1").Range("F" & i + 1).Value
            Sheets("Database").Range("Y5").Value = Sheets("Sheet1").Range("G" & i + 1).Value
            Sheets("Database").Range("Z5").Value = "No"
            Sheets("Database").Range("AA5").Value = "No"
          End If
                
          If Sheets("Sheet1").Range("A7").Value = "Abuja MoU " Then
            Sheets("Database").Range("W5").Value = Sheets("Sheet1").Range("E" & i + 1).Value
            Sheets("Database").Range("X5").Value = Sheets("Sheet1").Range("F" & i + 1).Value
            Sheets("Database").Range("Y5").Value = Sheets("Sheet1").Range("G" & i + 1).Value
            Sheets("Database").Range("Z5").Value = "No"
            Sheets("Database").Range("AA5").Value = "No"
          End If
                
          If Sheets("Sheet1").Range("A7").Value = "Paris MoU " Then
            Sheets("Database").Range("W5").Value = Sheets("Sheet1").Range("O" & i + 1).Value
            Sheets("Database").Range("X5").Value = Sheets("Sheet1").Range("P" & i + 1).Value
            Sheets("Database").Range("Y5").Value = Sheets("Sheet1").Range("Q" & i + 1).Value
            Sheets("Database").Range("Z5").Value = "No"
            Sheets("Database").Range("AA5").Value = "No"
          End If
                
          If Sheets("Sheet1").Range("A7").Value = "Caribbean MoU " Then
            Sheets("Database").Range("W5").Value = Sheets("Sheet1").Range("O" & i + 1).Value
            Sheets("Database").Range("X5").Value = Sheets("Sheet1").Range("P" & i + 1).Value
            Sheets("Database").Range("Y5").Value = Sheets("Sheet1").Range("Q" & i + 1).Value
            Sheets("Database").Range("Z5").Value = "No"
            Sheets("Database").Range("AA5").Value = "No"
          End If
                
          If Sheets("Sheet1").Range("A7").Value = "Tokyo MoU " Then
            Sheets("Database").Range("W5").Value = Sheets("Sheet1").Range("E" & i + 1).Value
            Sheets("Database").Range("X5").Value = Sheets("Sheet1").Range("F" & i + 1).Value
            Sheets("Database").Range("Y5").Value = Sheets("Sheet1").Range("G" & i + 1).Value
            Sheets("Database").Range("Z5").Value = "No"
            Sheets("Database").Range("AA5").Value = "No"
          End If
                
          If Sheets("Sheet1").Range("A7").Value = "Indian Ocean MoU " Then
            Sheets("Database").Range("W5").Value = Sheets("Sheet1").Range("J" & i + 1).Value
            Sheets("Database").Range("X5").Value = Sheets("Sheet1").Range("K" & i + 1).Value
            Sheets("Database").Range("Y5").Value = Sheets("Sheet1").Range("L" & i + 1).Value
            Sheets("Database").Range("Z5").Value = "No"
            Sheets("Database").Range("AA5").Value = "No"
          End If
      Next i
    End If
    If k > 0 Then
      i = 1
        For i = 1 To k
          Sheets("Database").Select
          Rows("5:5").Select
          Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
          Application.Wait (Now + TimeValue("00:00:01"))
            
          Sheets("Database").Select
          Sheets("Database").Range("A2:AB2").Select
          Selection.Copy
          Sheets("Database").Range("A5:AB5").Select
          Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
              :=False, Transpose:=False
            
          Sheets("Database").Range("V5").Value = Sheets("Database").Range("A5").Value & "_" & (i + j)
                
          If Sheets("Sheet1").Range("A7").Value = "US Coast Guard " Then
              Sheets("Database").Range("W5").Value = Sheets("Sheet1").Range("O" & i + 1).Value
              Sheets("Database").Range("X5").Value = Sheets("Sheet1").Range("P" & i + 1).Value
              Sheets("Database").Range("Y5").Value = Sheets("Sheet1").Range("Q" & i + 1).Value
              Sheets("Database").Range("Z5").Value = "Yes"
              Sheets("Database").Range("AA5").Value = Sheets("Sheet1").Range("S" & i + 1).Value
          End If
                
          If Sheets("Sheet1").Range("A7").Value = "Vina Del Mar MoU " Then
            Sheets("Database").Range("W5").Value = Sheets("Sheet1").Range("O" & i + 1).Value
            Sheets("Database").Range("X5").Value = Sheets("Sheet1").Range("P" & i + 1).Value
            Sheets("Database").Range("Y5").Value = Sheets("Sheet1").Range("Q" & i + 1).Value
            Sheets("Database").Range("Z5").Value = "Yes"
            Sheets("Database").Range("AA5").Value = Sheets("Sheet1").Range("S" & i + 1).Value
          End If
                
          If Sheets("Sheet1").Range("A7").Value = "Black Sea MoU " Then
            Sheets("Database").Range("W5").Value = Sheets("Sheet1").Range("J" & i + 1).Value
            Sheets("Database").Range("X5").Value = Sheets("Sheet1").Range("K" & i + 1).Value
            Sheets("Database").Range("Y5").Value = Sheets("Sheet1").Range("L" & i + 1).Value
            Sheets("Database").Range("Z5").Value = "Yes"
            Sheets("Database").Range("AA5").Value = Sheets("Sheet1").Range("N" & i + 1).Value
          End If
                
          If Sheets("Sheet1").Range("A7").Value = "Abuja MoU " Then
            Sheets("Database").Range("W5").Value = Sheets("Sheet1").Range("J" & i + 1).Value
            Sheets("Database").Range("X5").Value = Sheets("Sheet1").Range("K" & i + 1).Value
            Sheets("Database").Range("Y5").Value = Sheets("Sheet1").Range("L" & i + 1).Value
            Sheets("Database").Range("Z5").Value = "Yes"
            Sheets("Database").Range("AA5").Value = Sheets("Sheet1").Range("N" & i + 1).Value
          End If
                
          If Sheets("Sheet1").Range("A7").Value = "Paris MoU " Then
            Sheets("Database").Range("W5").Value = Sheets("Sheet1").Range("T" & i + 1).Value
            Sheets("Database").Range("X5").Value = Sheets("Sheet1").Range("U" & i + 1).Value
            Sheets("Database").Range("Y5").Value = Sheets("Sheet1").Range("V" & i + 1).Value
            Sheets("Database").Range("Z5").Value = "Yes"
            Sheets("Database").Range("AA5").Value = Sheets("Sheet1").Range("X" & i + 1).Value
          End If
                
          If Sheets("Sheet1").Range("A7").Value = "Caribbean MoU " Then
            Sheets("Database").Range("W5").Value = Sheets("Sheet1").Range("T" & i + 1).Value
            Sheets("Database").Range("X5").Value = Sheets("Sheet1").Range("U" & i + 1).Value
            Sheets("Database").Range("Y5").Value = Sheets("Sheet1").Range("V" & i + 1).Value
            Sheets("Database").Range("Z5").Value = "Yes"
            Sheets("Database").Range("AA5").Value = Sheets("Sheet1").Range("X" & i + 1).Value
          End If
                
          If Sheets("Sheet1").Range("A7").Value = "Tokyo MoU " Then
            Sheets("Database").Range("W5").Value = Sheets("Sheet1").Range("J" & i + 1).Value
            Sheets("Database").Range("X5").Value = Sheets("Sheet1").Range("K" & i + 1).Value
            Sheets("Database").Range("Y5").Value = Sheets("Sheet1").Range("L" & i + 1).Value
            Sheets("Database").Range("Z5").Value = "Yes"
            Sheets("Database").Range("AA5").Value = Sheets("Sheet1").Range("N" & i + 1).Value
          End If
                
          If Sheets("Sheet1").Range("A7").Value = "Indian Ocean MoU " Then
            Sheets("Database").Range("W5").Value = Sheets("Sheet1").Range("O" & i + 1).Value
            Sheets("Database").Range("X5").Value = Sheets("Sheet1").Range("P" & i + 1).Value
            Sheets("Database").Range("Y5").Value = Sheets("Sheet1").Range("Q" & i + 1).Value
            Sheets("Database").Range("Z5").Value = "Yes"
            Sheets("Database").Range("AA5").Value = Sheets("Sheet1").Range("S" & i + 1).Value
          End If
       Next i
    End If
    If j = 0 Then
      Sheets("Database").Select
      Rows("5:5").Select
      Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
      Application.Wait (Now + TimeValue("00:00:01"))
            
      Sheets("Database").Select
      Sheets("Database").Range("A2:AB2").Select
      Selection.Copy
      Sheets("Database").Range("A5:AB5").Select
      Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
          :=False, Transpose:=False
    End If
        
    Sheets("User Sheet").Range("K4").Value = "Error0" 'Error handling


Navgate back to the vessel

    Application.Wait (Now + TimeValue("00:00:02"))
    Set idoc = IE.document
    Set IEobj = idoc.getElementById("P_ENTREE_ENTETE")
    IEobj.Value = "9713337"
    Application.Wait (Now + TimeValue("00:00:01"))
    Set Buttons = idoc.getElementsByClassName("btn btn-default")
    Application.Wait (Now + TimeValue("00:00:01"))
    For Each Button In Buttons
      If Button.className = "btn btn-default" Then
        Button.Click
        Exit For
      End If
      Next Button
      With IE
        While IE.ReadyState <> 4
          DoEvents
        Wend
      End With

Clear Sheet1 to start again the loop

    Sheets("Sheet1").Select
    Sheets("Sheet1").Range("A1:AZ154").Select
    Selection.ClearContents
    Sheets("User Sheet").Range("B3").Value = P_INSP_Start
    P_INSP_Start = P_INSP_Start + 1
    
    counter = counter + 1 'count how many inspections I searched with the same credentials
    
    If counter > 9 Then 'Handle the server not to overload from my extraction
        counter = 0
        Application.Wait (Now + TimeValue("00:00:01"))
        'Log Out Equasis and Exit IE
        With IE
           .navigate ("Web page new connection")
            While IE.ReadyState <> 4
                DoEvents
            Wend
        End With
        Application.Wait (Now + TimeValue("00:00:02"))
        IE.Quit
        Application.Wait (Now + TimeValue("00:00:01"))
        Set IE = Nothing
        loginCounter = loginCounter + 1
        GoTo Start
    End If
    Sheets("User Sheet").Range("K4").Value = ""
    
    If counter > 15 Then 'Extreme case for error handling when there is no data for the PSC inspection id I placed
      NoInsp:
      CounterNoInsp = CounterNoInsp + 1
      counter = counter + 1
      If counter > 9 Then
        counter = 0
        loginCounter = loginCounter + 1
        'Log Out web site and Exit IE
        With IE
          .navigate ("Web page new connection")
          While IE.ReadyState <> 4
            DoEvents
          Wend
        End With
        Application.Wait (Now + TimeValue("00:00:02"))
        IE.Quit
        Application.Wait (Now + TimeValue("00:00:01"))
        Set IE = Nothing
      End If
      Sheets("Sheet1").Select
      Sheets("Sheet1").Range("A1:AZ154").Select
      Selection.ClearContents
      Sheets("User Sheet").Range("B3").Value = P_INSP_Start
      P_INSP_Start = P_INSP_Start + 1
        
      If counter > 0 Then
        refresher = 1
        With IE
          .navigate ("Web page Home Page")
          While IE.ReadyState <> 4
            DoEvents
          Wend
        End With
        Application.Wait (Now + TimeValue("00:00:01"))
      End If
        
      CountNoInsp = CountNoInsp + 1 'counter on how many times in a row the inspection not found
      cc = cc + 1
      DoEvents
      If CountNoInsp > 1000 Then 'if 1000 times in a row we missed an inspection it could be an alert that one of our emails is banned
        Unload StatusBar 'Close the Status Bar
        Sheets("User Sheet").Range("K4").Value = "" 'Delete Error Handling entry
        Sheets("User Sheet").Range("M1").Value = "" 'Delete Error Handling entry
        MsgBox ("Check what is going on, we had multiple fails with this email: " & Username)
        Exit Sub
      End If
        Call RunStatusBar(cc, count)
        GoTo Start
    End If
    CountNoInsp = 0
    
    cc = cc + 1
    DoEvents
    Call RunStatusBar(cc, count)

When the 'Exit'button is pressed by the end-user on Status Bar

    If Sheets("User Sheet").Range("M1").Value = "Exit" Then
      Sheets("User Sheet").Range("M1").Value = ""
      Sheets("Database").Select
      Sheets("Sheet1").Visible = False
      Unload StatusBar
      'Log Out and Exit IE
      With IE
        .navigate ("Web page new connection")
        While IE.ReadyState <> 4
          DoEvents
        Wend
      End With
      Application.Wait (Now + TimeValue("00:00:01"))
      IE.Quit
      Application.Wait (Now + TimeValue("00:00:01"))
      Set IE = Nothing
      MsgBox "You stopped the procedure"
      Exit Sub
    End If

End Loop

    Loop
    Sheets("Database").Select
    Sheets("Sheet1").Visible = False 'Hide the Back-end page
    Sheets("Emails").Visible = False
    Sheets("Non Found IDs").Visible = False

Disconnect from webpage, close IE, and finish the procedure

    With IE
      .navigate ("Web page new connection")
      While IE.ReadyState <> 4
        DoEvents
      Wend
    End With
    Application.Wait (Now + TimeValue("00:00:02"))
    IE.Quit
    Application.Wait (Now + TimeValue("00:00:01"))
    Set IE = Nothing

    Unload StatusBar 'Close the Status Bar
    Sheets("User Sheet").Range("K4").Value = "" 'Delete Error Handling entry
    Sheets("User Sheet").Range("M1").Value = "" 'Delete Error Handling entry
    MsgBox "You are ready! We extracted " & count - CounterNoInsp & " inspections out of total " & count & ". Check the extracted PSC inspections.."

    End Sub

OpenStatusBar procedure

    Sub OpenStatusBar()

    With StatusBar
      .Bar.Width = 0
      .Frame.Caption = "0% Complete"
      .Show vbModeless
    End With

    End Sub

RunStatusBar procedure

    Sub RunStatusBar(row As Integer, total As Integer)

    With StatusBar
      .Bar.Width = 312 * (row / total)
      .Frame.Caption = Round((row / total) * 100, 0) & "% Complete"
    End With

    End Sub

Procedure for the NonFoundIDs

    'Dim variables
    Dim k As Integer
    Dim P_INSP_Start As String
    Dim P_INSP_End As String
    Dim counter As Long
    Dim loginCounter As Integer
    Dim count As Integer, cc As Integer
    Dim Username, Password As String
    Dim idoc As MSHTML.HTMLDocument
    Dim IEobj As Object
    Dim Button As MSHTML.IHTMLElement
    Dim Buttons As MSHTML.IHTMLElementCollection
    Dim PressVessel As String, PV1 As String, PV2 As String, PV3 As String
    Dim string1 As String, string2 As String, string3 As String
    Dim ws As Worksheet
    Dim hTable As Object
    Dim td As Object, tr As Object, th As Object, r As Long, c As Long
    Dim i As Integer
    Dim j As Integer
    Dim IMO As String
    Dim Result As Integer
    Dim Result0 As Integer
    Dim CountNoInsp As Integer
    Dim refresher As Integer
    Dim ii As Integer
    Dim CounterNoInsp As Integer
    
    Result = 999

    Result0 = MsgBox("The procedure will start now; Please do not use any Excel file in the meantime. The total execution time estimated at: " & Sheets("User Sheet").Range("K1").Value & ". Press YES to continue..", vbQuestion + vbYesNo)
    If Result0 = vbNo Then
    Sheets("User Sheet").Select
    Exit Sub
    End If

    Result0 = MsgBox("Do you want to see the Excel updating?", vbQuestion + vbYesNo)
    If Result0 = vbNo Then
    Application.ScreenUpdating = False 'Not to update the Excel while the function is running..
    Else
    Application.ScreenUpdating = True
    End If

    Sheets("User Sheet").Select
    Sheets("Sheet1").Visible = True
    Sheets("Emails").Visible = True
    Sheets("Non Found IDs").Visible = True

    'Randomize the Emails
    ActiveWorkbook.Worksheets("Emails").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Emails").AutoFilter.Sort.SortFields.Add2 Key:=Range("C1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("Emails").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    'Find how many PSC inspections we will extract
    counter = 0
    loginCounter = 0
    count = Sheets("User Sheet").Range("J1").Value 'Count how many loops we will have
    cc = 0
    ii = 0

    If Result = 999 Then
    Result = MsgBox("Do you want to see the Internet Explorer browser while web-scrapping?", vbQuestion + vbYesNo)
    End If

    OpenStatusBar 'Open the Status Bar for loading..

    Start:

    counter = counter + 1
    
    If refresher = 0 Then
    'Create Internet Explorer with Question handling
            If Result = vbYes Then
                Set IE = New InternetExplorer
                IE.Visible = True 'Present Internet Explorer
            Else
                Set IE = New InternetExplorer
                IE.Visible = False 'Hide Internet Explorer
            End If
    End If
            
    refresher = 0

    'Navigate to Equasis
    With IE
        .navigate ("https://www.equasis.org/EquasisWeb/public/HomePage")
        While IE.ReadyState <> 4
            DoEvents
        Wend
    End With
    Application.Wait (Now + TimeValue("00:00:01"))
        
    
    'Equasis credentials
    Sheets("Emails").Select
    Username = Sheets("Emails").Range("A" & loginCounter + 2).Value
    Password = Sheets("Emails").Range("B" & loginCounter + 2).Value
    Sheets("User Sheet").Select

    'LogIn to Equasis
    Application.Wait (Now + TimeValue("00:00:01"))
    Err.Clear
    On Error Resume Next
    Set IEobj = IE.document.getElementById("home-login")
    IEobj.Value = Username
    Application.Wait (Now + TimeValue("00:00:01"))
    Set IEobj = IE.document.getElementById("home-password")
    IEobj.Value = Password
    Application.Wait (Now + TimeValue("00:00:01"))
    Set idoc = IE.document
    Set Buttons = idoc.getElementsByClassName("pull-right btn btn-lg gris-bleu-copyright")
    For Each Button In Buttons
        If Button.className = "pull-right btn btn-lg gris-bleu-copyright" Then
            Button.Click
            Exit For
        End If
    Next Button
    Application.Wait (Now + TimeValue("00:00:01"))
    With IE
        While IE.ReadyState <> 4
            DoEvents
        Wend
    End With
    Application.Wait (Now + TimeValue("00:00:01"))

    Sheets("User Sheet").Range("K4").Value = "Error99" 'Error handling

    'Search Vessel
    Application.Wait (Now + TimeValue("00:00:01"))
    Set idoc = IE.document
    Set IEobj = idoc.getElementById("P_ENTREE_HOME")
    Application.Wait (Now + TimeValue("00:00:01"))
    IMO = Sheets("Database").Range("C10").Value
    IEobj.Value = IMO
    Application.Wait (Now + TimeValue("00:00:01"))
    Set idoc = IE.document
    Set Buttons = idoc.getElementsByClassName("btn btn-default")
    Application.Wait (Now + TimeValue("00:00:01"))
    For Each Button In Buttons
        If Button.className = "btn btn-default" Then
            Button.Click
            Exit For
        End If
    Next Button
    Application.Wait (Now + TimeValue("00:00:01"))
    With IE
        While IE.ReadyState <> 4
            DoEvents
        Wend
    End With
    Application.Wait (Now + TimeValue("00:00:02"))
 
    'Start the procedure
    Do Until ii = count
    ii = ii + 1
    P_INSP_Start = Sheets("User Sheet").Range("I" & ii + 1).Value
    Sheets("Database").Range("A2").Value = P_INSP_Start
    Application.Wait (Now + TimeValue("00:00:02"))
    
    Err.Clear
    On Error Resume Next

    'Press vessel-enable onclick
        IMO = Sheets("Database").Range("C10").Value
        Set idoc = IE.document
        PV1 = ("document.formShip.P_IMO.value='")
        PV2 = PV1 & IMO
        PV3 = ("';document.formShip.submit();")
        PressVessel = PV2 & PV3
        Call idoc.parentWindow.execScript(PressVessel, "JavaScript")
        Application.Wait (Now + TimeValue("00:00:01"))
        With IE
            While IE.ReadyState <> 4
                DoEvents
            Wend
        End With
        Application.Wait (Now + TimeValue("00:00:02"))
    
    'Press Inspections-enable onclick
        Set idoc = IE.document
        Call idoc.parentWindow.execScript("document.formOngletShip.action ='ShipInspection?fs=ShipInfo';document.formOngletShip.submit();", "JavaScript")

    Application.Wait (Now + TimeValue("00:00:02"))
    Err.Clear
    On Error Resume Next
    
    'Access specific PSC inspection based on "P_INSP.value= <number>"
        
        Set idoc = IE.document
        string1 = ("document.formShipInspection.P_INSP.value='")
        string2 = string1 & P_INSP_Start
        string3 = string2 & ("';document.formShipInspection.action='DetailsPSC?fs=ShipInspection';document.formShipInspection.submit();")
        Call idoc.parentWindow.execScript(string3, "JavaScript")
        Application.Wait (Now + TimeValue("00:00:01"))
        With IE
            While IE.ReadyState <> 4
                DoEvents
            Wend
        End With

    Application.Wait (Now + TimeValue("00:00:02"))
    Err.Clear
    On Error Resume Next

    'Extract PSC info
        Set idoc = IE.document
        Application.Wait (Now + TimeValue("00:00:01"))
        Sheets("Sheet1").Range("A1").Value = idoc.getElementsByClassName("color-gris-bleu-copyright")(0).innerText 'VesselName and IMO
        Sheets("Sheet1").Range("A2").Value = idoc.getElementsByClassName("col-lg-4 col-md-4 col-sm-6 col-xs-6")(2).innerText 'Flag
        Sheets("Sheet1").Range("A3").Value = idoc.getElementsByClassName("col-lg-4 col-md-4 col-sm-6 col-xs-6")(8).innerText 'GRT
        Sheets("Sheet1").Range("A4").Value = idoc.getElementsByClassName("col-lg-4 col-md-4 col-sm-6 col-xs-6")(11).innerText 'DWT
        Sheets("Sheet1").Range("A5").Value = idoc.getElementsByClassName("col-lg-4 col-md-4 col-sm-6 col-xs-6")(13).innerText 'TypeOfShip
        Sheets("Sheet1").Range("A6").Value = idoc.getElementsByClassName("col-lg-4 col-md-4 col-sm-6 col-xs-6")(16).innerText 'YearOfBuild
        Sheets("Sheet1").Range("A7").Value = idoc.getElementsByClassName("col-lg-6 col-md-6 col-sm-6 col-xs-6")(3).innerText 'PSC Organization
        Sheets("Sheet1").Range("A8").Value = idoc.getElementsByClassName("col-lg-6 col-md-6 col-sm-6 col-xs-6")(5).innerText 'Authority
        Sheets("Sheet1").Range("A9").Value = idoc.getElementsByClassName("col-lg-6 col-md-6 col-sm-6 col-xs-6")(7).innerText 'Port
        Sheets("Sheet1").Range("A10").Value = idoc.getElementsByClassName("col-lg-6 col-md-6 col-sm-6 col-xs-6")(9).innerText 'TypeOfInsp
        Sheets("Sheet1").Range("A11").Value = idoc.getElementsByClassName("col-lg-6 col-md-6 col-sm-6 col-xs-6")(11).innerText 'Date
        Sheets("Sheet1").Range("A12").Value = idoc.getElementsByClassName("col-lg-6 col-md-6 col-sm-6 col-xs-6")(13).innerText 'Detention
        Sheets("Sheet1").Range("A13").Value = idoc.getElementsByClassName("col-lg-6 col-md-6 col-sm-6 col-xs-6")(15).innerText 'NumberOfDeficiencies
        
        If Sheets("Sheet1").Range("A1").Value = "" Then
            Sheets("Non Found IDs").Select
            Rows("2:2").Select
            Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            Application.Wait (Now + TimeValue("00:00:01"))
            Sheets("Non Found IDs").Range("A2").Value = P_INSP_Start
            GoTo NoInsp
        End If
        
    'Extract Statutory surveys at the time of the inspection
        Set ws = ThisWorkbook.Worksheets("Sheet1")
        Application.Wait (Now + TimeValue("00:00:01"))
        Err.Clear
        On Error Resume Next

        Set hTable = idoc.getElementsByClassName("tableLS table table-striped table-responsive")(0)
    
        r = 0
        For Each tr In hTable.getElementsByTagName("tr")
            r = r + 1: c = 5
            For Each th In tr.getElementsByTagName("th")
                ws.Cells(r, c) = th.innerText
                c = c + 1
            Next
                For Each td In tr.getElementsByTagName("td")
                ws.Cells(r, c) = td.innerText
                c = c + 1
            Next
        Next
        Application.Wait (Now + TimeValue("00:00:01"))

     Sheets("User Sheet").Range("K4").Value = "Error1" 'Error handling

    'Extract Classification surveys at the time of the inspection
        Err.Clear
        On Error Resume Next

        Set hTable = idoc.getElementsByClassName("tableLS table table-striped table-responsive")(1)
    
        r = 0
        For Each tr In hTable.getElementsByTagName("tr")
            r = r + 1: c = 10
            For Each th In tr.getElementsByTagName("th")
                ws.Cells(r, c) = th.innerText
                c = c + 1
            Next
            For Each td In tr.getElementsByTagName("td")
                ws.Cells(r, c) = td.innerText
                c = c + 1
            Next
        Next
        Application.Wait (Now + TimeValue("00:00:01"))

    Sheets("User Sheet").Range("K4").Value = "Error2" 'Error handling

    If Sheets("Sheet1").Range("A13").Value = 0 Then
        GoTo NoDef
    End If

    'Extract Deficiencies per category
        Err.Clear
        On Error Resume Next
        Set hTable = idoc.getElementsByClassName("tableLS table table-striped table-responsive")(2)
        
        r = 0
        For Each tr In hTable.getElementsByTagName("tr")
            r = r + 1: c = 15
            For Each th In tr.getElementsByTagName("th")
                ws.Cells(r, c) = th.innerText
                c = c + 1
            Next
            For Each td In tr.getElementsByTagName("td")
                ws.Cells(r, c) = td.innerText
                c = c + 1
            Next
        Next
        Application.Wait (Now + TimeValue("00:00:01"))

              Sheets("User Sheet").Range("K4").Value = "Error3" 'Error handling

    'Extract Grounds for detention
        Err.Clear
        On Error Resume Next
        If Sheets("Sheet1").Range("A12").Value Like "*Yes*" Then

            Set hTable = idoc.getElementsByClassName("tableLS table table-striped table-responsive")(3)
    
            r = 0
            For Each tr In hTable.getElementsByTagName("tr")
                r = r + 1: c = 20
                For Each th In tr.getElementsByTagName("th")
                    ws.Cells(r, c) = th.innerText
                    c = c + 1
                Next
                For Each td In tr.getElementsByTagName("td")
                    ws.Cells(r, c) = td.innerText
                    c = c + 1
                Next
            Next
    
        End If
        Application.Wait (Now + TimeValue("00:00:01"))

    Sheets("User Sheet").Range("K4").Value = "Error8" 'Error handling
    NoDef:

    'Extract Company History
        Err.Clear
        On Error Resume Next
        Set idoc = IE.document
        Application.Wait (Now + TimeValue("00:00:01"))
        Call idoc.parentWindow.execScript("document.formOngletShip.action ='ShipHistory?fs=ShipInspection';document.formOngletShip.submit();", "JavaScript")
        Application.Wait (Now + TimeValue("00:00:02"))
        With IE
            While IE.ReadyState <> 4
                DoEvents
            Wend
        End With
        Application.Wait (Now + TimeValue("00:00:02"))
        Set idoc = IE.document
    
        Err.Clear
        On Error Resume Next
        Set hTable = idoc.getElementsByClassName("tableLS table table-striped table-responsive")(0)
    
        r = 0
        For Each tr In hTable.getElementsByTagName("tr")
            r = r + 1: c = 26
            For Each th In tr.getElementsByTagName("th")
                ws.Cells(r, c) = th.innerText
                c = c + 1
            Next
            For Each td In tr.getElementsByTagName("td")
                ws.Cells(r, c) = td.innerText
                c = c + 1
            Next
        Next
        Application.Wait (Now + TimeValue("00:00:01"))

    Sheets("User Sheet").Range("K4").Value = "Error4" 'Error handling

    'On Error GoTo Er5
        Err.Clear
        On Error Resume Next
        Set hTable = idoc.getElementsByClassName("tableLS table table-striped table-responsive")(1)
        Application.Wait (Now + TimeValue("00:00:01"))
    
        r = 0
        For Each tr In hTable.getElementsByTagName("tr")
            r = r + 1: c = 31
            For Each th In tr.getElementsByTagName("th")
                ws.Cells(r, c) = th.innerText
                c = c + 1
            Next
            For Each td In tr.getElementsByTagName("td")
                ws.Cells(r, c) = td.innerText
                c = c + 1
            Next
        Next
        Application.Wait (Now + TimeValue("00:00:01"))

    Sheets("User Sheet").Range("K4").Value = "Error5" 'Error handling
    
    'On Error GoTo Er6
        Err.Clear
        On Error Resume Next
        Set hTable = idoc.getElementsByClassName("tableLS table table-striped table-responsive")(2)
        Application.Wait (Now + TimeValue("00:00:01"))
    
        r = 0
        For Each tr In hTable.getElementsByTagName("tr")
            r = r + 1: c = 36
            For Each th In tr.getElementsByTagName("th")
                ws.Cells(r, c) = th.innerText
                c = c + 1
            Next
            For Each td In tr.getElementsByTagName("td")
                ws.Cells(r, c) = td.innerText
                c = c + 1
            Next
        Next
        Application.Wait (Now + TimeValue("00:00:01"))

    Sheets("User Sheet").Range("K4").Value = "Error6" 'Error handling

    'On Error GoTo Er7
        Err.Clear
        On Error Resume Next
        Set hTable = idoc.getElementsByClassName("tableLS table table-striped table-responsive")(3)
        Application.Wait (Now + TimeValue("00:00:01"))
    
        r = 0
        For Each tr In hTable.getElementsByTagName("tr")
            r = r + 1: c = 41
            For Each th In tr.getElementsByTagName("th")
                ws.Cells(r, c) = th.innerText
                c = c + 1
            Next
            For Each td In tr.getElementsByTagName("td")
                ws.Cells(r, c) = td.innerText
                c = c + 1
            Next
        Next
        Application.Wait (Now + TimeValue("00:00:01"))

    Sheets("User Sheet").Range("K4").Value = "Error7" 'Error handling

    'Data entry from Sheet1 to Database
        'Find DOC Issuer & SMC Issuer
            Application.Wait (Now + TimeValue("00:00:01"))
            i = 2
            Sheets("Database").Range("O2").Value = ""
            Sheets("Database").Range("P2").Value = ""
            Sheets("Database").Range("Q2").Value = ""
            Sheets("Database").Range("R2").Value = ""
            Do
                If Sheets("Sheet1").Range("E" & i).Value Like "*DoC*" Then
                    Sheets("Database").Range("O2").Value = Sheets("Sheet1").Range("F" & i).Value
                    Sheets("Database").Range("P2").Value = Sheets("Sheet1").Range("G" & i).Value
                    Exit Do
                End If
                If Sheets("Sheet1").Range("E" & i).Value = "" Then
                    Exit Do
                End If
                i = i + 1
            Loop
            i = 2
            Do
                If Sheets("Sheet1").Range("E" & i).Value Like "*SMC*" Then
                    Sheets("Database").Range("Q2").Value = Sheets("Sheet1").Range("F" & i).Value
                    Sheets("Database").Range("R2").Value = Sheets("Sheet1").Range("G" & i).Value
                    Exit Do
                End If
                If Sheets("Sheet1").Range("E" & i).Value = "" Then
                    Exit Do
                End If
                i = i + 1
            Loop
        
        'Find ISM Manager
            Application.Wait (Now + TimeValue("00:00:01"))
            Sheets("Sheet1").Select
            Call Extract_Date
            Call Extract_PSC_Date
            Call Check_Dates
            i = 2
            Sheets("Database").Range("U2").Value = ""
            Do
                If Sheets("Sheet1").Range("AP" & i).Value Like "*ISM*" And Sheets("Sheet1").Range("AT" & i).Value = True Then
                    Sheets("Database").Range("U2").Value = Sheets("Sheet1").Range("AO" & i).Value
                    Exit Do
                End If
                If Sheets("Sheet1").Range("AP" & i).Value = "" Then
                    Exit Do
                End If
                i = i + 1
            Loop

    'Deficiencies Entry
        i = 2
        j = 0
        Sheets("Sheet1").Select
    
        If Sheets("Sheet1").Range("A7").Value = "US Coast Guard " Then
            Do
                If Sheets("Sheet1").Range("J" & i).Value <> "" Then
                    j = j + 1
                    i = i + 1
                End If
                If Sheets("Sheet1").Range("J" & i).Value = "" Then
                    Exit Do
                End If
            Loop
            
            i = 2
            k = 0
            Do
                If Sheets("Sheet1").Range("O" & i).Value <> "" Then
                    k = k + 1
                    i = i + 1
                End If
                If Sheets("Sheet1").Range("O" & i).Value = "" Then
                    Exit Do
                End If
            Loop
        End If
        
        If Sheets("Sheet1").Range("A7").Value = "Vina Del Mar MoU " Then
            Do
                If Sheets("Sheet1").Range("J" & i).Value <> "" Then
                    j = j + 1
                    i = i + 1
                End If
                If Sheets("Sheet1").Range("J" & i).Value = "" Then
                    Exit Do
                End If
            Loop
            
            i = 2
            k = 0
            Do
                If Sheets("Sheet1").Range("O" & i).Value <> "" Then
                    k = k + 1
                    i = i + 1
                End If
                If Sheets("Sheet1").Range("O" & i).Value = "" Then
                    Exit Do
                End If
            Loop
        End If
        
        If Sheets("Sheet1").Range("A7").Value = "Black Sea MoU " Then
            Do
                If Sheets("Sheet1").Range("E" & i).Value <> "" Then
                    j = j + 1
                    i = i + 1
                End If
                If Sheets("Sheet1").Range("E" & i).Value = "" Then
                    Exit Do
                End If
            Loop
            
            i = 2
            k = 0
            Do
                If Sheets("Sheet1").Range("J" & i).Value <> "" Then
                    k = k + 1
                    i = i + 1
                End If
                If Sheets("Sheet1").Range("J" & i).Value = "" Then
                    Exit Do
                End If
            Loop
        End If
        
        If Sheets("Sheet1").Range("A7").Value = "Abuja MoU " Then
            Do
                If Sheets("Sheet1").Range("E" & i).Value <> "" Then
                    j = j + 1
                    i = i + 1
                End If
                If Sheets("Sheet1").Range("E" & i).Value = "" Then
                    Exit Do
                End If
            Loop
            
            i = 2
            k = 0
            Do
                If Sheets("Sheet1").Range("J" & i).Value <> "" Then
                    k = k + 1
                    i = i + 1
                End If
                If Sheets("Sheet1").Range("J" & i).Value = "" Then
                    Exit Do
                End If
            Loop
        End If
        
        If Sheets("Sheet1").Range("A7").Value = "Tokyo MoU " Then
            Do
                If Sheets("Sheet1").Range("E" & i).Value <> "" Then
                    j = j + 1
                    i = i + 1
                End If
                If Sheets("Sheet1").Range("E" & i).Value = "" Then
                    Exit Do
                End If
            Loop
            
            i = 2
            k = 0
            Do
                If Sheets("Sheet1").Range("J" & i).Value <> "" Then
                    k = k + 1
                    i = i + 1
                End If
                If Sheets("Sheet1").Range("J" & i).Value = "" Then
                    Exit Do
                End If
            Loop
        End If
        
        If Sheets("Sheet1").Range("A7").Value = "Indian Ocean MoU " Then
            Do
                If Sheets("Sheet1").Range("J" & i).Value <> "" Then
                    j = j + 1
                    i = i + 1
                End If
                If Sheets("Sheet1").Range("J" & i).Value = "" Then
                    Exit Do
                End If
            Loop
            
            i = 2
            k = 0
            Do
                If Sheets("Sheet1").Range("O" & i).Value <> "" Then
                    k = k + 1
                    i = i + 1
                End If
                If Sheets("Sheet1").Range("O" & i).Value = "" Then
                    Exit Do
                End If
            Loop
        End If
        
        If Sheets("Sheet1").Range("A7").Value = "Paris MoU " Then
            Do
                If Sheets("Sheet1").Range("O" & i).Value <> "" Then
                    j = j + 1
                    i = i + 1
                End If
                If Sheets("Sheet1").Range("O" & i).Value = "" Then
                    Exit Do
                End If
            Loop
    
        i = 2
        k = 0
        Do
            If Sheets("Sheet1").Range("T" & i).Value <> "" Then
                k = k + 1
                i = i + 1
            End If
            If Sheets("Sheet1").Range("T" & i).Value = "" Then
                Exit Do
            End If
        Loop
        End If
        
        If Sheets("Sheet1").Range("A7").Value = "Caribbean MoU " Then
            Do
                If Sheets("Sheet1").Range("O" & i).Value <> "" Then
                    j = j + 1
                    i = i + 1
                End If
                If Sheets("Sheet1").Range("O" & i).Value = "" Then
                    Exit Do
                End If
            Loop
    
        i = 2
        k = 0
        Do
            If Sheets("Sheet1").Range("T" & i).Value <> "" Then
                k = k + 1
                i = i + 1
            End If
            If Sheets("Sheet1").Range("T" & i).Value = "" Then
                Exit Do
            End If
        Loop
        End If
        
        Sheets("Database").Select
        If j > 0 Then
            i = 1
            For i = 1 To j
                Sheets("Database").Select
                Rows("5:5").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                Application.Wait (Now + TimeValue("00:00:01"))
            
                Sheets("Database").Select
                Sheets("Database").Range("A2:AB2").Select
                Selection.Copy
                Sheets("Database").Range("A5:AB5").Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                        :=False, Transpose:=False
            
                Sheets("Database").Range("V5").Value = Sheets("Database").Range("A5").Value & "_" & i
                
                If Sheets("Sheet1").Range("A7").Value = "US Coast Guard " Then
                    Sheets("Database").Range("W5").Value = Sheets("Sheet1").Range("J" & i + 1).Value
                    Sheets("Database").Range("X5").Value = Sheets("Sheet1").Range("K" & i + 1).Value
                    Sheets("Database").Range("Y5").Value = Sheets("Sheet1").Range("L" & i + 1).Value
                    Sheets("Database").Range("Z5").Value = "No"
                    Sheets("Database").Range("AA5").Value = "No"
                End If
                
                If Sheets("Sheet1").Range("A7").Value = "Vina Del Mar MoU " Then
                    Sheets("Database").Range("W5").Value = Sheets("Sheet1").Range("J" & i + 1).Value
                    Sheets("Database").Range("X5").Value = Sheets("Sheet1").Range("K" & i + 1).Value
                    Sheets("Database").Range("Y5").Value = Sheets("Sheet1").Range("L" & i + 1).Value
                    Sheets("Database").Range("Z5").Value = "No"
                    Sheets("Database").Range("AA5").Value = "No"
                End If
                
                If Sheets("Sheet1").Range("A7").Value = "Black Sea MoU " Then
                    Sheets("Database").Range("W5").Value = Sheets("Sheet1").Range("E" & i + 1).Value
                    Sheets("Database").Range("X5").Value = Sheets("Sheet1").Range("F" & i + 1).Value
                    Sheets("Database").Range("Y5").Value = Sheets("Sheet1").Range("G" & i + 1).Value
                    Sheets("Database").Range("Z5").Value = "No"
                    Sheets("Database").Range("AA5").Value = "No"
                End If
                
                If Sheets("Sheet1").Range("A7").Value = "Abuja MoU " Then
                    Sheets("Database").Range("W5").Value = Sheets("Sheet1").Range("E" & i + 1).Value
                    Sheets("Database").Range("X5").Value = Sheets("Sheet1").Range("F" & i + 1).Value
                    Sheets("Database").Range("Y5").Value = Sheets("Sheet1").Range("G" & i + 1).Value
                    Sheets("Database").Range("Z5").Value = "No"
                    Sheets("Database").Range("AA5").Value = "No"
                End If
                
                If Sheets("Sheet1").Range("A7").Value = "Paris MoU " Then
                    Sheets("Database").Range("W5").Value = Sheets("Sheet1").Range("O" & i + 1).Value
                    Sheets("Database").Range("X5").Value = Sheets("Sheet1").Range("P" & i + 1).Value
                    Sheets("Database").Range("Y5").Value = Sheets("Sheet1").Range("Q" & i + 1).Value
                    Sheets("Database").Range("Z5").Value = "No"
                    Sheets("Database").Range("AA5").Value = "No"
                End If
                
                If Sheets("Sheet1").Range("A7").Value = "Caribbean MoU " Then
                    Sheets("Database").Range("W5").Value = Sheets("Sheet1").Range("O" & i + 1).Value
                    Sheets("Database").Range("X5").Value = Sheets("Sheet1").Range("P" & i + 1).Value
                    Sheets("Database").Range("Y5").Value = Sheets("Sheet1").Range("Q" & i + 1).Value
                    Sheets("Database").Range("Z5").Value = "No"
                    Sheets("Database").Range("AA5").Value = "No"
                End If
                
                If Sheets("Sheet1").Range("A7").Value = "Tokyo MoU " Then
                    Sheets("Database").Range("W5").Value = Sheets("Sheet1").Range("E" & i + 1).Value
                    Sheets("Database").Range("X5").Value = Sheets("Sheet1").Range("F" & i + 1).Value
                    Sheets("Database").Range("Y5").Value = Sheets("Sheet1").Range("G" & i + 1).Value
                    Sheets("Database").Range("Z5").Value = "No"
                    Sheets("Database").Range("AA5").Value = "No"
                End If
                
                If Sheets("Sheet1").Range("A7").Value = "Indian Ocean MoU " Then
                    Sheets("Database").Range("W5").Value = Sheets("Sheet1").Range("J" & i + 1).Value
                    Sheets("Database").Range("X5").Value = Sheets("Sheet1").Range("K" & i + 1).Value
                    Sheets("Database").Range("Y5").Value = Sheets("Sheet1").Range("L" & i + 1).Value
                    Sheets("Database").Range("Z5").Value = "No"
                    Sheets("Database").Range("AA5").Value = "No"
                End If
            Next i
        End If
        If k > 0 Then
            i = 1
            For i = 1 To k
                Sheets("Database").Select
                Rows("5:5").Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                Application.Wait (Now + TimeValue("00:00:01"))
            
                Sheets("Database").Select
                Sheets("Database").Range("A2:AB2").Select
                Selection.Copy
                Sheets("Database").Range("A5:AB5").Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                        :=False, Transpose:=False
            
                Sheets("Database").Range("V5").Value = Sheets("Database").Range("A5").Value & "_" & (i + j)
                
                If Sheets("Sheet1").Range("A7").Value = "US Coast Guard " Then
                    Sheets("Database").Range("W5").Value = Sheets("Sheet1").Range("O" & i + 1).Value
                    Sheets("Database").Range("X5").Value = Sheets("Sheet1").Range("P" & i + 1).Value
                    Sheets("Database").Range("Y5").Value = Sheets("Sheet1").Range("Q" & i + 1).Value
                    Sheets("Database").Range("Z5").Value = "Yes"
                    Sheets("Database").Range("AA5").Value = Sheets("Sheet1").Range("S" & i + 1).Value
                End If
                
                If Sheets("Sheet1").Range("A7").Value = "Vina Del Mar MoU " Then
                    Sheets("Database").Range("W5").Value = Sheets("Sheet1").Range("O" & i + 1).Value
                    Sheets("Database").Range("X5").Value = Sheets("Sheet1").Range("P" & i + 1).Value
                    Sheets("Database").Range("Y5").Value = Sheets("Sheet1").Range("Q" & i + 1).Value
                    Sheets("Database").Range("Z5").Value = "Yes"
                    Sheets("Database").Range("AA5").Value = Sheets("Sheet1").Range("S" & i + 1).Value
                End If
                
                If Sheets("Sheet1").Range("A7").Value = "Black Sea MoU " Then
                    Sheets("Database").Range("W5").Value = Sheets("Sheet1").Range("J" & i + 1).Value
                    Sheets("Database").Range("X5").Value = Sheets("Sheet1").Range("K" & i + 1).Value
                    Sheets("Database").Range("Y5").Value = Sheets("Sheet1").Range("L" & i + 1).Value
                    Sheets("Database").Range("Z5").Value = "Yes"
                    Sheets("Database").Range("AA5").Value = Sheets("Sheet1").Range("N" & i + 1).Value
                End If
                
                If Sheets("Sheet1").Range("A7").Value = "Abuja MoU " Then
                    Sheets("Database").Range("W5").Value = Sheets("Sheet1").Range("J" & i + 1).Value
                    Sheets("Database").Range("X5").Value = Sheets("Sheet1").Range("K" & i + 1).Value
                    Sheets("Database").Range("Y5").Value = Sheets("Sheet1").Range("L" & i + 1).Value
                    Sheets("Database").Range("Z5").Value = "Yes"
                    Sheets("Database").Range("AA5").Value = Sheets("Sheet1").Range("N" & i + 1).Value
                End If
                
                If Sheets("Sheet1").Range("A7").Value = "Paris MoU " Then
                    Sheets("Database").Range("W5").Value = Sheets("Sheet1").Range("T" & i + 1).Value
                    Sheets("Database").Range("X5").Value = Sheets("Sheet1").Range("U" & i + 1).Value
                    Sheets("Database").Range("Y5").Value = Sheets("Sheet1").Range("V" & i + 1).Value
                    Sheets("Database").Range("Z5").Value = "Yes"
                    Sheets("Database").Range("AA5").Value = Sheets("Sheet1").Range("X" & i + 1).Value
                End If
                
                If Sheets("Sheet1").Range("A7").Value = "Caribbean MoU " Then
                    Sheets("Database").Range("W5").Value = Sheets("Sheet1").Range("T" & i + 1).Value
                    Sheets("Database").Range("X5").Value = Sheets("Sheet1").Range("U" & i + 1).Value
                    Sheets("Database").Range("Y5").Value = Sheets("Sheet1").Range("V" & i + 1).Value
                    Sheets("Database").Range("Z5").Value = "Yes"
                    Sheets("Database").Range("AA5").Value = Sheets("Sheet1").Range("X" & i + 1).Value
                End If
                
                If Sheets("Sheet1").Range("A7").Value = "Tokyo MoU " Then
                    Sheets("Database").Range("W5").Value = Sheets("Sheet1").Range("J" & i + 1).Value
                    Sheets("Database").Range("X5").Value = Sheets("Sheet1").Range("K" & i + 1).Value
                    Sheets("Database").Range("Y5").Value = Sheets("Sheet1").Range("L" & i + 1).Value
                    Sheets("Database").Range("Z5").Value = "Yes"
                    Sheets("Database").Range("AA5").Value = Sheets("Sheet1").Range("N" & i + 1).Value
                End If
                
                If Sheets("Sheet1").Range("A7").Value = "Indian Ocean MoU " Then
                    Sheets("Database").Range("W5").Value = Sheets("Sheet1").Range("O" & i + 1).Value
                    Sheets("Database").Range("X5").Value = Sheets("Sheet1").Range("P" & i + 1).Value
                    Sheets("Database").Range("Y5").Value = Sheets("Sheet1").Range("Q" & i + 1).Value
                    Sheets("Database").Range("Z5").Value = "Yes"
                    Sheets("Database").Range("AA5").Value = Sheets("Sheet1").Range("S" & i + 1).Value
                End If
            Next i
        End If
        If j = 0 Then
            Sheets("Database").Select
            Rows("5:5").Select
            Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            Application.Wait (Now + TimeValue("00:00:01"))
            
            Sheets("Database").Select
            Sheets("Database").Range("A2:AB2").Select
            Selection.Copy
            Sheets("Database").Range("A5:AB5").Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
        End If
        
    Sheets("User Sheet").Range("K4").Value = "Error0" 'Error handling

    'Navigate again to vessel
        Application.Wait (Now + TimeValue("00:00:02"))
        Set idoc = IE.document
        Set IEobj = idoc.getElementById("P_ENTREE_ENTETE")
        IEobj.Value = "9713337"
        Application.Wait (Now + TimeValue("00:00:01"))
        Set Buttons = idoc.getElementsByClassName("btn btn-default")
        Application.Wait (Now + TimeValue("00:00:01"))
        For Each Button In Buttons
            If Button.className = "btn btn-default" Then
                Button.Click
                Exit For
            End If
        Next Button
        With IE
            While IE.ReadyState <> 4
                DoEvents
            Wend
        End With
    
    'Clear Sheet1 to start again the loop
        Sheets("Sheet1").Select
        Sheets("Sheet1").Range("A1:AZ154").Select
        Selection.ClearContents
        Sheets("User Sheet").Range("B3").Value = P_INSP_Start
        P_INSP_Start = P_INSP_Start + 1
    
    counter = counter + 1 'count how many inspections I searched with the same credentials
    
    If counter > 9 Then 'Handle the server not to overload from my extraction
        counter = 0
        Application.Wait (Now + TimeValue("00:00:01"))
        'Log Out Equasis and Exit IE
        With IE
           .navigate ("https://www.equasis.org/EquasisWeb/public/HomePage?fs=DetailsPSC&P_ACTION=NEW_CONNECTION")
            While IE.ReadyState <> 4
                DoEvents
            Wend
        End With
        Application.Wait (Now + TimeValue("00:00:02"))
        IE.Quit
        Application.Wait (Now + TimeValue("00:00:01"))
        Set IE = Nothing
        loginCounter = loginCounter + 1
        GoTo Start
    End If
    Sheets("User Sheet").Range("K4").Value = ""
    
    If counter > 15 Then 'Extreme case for error handling when there is no data for the PSC inspection id I placed
    NoInsp:
        CounterNoInsp = CounterNoInsp + 1
        counter = counter + 1
        If counter > 9 Then
            counter = 0
            loginCounter = loginCounter + 1
            'Log Out Equasis and Exit IE
            With IE
            .navigate ("https://www.equasis.org/EquasisWeb/public/HomePage?fs=DetailsPSC&P_ACTION=NEW_CONNECTION")
                While IE.ReadyState <> 4
                    DoEvents
                Wend
            End With
            Application.Wait (Now + TimeValue("00:00:02"))
            IE.Quit
            Application.Wait (Now + TimeValue("00:00:01"))
            Set IE = Nothing
        End If
        Sheets("Sheet1").Select
        Sheets("Sheet1").Range("A1:AZ154").Select
        Selection.ClearContents
        Sheets("User Sheet").Range("B3").Value = P_INSP_Start
        P_INSP_Start = P_INSP_Start + 1
        
        If counter > 0 Then
            refresher = 1
            With IE
                .navigate ("https://www.equasis.org/EquasisWeb/public/HomePage")
                While IE.ReadyState <> 4
                    DoEvents
                Wend
            End With
            Application.Wait (Now + TimeValue("00:00:01"))
        End If
        
        CountNoInsp = CountNoInsp + 1 'counter on how many times in a row the inspection not found
        cc = cc + 1
        DoEvents
        If CountNoInsp > 1000 Then 'if 1000 times in a row we missed an inspection it could be an alert that one of our emails is banned
            Unload StatusBar 'Close the Status Bar
            Sheets("User Sheet").Range("K4").Value = "" 'Delete Error Handling entry
            Sheets("User Sheet").Range("M1").Value = "" 'Delete Error Handling entry
            MsgBox ("Check what is going on, we had multiple fails with this email: " & Username)
            Exit Sub
        End If
        Call RunStatusBar(cc, count)
        GoTo Start
    End If
    CountNoInsp = 0
    
    cc = cc + 1
    DoEvents
    Call RunStatusBar(cc, count)
    
    'When the <Exit> button is pressed on Status Bar
        If Sheets("User Sheet").Range("M1").Value = "Exit" Then
            Sheets("User Sheet").Range("M1").Value = ""
            Sheets("Database").Select
            Sheets("Sheet1").Visible = False
            Unload StatusBar
            'Log Out Equasis and Exit IE
            With IE
                .navigate ("https://www.equasis.org/EquasisWeb/public/HomePage?fs=DetailsPSC&P_ACTION=NEW_CONNECTION")
                While IE.ReadyState <> 4
                    DoEvents
                Wend
            End With
            Application.Wait (Now + TimeValue("00:00:01"))
            IE.Quit
            Application.Wait (Now + TimeValue("00:00:01"))
            Set IE = Nothing
            MsgBox "You stopped the procedure"
            Exit Sub
        End If
    Loop

    Sheets("Database").Select
    Sheets("Sheet1").Visible = False 'Hide the Back-end page
    Sheets("Emails").Visible = False
    Sheets("Non Found IDs").Visible = False

    'Disconnect from Equasis and close Internet Explorer
    With IE
        .navigate ("https://www.equasis.org/EquasisWeb/public/HomePage?fs=DetailsPSC&P_ACTION=NEW_CONNECTION")
        While IE.ReadyState <> 4
            DoEvents
        Wend
    End With
    Application.Wait (Now + TimeValue("00:00:02"))
    IE.Quit
    Application.Wait (Now + TimeValue("00:00:01"))
    Set IE = Nothing

    Unload StatusBar 'Close the Status Bar
    Sheets("User Sheet").Range("K4").Value = "" 'Delete Error Handling entry
    Sheets("User Sheet").Range("M1").Value = "" 'Delete Error Handling entry
    Sheets("User Sheet").Range("I2:I1000").Delete

    MsgBox "You are ready! We extracted " & count - CounterNoInsp & " inspections out of total " & count & ". Check the extracted PSC inspections.."

    End Sub

Procedure Extract_Date

    Sub Extract_Date()
    Range("AS2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(LEFT(RC[-2],5)=""since"",RIGHT(RC[-2],11),RIGHT(RC[-2],8))"
    Range("AS2").Select
    Selection.AutoFill Destination:=Range("AS2:AS14")
    Range("AS2:AS14").Select
    Range("AS2").Select
    End Sub

Procedure Extract_PSC_Date

    Sub Extract_PSC_Date()
    Range("AS1").Select
    ActiveCell.FormulaR1C1 = "=TRIM(R[10]C[-44])"
    Range("AS2").Select
    End Sub

Procedure Check_Dates

    Sub Check_Dates()
    ActiveCell.FormulaR1C1 = ""
    Range("AT2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-1]<R[-1]C[-1]"
    Range("AT2").Select
    Selection.AutoFill Destination:=Range("AT2:AT14")
    Range("AT2:AT14").Select
    Range("AT3:AT14").Select
    Range("AT14").Activate
    ActiveCell.FormulaR1C1 = ""
    Range("AT3:AT13").Select
    Selection.ClearContents
    Range("AT2").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]<R1C45"
    Range("AT2").Select
    Selection.AutoFill Destination:=Range("AT2:AT14")
    Range("AT2:AT14").Select
    Range("AT2").Select
    End Sub

Procedure for Cancel Click on Status Bar

    Private Sub Cancel_Click()

    Sheets("User Sheet").Range("M1").Value = "Exit"

    End Sub

    Private Sub UserForm_Click()

    Sheets("User Sheet").Range("M1").Value = "Exit"

    End Sub

How the Status Bar is looking

![image](https://user-images.githubusercontent.com/86843206/179949534-47969921-4bd3-4e3b-a230-294b5c5e45b1.png)


-Here it ends, feel free to contact me if you want to ask you anything or you want to navigate you through






