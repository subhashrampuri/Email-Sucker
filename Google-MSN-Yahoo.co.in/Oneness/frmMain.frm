VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9165
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.StatusBar stsBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   26
      Top             =   2940
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   16113
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame5 
      Caption         =   "Stats"
      Height          =   1725
      Left            =   4725
      TabIndex        =   19
      Top             =   1215
      Width           =   4335
      Begin VB.Label lblElapsedTime 
         Caption         =   "Elapsed:"
         Height          =   240
         Left            =   2385
         TabIndex        =   24
         Top             =   270
         Width           =   1725
      End
      Begin VB.Label lblPagesSacnned 
         Caption         =   "Pages Scanned:"
         Height          =   240
         Left            =   2385
         TabIndex        =   23
         Top             =   585
         Width           =   1635
      End
      Begin VB.Label lblEmailsExtracted 
         Caption         =   "Emails Extracted:"
         Height          =   240
         Left            =   270
         TabIndex        =   22
         Top             =   945
         Width           =   2400
      End
      Begin VB.Label lblEndTime 
         Caption         =   "End Time:"
         Height          =   195
         Left            =   270
         TabIndex        =   21
         Top             =   585
         Width           =   1815
      End
      Begin VB.Label lblStartTime 
         Caption         =   "Start Time:"
         Height          =   240
         Left            =   270
         TabIndex        =   20
         Top             =   270
         Width           =   1680
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Logging Details [ODBC]"
      Height          =   1050
      Left            =   4725
      TabIndex        =   14
      Top             =   90
      Width           =   4335
      Begin VB.CommandButton cmdDropTables 
         Caption         =   "Drop Tables"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3015
         TabIndex        =   18
         Top             =   675
         Width           =   1185
      End
      Begin VB.CommandButton cmdTables 
         Caption         =   "Create Tables"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1620
         TabIndex        =   17
         Top             =   675
         Width           =   1185
      End
      Begin VB.CommandButton cmdTestODBC 
         Caption         =   "&Test Connection"
         Height          =   285
         Left            =   135
         TabIndex        =   16
         Top             =   675
         Width           =   1455
      End
      Begin VB.TextBox txtsODBC 
         Height          =   330
         Left            =   135
         TabIndex        =   15
         Top             =   315
         Width           =   4065
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Search engine details"
      Height          =   1860
      Left            =   180
      TabIndex        =   3
      Top             =   90
      Width           =   4425
      Begin VB.Frame Frame3 
         Caption         =   "Records"
         Height          =   555
         Left            =   2700
         TabIndex        =   10
         Top             =   1170
         Width           =   1500
         Begin VB.TextBox txtiEnd 
            Height          =   285
            Left            =   810
            TabIndex        =   13
            Text            =   "1000"
            Top             =   225
            Width           =   465
         End
         Begin VB.TextBox txtiStart 
            Height          =   285
            Left            =   90
            TabIndex        =   11
            Text            =   "10"
            Top             =   225
            Width           =   465
         End
         Begin VB.Label Label3 
            Caption         =   "-"
            Height          =   195
            Left            =   630
            TabIndex        =   12
            Top             =   270
            Width           =   375
         End
      End
      Begin VB.TextBox txtsIncValue 
         Height          =   330
         Left            =   1980
         TabIndex        =   9
         Text            =   "10"
         Top             =   1350
         Width           =   600
      End
      Begin VB.TextBox txtsIncVariable 
         Height          =   330
         Left            =   1980
         TabIndex        =   7
         Top             =   855
         Width           =   2220
      End
      Begin VB.TextBox txtsURL 
         Height          =   330
         Left            =   720
         TabIndex        =   4
         Top             =   360
         Width           =   3480
      End
      Begin VB.Label Label2 
         Caption         =   "Incrementing Factor:"
         Height          =   195
         Left            =   270
         TabIndex        =   8
         Top             =   1395
         Width           =   1635
      End
      Begin VB.Label Label1 
         Caption         =   "Incrementing Variable:"
         Height          =   195
         Left            =   270
         TabIndex        =   6
         Top             =   855
         Width           =   1635
      End
      Begin VB.Label lblCaption 
         Caption         =   "URI:"
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   5
         Top             =   405
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Extract E-mails"
      Height          =   915
      Left            =   180
      TabIndex        =   0
      Top             =   2025
      Width           =   4425
      Begin VB.CommandButton cmdStop 
         Caption         =   "&Stop"
         Height          =   330
         Left            =   1035
         TabIndex        =   2
         Top             =   225
         Width           =   780
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "&Start"
         Height          =   330
         Left            =   135
         TabIndex        =   1
         Top             =   225
         Width           =   825
      End
      Begin VB.Label lblStatus 
         Height          =   285
         Left            =   135
         TabIndex        =   25
         Top             =   585
         Width           =   4155
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_iEmailsExtracted As Long
Dim m_iPagesScanned As Long
Dim m_vStartTime
Dim m_bStop As Boolean
Dim m_objConnection As ADODB.Connection
Private Function fbExtractEmail(ByVal sTemp As String)
Dim sFirst As String
Dim sEnd As String
Dim iLoop As Long
Dim iFirst As Long
Dim iEnd As Long
Dim sEmail As String
    For iLoop = 1 To Len(sTemp)
        If Mid(sTemp, iLoop, 1) = "@" Then
                iFirst = fiStartingPoint(sTemp, iLoop)
                iEnd = fiEndingPoint(sTemp, iLoop)
                sEmail = Mid(sTemp, iFirst, iEnd - iFirst)
                If Mid(sEmail, Len(sEmail), 1) = "." Then
                   sEmail = Mid(sEmail, 1, Len(sEmail) - 1)
                   If fbCheckEmail(sEmail) Then
                        lblStatus.Caption = "logging:" & sEmail
                        DoEvents
                        pLogEmail (sEmail)
                   End If
                Else
                   If fbCheckEmail(sEmail) Then
                        lblStatus.Caption = "logging:" & sEmail
                        DoEvents
                        pLogEmail (sEmail)
                   End If
                End If
       End If
    Next
End Function
Private Sub pLogEmail(ByVal sTemp As String)
    
    If cmdStart.Caption = "&Start" Then
        cmdStart.Caption = "&Start [|]"
    ElseIf cmdStart.Caption = "&Start [|]" Then
        cmdStart.Caption = "&Start [/]"
    ElseIf cmdStart.Caption = "&Start [/]" Then
        cmdStart.Caption = "&Start [-]"
    ElseIf cmdStart.Caption = "&Start [-]" Then
        cmdStart.Caption = "&Start [\]"
    ElseIf cmdStart.Caption = "&Start [\]" Then
        cmdStart.Caption = "&Start [|]"
    End If
          
    m_iEmailsExtracted = m_iEmailsExtracted + 1
    lblEmailsExtracted = "Emails Extracted: [" & m_iEmailsExtracted & "]"
    If DateDiff("s", m_vStartTime, Time) < 60 Then
        lblElapsedTime.Caption = "Elapsed: [" & DateDiff("s", m_vStartTime, Time) & " Sec(s)]"
    ElseIf DateDiff("s", m_vStartTime, Time) > 60 And DateDiff("s", m_vStartTime, Time) < 3600 Then
        lblElapsedTime.Caption = "Elapsed: [" & Format((DateDiff("s", m_vStartTime, Time) / 60), "#.#") & " Min(s)]"
    ElseIf (Time - m_vStartTime) > 3600 Then
        lblElapsedTime.Caption = "Elapsed: [" & Format(((DateDiff("s", m_vStartTime, Time) / 60) / 60), "#.#") & " Hr(s)]"
    End If
    DoEvents
    fbAddEmail sTemp
       
End Sub
Private Function fiStartingPoint(ByVal sTemp As String, ByVal iTemp As Long) As Long
    Dim iLoop As Long
    For iLoop = iTemp To 1 Step -1
        If fbStartEndCharacters(Mid(sTemp, iLoop - 1, 1)) Then
            fiStartingPoint = iLoop
            Exit For
        End If
    Next
End Function
Private Function fiEndingPoint(ByVal sTemp As String, ByVal iTemp As Long) As Long
    Dim iLoop As Long
    For iLoop = iTemp To Len(sTemp)
        If fbStartEndCharacters(Mid(sTemp, iLoop, 1)) Then
            fiEndingPoint = iLoop
            Exit For
        End If
    Next
End Function
Private Sub cmdStart_Click()
    Dim vLinkHolder
    Dim vVariableHolder
    Dim vIncrementingVariable
    Dim bIncrementVariable As Boolean
    Dim i As Integer
    Dim sExtractLink As String
   
   
    vLinkHolder = Split(Trim(txtsURL), "?")
       
    If UBound(vLinkHolder) < 1 Then
        MsgBox "Please enter proper URL with variables!", vbExclamation, App.Title
        Exit Sub
    End If
    
    vVariableHolder = Split(vLinkHolder(1), "&")
      
    If Len(Trim(txtsIncVariable.Text)) = 0 Then
        MsgBox "Please enter incrementing variable number!", vbExclamation, App.Title
        Exit Sub
    End If
    
    
    For i = 0 To UBound(vVariableHolder)
        vIncrementingVariable = Split(vVariableHolder(i), "=")
        If vIncrementingVariable(0) = Trim(txtsIncVariable.Text) Then
            bIncrementVariable = True
        End If
    Next
    
    
    If bIncrementVariable Then
        If Not fbGetConnection Then
            Exit Sub
        End If
        cmdStart.Enabled = False
        m_vStartTime = Time
        lblStartTime.Caption = "Start Time: [" & Time & "]"
        pStartExtract vLinkHolder, vVariableHolder
    Else
        MsgBox "No Incrementing variable found!", vbExclamation, App.Title
    End If
    lblEndTime.Caption = "End Time: [" & Time & "]"
    cmdStart.Caption = "&Start"
    cmdStart.Enabled = True
    m_objConnection.Close
    Set m_objConnection = Nothing
    lblStatus.Caption = ""
    
End Sub
Private Sub pStartExtract(ByVal vLinkHolder, ByVal vVariableHolder)
    Dim objXMLHTTP As New XMLHTTP40
    Dim sText As String
    Dim sURL As String
    Dim sURLFinal As String
    Dim vIncrementingVariable
    Dim i As Integer
    
    m_bStop = False
    sURL = sURL & vLinkHolder(0) & "?"
    For i = 0 To UBound(vVariableHolder)
        vIncrementingVariable = Split(vVariableHolder(i), "=")
        If vIncrementingVariable(0) = Trim(txtsIncVariable.Text) Then
            
        Else
            If i = UBound(vVariableHolder) Then
                sURL = sURL & vVariableHolder(i)
            Else
                sURL = sURL & vVariableHolder(i) & "&"
            End If
        End If
    Next
    m_iPagesScanned = 0
     
    For i = txtiStart To txtiEnd Step txtsIncValue
        If m_bStop Then
            Exit For
        End If
        sURLFinal = sURL & "&" & Trim(txtsIncVariable.Text) & "=" & i
        objXMLHTTP.open "GET", sURLFinal, False
        stsBar.Panels(1).Text = "Requesting page: " & sURLFinal
        objXMLHTTP.sEnd
        sText = objXMLHTTP.responseText
        sText = SF_removeAll(sText, "<b>")
        sText = SF_removeAll(sText, "</b>")
        m_iPagesScanned = m_iPagesScanned + 1
        lblPagesSacnned.Caption = "Pages Scanned: [" & m_iPagesScanned & "]"
        fbExtractEmail sText
        DoEvents
    Next i
    
    
    
End Sub
Private Sub cmdStop_Click()
    m_bStop = True
End Sub

Private Sub cmdTables_Click()
    If fbCreate Then
        MsgBox "Table Created successfully!", vbInformation, App.Title
    End If
End Sub

Private Sub cmdTestODBC_Click()
On Error GoTo LocalErr
    Dim objCon As New ADODB.Connection
    objCon.open Trim(txtsODBC)
    objCon.Close
    MsgBox "Test Successfull.", vbInformation, App.Title
    
Exit Sub
LocalErr:
    MsgBox Err.Description, vbCritical, App.Title
    
End Sub
Private Function fbStartEndCharacters(ByVal sChar As String) As Boolean
    Dim iChar As Integer
    
    For iChar = 1 To 44
        If sChar = Chr(iChar) Then
            fbStartEndCharacters = True
            Exit Function
        End If
    Next
    
    For iChar = 58 To 63
        If sChar = Chr(iChar) Then
            fbStartEndCharacters = True
            Exit Function
        End If
    Next

    For iChar = 91 To 96
        If sChar = Chr(iChar) Then
            fbStartEndCharacters = True
            Exit Function
        End If
    Next

    For iChar = 123 To 250
        If sChar = Chr(iChar) Then
            fbStartEndCharacters = True
            Exit Function
        End If
    Next

    If sChar = "/" Then
        fbStartEndCharacters = True
    End If
End Function
Private Sub Command3_Click()
    MsgBox Asc("-")
End Sub



Private Sub Form_Load()
    pInitalizeComponent
End Sub
Private Sub pInitalizeComponent()
    txtsODBC = "DSN=emaillog;UID=email-log;PWD=emaillog"
    Me.Caption = App.Title
End Sub
Private Function fbCheckEmail(ByVal sEmail As String) As Boolean
Dim bResult As Boolean
Dim sStr As String
Dim iAtPos As Integer
Dim iDotPos As Integer

bResult = False
    
    sStr = sEmail
    iAtPos = InStr(1, sStr, "@")
    
    If iAtPos > 0 Then
        iDotPos = InStr(iAtPos, sStr, ".")
        If ((iDotPos > iAtPos + 1) And (Len(sStr) > (iDotPos + 1))) Then
            bResult = True
        Else
            bResult = False
        End If
   Else
        bResult = False
   End If
                               
   fbCheckEmail = bResult
End Function

Private Function fbGetConnection() As Boolean

On Error GoTo LocalErr
    
    Set m_objConnection = New ADODB.Connection
    m_objConnection.open (txtsODBC)

    fbGetConnection = True
    Exit Function
LocalErr:
    If Err.Number = -2147467259 Then
        MsgBox "DSN Not found. Please contact system adminstrator", vbCritical, App.Title
    Else
        MsgBox Err.Description & " " & Err.Number, vbCritical, App.Title
    End If
    fbGetConnection = False
End Function

Private Function fbCreate() As Boolean
    Dim bCreate As Boolean
    bCreate = True
    Dim objCon As New ADODB.Connection
    objCon.open Trim(txtsODBC.Text)
    If Not fbDropTable(objCon) And fbDropProcedure(objCon) Then
        bCreate = False
    End If
    
    If bCreate Then
        If Not fbCreateTable(objCon) And fbCreateProcedure(objCon) Then
            fbCreate = False
        End If
    End If
    
    objCon.Close
    Set objCon = Nothing
    bCreate = True
    fbCreate = bCreate
End Function
Private Function fbDrop() As Boolean
    Dim bDrop As Boolean
    bDrop = True
    Dim objCon As New ADODB.Connection
    objCon.open Trim(txtsODBC.Text)
    
    If Not fbDropTable(objCon) And fbDropProcedure(objCon) Then
        bDrop = False
    End If
    objCon.Close
    Set objCon = Nothing
    
    bDrop = True
    fbDrop = bDrop
End Function
Private Function fbCreateTable(objCon As ADODB.Connection) As Boolean
    Dim sSelect As String
    Dim adoCmd As New ADODB.Command
    
  
    
    sSelect = "CREATE TABLE [tblEmailLog] (" _
        & " [ema_lID] [numeric](18, 0) IDENTITY (1, 1) NOT NULL ," _
        & " [ema_sName] [varchar] (100) NOT NULL ," _
        & " CONSTRAINT [PK_tblEmailLog] PRIMARY KEY  CLUSTERED" _
        & " ([ema_lID])  ON [PRIMARY] ," _
        & " CONSTRAINT [IX_tblEmailLog] UNIQUE  NONCLUSTERED" _
        & " ([ema_sName]" _
        & "  )  ON [PRIMARY]" _
        & " ) ON [PRIMARY]"
        
    adoCmd.ActiveConnection = objCon
    adoCmd.CommandType = adCmdText
    adoCmd.CommandText = sSelect
    adoCmd.Execute
   
    Set adoCmd = Nothing
   
    fbCreateTable = True
Exit Function
LocalErr:
    fbCreateTable = False
End Function
Private Function fbCreateProcedure(objCon As ADODB.Connection) As Boolean
    Dim sSelect As String
    Dim adoCmd As New ADODB.Command
    
    sSelect = "CREATE PROCEDURE [dbo].[sp_itblEmailLog]" _
        & " (@ema_sName as varchar(100)) " _
        & " as INSERT INTO tblEmailLog(ema_sName)  VALUES " _
        & " (@ema_sName)"
        
    adoCmd.ActiveConnection = objCon
    adoCmd.CommandType = adCmdText
    adoCmd.CommandText = sSelect
    adoCmd.Execute
    Set adoCmd = Nothing
    fbCreateProcedure = True
Exit Function
LocalErr:
    fbCreateProcedure = False
End Function
Private Function fbAddEmail(ByVal sEmail As String) As Boolean
    Dim adoCmd As New ADODB.Command
On Error GoTo LocalErr
    adoCmd.ActiveConnection = m_objConnection
    adoCmd.Prepared = True
    adoCmd.CommandType = adCmdStoredProc
    adoCmd.CommandText = "sp_itblEmailLog"
    adoCmd.Parameters.Append adoCmd.CreateParameter("@ema_sName", adVarChar, adParamInput, 100, sEmail)
    adoCmd.Execute
    fbAddEmail = True
Exit Function
LocalErr:
    Resume Next
End Function

Private Function fbDropProcedure(objCon As ADODB.Connection) As Boolean
  Dim objCmd As New ADODB.Command
  Dim sSelect As String
  
 On Error GoTo LocalErr
 
  sSelect = "IF EXISTS (SELECT name FROM sysobjects " _
    & " WHERE name = 'sp_itblEmailLog' AND type = 'P')" _
    & " DROP PROCEDURE sp_itblEmailLog"
      
    objCmd.ActiveConnection = objCon
    objCmd.CommandType = adCmdText
    objCmd.CommandText = sSelect
    objCmd.Execute
      
    If Err.Number <> 0 Then
        fbDropProcedure = False
    Else
        fbDropProcedure = True
    End If
    Set objCmd = Nothing
 Exit Function
LocalErr:
    fbDropProcedure = False
End Function
Private Function fbDropTable(objCon As ADODB.Connection) As Boolean
  Dim objCmd As New ADODB.Command
  Dim sSelect As String
  
 On Error GoTo LocalErr
 
    sSelect = "IF EXISTS (SELECT name FROM sysobjects " _
    & " WHERE name = 'tblEmailLog' AND type = 'U')" _
    & " DROP TABLE tblEmailLog"
    
    objCmd.ActiveConnection = objCon
    objCmd.CommandType = adCmdText
    objCmd.CommandText = sSelect
    objCmd.Execute
     
    If Err.Number <> 0 Then
        fbDropTable = False
    Else
        fbDropTable = True
    End If
    
    Set objCmd = Nothing
 Exit Function
LocalErr:
    fbDropTable = False
End Function

