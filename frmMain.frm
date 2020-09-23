VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Commentor 1.1"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5610
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   5610
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLoadDir 
      Cancel          =   -1  'True
      Caption         =   "Load folder"
      Height          =   255
      Left            =   1890
      TabIndex        =   13
      Top             =   6540
      Width           =   1200
   End
   Begin VB.Frame fraDescription 
      Caption         =   "Description (Valid : %s, %t, %n, %p, %r)"
      Height          =   1005
      Left            =   30
      TabIndex        =   10
      Top             =   2340
      Width           =   5565
      Begin VB.TextBox txtDescription 
         Height          =   705
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Text            =   "frmMain.frx":0442
         Top             =   210
         Width           =   5355
      End
   End
   Begin VB.Frame fraErrorHandling 
      Caption         =   "Error handling (Valid : %s, %t, %n, %r)"
      Height          =   3180
      Left            =   30
      TabIndex        =   5
      Top             =   3330
      Width           =   5535
      Begin VB.CheckBox chkErrorHandling 
         Caption         =   "Add error handling to file"
         Height          =   225
         Left            =   3270
         TabIndex        =   9
         Top             =   2880
         Width           =   2145
      End
      Begin VB.TextBox txtErrorHandling 
         Height          =   2055
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Text            =   "frmMain.frx":0473
         Top             =   780
         Width           =   5325
      End
      Begin VB.TextBox txtErrorHandlingTop 
         Height          =   525
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Text            =   "frmMain.frx":06E2
         Top             =   240
         Width           =   5325
      End
   End
   Begin VB.Frame fraComments 
      Caption         =   "Comments at start of functions (Valid : %s, %t, %n, %p, %r, %d)"
      Height          =   2265
      Left            =   30
      TabIndex        =   3
      Top             =   60
      Width           =   5535
      Begin VB.ComboBox cboLang 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         ItemData        =   "frmMain.frx":0705
         Left            =   120
         List            =   "frmMain.frx":0715
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1920
         Width           =   1665
      End
      Begin VB.CheckBox chkComments 
         Caption         =   "Add comments to file"
         Height          =   195
         Left            =   3600
         TabIndex        =   8
         Top             =   1980
         Width           =   1815
      End
      Begin VB.TextBox txtComments 
         Height          =   1665
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Text            =   "frmMain.frx":073E
         Top             =   240
         Width           =   5325
      End
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   255
      Left            =   4350
      TabIndex        =   2
      Top             =   6540
      Width           =   1200
   End
   Begin VB.CommandButton cmdLoadFile 
      Caption         =   "Load file"
      Default         =   -1  'True
      Height          =   255
      Left            =   3120
      TabIndex        =   1
      Top             =   6540
      Width           =   1200
   End
   Begin VB.CheckBox chkBackup 
      Caption         =   "Make backup"
      Height          =   255
      Left            =   30
      TabIndex        =   0
      Top             =   6540
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog cdlgFile 
      Left            =   30
      Top             =   6900
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'__________________________________________________
' Author : Jean-Philippe Leconte
' File : frmMain.frm
' Date : 11 october 2000 03:48
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private Type BROWSEINFO
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Const BACKUP_EXT = ".backup"
Private Const PROCSCOPE = "%s"
Private Const PROCTYPE = "%t"
Private Const PROCNAME = "%n"
Private Const PROCPARAM = "%p"
Private Const PROCRETURN = "%r"
Private Const PROCDESC = "%d"

Private Const LANG_FRA = "Français"
Private Const LANG_ENG = "English"
Private Const LANG_ESP = "Español"
Private Const LANG_DEU = "Deutsch"

Private vntScopes As Variant
Private vntMids As Variant
Private vntProcedures As Variant
Private vntProcedureEnds As Variant
Private vntEnds As Variant

Private lModProcedures As Long
Private lProcedures As Long

'__________________________________________________
' Scope : Private
' Type : Sub
' Name : cmdLoadDir_Click
' Parameters :
' Returns : Nothing
' Description : The Sub uses parameters for cmdLoadDir_Click and returns Nothing.
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub cmdLoadDir_Click()
    On Error GoTo ErrorHandle_cmdLoadDir_Click
    Dim sFolder As String
    Dim sFilename As String
    Dim sMsgBox As String
    Dim sFiles As String
    
    MsgBox "By precaution, this option will always make backups of files modified", vbInformation, "Commentor"
    
    sFolder = BrowseFolder
    If Not Right(sFolder, 1) = "\" Then sFolder = sFolder + "\"
    
    sFilename = Dir(sFolder + "*.*")
    While Len(sFilename) > 0
        If Right(sFilename, 4) = ".bas" Or Right(sFilename, 4) = ".frm" Or Right(sFilename, 4) = ".ctl" Or Right(sFilename, 4) = ".cls" Then sFiles = sFiles + ModifyFile(sFolder + sFilename, CBool(chkComments.Value), CBool(chkErrorHandling.Value), True, True) + vbCrLf
        sFilename = Dir
    Wend
    
    If Len(sFiles) > 0 Then
        sMsgBox = IIf(CBool(chkComments.Value), "comments", "")
        sMsgBox = sMsgBox + IIf(Len(sMsgBox) > 0 And CBool(chkErrorHandling.Value), " and ", "") + IIf(CBool(chkErrorHandling.Value), "error handling", "")
        MsgBox "Finished adding " + IIf(Len(sMsgBox) > 0, sMsgBox, "nothing") + " to " + vbCrLf + vbCrLf + Right(sFiles, Len(sFiles) - Len(vbCrLf)) + vbCrLf + vbCrLf + "With a total of " + CStr(lModProcedures) + " procedures modified on a great total of " + CStr(lProcedures), vbApplicationModal, "Commentor"
    End If

ErrorHandle_cmdLoadDir_Click:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Sub"
    sErrorName = "cmdLoadDir_Click"
    sErrorReturns = "Nothing"
  
    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical, App.Title
          Err.Clear
          Resume Next
    End Select
End Sub

'__________________________________________________
' Scope : Private
' Type : Sub
' Name : cmdLoadFile_Click
' Parameters :
' Returns : Nothing
' Description : The Sub uses parameters for cmdLoadFile_Click and returns Nothing.
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub cmdLoadFile_Click()
    On Error GoTo ErrorHandle_cmdLoadFile_Click
    
    cdlgFile.DialogTitle = "Visual Basic file to load"
    cdlgFile.DefaultExt = ".*"
    cdlgFile.Filter = "All files (*.*)|*.*|Forms (*.frm)|*.frm|Modules (*.bas)|*.bas|Classes (*.cls)|*.cls|User controls (*.ctl)|*.ctl"
    cdlgFile.CancelError = True
    cdlgFile.Flags = cdlOFNHideReadOnly
    cdlgFile.ShowOpen
    If Len(cdlgFile.FileName) > 0 And Len(Dir(cdlgFile.FileName)) > 0 Then
        ModifyFile cdlgFile.FileName, CBool(chkComments.Value), CBool(chkErrorHandling.Value), CBool(chkBackup.Value)
    Else
        MsgBox "Cannot find file", vbCritical, "Commentor"
    End If

ErrorHandle_cmdLoadFile_Click:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Sub"
    sErrorName = "cmdLoadFile_Click"
    sErrorReturns = "Nothing"
  
    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case 32755
          'Cancel was pressed
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical Or vbApplicationModal, App.Title
          Err.Clear
          Resume Next
    End Select
End Sub

'__________________________________________________
' Scope : Private
' Type : Sub
' Name : cmdQuit_Click
' Parameters :
' Returns : Nothing
' Description : The Sub uses parameters for cmdQuit_Click and returns Nothing.
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub cmdQuit_Click()
    On Error GoTo ErrorHandle_cmdQuit_Click
    Unload Me

ErrorHandle_cmdQuit_Click:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Sub"
    sErrorName = "cmdQuit_Click"
    sErrorReturns = "Nothing"
  
    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical, App.Title
          Err.Clear
          Resume Next
    End Select
End Sub

'__________________________________________________
' Scope : Private
' Type : Function
' Name : ModifyFile
' Parameters :
'         ByVal sFilename As String
'         ByVal bComments As Boolean
'         ByVal bErrorHandling As Boolean
'         ByVal bMakeBackup As Boolean
'         Optional bDoNotDisplay As Boolean = False
' Returns : String
' Description : The Function uses parameters ByVal sFilename As String, ByVal bComments As Boolean, ByVal bErrorHandling As Boolean, ByVal bMakeBackup As Boolean and Optional bDoNotDisplay As Boolean = False for ModifyFile and returns String.
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Function ModifyFile(ByVal sFilename As String, ByVal bComments As Boolean, ByVal bErrorHandling As Boolean, ByVal bMakeBackup As Boolean, Optional bDoNotDisplay As Boolean = False) As String
    On Error GoTo ErrorHandle_ModifyFile
    Dim sScope As String
    Dim sMid As String
    Dim sType As String
    Dim sName As String
    Dim vntParameters As Variant
    Dim sReturn As String
    Dim sDescription As String
    Dim sEnd As String
    
    Dim bStartErrorHandling As Boolean
    Dim sFile As String
    Dim vntFile As Variant
    Dim iOpen As Integer
    Dim lCount As Long
    Dim lLBound As Long
    Dim lUBound As Long
    Dim lChar As Long
    Dim lPos As Long
    
    Dim sMsgBox As String
    Dim bUp As Long
    
    vntScopes = Array("Private", "Public", "Global", "Friend", "Protected")
    vntMids = Array("Static")
    vntProcedures = Array("Function", "Sub", "Property Get", "Property Let", "Property Set")
    vntEnds = Array("End")
    vntProcedureEnds = Array("Function", "Sub", "Property")
        
    If bMakeBackup Then FileCopy sFilename, sFilename + BACKUP_EXT
    
    iOpen = FreeFile(1)
    Open sFilename For Input As iOpen
        sFile = Input(LOF(iOpen), iOpen)
    Close iOpen
    
    vntFile = Split(sFile, vbCrLf)
    lChar = 1
    
    lLBound = LBound(vntFile)
    lUBound = UBound(vntFile)
    For lCount = lLBound To lUBound
        lPos = 1
        sScope = ""
        sMid = ""
        sType = ""
        sName = ""
        sReturn = ""
        vntParameters = Null
        sDescription = ""
        sEnd = ""
        bUp = False
        
        sScope = CheckScope(vntFile(lCount), lPos)
        lPos = lPos + IIf(Len(sScope) = 0, 0, Len(sScope) + 1)
        sMid = CheckMid(vntFile(lCount), lPos)
        lPos = lPos + IIf(Len(sMid) = 0, 0, Len(sMid) + 1)
        sType = CheckProcedure(vntFile(lCount), lPos)
        If Len(sType) > 0 Then
            lProcedures = lProcedures + 1
            lPos = lPos + Len(sType) + 1
            sName = GetName(vntFile(lCount), lPos)
            If Len(sName) > 0 Then
                lPos = lPos + Len(sName) + 1
                sReturn = GetReturn(vntFile(lCount))
                If bComments Then
                    vntParameters = GetParams(vntFile(lCount), lPos)
                    sDescription = MakeDescription(txtDescription.Text, sScope + IIf(Len(sScope) > 0 And Len(sMid) > 0, " ", "") + sMid, PROCSCOPE, sType, PROCTYPE, sName, PROCNAME, vntParameters, PROCPARAM, sReturn, PROCRETURN)
                    lChar = lChar + AddComments(sFile, txtComments.Text, lChar, sScope + IIf(Len(sScope) > 0 And Len(sMid) > 0, " ", "") + sMid, PROCSCOPE, sType, PROCTYPE, sName, PROCNAME, vntParameters, PROCPARAM, sReturn, PROCRETURN, sDescription, PROCDESC)
                    lModProcedures = lModProcedures + 1
                    bUp = True
                End If
                If bErrorHandling Then
                    bStartErrorHandling = True
                    sLastScope = sScope
                    sLastMid = sMid
                    sLastType = sType
                    sLastName = sName
                    sLastReturn = sReturn
                    lChar = lChar + AddErrorHandlingTop(sFile, txtErrorHandlingTop.Text, lChar + Len(vntFile(lCount)) + Len(vbCrLf), sScope + IIf(Len(sScope) > 0 And Len(sMid) > 0, " ", "") + sMid, PROCSCOPE, sType, PROCTYPE, sName, PROCNAME, sReturn, PROCRETURN)
                    If Not bUp Then lModProcedures = lModProcedures + 1
                End If
            End If
        End If
        If bErrorHandling And bStartErrorHandling Then
            lPos = 1
            sEnd = CheckEnd(vntFile(lCount))
            lPos = lPos + IIf(Len(sEnd) = 0, 0, Len(sEnd) + 1)
            If Len(sEnd) > 0 And Len(CheckProcedureEnd(vntFile(lCount), lPos)) > 0 Then
                lChar = lChar + AddErrorHandling(sFile, txtErrorHandling.Text, lChar, sLastScope + IIf(Len(sLastScope) > 0 And Len(sLastMid) > 0, " ", "") + sLastMid, PROCSCOPE, sLastType, PROCTYPE, sLastName, PROCNAME, sLastReturn, PROCRETURN)
                bStartErrorHandling = False
                sLastScope = ""
                sLastMid = ""
                sLastType = ""
                sLastName = ""
                sLastReturn = ""
            End If
        End If
        lChar = lChar + Len(vntFile(lCount)) + Len(vbCrLf)
    Next lCount
    
    iOpen = FreeFile(1)
    Open sFilename For Output As iOpen
        Print #iOpen, sFile
    Close iOpen
    
    If Not bDoNotDisplay Then
        sMsgBox = IIf(bComments, "comments", "")
        sMsgBox = sMsgBox + IIf(Len(sMsgBox) > 0 And bErrorHandling, " and ", "") + IIf(bErrorHandling, "error handling", "")
        MsgBox "Finished adding " + IIf(Len(sMsgBox) > 0, sMsgBox, "nothing") + " to " + sFilename + vbCrLf + vbCrLf + "With a total of " + CStr(lModProcedures) + " procedures modified on a great total of " + CStr(lProcedures), vbApplicationModal, "Commentor"
    End If
    
    ModifyFile = sFilename

ErrorHandle_ModifyFile:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Function"
    sErrorName = "ModifyFile"
    sErrorReturns = "String"
  
    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical, App.Title
          Err.Clear
          Resume Next
    End Select
End Function

'__________________________________________________
' Scope : Private
' Type : Function
' Name : CheckScope
' Parameters :
'         ByVal sLine As String
'         ByVal lPos As Long
' Returns : String
' Description : The Function uses parameters ByVal sLine As String and ByVal lPos As Long for CheckScope and returns String.
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Function CheckScope(ByVal sLine As String, ByVal lPos As Long) As String
    On Error GoTo ErrorHandle_CheckScope
    Dim sFound As String
    Dim lCount As Long
    
    While Not lCount > UBound(vntScopes)
        If InStr(lPos, sLine, vntScopes(lCount)) = lPos Then sFound = vntScopes(lCount)
        lCount = lCount + 1
    Wend
    
    CheckScope = sFound

ErrorHandle_CheckScope:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Function"
    sErrorName = "CheckScope"
    sErrorReturns = "String"
  
    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical, App.Title
          Err.Clear
          Resume Next
    End Select
End Function

'__________________________________________________
' Scope : Private
' Type : Function
' Name : CheckMid
' Parameters :
'         ByVal sLine As String
'         ByVal lPos As Long
' Returns : String
' Description : The Function uses parameters ByVal sLine As String and ByVal lPos As Long for CheckMid and returns String.
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Function CheckMid(ByVal sLine As String, ByVal lPos As Long) As String
    On Error GoTo ErrorHandle_CheckMid
    Dim sFound As String
    Dim lCount As Long
    
    While Not lCount > UBound(vntMids)
        If InStr(lPos, sLine, vntMids(lCount)) = lPos Then sFound = vntMids(lCount)
        lCount = lCount + 1
    Wend
    
    CheckMid = sFound

ErrorHandle_CheckMid:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Function"
    sErrorName = "CheckMid"
    sErrorReturns = "String"
  
    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical, App.Title
          Err.Clear
          Resume Next
    End Select
End Function

'__________________________________________________
' Scope : Private
' Type : Function
' Name : CheckProcedure
' Parameters :
'         ByVal sLine As String
'         ByVal lPos As Long
' Returns : String
' Description : The Function uses parameters ByVal sLine As String and ByVal lPos As Long for CheckProcedure and returns String.
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Function CheckProcedure(ByVal sLine As String, ByVal lPos As Long) As String
    On Error GoTo ErrorHandle_CheckProcedure
    Dim sFound As String
    Dim lCount As Long
    
    While Not lCount > UBound(vntProcedures)
        If InStr(lPos, sLine, vntProcedures(lCount)) = lPos Then sFound = vntProcedures(lCount)
        lCount = lCount + 1
    Wend
    
    CheckProcedure = sFound

ErrorHandle_CheckProcedure:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Function"
    sErrorName = "CheckProcedure"
    sErrorReturns = "String"
  
    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical, App.Title
          Err.Clear
          Resume Next
    End Select
End Function

'__________________________________________________
' Scope : Private
' Type : Function
' Name : CheckEnd
' Parameters :
'         ByVal sLine As String
' Returns : String
' Description : The Function uses parameters ByVal sLine As String for CheckEnd and returns String.
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Function CheckEnd(ByVal sLine As String) As String
    On Error GoTo ErrorHandle_CheckEnd
    Dim sFound As String
    Dim lCount As Long
    
    While Not lCount > UBound(vntEnds)
        If Left(sLine, Len(vntEnds(lCount))) = vntEnds(lCount) Then sFound = vntEnds(lCount)
        lCount = lCount + 1
    Wend
    
    CheckEnd = sFound

ErrorHandle_CheckEnd:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Function"
    sErrorName = "CheckEnd"
    sErrorReturns = "String"
  
    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical, App.Title
          Err.Clear
          Resume Next
    End Select
End Function

'__________________________________________________
' Scope : Private
' Type : Function
' Name : CheckProcedureEnd
' Parameters :
'         ByVal sLine As String
'         ByVal lPos As Long
' Returns : String
' Description : The Function uses parameters ByVal sLine As String and ByVal lPos As Long for CheckProcedureEnd and returns String.
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Function CheckProcedureEnd(ByVal sLine As String, ByVal lPos As Long) As String
    On Error GoTo ErrorHandle_CheckProcedureEnd
    Dim sFound As String
    Dim lCount As Long
    
    While Not lCount > UBound(vntProcedureEnds)
        If InStr(lPos, sLine, vntProcedureEnds(lCount)) = lPos Then sFound = vntProcedureEnds(lCount)
        lCount = lCount + 1
    Wend
    
    CheckProcedureEnd = sFound

ErrorHandle_CheckProcedureEnd:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Function"
    sErrorName = "CheckProcedureEnd"
    sErrorReturns = "String"
  
    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical, App.Title
          Err.Clear
          Resume Next
    End Select
End Function

'__________________________________________________
' Scope : Private
' Type : Function
' Name : GetName
' Parameters :
'         ByVal sLine As String
'         ByVal lPos As Long
' Returns : String
' Description : The Function uses parameters ByVal sLine As String and ByVal lPos As Long for GetName and returns String.
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Function GetName(ByVal sLine As String, ByVal lPos As Long) As String
    On Error GoTo ErrorHandle_GetName
    Dim sFound As String
    Dim lEnd As Long
    
    lEnd = InStr(lPos, sLine, "(")
    If lEnd > 0 Then sFound = Mid(sLine, lPos, lEnd - lPos)
    
    GetName = sFound

ErrorHandle_GetName:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Function"
    sErrorName = "GetName"
    sErrorReturns = "String"
  
    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical, App.Title
          Err.Clear
          Resume Next
    End Select
End Function

'__________________________________________________
' Scope : Private
' Type : Function
' Name : GetParams
' Parameters :
'         ByVal sLine As String
'         ByVal lPos As Long
' Returns : Variant
' Description : The Function uses parameters ByVal sLine As String and ByVal lPos As Long for GetParams and returns Variant.
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Function GetParams(ByVal sLine As String, ByVal lPos As Long) As Variant
    On Error GoTo ErrorHandle_GetParams
    Dim sFound As String
    Dim lEnd As Long
    
    lEnd = InStrRev(sLine, ")")
    If lEnd > lPos Then sFound = Mid(sLine, lPos, lEnd - lPos)
    
    GetParams = Split(sFound, ", ")

ErrorHandle_GetParams:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Function"
    sErrorName = "GetParams"
    sErrorReturns = "Variant"
  
    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical, App.Title
          Err.Clear
          Resume Next
    End Select
End Function

'__________________________________________________
' Scope : Private
' Type : Function
' Name : GetReturn
' Parameters :
'         ByVal sLine As String
' Returns : String
' Description : The Function uses parameters ByVal sLine As String for GetReturn and returns String.
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Function GetReturn(ByVal sLine As String) As String
    On Error GoTo ErrorHandle_GetReturn
    GetReturn = IIf(Right(Trim(Mid(sLine, InStrRev(sLine, " ") + 1)), 1) = ")", "Nothing", Trim(Mid(sLine, InStrRev(sLine, " ") + 1)))

ErrorHandle_GetReturn:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Function"
    sErrorName = "GetReturn"
    sErrorReturns = "String"
  
    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical, App.Title
          Err.Clear
          Resume Next
    End Select
End Function

'__________________________________________________
' Scope : Private
' Type : Function
' Name : MakeDescription
' Parameters :
'         ByVal sTemplate As String
'         ByVal sScope As String
'         ByVal sScopeParam As String
'         ByVal sType As String
'         ByVal sTypeParam As String
'         ByVal sName As String
'         ByVal sNameParam As String
'         ByVal vntParameters As Variant
'         ByVal sParametersParam As String
'         ByVal sReturn As String
'         ByVal sReturnParam As String
' Returns : String
' Description : The Function uses parameters ByVal sTemplate As String, ByVal sScope As String, ByVal sScopeParam As String, ByVal sType As String, ByVal sTypeParam As String, ByVal sName As String, ByVal sNameParam As String, ByVal vntParameters As Variant, ByVal sParametersParam As String, ByVal sReturn As String and ByVal sReturnParam As String for MakeDescription and returns String.
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Function MakeDescription(ByVal sTemplate As String, ByVal sScope As String, ByVal sScopeParam As String, ByVal sType As String, ByVal sTypeParam As String, ByVal sName As String, ByVal sNameParam As String, ByVal vntParameters As Variant, ByVal sParametersParam As String, ByVal sReturn As String, ByVal sReturnParam As String) As String
    On Error GoTo ErrorHandle_MakeDescription
    Dim sResult As String
    Dim sParams As String
    Dim sAnd As String
    
    sResult = sTemplate
    
    sResult = Replace(sResult, sScopeParam, sScope)
    sResult = Replace(sResult, sTypeParam, sType)
    sResult = Replace(sResult, sNameParam, sName)
    For lCount = LBound(vntParameters) To UBound(vntParameters) - 1
        sParams = sParams + vntParameters(lCount) + ", "
    Next lCount
    sAnd = " and "
    If cboLang.Text = LANG_ENG Then sAnd = " and "
    If cboLang.Text = LANG_FRA Then sAnd = " et "
    If cboLang.Text = LANG_ESP Then sAnd = " e "
    If cboLang.Text = LANG_DEU Then sAnd = " und "
    If Not UBound(vntParameters) = -1 Then If UBound(vntParameters) > LBound(vntParameters) Then sParams = Left(sParams, Len(sParams) - 2) + sAnd + vntParameters(UBound(vntParameters)) Else sParams = vntParameters(LBound(vntParameters))
    sResult = Replace(sResult, sParametersParam, sParams)
    sResult = Replace(sResult, sReturnParam, sReturn)
    
    MakeDescription = sResult

ErrorHandle_MakeDescription:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Function"
    sErrorName = "MakeDescription"
    sErrorReturns = "String"
  
    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical, App.Title
          Err.Clear
          Resume Next
    End Select
End Function

'__________________________________________________
' Scope : Private
' Type : Function
' Name : AddComments
' Parameters :
'         ByRef sFile As String
'         ByVal sTemplate As String
'         ByVal lPos As Long
'         ByVal sScope As String
'         ByVal sScopeParam As String
'         ByVal sType As String
'         ByVal sTypeParam As String
'         ByVal sName As String
'         ByVal sNameParam As String
'         ByVal vntParameters As Variant
'         ByVal sParametersParam As String
'         ByVal sReturn As String
'         ByVal sReturnParam As String
'         ByVal sDescription As String
'         ByVal sDescriptionParam As String
' Returns : Long
' Description : The Function uses parameters ByRef sFile As String, ByVal sTemplate As String, ByVal lPos As Long, ByVal sScope As String, ByVal sScopeParam As String, ByVal sType As String, ByVal sTypeParam As String, ByVal sName As String, ByVal sNameParam As String, ByVal vntParameters As Variant, ByVal sParametersParam As String, ByVal sReturn As String, ByVal sReturnParam As String, ByVal sDescription As String and ByVal sDescriptionParam As String for AddComments and returns Long.
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Function AddComments(ByRef sFile As String, ByVal sTemplate As String, ByVal lPos As Long, ByVal sScope As String, ByVal sScopeParam As String, ByVal sType As String, ByVal sTypeParam As String, ByVal sName As String, ByVal sNameParam As String, ByVal vntParameters As Variant, ByVal sParametersParam As String, ByVal sReturn As String, ByVal sReturnParam As String, ByVal sDescription As String, ByVal sDescriptionParam As String) As Long
    On Error GoTo ErrorHandle_AddComments
    Dim lCount As Long
    Dim sResult As String
    Dim sParamLine As String
    Dim sParams As String
    Dim lParamPos As Long
    Dim lLastCrLf As Long
    Dim lNextCrLf As Long
    
    sResult = sTemplate
    
    sResult = Replace(sResult, sScopeParam, sScope)
    sResult = Replace(sResult, sTypeParam, sType)
    sResult = Replace(sResult, sNameParam, sName)
    lParamPos = InStr(sResult, sParametersParam)
    If lParamPos > 0 Then
        lLastCrLf = InStrRev(sResult, vbCrLf, lParamPos)
        lNextCrLf = InStr(lParamPos, sResult, vbCrLf)
        If lNextCrLf - lLastCrLf > Len(sParametersParam) Then
            sParamLine = Mid(sResult, IIf(lLastCrLf > 0, lLastCrLf, 1), lNextCrLf - IIf(lLastCrLf > 0, lLastCrLf, 1))
            For lCount = LBound(vntParameters) To UBound(vntParameters)
                sParams = sParams + Replace(sParamLine, sParametersParam, vntParameters(lCount))
            Next lCount
            sResult = Replace(sResult, sParamLine, sParams)
        Else
            For lCount = LBound(vntParameters) To UBound(vntParameters) - 1
                sParams = sParams + vntParameters(lCount) + ", "
            Next lCount
            sAnd = " and "
            If cboLang.Text = LANG_ENG Then sAnd = " and "
            If cboLang.Text = LANG_FRA Then sAnd = " et "
            If cboLang.Text = LANG_ESP Then sAnd = " e "
            If cboLang.Text = LANG_DEU Then sAnd = " und "
            If Not UBound(vntParameters) = -1 Then If UBound(vntParameters) > LBound(vntParameters) Then sParams = Left(sParams, Len(sParams) - 2) + sAnd + vntParameters(UBound(vntParameters)) Else sParams = vntParameters(LBound(vntParameters))
            sResult = Replace(sResult, sParametersParam, sParams)
        End If
    End If
    sResult = Replace(sResult, sReturnParam, sReturn)
    sResult = Replace(sResult, sDescriptionParam, sDescription)
    
    sFile = Left(sFile, lPos - 1) + sResult + Mid(sFile, lPos)
    
    AddComments = Len(sResult)

ErrorHandle_AddComments:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Function"
    sErrorName = "AddComments"
    sErrorReturns = "Long"
  
    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical, App.Title
          Err.Clear
          Resume Next
    End Select
End Function

'__________________________________________________
' Scope : Private
' Type : Function
' Name : AddErrorHandlingTop
' Parameters :
'         ByRef sFile As String
'         ByVal sTemplate As String
'         ByVal lPos As String
'         ByVal sScope As String
'         ByVal sScopeParam As String
'         ByVal sType As String
'         ByVal sTypeParam As String
'         ByVal sName As String
'         ByVal sNameParam As String
'         ByVal sReturn As String
'         ByVal sReturnParam As String
' Returns : Long
' Description : The Function uses parameters ByRef sFile As String, ByVal sTemplate As String, ByVal lPos As String, ByVal sScope As String, ByVal sScopeParam As String, ByVal sType As String, ByVal sTypeParam As String, ByVal sName As String, ByVal sNameParam As String, ByVal sReturn As String and ByVal sReturnParam As String for AddErrorHandlingTop and returns Long.
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Function AddErrorHandlingTop(ByRef sFile As String, ByVal sTemplate As String, ByVal lPos As String, ByVal sScope As String, ByVal sScopeParam As String, ByVal sType As String, ByVal sTypeParam As String, ByVal sName As String, ByVal sNameParam As String, ByVal sReturn As String, ByVal sReturnParam As String) As Long
    On Error GoTo ErrorHandle_AddErrorHandlingTop
    Dim sResult As String
    
    sResult = sTemplate
    
    sResult = Replace(sResult, sScopeParam, sScope)
    sResult = Replace(sResult, sTypeParam, sType)
    sResult = Replace(sResult, sNameParam, sName)
    sResult = Replace(sResult, sReturnParam, sReturn)
    
    sFile = Left(sFile, lPos - 1) + sResult + Mid(sFile, lPos)
    
    AddErrorHandlingTop = Len(sResult)

ErrorHandle_AddErrorHandlingTop:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Function"
    sErrorName = "AddErrorHandlingTop"
    sErrorReturns = "Long"
  
    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical, App.Title
          Err.Clear
          Resume Next
    End Select
End Function

'__________________________________________________
' Scope : Private
' Type : Function
' Name : AddErrorHandling
' Parameters :
'         ByRef sFile As String
'         ByVal sTemplate As String
'         ByVal lPos As String
'         ByVal sScope As String
'         ByVal sScopeParam As String
'         ByVal sType As String
'         ByVal sTypeParam As String
'         ByVal sName As String
'         ByVal sNameParam As String
'         ByVal sReturn As String
'         ByVal sReturnParam As String
' Returns : Long
' Description : The Function uses parameters ByRef sFile As String, ByVal sTemplate As String, ByVal lPos As String, ByVal sScope As String, ByVal sScopeParam As String, ByVal sType As String, ByVal sTypeParam As String, ByVal sName As String, ByVal sNameParam As String, ByVal sReturn As String and ByVal sReturnParam As String for AddErrorHandling and returns Long.
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Function AddErrorHandling(ByRef sFile As String, ByVal sTemplate As String, ByVal lPos As String, ByVal sScope As String, ByVal sScopeParam As String, ByVal sType As String, ByVal sTypeParam As String, ByVal sName As String, ByVal sNameParam As String, ByVal sReturn As String, ByVal sReturnParam As String) As Long
    On Error GoTo ErrorHandle_AddErrorHandling
    Dim sResult As String
    
    sResult = sTemplate
    
    sResult = Replace(sResult, sScopeParam, sScope)
    sResult = Replace(sResult, sTypeParam, sType)
    sResult = Replace(sResult, sNameParam, sName)
    sResult = Replace(sResult, sReturnParam, sReturn)
    
    sFile = Left(sFile, lPos - 1) + sResult + Mid(sFile, lPos)
    
    AddErrorHandling = Len(sResult)

ErrorHandle_AddErrorHandling:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Function"
    sErrorName = "AddErrorHandling"
    sErrorReturns = "Long"
  
    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical, App.Title
          Err.Clear
          Resume Next
    End Select
End Function

'__________________________________________________
' Scope : Private
' Type : Sub
' Name : Form_Load
' Parameters :
' Returns : Nothing
' Description : The Sub uses parameters for Form_Load and returns Nothing.
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub Form_Load()
    On Error GoTo ErrorHandle_Form_Load
    cboLang.ListIndex = 0

ErrorHandle_Form_Load:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Sub"
    sErrorName = "Form_Load"
    sErrorReturns = "Nothing"
  
    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical, App.Title
          Err.Clear
          Resume Next
    End Select
End Sub

'__________________________________________________
' Scope : Private
' Type : Sub
' Name : Form_Unload
' Parameters :
'         Cancel As Integer
' Returns : Nothing
' Description : The Sub uses parameters Cancel As Integer for Form_Unload and returns Nothing.
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrorHandle_Form_Unload
    MsgBox "Commentor is made by Jean-Philippe Leconte." + vbCrLf + "Source code for this program is available and is released under the GNU Public License." + vbCrLf + "Please send any change to the program to insomniaque@mail.com." + vbCrLf + vbCrLf + "N.B. Please give him credits for his work, because... :)", vbApplicationModal, "About Commentor"

ErrorHandle_Form_Unload:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Sub"
    sErrorName = "Form_Unload"
    sErrorReturns = "Nothing"
  
    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical, App.Title
          Err.Clear
          Resume Next
    End Select
End Sub

'__________________________________________________
' Scope : Private
' Type : Function
' Name : BrowseFolder
' Parameters :
' Returns : String
' Description : The Function uses parameters for BrowseFolder and returns String.
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Private Function BrowseFolder() As String
    On Error GoTo ErrorHandle_BrowseFolder
    Dim lIDList As Long
    Dim sPath As String
    Dim uBrowse As BROWSEINFO
    
    uBrowse.hWndOwner = Me.hWnd
    uBrowse.lpszTitle = StrPtr("Choose folder" + vbNullChar)
    uBrowse.ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    lIDList = SHBrowseForFolder(uBrowse)
    If lIDList Then
        sPath = Space(MAX_PATH)
        SHGetPathFromIDList lIDList, sPath
        sPath = Left(sPath, InStr(sPath, vbNullChar) - 1)
    End If
    
    BrowseFolder = sPath

ErrorHandle_BrowseFolder:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Function"
    sErrorName = "BrowseFolder"
    sErrorReturns = "String"
  
    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical, App.Title
          Err.Clear
          Resume Next
    End Select
End Function






