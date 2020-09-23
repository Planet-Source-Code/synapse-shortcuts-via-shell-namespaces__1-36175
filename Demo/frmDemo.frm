VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDemo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shell Shortcuts Demo"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5865
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDemo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   5865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create Shortcut"
      Height          =   375
      Left            =   4185
      TabIndex        =   12
      Top             =   2085
      Width           =   1515
   End
   Begin VB.Frame Frame1 
      Caption         =   "Location:"
      Height          =   690
      Left            =   135
      TabIndex        =   8
      Top             =   1320
      Width           =   5580
      Begin VB.OptionButton optLocation 
         Appearance      =   0  'Flat
         Caption         =   "Network Neighbourhood"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   3105
         TabIndex        =   11
         Top             =   315
         Width           =   2175
      End
      Begin VB.OptionButton optLocation 
         Appearance      =   0  'Flat
         Caption         =   "My Computer"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   1470
         TabIndex        =   10
         Top             =   300
         Width           =   1395
      End
      Begin VB.OptionButton optLocation 
         Appearance      =   0  'Flat
         Caption         =   "Desktop"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   285
         Value           =   -1  'True
         Width           =   1050
      End
   End
   Begin VB.CommandButton cmdIconResource 
      Caption         =   "...."
      Height          =   315
      Left            =   5235
      TabIndex        =   7
      Top             =   525
      Width           =   420
   End
   Begin VB.TextBox txtIconResource 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1185
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   555
      Width           =   3960
   End
   Begin VB.TextBox txtAppName 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   1455
      TabIndex        =   3
      Top             =   930
      Width           =   3180
   End
   Begin VB.CommandButton cmdBrowseFile 
      Caption         =   "...."
      Height          =   315
      Left            =   4995
      TabIndex        =   2
      Top             =   165
      Width           =   420
   End
   Begin VB.TextBox txtFilename 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   855
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   180
      Width           =   4065
   End
   Begin MSComDlg.CommonDialog cdl1 
      Left            =   165
      Top             =   2025
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select executabale file"
      Filter          =   "EXEcutable files (*.exe) | *.exe"
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Icon resource:"
      Height          =   195
      Left            =   75
      TabIndex        =   5
      Top             =   570
      Width           =   1050
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Application name:"
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   960
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Filename:"
      Height          =   270
      Left            =   60
      TabIndex        =   1
      Top             =   195
      Width           =   705
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'/*******************************************************
'*  Shell Shortcuts Demo                                *
'*  Written by Jahufar Sadique [AKA sYNAPSE]            *
'*  Last updated on June 22nd, 2002                     *
'*******************************************************/

'/*************************************************
'* See 'Shortcuts Via Shell Namespaces.doc' for   *
'* detailed explanation on how this works.        *
'*                                                *
'* IMPORTANT: There's no code to undo any changes *
'* made to to registry. The program will, however *
'* log all its actions to a file:                 *
'*      app.path + "\ShellLog.txt".               *
'*                                                *
'* You may have to refer to this file to undo     *
'* any changes manually.                          *
'*                                                *
'* NOTE: CreateGUID() function was written by     *
'* a guy called Dion Wiggins.                     *
'*************************************************/

Option Explicit

Private Type GUID
     Data1 As Long
     Data2 As Long
     Data3 As Long
     Data4(8) As Byte
End Type

Private Enum hKey
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
End Enum

Private Enum dwType
    REG_SZ = 1
    REG_DWORD = 4
End Enum

'//registry API constants:
Private Const KEY_ALL_ACCESS = &H3F
Private Const REG_OPTION_NON_VOLATILE = 0
Private Const ERROR_SUCCESS = 0&
'//user defined NameSpace constants:
Private Const NS_DESKTOP = "software\microsoft\windows\currentversion\explorer\desktop\namespace\"
Private Const NS_MYCOMPUTER = "software\microsoft\windows\currentversion\explorer\mycomputer\namespace\"
Private Const NS_NETHOOD = "software\microsoft\windows\currentversion\explorer\networkneighborhood\namespace\"
'//user defined app errorbase:
Private Const APP_ERROR_BASE = 1981

Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, PhkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegDeleteValue& Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String)
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, PhkResult As Long, lpdwDisposition As Long) As Long

Private Declare Function CoCreateGuid Lib "ole32.dll" (pguid As GUID) As Long
Private Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As Any, ByVal lpstrClsId As Long, ByVal cbMax As Long) As Long

'Author:    Dion Wiggins
'Purpose:   Creates a GUID
'Notes:
'Inputs:
'   - strRemoveChars    The characters to remove from the GUID (usually the {}- characters)
'History
'Date           Author          Description
'1 Jun 1999     Dion Wiggins    Created
Public Function CreateGUID(Optional strRemoveChars As String = "{}-") As String
    Dim udtGUID As GUID
    Dim strGUID As String
    Dim bytGUID() As Byte
    Dim lngLen As Long
    Dim lngRetVal As Long
    Dim lngPos As Long
    
    'Initialize
    lngLen = 40
    bytGUID = String(lngLen, 0)
    
    'Create the GUID
    CoCreateGuid udtGUID
    
    'Convert the structure into a displayable string
    lngRetVal = StringFromGUID2(udtGUID, VarPtr(bytGUID(0)), lngLen)
    strGUID = bytGUID
    If (Asc(Mid$(strGUID, lngRetVal, 1)) = 0) Then
        lngRetVal = lngRetVal - 1
    End If
    
    'Trim the trailing characters
    strGUID = Left$(strGUID, lngRetVal)
    
    'Remove the unwanted characters
    For lngPos = 1 To Len(strRemoveChars)
        strGUID = Replace(strGUID, Mid(strRemoveChars, lngPos, 1), "")
    Next
    
    CreateGUID = strGUID
End Function

Private Sub cmdIconResource_Click()
    With cdl1
        .FileName = ""
        .Filter = "Icon files (*.ico) | *.ico"
        .DialogTitle = "Select icon.."
        .ShowOpen
        If .FileName = "" Then Exit Sub
        txtIconResource = .FileName
    End With
End Sub

Private Sub cmdBrowseFile_Click()
    
    With cdl1
        .FileName = ""
        .Filter = "Executable files (*.exe) | *.exe"
        .DialogTitle = "Select file.."
        .ShowOpen
        If .FileName = "" Then Exit Sub
        txtFilename = .FileName
    End With
     
End Sub

Private Sub cmdCreate_Click()
    On Error GoTo whoops
    
    Dim strNameSpace As String
    
    If txtFilename = "" Or txtIconResource = "" Or txtAppName = "" Then
        MsgBox "Not enough information to create shortcut!", vbExclamation, "Error"
        Exit Sub
    End If
    
    Dim strLog As String
    Dim strCLSID As String
    
    strLog = strLog + "DateTime: " & Now & vbCrLf
    strLog = strLog + "Application name: " & txtAppName & vbCrLf
    strLog = strLog + "Application path: " & txtFilename & vbCrLf
    strLog = strLog + "Icon resource: " & txtIconResource & vbCrLf
    
    If optLocation(0).Value Then
        strLog = strLog + "Location: Desktop" & vbCrLf
        strNameSpace = NS_DESKTOP
    End If
    
    If optLocation(1).Value Then
        strLog = strLog + "Location: My Computer" & vbCrLf
        strNameSpace = NS_MYCOMPUTER
    End If
    
    If optLocation(2).Value Then
        strLog = strLog + "Location: Network Neighbourhood" & vbCrLf
        strNameSpace = NS_NETHOOD
    End If
    
    strCLSID = CreateGUID("")
    strLog = strLog + "CLSID: " & strCLSID & vbCrLf
    
    '//step 2:
    If CreateNewKey(HKEY_CLASSES_ROOT, "CLSID\" + strCLSID) <> ERROR_SUCCESS Then Err.Raise APP_ERROR_BASE + 1, , "Could not create ClassID key"
    If SetKeyValue(HKEY_CLASSES_ROOT, "CLSID\" + strCLSID, "", txtAppName, REG_SZ) <> ERROR_SUCCESS Then Err.Raise APP_ERROR_BASE + 2, , "Could not set AppName"
    '//step 3:
    If CreateNewKey(HKEY_CLASSES_ROOT, "CLSID\" + strCLSID + "\DefaultIcon") <> ERROR_SUCCESS Then Err.Raise APP_ERROR_BASE + 3, , "Could not create DefaultIcon key"
    If SetKeyValue(HKEY_CLASSES_ROOT, "CLSID\" + strCLSID + "\DefaultIcon", "", txtIconResource, REG_SZ) <> ERROR_SUCCESS Then Err.Raise APP_ERROR_BASE + 4, , "Could not set DefaultIcon"
    '//step 4:
    If CreateNewKey(HKEY_CLASSES_ROOT, "CLSID\" + strCLSID + "\Shell") <> ERROR_SUCCESS Then Err.Raise APP_ERROR_BASE + 5, , "Could not create Shell key"
    '//step 5:
    If CreateNewKey(HKEY_CLASSES_ROOT, "CLSID\" + strCLSID + "\Shell\Open") <> ERROR_SUCCESS Then Err.Raise APP_ERROR_BASE + 6, , "Could not create Open key"
    '//step 6:
    If CreateNewKey(HKEY_CLASSES_ROOT, "CLSID\" + strCLSID + "\Shell\Open\Command") <> ERROR_SUCCESS Then Err.Raise APP_ERROR_BASE + 7, , "Could not create Command key"
    If SetKeyValue(HKEY_CLASSES_ROOT, "CLSID\" + strCLSID + "\Shell\Open\Command", "", txtFilename, REG_SZ) <> ERROR_SUCCESS Then Err.Raise APP_ERROR_BASE + 8, , "Could not set AppName"
    '//step 7:
    If CreateNewKey(HKEY_LOCAL_MACHINE, strNameSpace + strCLSID) <> ERROR_SUCCESS Then Err.Raise APP_ERROR_BASE + 9, , "Could not create NameSpace key"
    If SetKeyValue(HKEY_LOCAL_MACHINE, strNameSpace + strCLSID, "", txtAppName, REG_SZ) <> ERROR_SUCCESS Then Err.Raise APP_ERROR_BASE + 10, , "Could not set NameSpace"
             
    MsgBox "The shortcut was successfuly created!", vbInformation, "Success!"
    
    '//write to log file:
    On Error GoTo FileErr
    Open App.Path & "\ShellLog.txt" For Append As #1
    If FileLen(App.Path & "\ShellLog.txt") = 0 Then strLog = "Log file created on " & Now() & vbCrLf & vbCrLf & strLog
    Print #1, strLog & vbCrLf
    Close #1
    Exit Sub
    
FileErr:
    MsgBox "Unable to create|write log file: " & Err.Description, vbCritical, "Error"
    Exit Sub
    
whoops:
    MsgBox "Could not create the shortcut: " & Err.Description, vbCritical, "Error"
End Sub

Private Function SetValueEx(ByVal hKey As hKey, ValueName As String, DataType As dwType, Value As Variant) As Long
    Dim lValue As Long
    Dim sValue As String

    Select Case DataType
        Case REG_SZ
            sValue = Value
            SetValueEx = RegSetValueExString(hKey, ValueName, 0&, DataType, sValue, Len(sValue))
        Case REG_DWORD
            lValue = Value
            SetValueEx = RegSetValueExLong(hKey, ValueName, 0&, DataType, lValue, 4)
        End Select

End Function

Private Function SetKeyValue(PredefinedKey As hKey, KeyName As String, ValueName As String, ValueSetting As Variant, DataType As dwType) As Long
    Dim lRetVal As Long
    Dim hKey As Long
    
    lRetVal = RegOpenKeyEx(PredefinedKey, KeyName, 0, KEY_ALL_ACCESS, hKey)
    
    If ValueSetting = "" Then
        DeleteValue HKEY_LOCAL_MACHINE, KeyName, ValueName
        RegCloseKey (hKey)
        Exit Function
    End If
    
    lRetVal = SetValueEx(hKey, ValueName, DataType, ValueSetting)
    SetKeyValue = lRetVal
    RegCloseKey (hKey)

End Function

Private Function CreateNewKey(PredefinedKey As hKey, NewKeyName As String) As Long
    Dim hNewKey As Long         'handle to the new key
    Dim lRetVal  As Long
    
    CreateNewKey = RegCreateKeyEx(PredefinedKey, NewKeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hNewKey, lRetVal)
    RegCloseKey (hNewKey)
End Function

Private Function DeleteValue(PredefinedKey As hKey, KeyName As String, ValueName As String)
       Dim lRetVal As Long
       Dim hKey As Long

       lRetVal = RegOpenKeyEx(PredefinedKey, KeyName, 0, KEY_ALL_ACCESS, hKey)
       lRetVal = RegDeleteValue(hKey, ValueName)
       RegCloseKey (hKey)
End Function

