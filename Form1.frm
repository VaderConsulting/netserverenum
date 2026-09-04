VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   8430
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView ListView1 
      Height          =   3855
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   6800
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   300
      Left            =   5265
      TabIndex        =   3
      Top             =   4065
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&NetServerEnum"
      Default         =   -1  'True
      Height          =   300
      Left            =   3615
      TabIndex        =   2
      Top             =   4065
      Width           =   1635
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   1380
      TabIndex        =   1
      Text            =   "justicex"
      Top             =   4065
      Width           =   2205
   End
   Begin VB.Label Label1 
      Caption         =   "&Domain Name:"
      Height          =   300
      Left            =   180
      TabIndex        =   0
      Top             =   4065
      Width           =   1140
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SV_TYPE_NT                    As Long = &H1000

Private Type Server_Info_100
    sv100_PlatformID    As Long
    sv100_Servername    As Long
End Type

'used for pointer conversion
Private Type MungeLong
  x As Long
  Dummy As Integer
End Type

Private Type MungeInt
  XLo As Integer
  XHi As Integer
  Dummy As Integer
End Type

Private Declare Function NetServerEnum Lib "NETAPI32.DLL" (ByVal Servername As Long, _
                                                    ByVal Level As Long, _
                                                    ByRef buffer As Long, _
                                                    ByRef PrefMaxLen As Long, _
                                                    ByRef EntriesRead As Long, _
                                                    ByRef TotalEntries As Long, _
                                                    ByRef ServerType As Long, _
                                                    DomainName As Byte, _
                                                    ByRef ResumeHandle As Long) As Long
                                                    
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (pTo As Any, uFrom As Any, ByVal lSize As Long)
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function lstrcpyW Lib "kernel32" (lpString1 As Byte, ByVal lpString2 As Long) As Long
Private Declare Function PtrToInt Lib "kernel32" Alias "lstrcpynW" (RetVal As Any, ByVal Ptr As Long, ByVal nCharCount As Long) As Long
Private Declare Function PtrToStr Lib "kernel32" Alias "lstrcpyW" (RetVal As Byte, ByVal Ptr As Long) As Long
Private Declare Function StrLen Lib "kernel32" Alias "lstrlenW" (ByVal Ptr As Long) As Long

Private Sub Command1_Click()
    Dim strDomainName       As String
    
    Me.MousePointer = vbHourglass
    
    ListView1.ListItems.Clear
    
    strDomainName = Text1.Text
    If Len(Trim$(strDomainName)) = 0 Then
        MsgBox "Please enter a Doamin Name to get the computers for."
        Exit Sub
    End If
    
    Call GetActiveMachinesList(strDomainName, ListView1)
    
    Me.MousePointer = vbNormal
End Sub

Private Sub GetActiveMachinesList(strDomain As String, lvList As ListView)
    Dim lpBufPtr            As Long
    Dim EntriesRead         As Long
    Dim TotalEntries        As Long
    Dim lMaxLenPref         As Long
    Dim lReturn             As Long
    Dim lLevel              As Long
    Dim lResumeHandle       As Long
    Dim ServerArray()       As Byte
    Dim DomainArray()       As Byte
    Dim lvi                 As ListItem
    Dim lIndex              As Long
    Dim CompData            As String
    
    ServerArray = vbNullChar
    DomainArray = strDomain & vbNullChar
    
    lMaxLenPref = &HFFFFFFFF
    lLevel = 100
        
    lReturn = NetServerEnum(0&, lLevel, lpBufPtr, lMaxLenPref, EntriesRead, TotalEntries, SV_TYPE_NT, DomainArray(0), lResumeHandle)
    If lReturn = 0 Then
        Debug.Print "TotalEntries: " & TotalEntries, "EntriesRead: " & EntriesRead
        For lIndex = 1 To EntriesRead * 2 Step 2
            CompData = GetPointedField(lpBufPtr, lIndex + 1)
            Set lvi = ListView1.ListItems.Add(, , CompData)
            Set lvi = Nothing
        Next
    Else
        Debug.Print "Error... NetServerEnum returned: " & lReturn
    End If
                        
End Sub

Private Function PointerToStringW(lpStringW As Long) As String
   Dim buffer() As Byte
   Dim nLen As Long
   
   If lpStringW Then
      nLen = lstrlenW(lpStringW) * 2
      If nLen Then
         ReDim buffer(0 To (nLen - 1)) As Byte
         CopyMem buffer(0), ByVal lpStringW, nLen
         PointerToStringW = buffer
      End If
   End If
End Function

Function GetPointedField(bufptr, i) As Variant
    Dim GNArray(255) As Byte
    Dim TempPtr As MungeLong
    Dim TempStr As MungeInt
    Dim Result As Long
    
    ' Get pointer to string from beginning of buffer
    ' Copy 4 byte block of memory in 2 steps
    Result = PtrToInt(TempStr.XLo, bufptr + (i - 1) * 4, 2)
    Result = PtrToInt(TempStr.XHi, bufptr + (i - 1) * 4 + 2, 2)
    LSet TempPtr = TempStr ' munge 2 Integers to a Long
    ' Copy string to array and convert to a string
    Result = PtrToStr(GNArray(0), TempPtr.x)
    GetPointedField = Left(GNArray, StrLen(TempPtr.x))
End Function

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Load()

End Sub

Private Sub Text1_GotFocus()
    SelectAllOnFocus Text1
End Sub

Private Sub SelectAllOnFocus(TextBoxItem As TextBox)
'Desc : HighLights the Text in a TextBox Control When It Gets Focus
    On Error Resume Next
    
    If TextBoxItem.Text <> "" Then
        TextBoxItem.SelStart = 0
        TextBoxItem.SelLength = Len(TextBoxItem.Text)
    End If
    
End Sub

