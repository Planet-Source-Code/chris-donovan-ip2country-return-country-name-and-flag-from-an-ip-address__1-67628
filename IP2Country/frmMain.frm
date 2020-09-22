VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IP2Country Test"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   705
   ClientWidth     =   9975
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   9975
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrStart 
      Interval        =   300
      Left            =   3600
      Top             =   0
   End
   Begin VB.Frame Frame3 
      Caption         =   "Speed Test"
      Height          =   2295
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   3135
      Begin VB.CommandButton cmdSpeedTest 
         Caption         =   "Go"
         Height          =   315
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   2655
      End
      Begin VB.ComboBox cboIP 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "(If less than 10 ms, duration = 0)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1960
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "Number of IP's:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1040
         Width           =   1095
      End
      Begin VB.Label lblElapsed 
         Caption         =   "Time Elapsed (sec):"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   2895
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Get Country From IP"
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3135
      Begin VB.PictureBox picFlag 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1360
         ScaleHeight     =   375
         ScaleWidth      =   615
         TabIndex        =   13
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox txtIP 
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton cmdGetCountry 
         Caption         =   "Go"
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "IP Address:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1000
         Width           =   855
      End
      Begin VB.Label lblCountry 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   2895
      End
   End
   Begin VB.Timer tmrRefreshTable 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   3120
      Top             =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "TCP Table (refreshing every 3 seconds)"
      Height          =   4935
      Left            =   3480
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin MSComctlLib.ListView ListView 
         Height          =   4095
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   7223
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgFlagList"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "#"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "IP"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Country"
            Object.Width           =   5998
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "(If less than 10 ms, duration = 0)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   14
         Top             =   4560
         Width           =   2775
      End
      Begin VB.Label lblRefresh 
         Caption         =   "Refresh Time (sec):"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   4560
         Width           =   3375
      End
   End
   Begin MSComctlLib.ImageList imgFlagList 
      Left            =   3120
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuConverter 
         Caption         =   "IP List Converter"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'The TCP Table code is simply used to show how to use
'IP2Country with a Listview and it was taken from:
'http://www.vbip.com/iphelper/get_tcp_table.asp
'
'The IPConvert function was taken from:
'http://www.freevbcode.com/ShowCode.Asp?ID=5512


Private Type MIB_TCPROW
    dwState As Long
    dwLocalAddr As Long
    dwLocalPort As Long
    dwRemoteAddr As Long
    dwRemotePort As Long
End Type

Private Const ERROR_BUFFER_OVERFLOW = 111&
Private Const ERROR_INVALID_PARAMETER = 87
Private Const ERROR_NO_DATA = 232&
Private Const ERROR_NOT_SUPPORTED = 50&
Private Const ERROR_SUCCESS = 0&

Private Declare Function GetTCPTable Lib "iphlpapi.dll" Alias "GetTcpTable" (ByRef pTcpTable As Any, ByRef pdwSize As Long, ByVal bOrder As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef pDest As Any, ByRef pSource As Any, ByVal Length As Long)

Dim cIP As IP2Country.clsIP2Country

Private Sub cmdGetCountry_Click()
  If ValidIP(txtIP.Text) Then
    lblCountry.Caption = cIP.GetCountryFromIP(txtIP.Text) 'Show country name based on IP
    picFlag.Picture = imgFlagList.ListImages(cIP.Flag).Picture 'Load the corresponding flag in the Picturebox
  Else
    lblCountry.Caption = "INVALID IP !!"
    picFlag.Picture = LoadPicture() 'Empty the picturebox on invalid IP
  End If
End Sub


'The IP list is loaded here and not in the splash screen first, because unloading the
'splash screen also removes the IP list from memory and I don't like hiding forms.
Private Sub Start()
  
  tmrStart.Enabled = False
  
  If Dir(App.Path & "\IPList.dat", vbNormal) <> "" Then
    frmSplash.Show 'Don't show this modally or the IPList/ImageList won't be loaded
    
      Pause (200) 'Add a small pause or frmSplash won't have a chance to show up properly,
                  'because LoadIPList halts the program right away for a second while loading the IP list.
       
        Call cIP.LoadIPList(App.Path & "\IPList.dat")
    
          frmSplash.Label1.Caption = "Loading Image list..."
          Call cIP.LoadImageList(imgFlagList)
    
        Pause (1500) 'Give the user the change to see the 'image loading' message.
    
    Unload frmSplash
    RefreshTable 'Load the TCP table
    tmrRefreshTable.Enabled = True
  Else
    MsgBox "IPList.dat not found. Make sure it is in the program folder." & vbCrLf & vbCrLf & _
      "IP2Country will now close.", vbExclamation, "IP2Country"
    Unload Me
  End If

End Sub

Private Sub cmdSpeedTest_Click()
  Dim i&, StartTime&, EndTime&

  cmdSpeedTest.Caption = "Running..."
  
    StartTime = GetTickCount 'Start timer
      For i = 1 To cboIP.Text
        Call cIP.GetCountryFromIP(GetRandomIP) 'Get county name from a new random IP in every loop
      Next i
    EndTime = GetTickCount 'End timer
  
  cmdSpeedTest.Caption = "Go"
  
  lblElapsed.Caption = "Time Elapsed (sec): " & (EndTime - StartTime) / 1000
End Sub

Private Sub Form_Load()
  Set cIP = New IP2Country.clsIP2Country
  
  cboIP.AddItem "1"
  cboIP.AddItem "100"
  cboIP.AddItem "1000"
  cboIP.AddItem "10000"
  cboIP.AddItem "100000"
  cboIP.AddItem "1000000"
  cboIP.Text = "100000"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set cIP = Nothing
  
  Dim f As Form
    For Each f In Forms
        Unload f
    Next
  Set f = Nothing

End Sub

Private Sub RefreshTable()

    Dim arrBuffer() As Byte
    Dim lngSize As Long
    Dim lngRetVal As Long
    Dim i As Long
    Dim lngRows As Long
    Dim TcpTableRow As MIB_TCPROW
    Dim lvItem As ListItem
    Dim IPString As String
    Dim StartTime&, EndTime&
    
    StartTime = GetTickCount 'Start timer
    
    ListView.ListItems.Clear

    lngSize = 0

    lngRetVal = GetTCPTable(ByVal 0&, lngSize, 0)
    
    If lngRetVal = ERROR_NOT_SUPPORTED Then

        MsgBox "IP Helper is not supported by this system."
        Exit Sub
        '
    End If

    ReDim arrBuffer(0 To lngSize - 1) As Byte

    lngRetVal = GetTCPTable(arrBuffer(0), lngSize, 0)

    If lngRetVal = ERROR_SUCCESS Then

        CopyMemory lngRows, arrBuffer(0), 4

        For i = 1 To lngRows

            CopyMemory TcpTableRow, arrBuffer(4 + (i - 1) * Len(TcpTableRow)), Len(TcpTableRow)

            If Not (GetIpFromLong(TcpTableRow.dwRemoteAddr) = "0.0.0.0" Or GetIpFromLong(TcpTableRow.dwLocalAddr) = "127.0.0.1") Then

                With TcpTableRow
                    IPString = GetIpFromLong(.dwRemoteAddr)
                    Set lvItem = ListView.ListItems.Add() 'Don't add anything to first column - column width is 0
                    lvItem.SubItems(1) = IPString 'The IP address from the tcp table
                    On Error Resume Next 'Keep going if imagelist key does not exist
                    
                    'cIP.GetCountryFromIP(IPString) = country name
                    'cIP.Flag = corresponding country code / imagelist key
                    lvItem.ListSubItems.Add , , cIP.GetCountryFromIP(IPString), cIP.Flag
                    
                    'If the imagelist key does not exist, then show county name or '--NOT IN IP LIST--'
                    'message and show the 'question mark' flag - (error 35601 = element not found)
                    If Err.Number = 35601 Then lvItem.ListSubItems.Add , , cIP.GetCountryFromIP(IPString), "UNKNOWN"
                End With

            End If

        Next i

    End If
    
    EndTime = GetTickCount 'End timer
    lblRefresh.Caption = "Refresh Time (sec): " & (EndTime - StartTime) / 1000 'Show refresh duration
End Sub

Private Function GetIpFromLong(lngIPAddress As Long) As String
  Dim arrIpParts(3) As Byte
  CopyMemory arrIpParts(0), lngIPAddress, 4
  GetIpFromLong = CStr(arrIpParts(0)) & "." & CStr(arrIpParts(1)) & "." & CStr(arrIpParts(2)) & "." & CStr(arrIpParts(3))
End Function

Private Sub mnuConverter_Click()
  frmConverter.Show vbModal
End Sub

Private Sub mnuExit_Click()
  Unload Me
End Sub

Private Sub tmrRefreshTable_Timer()
  RefreshTable
End Sub

Private Sub tmrStart_Timer()
  Start 'Show frmSplash, load IP list and ImageList
End Sub
