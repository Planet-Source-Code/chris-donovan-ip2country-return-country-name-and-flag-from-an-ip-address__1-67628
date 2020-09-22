VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmConverter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IP List Converter"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   0
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Convert"
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   5655
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   580
         Left            =   120
         ScaleHeight     =   585
         ScaleWidth      =   5415
         TabIndex        =   6
         Top             =   720
         Width           =   5415
         Begin VB.TextBox txtCSV 
            Height          =   285
            Left            =   400
            TabIndex        =   7
            Top             =   100
            Width           =   4575
         End
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   215
         Left            =   120
         TabIndex        =   4
         Top             =   1620
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   370
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.CommandButton cmdConvert 
         Caption         =   "Load IP List"
         Height          =   330
         Left            =   2280
         TabIndex        =   2
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label lblUpdate 
         Caption         =   "Waiting..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   5295
      End
   End
   Begin VB.Label lblHyperlink 
      Alignment       =   2  'Center
      Caption         =   "Download Latest IP-to-Country Database"
      Height          =   255
      Left            =   120
      MouseIcon       =   "frmConverter.frx":0000
      TabIndex        =   5
      Top             =   1060
      Width           =   5655
   End
   Begin VB.Label Label1 
      Caption         =   $"frmConverter.frx":030A
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frmConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
                        ByVal hwnd As Long, _
                        ByVal lpOperation As String, _
                        ByVal lpFile As String, _
                        ByVal lpParameters As String, _
                        ByVal lpDirectory As String, _
                        ByVal nShowCmd As Long) _
                        As Long

Private Sub Form_Load()
    lblHyperlink.ForeColor = vbBlue
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHyperlink.FontUnderline = False
    lblHyperlink.MousePointer = 0
End Sub

Private Sub lblHyperlink_Click()
    ShellExecute Me.hwnd, "open", "http://ip-to-country.webhosting.info/node/view/6", vbNullString, vbNullString, vbNormal
End Sub

Private Sub lblHyperlink_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblHyperlink.FontUnderline = True
    lblHyperlink.MousePointer = 99
    lblHyperlink.ToolTipText = "http://ip-to-country.webhosting.info/node/view/6"
End Sub

Private Sub cmdConvert_Click()
    
  On Error GoTo ErrHandler
    
    With CommonDialog
        .Flags = cdlOFNExplorer + cdlOFNHideReadOnly + cdlOFNFileMustExist
        .Filter = "ip-to-country (*.csv)|*.csv"
        .DialogTitle = "ip-to-country.csv"
        .InitDir = App.Path
        .CancelError = True
        .ShowOpen
        If .FileName <> "" Then
          If Dir(App.Path & "\IPList.dat", vbNormal) <> "" Then
              If MsgBox("IPList.dat already exists. Do you want to overwrite it?", vbQuestion + vbYesNo, "Confirmation") = vbYes Then
                txtCSV.Text = .FileName
                ConvertIPList (.FileName)
              End If
          Else
            txtCSV.Text = .FileName
            ConvertIPList (.FileName)
          End If
        End If
    
    End With

ErrHandler:
  'Cancel was selected
End Sub

Private Function ConvertIPList(sFileName As String)
  Dim ff As Integer
  Dim ln As String
  Dim arr() As String
  Dim sArray() As String
  
  On Error GoTo ErrHandler
  
  ff = FreeFile
  ProgressBar1.Max = FileLen(sFileName) 'Use progbar based on bytes read - counting text file lines takes too long
  lblUpdate.Caption = "Converting IP list..."
  
  ReDim sArray(0)
  Open sFileName For Input As #ff
    Do While Not EOF(ff)
      Line Input #ff, ln
      DoEvents
      ProgressBar1.Value = ProgressBar1.Value + Len(ln) + 2 'Update progressbar based on bytes read
      ln = Replace(ln, Chr(34), "") 'Remove quotation marks
      arr() = Split(ln, ",") 'Split items on the comma
      ln = arr(4) & ":" & arr(3) & ":" & IPConvert(arr(0)) & ":" & IPConvert(arr(1)) 'Put useful items back in different order
      sArray(UBound(sArray)) = ln 'Add to a new array
      ReDim Preserve sArray(UBound(sArray) + 1)
    Loop
  Close #ff
  ProgressBar1.Value = 0
  
  
  QuickSort sArray, LBound(sArray), UBound(sArray) 'Sort the array


  'We could simply dump the entire array into a text file with: "Print #1, Join(sArray, vbCrLf)"
  'but the ReDim code above adds an empty line and QuickSort puts it at the start of the array,
  'so the updated IPList.dat file also starts with an empty line. The code below won't add that
  'empty line, it just takes slightly longer to save the array to the file
  Dim i As Long
  ProgressBar1.Max = UBound(sArray)
  lblUpdate.Caption = "Saving IP list..."
  Open App.Path & "\IPList.dat" For Output As #1
    Do While i <= UBound(sArray)
      DoEvents
      ProgressBar1.Value = i
      If sArray(i) <> "" Then 'Make sure it's not an empty line...
        Print #1, sArray(i) 'then write it to the IPList.dat file
      End If
      i = i + 1
    Loop
  Close #1
  ProgressBar1.Value = 0
  lblUpdate.Caption = "Waiting..."
  txtCSV.Text = ""

  MsgBox "Conversion done!", vbInformation, "Done"
Exit Function

ErrHandler:
  ProgressBar1.Value = 0
  lblUpdate.Caption = "Waiting..."
  txtCSV.Text = ""
  
  If Err.Number = 9 Then
    MsgBox "There was an error reading the IP list on line: " & UBound(sArray) + 1 & vbCrLf & "Make sure it is in the correct format." _
    & vbCrLf & vbCrLf & """33996344""" & ", " & """33996351""" & ", " & """GB""" & ", " & """GBR""" & ", " & """UNITED KINGDOM""", vbExclamation, "Error"
  Else
    MsgBox "An error has occured!" & vbCrLf & vbCrLf & Err.Description, vbOKOnly + vbExclamation, "Error: " & Err.Number
  End If
  
  Err.Clear
End Function

Private Sub QuickSort(c() As String, ByVal First As Long, ByVal Last As Long)
  Dim Low As Long, High As Long
  Dim MidValue As String
    
  Low = First
  High = Last
  MidValue = c((First + Last) \ 2)
    
  Do
      While c(Low) < MidValue
          Low = Low + 1
      Wend
        
      While c(High) > MidValue
          High = High - 1
      Wend
        
      If Low <= High Then
          Swap c(Low), c(High)
          Low = Low + 1
          High = High - 1
      End If
  Loop While Low <= High
    
  If First < High Then QuickSort c, First, High
  If Low < Last Then QuickSort c, Low, Last
End Sub

Private Sub Swap(ByRef a As String, ByRef b As String)
  Dim T As String
    
  T = a
  a = b
  b = T
End Sub
