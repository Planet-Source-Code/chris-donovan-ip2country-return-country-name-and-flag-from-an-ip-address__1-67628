VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIP2Country"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Use and modify as you wish
'Thanks to Thread1 for his help

'=================================================
'Country name
'=================================================

Private Type ITEM
  CountryName As String
  CountryCode As String
  IPStart As Long
  IPEnd As Long
End Type

Private Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long
Private Declare Function ntohl Lib "wsock32.dll" (ByVal netlong As Long) As Long

Private m_items() As ITEM

'=================================================
'Country flag
'=================================================

Type GUID
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(7) As Byte
End Type
  
Private Declare Function CreateStreamOnHGlobal Lib "ole32.dll" (ByRef hGlobal As Any, ByVal fDeleteOnResume As Long, ByRef ppstr As Any) As Long
Private Declare Function OleLoadPicture Lib "olepro32.dll" (ByVal lpStream As IUnknown, ByVal lSize As Long, ByVal fRunMode As Long, ByRef riid As GUID, ByRef lplpObj As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpsz As Long, ByRef pclsid As GUID) As Long
  
' HKEY_CLASSES_ROOT\Interface\{7BF80980-BF32-101A-8BBB-00AA00300CAB} = IPicture
Private Const SIPICTURE As String = "{7BF80980-BF32-101A-8BBB-00AA00300CAB}"
Private GetCountryCode As Variant
Public Flag As String

'=================================================
'Country name
'=================================================

Public Function LoadIPList(sFilepath As String)
  ReadFile m_items, sFilepath           'Read the IPList.dat file
  QuickSort m_items, 1, UBound(m_items) 'Sort the array
End Function

Public Function GetCountryFromIP(sIP As String) As String
  GetCountryFromIP = Country(m_items, sIP) 'Return country name from an IP address
End Function

'Read the IPList.dat file
Private Function ReadFile(r() As ITEM, ByVal Filepath As String)
Dim f%, s$, i&
Dim x0&, x1&, x2&
  
  On Error GoTo ErrHandler
  
  f = FreeFile()
  Open Filepath For Input As #f
  
  i = 0
  ReDim r(i) As ITEM
  
  Do Until EOF(f)
  
    Line Input #f, s
    
      If s <> "" Then
        i = i + 1
        ReDim Preserve r(i) As ITEM
          With r(i) 'Split country name, country code, start IP range, end IP range
            x1 = InStr(1, s, ":")
            .CountryName = Left$(s, x1 - 1) 'County name
            x0 = Len(.CountryName)
            x2 = InStr(x0 + 2, s, ":")
            .CountryCode = Mid$(s, x1 + 1, x2 - x1 - 1) 'Country code
            .IPStart = ntohl(inet_addr(Mid$(s, x2 + 1, InStr(Len(.CountryCode) + x0 + 4, s, ":") - x2 - 1))) 'Start IP range
            .IPEnd = ntohl(inet_addr(Mid$(s, InStrRev(s, ":") + 1))) 'End IP range
          End With
      End If
  Loop
  
  Close #f
  
Exit Function
ErrHandler:
  Err.Clear
  MsgBox "There was an error reading the IP list on line: " & i & vbCrLf & "Make sure it is in the correct format." & _
  vbCrLf & vbCrLf & "Country Name : Country Code : Start IP Range : End IP Range  " & vbCrLf & vbCrLf & _
  "UNITED KINGDOM:GBR:128.16.0.0:128.16.255.255", vbExclamation, "Error"
End Function

'Sort the array based on start IP ranges
Private Function QuickSort(r() As ITEM, ByVal lo As Long, ByVal up As Long)
Dim i&, j&, tempr As ITEM
  
  While up > lo
  
    i = lo
    j = up
    
    tempr = r(lo)
    
    While i < j
      While r(j).IPStart > tempr.IPStart
        j = j - 1
      Wend
      r(i) = r(j)
      While (i < j) And r(i).IPStart <= tempr.IPStart
        i = i + 1
      Wend
      r(j) = r(i)
    Wend
    
    r(i) = tempr
    
    QuickSort r, lo, i - 1
    lo = i + 1
    
  Wend
  
End Function

'Search IP ranges and return the country name and country code the IP belongs to
'Country = country name
'Flag = country code (to call the image key in the ImageList)
Private Function Country(r() As ITEM, ByVal ip As String) As String
Dim h&, m&, l&
Dim n&
  
  On Error GoTo ErrHandler
  
  n = ntohl(inet_addr(ip))
  
  l = 1
  h = UBound(r)
    
  Do Until l > h
    m = (h - l) \ 2 + l
    If n < r(m).IPStart Then
      h = m - 1
    Else
      If n > r(m).IPEnd Then
        l = m + 1
      Else
          Country = r(m).CountryName
          Flag = r(m).CountryCode
        Exit Function
      End If
    End If
  Loop
  
  Country = "--NOT IN IP LIST--"
  Flag = "UNKNOWN"
  
  Exit Function
ErrHandler:
  Err.Clear
  
  Country = "IP LIST IS NOT LOADED !!"

End Function

'=================================================
'Country flag
'=================================================

Private Function PictureFromRes(ByRef b() As Byte) As IPicture
  On Error GoTo ErrHandler
     
  Dim istrm As IUnknown
  Dim tGuid As GUID
     
  If Not CreateStreamOnHGlobal(b(LBound(b)), False, istrm) Then
    CLSIDFromString StrPtr(SIPICTURE), tGuid
    OleLoadPicture istrm, UBound(b) - LBound(b) + 1, False, tGuid, PictureFromRes
  End If
     
  Set istrm = Nothing
  Exit Function
ErrHandler:
  Debug.Print "Could not convert to IPicture!"
End Function

Private Function GetCountryFlag(ByVal lResID As Long) As StdPicture
  Set GetCountryFlag = PictureFromRes(LoadResData(lResID, "CUSTOM"))
End Function

Public Function LoadImageList(imgList As Object)
Dim intIndex%, i%

  'These are the country codes matching the image ID's in the resource file
  'GetCountryCode(1) = "AFG" --> image ID in resource file = 1 --> Afghan(istan) flag
  'GetCountryCode(2) = "ALB" --> image ID in resource file = 2 --> Albania(n) flag
  'GetCountryCode(198) = "GBR" --> image ID in resource file = 198 --> United Kingdom flag
  'etc
  
  'When loaded into ImageList, image 1 (Afghan flag) gets image key "AFG"
  'When loaded into ImageList, image 2 (Albanian flag) gets image key "ALB"
  'etc
  GetCountryCode = Array("", "AFG", "ALB", "DZA", "ASM", "AND", "AGO", "AIA", "ATA", "ATG", "ARG", "ARM", _
                         "ABW", "AUS", "AUT", "AZE", "BHS", "BHR", "BGD", "BRB", "BLR", "BEL", "BLZ", "BEN", _
                         "BMU", "BTN", "BOL", "BIH", "BWA", "BRA", "IOT", "BRN", "BGR", "BFA", "BDI", "KHM", _
                         "CMR", "CAN", "CPV", "CYM", "CAF", "TCD", "CHL", "CHN", "COL", "COM", "COG", "COK", _
                         "CRI", "CIV", "HRV", "CUB", "CYP", "CZE", "DNK", "DJI", "DMA", "DOM", "ECU", "EGY", _
                         "SLV", "GNQ", "ERI", "EST", "ETH", "FLK", "FRO", "FSM", "FJI", "FIN", "FRA", "GUF", _
                         "PYF", "GAB", "GMB", "GEO", "DEU", "GHA", "GIB", "GRC", "GRL", "GRD", "GLP", "GUM", _
                         "GTM", "GIN", "GNB", "GUY", "HTI", "VAT", "HND", "HKG", "HUN", "ISL", "IND", "IDN", _
                         "IRQ", "IRL", "IRN", "ISR", "ITA", "JAM", "JPN", "JOR", "KAZ", "KEN", "KIR", "KWT", _
                         "KGZ", "LAO", "LVA", "LBN", "LSO", "LBR", "LBY", "LIE", "LTU", "LUX", "MAC", "MDG", _
                         "MWI", "MYS", "MDV", "MLI", "MLT", "MTQ", "MRT", "MUS", "MEX", "MCO", "MNG", "MAR", _
                         "MOZ", "NAM", "NRU", "NPL", "ANT", "NLD", "NCL", "NZL", "NIC", "NER", "NGA", "MNP", _
                         "NOR", "OMN", "PAK", "PSE", "PAN", "PNG", "PRY", "PER", "PHL", "POL", "PRT", "PRI", _
                         "QAT", "KOR", "MDA", "ROM", "RUS", "RWA", "KNA", "LCA", "VCT", "SMR", "STP", "SAU", _
                         "SEN", "SCG", "SYC", "SLE", "SGP", "SVK", "SVN", "SLB", "SOM", "ZAF", "ESP", "LKA", _
                         "SUR", "SWE", "CHE", "SYR", "TWN", "TJK", "THA", "COD", "MKD", "TGO", "TTO", "TUN", _
                         "TUR", "TKM", "TUV", "UGA", "UKR", "ARE", "GBR", "TZA", "USA", "URY", "UZB", "VUT", _
                         "VEN", "VNM", "VGB", "YEM", "ZMB", "ZWE", "UNKNOWN") '210 country codes used as ImageList Keys
  
  'Load the 210 flags into the ImageList
  For i = 1 To 210
    intIndex = imgList.ListImages.Count + 1
    imgList.ImageHeight = 14
    imgList.ImageWidth = 22
    imgList.ListImages.Add intIndex, GetCountryCode(i), GetCountryFlag(i)
  Next i

  Erase GetCountryCode 'Clear the GetCountryCode array, because we don't need it anymore
End Function
