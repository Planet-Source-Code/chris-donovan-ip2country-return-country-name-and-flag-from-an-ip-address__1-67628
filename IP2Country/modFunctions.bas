Attribute VB_Name = "modFunctions"
Option Explicit

Public Declare Function GetTickCount Lib "kernel32" () As Long

'Get a random IP for the speed test
Public Function GetRandomIP() As String
  Dim a%, b%, c%, d%
  
  Randomize
      
      a = CInt(Rnd() * 254 + 1)
      b = CInt(Rnd() * 254 + 1)
      c = CInt(Rnd() * 254 + 1)
      d = CInt(Rnd() * 254 + 1)
  
  GetRandomIP = a & "." & b & "." & c & "." & d

End Function

'Pause the app without freezing it ('Sleep' freezes the app)
Public Function Pause(HowLong As Long)
  Dim Start&
  Start = GetTickCount()
  
  Do
    DoEvents
  Loop Until Start + HowLong < GetTickCount
End Function

'Simple function to check if IP is valid --> if it's not higher than 255.255.255.255
Public Function ValidIP(ByVal strIPAddress As String) As Boolean
  Dim sArray As Variant
  
  On Error GoTo ErrHandler
    
    sArray = Split(strIPAddress, ".")
      If sArray(0) > 255 Or sArray(1) > 255 Or sArray(2) > 255 Or sArray(3) > 255 Then
        ValidIP = False
      Else
        ValidIP = True
      End If

Exit Function
ErrHandler:
  ValidIP = False
End Function

'Convert a Long IP to a Dotted one (or back)
Public Function IPConvert(IPAddress As Variant) As Variant

    Dim X       As Integer
    Dim pos     As Integer
    Dim PrevPos As Integer
    Dim Num     As Integer

    If IsNumeric(IPAddress) Then
        IPConvert = "0.0.0.0"
        For X = 1 To 4
            Num = Int(IPAddress / 256 ^ (4 - X))
            IPAddress = IPAddress - (Num * 256 ^ (4 - X))
            If Num > 255 Then
                IPConvert = "0.0.0.0"
                Exit Function
            End If

            If X = 1 Then
                IPConvert = Num
            Else
                IPConvert = IPConvert & "." & Num
            End If
        Next
    ElseIf UBound(Split(IPAddress, ".")) = 3 Then
'        On Error Resume Next
        For X = 1 To 4
            pos = InStr(PrevPos + 1, IPAddress, ".", 1)
            If X = 4 Then pos = Len(IPAddress) + 1
            Num = Int(Mid(IPAddress, PrevPos + 1, pos - PrevPos - 1))
            If Num > 255 Then
                IPConvert = "0"
                Exit Function
            End If
            PrevPos = pos
            IPConvert = ((Num Mod 256) * (256 ^ (4 - X))) + IPConvert
        Next
    End If

End Function



