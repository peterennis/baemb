Option Compare Database
Option Explicit

' jleach@dymeng, Feb 2015
' handles some registry stuff via WMI
' v1.0
' all breaking changes must increment the major version number
'
' https://msdn.microsoft.com/en-us/library/aa394600(v=vs.85).aspx
'
' should probably set it up to run silent... maybe v1.1

Public Enum RegistryHives
  HKCR = &H80000000   'HKEY_CLASSES_ROOT
  HKCU = &H80000001   'HKEY_CURRENT_USER
  HKLM = &H80000002   'HKEY_LOCAL_MACHINE
  HKU = &H80000003    'HKEY_USERS
  HKCC = &H80000005   'HKEY_CURRENT_CONFIG
  HKDD = &H80000006   'HKEY_DYN_DATA
End Enum

Public Enum RegValueTypes
  REG_ANY = 0
  REG_SZ = 1
  REG_EXPAND_SZ = 2
  REG_BINARY = 3
  REG_DWORD = 4
  REG_MULTI_SZ = 7
  REG_QWORD = 11
End Enum

Public Function GetDWORDValue(Hive As RegistryHives, ByVal KeyPath As String, ValueName As String) As Long
On Error GoTo Err_Proc
'=========================
  Dim ret As Long
  Dim reg As Object
'=========================

  If Left(KeyPath, 1) = "\" Then KeyPath = Mid(KeyPath, 2)
  If Right(KeyPath, 1) = "\" Then KeyPath = Left(KeyPath, Len(KeyPath) - 1)

  If Not Registry.ValueExists(Hive, KeyPath, ValueName) Then Err.Raise vbObjectError + 1, , "Value doesn't exists."
  If Not Registry.KeyExists(Hive, KeyPath) Then Err.Raise vbObjectError + 2, , "Key doesn't exist."

  Set reg = GetReg()
  
  reg.GetDWORDValue Hive, KeyPath, ValueName, ret

'=========================
Exit_Proc:
  Set reg = Nothing
  GetDWORDValue = ret
  Exit Function
Err_Proc:
  Err.Source = "Registry.GetDWORDValue"
  Select Case Err.Number
    Case Else
      MsgBox Err.Number & ": " & Err.Description
  End Select
  Resume Exit_Proc
  Resume
End Function

Public Function CreateDWORDValue(Hive As RegistryHives, ByVal KeyPath As String, ValueName As String, Value As Long) As Boolean
On Error GoTo Err_Proc
'=========================
  Dim ret As Boolean
  Dim reg As Object
'=========================

  'remove leading and trailing \ if they exist
  If Left(KeyPath, 1) = "\" Then KeyPath = Mid(KeyPath, 2)
  If Right(KeyPath, 1) = "\" Then KeyPath = Left(KeyPath, Len(KeyPath) - 1)

  If Registry.ValueExists(Hive, KeyPath, ValueName) Then Err.Raise vbObjectError + 1, , "Value already exists."
  If Not Registry.KeyExists(Hive, KeyPath) Then Err.Raise vbObjectError + 2, , "Key doesn't exist."
  
  Set reg = GetReg()
  
  ret = Not CBool(reg.SetDWORDValue(Hive, KeyPath, ValueName, Value))
  
'=========================
Exit_Proc:
  Set reg = Nothing
  CreateDWORDValue = ret
  Exit Function
Err_Proc:
  Err.Source = "Registry.CreateValue"
  Select Case Err.Number
    Case vbObjectError + 1
      MsgBox "Specified value already exists"
    Case vbObjectError + 2
      MsgBox "Key does not exist.  Please create it first."
    Case Else
      MsgBox Err.Number & ": " & Err.Description
  End Select
  Resume Exit_Proc
  Resume
End Function

Public Function ValueExists( _
      Hive As RegistryHives, _
      ByVal KeyPath As String, _
      ValueName As String, _
      Optional ValueType As RegValueTypes = RegValueTypes.REG_ANY _
) As Boolean
On Error GoTo Err_Proc
'=========================
  Dim ret As Boolean
  Dim reg As Object
  Dim Values() As Variant
  Dim ValueTypes() As Variant
  Dim v As Variant
  Dim vt As Variant
  Dim i As Integer
'=========================

  'remove leading and trailing \ if they exist
  If Left(KeyPath, 1) = "\" Then KeyPath = Mid(KeyPath, 2)
  If Right(KeyPath, 1) = "\" Then KeyPath = Left(KeyPath, Len(KeyPath) - 1)

  If Not Registry.KeyExists(Hive, KeyPath) Then GoTo Exit_Proc

  Set reg = GetReg()
  
  reg.EnumValues Hive, KeyPath, Values, ValueTypes
  
  For i = 0 To UBound(ValueTypes)
    If CStr(Values(i)) = ValueName Then
      If ValueType = REG_ANY Then
        ret = True
        GoTo Exit_Proc
      Else
        If CLng(ValueTypes(i)) = CLng(ValueType) Then
          ret = True
          GoTo Exit_Proc
        End If
      End If
    End If
  Next
  
'=========================
Exit_Proc:
  Set reg = Nothing
  ValueExists = ret
  Exit Function
Err_Proc:
  Err.Source = "Registry.ValueExists"
  Select Case Err.Number
    Case Else
      MsgBox Err.Number & ": " & Err.Description
  End Select
  Resume Exit_Proc
  Resume
End Function

Public Function CreateKey(Hive As RegistryHives, ByVal KeyPath As String) As Boolean
On Error GoTo Err_Proc
'=========================
  Dim ret As Boolean
  Dim reg As Object
'=========================

  'remove leading and trailing \ if they exist
  If Left(KeyPath, 1) = "\" Then KeyPath = Mid(KeyPath, 2)
  If Right(KeyPath, 1) = "\" Then KeyPath = Left(KeyPath, Len(KeyPath) - 1)
  
  Set reg = GetReg()
  
  ret = Not CBool(reg.CreateKey(Hive, KeyPath))

'=========================
Exit_Proc:
  Set reg = Nothing
  CreateKey = ret
  Exit Function
Err_Proc:
  Err.Source = "Registry.CreateKey"
  Select Case Err.Number
    Case Else
      MsgBox Err.Number & ": " & Err.Description
  End Select
  Resume Exit_Proc
  Resume
End Function

Public Function KeyExists(Hive As RegistryHives, ByVal KeyPath As String) As Boolean
On Error GoTo Err_Proc
'=========================
  Dim ret As Boolean
  Dim reg As Object
  Dim ParentPath As String
  Dim Subkey As String
  Dim x As Long
  Dim Subkeys() As Variant
  Dim v As Variant
'=========================

  If Right(KeyPath, 1) = "\" Then KeyPath = Left(KeyPath, Len(KeyPath) - 1) 'remove trailing \ if present
  
  ParentPath = Mid(KeyPath, 1, InStrRev(KeyPath, "\"))
  Subkey = Mid(KeyPath, InStrRev(KeyPath, "\") + 1)
  
  'remove leading \ if present
  If Left(ParentPath, 1) = "\" Then ParentPath = Mid(ParentPath, 2)
  
  Set reg = GetReg()
  
  If reg.EnumKey(Hive, ParentPath, Subkeys) <> 0 Then GoTo Exit_Proc
  
  For Each v In Subkeys
    If CStr(v) = Subkey Then
      ret = True
      GoTo Exit_Proc
    End If
  Next

'=========================
Exit_Proc:
  Set reg = Nothing
  KeyExists = ret
  Exit Function
Err_Proc:
  Err.Source = "Registry.KeyExists"
  Select Case Err.Number
    Case Else
      MsgBox Err.Number & ": " & Err.Description
  End Select
  Resume Exit_Proc
  Resume
End Function

Private Function GetReg() As Object
  Set GetReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
End Function