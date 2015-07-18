Option Compare Database
Option Explicit


'OPTIONAL: provides programmatic means to set HKCU key so the Web Browser Control runs under IE11 emulation mode
'
' v1.0
' jleach@dymeng, Feb 2015
' requires Dymeng.Registry module (v1.*)
'
' probably ought to set up with separate return values depending on various cases, but for a quick and dirty manual solution this works

Public Enum IEEmulationMode
  IEEmulation11 = 0
  'add others as desired
End Enum

Const IE_EMULATION_KEY = "Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION"
Const IE_EMULATION_VALUE = "msaccess.exe"
Const IE_EMULATION_MODE_11 = 11999


Public Function IsIEEmulationModeSet(IEMode As IEEmulationMode) As Boolean

  Dim IEModeValue As Long
  Dim lTemp As Long
  
  Select Case IEMode
    Case IEEmulationMode.IEEmulation11:   IEModeValue = 11999
    Case Else:                            Err.Raise vbObjectError, , "Emulation Mode Enum not recognized"
  End Select

  If Not Registry.KeyExists(HKCU, IE_EMULATION_KEY) Then
    IsIEEmulationModeSet = False
    Exit Function
  End If
  
  If Not Registry.ValueExists(HKCU, IE_EMULATION_KEY, IE_EMULATION_VALUE, REG_DWORD) Then
    IsIEEmulationModeSet = False
    Exit Function
  End If
  
  If Registry.GetDWORDValue(HKCU, IE_EMULATION_KEY, IE_EMULATION_VALUE) = IEModeValue Then
    IsIEEmulationModeSet = True
  Else
    IsIEEmulationModeSet = False
  End If

End Function

Public Function SetIEEmulationMode(IEMode As IEEmulationMode, Optional Silent As Boolean = False) As Boolean

  'check the keys and values, try to create if it doesn't exist, etc etc

  Dim IEModeValue As Long
  Dim lTemp As Long
  
  Select Case IEMode
    Case IEEmulationMode.IEEmulation11:   IEModeValue = 11999
    Case Else:                            Err.Raise vbObjectError, , "Emulation Mode Enum not recognized"
  End Select

  If Not Registry.KeyExists(HKCU, IE_EMULATION_KEY) Then
    If Not Registry.CreateKey(HKCU, IE_EMULATION_KEY) Then
      If Not Silent Then MsgBox "Unable to create registry key.", vbCritical, "Error"
      SetIEEmulationMode = False
      Exit Function
    End If
  End If
    
  If Registry.ValueExists(HKCU, IE_EMULATION_KEY, IE_EMULATION_VALUE, REG_DWORD) Then
    'check if it's what we want...
    lTemp = Registry.GetDWORDValue(HKCU, IE_EMULATION_KEY, IE_EMULATION_VALUE)
    If lTemp = IEModeValue Then
      SetIEEmulationMode = True
      Exit Function
    Else
      If Not Silent Then MsgBox "IE Emulation mode currently exists for Access but has a value of " & lTemp & " instead of " & IEModeValue
      Exit Function
    End If
  Else
    If Not Registry.CreateDWORDValue(HKCU, IE_EMULATION_KEY, IE_EMULATION_VALUE, IEModeValue) Then
      If Not Silent Then MsgBox "Unable to create value."
      SetIEEmulationMode = False
      Exit Function
    End If
  End If
  
  SetIEEmulationMode = True

End Function