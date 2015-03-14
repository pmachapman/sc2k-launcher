Attribute VB_Name = "ModuleMain"
' Make sure all variables are declared
Option Explicit

' Global constants
Public Const ProgramName = "SIMCITY.EXE"
Public Const REG_SZ As Long = 1
Public Const REG_DWORD As Long = 4

' Module constants
Private Const HKEY_CURRENT_USER = &H80000001
Private Const ERROR_NONE = 0
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_ALL_ACCESS = &H3F
Private Const REG_OPTION_NON_VOLATILE = 0

' Access the GetUserNameA function in advapi32.dll and call the function GetUserName.
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

' Access the registry functions from shell32.dll
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long

' The program entry point
Public Sub Main()
    ' First make sure that simcity.exe is in the current directory
    If Dir(ProgramName, vbNormal) = "" Then
        MsgBox "This program must be run from the same path as SimCity 2000", vbOKOnly + vbCritical, "Error"
        End
    End If
    
    ' Next, see if the registry keys exist
    If QueryValue("Software\Maxis\SimCity 2000\REGISTRATION", "Mayor Name") = "" Then
        ' They do not exist, so show the form
        FormMain.Show
    Else
        ' They do exist, so start SimCity 2000
        On Error GoTo ProgramNotFound
        Shell ProgramName
    End If
    
    ' Exit the routine before the error handlers
    Exit Sub
    
    ' Program not found error handler
ProgramNotFound:
    MsgBox "The program " & ProgramName & " could not be found.", vbOKOnly + vbCritical, "Error"
    End
End Sub

' Queries the registry key name in HKEY_CURRENT_USER for a value
Public Function QueryValue(sKeyName As String, sValueName As String) As Variant
    Dim lRetVal As Long      ' Result of the API functions
    Dim hKey As Long         ' Handle of opened key
    Dim vValue As Variant    ' Setting of queried value

    lRetVal = RegOpenKeyEx(HKEY_CURRENT_USER, sKeyName, 0&, KEY_QUERY_VALUE, hKey)
    lRetVal = QueryValueEx(hKey, sValueName, vValue)
    QueryValue = vValue
    RegCloseKey (hKey)
End Function

' Sets a registry key value in HKEY_CURRENT_USER
Public Sub SetKeyValue(sKeyName As String, sValueName As String, vValueSetting As Variant, lValueType As Long)
    Dim lRetVal As Long      ' Result of the SetValueEx function
    Dim hKey As Long         ' Handle of open key
    Dim lpdwDisposition As Long

    ' Open the specified key
    lRetVal = RegCreateKeyEx(HKEY_CURRENT_USER, sKeyName, 0&, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lpdwDisposition)
    lRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
    RegCloseKey (hKey)
End Sub

' Returns the logged in user's name
Public Function UserName() As String

    ' Declare variables
    Dim lpBuff As String * 25
    Dim ret As Long
    
    ' Get the user name minus any trailing spaces found in the name.
    ret = GetUserName(lpBuff, 25)
    UserName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
    
End Function

' Sets a registry key
Private Function SetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long
    Dim lValue As Long
    Dim sValue As String
    Select Case lType
        Case REG_SZ
            sValue = vValue & Chr$(0)
            SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
        Case REG_DWORD
            lValue = vValue
            SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
        End Select
End Function

' Queries a registry key for its value
Private Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
    Dim cch As Long
    Dim lrc As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String

    On Error GoTo QueryValueExError

    ' Determine the size and type of data to be read
    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
    If lrc <> ERROR_NONE Then Error 5

    Select Case lType
        ' For strings
        Case REG_SZ:
            sValue = String(cch, 0)
            lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
            If lrc = ERROR_NONE Then
                vValue = Left$(sValue, cch - 1)
            Else
                vValue = Empty
            End If
        ' For DWORDS
        Case REG_DWORD:
            lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
            If lrc = ERROR_NONE Then vValue = lValue
        Case Else
            'all other data types not supported
            lrc = -1
    End Select

QueryValueExExit:
    QueryValueEx = lrc
    Exit Function

QueryValueExError:
    Resume QueryValueExExit
End Function

