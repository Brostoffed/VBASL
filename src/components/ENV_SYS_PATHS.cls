VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ENV_SYS_PATHS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' ==========================================================================
' Class Name        : ENV_SYS_PATHS
' Module Type       : Class
' Module Path		:
' Author            : Kevin Brostoff
' Site:             : https://github.com/Brostoffed
' --------------------------------------------------------------------------
' Description       : This class is a container for methods dealing
'                     with Environments => System Paths
' --------------------------------------------------------------------------
' License           :
' ==========================================================================

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Windows APIs
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#If VBA7 Then
#Else
#End If
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Consumed Events
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Defined Events
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Implements Interfaces
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Enumerations & Types
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Class Variables
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Class: Initialization & Termination
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub Class_Initialize()
End Sub
' ------------------
Private Sub Class_Terminate()
End Sub

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Class Properties
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Using Events
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Public Subroutines
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Private Subroutines
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Public Functions
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Public Function all_users_profile() As String
    On Error GoTo ErrHandler:
    all_users_profile = LCase$(Environ("ALLUSERSPROFILE"))
Done:
    Exit Function
ErrHandler:
    Call LOGGER_.print_error("all_users_profile", Err, True)
End Function

Public Function app_data() As String
    On Error GoTo ErrHandler:
    app_data = LCase$(Environ("APPDATA"))
Done:
    Exit Function
ErrHandler:
    Call LOGGER_.print_error("app_data", Err, True)
End Function

Public Function local_app_data() As String
    On Error GoTo ErrHandler:
    local_app_data = LCase$(Environ("LOCALAPPDATA"))
Done:
    Exit Function
ErrHandler:
    Call LOGGER_.print_error("local_app_data", Err, True)
End Function

Public Function system_root() As String
    On Error GoTo ErrHandler:
    system_root = LCase$(Environ("SystemRoot"))
Done:
    Exit Function
ErrHandler:
    Call LOGGER_.print_error("system_root", Err, True)
End Function

Public Function home_drive() As String
    On Error GoTo ErrHandler:
    home_drive = LCase$(Environ("HOMEDRIVE"))
Done:
    Exit Function
ErrHandler:
    Call LOGGER_.print_error("home_drive", Err, True)
End Function

Public Function home_path() As String
    On Error GoTo ErrHandler:
    home_path = LCase$(Environ("HOMEPATH"))
Done:
    Exit Function
ErrHandler:
    Call LOGGER_.print_error("home_path", Err, True)
End Function

Public Function windows_directory() As String
    On Error GoTo ErrHandler:
    windows_directory = LCase$(Environ("windir"))
Done:
    Exit Function
ErrHandler:
    Call LOGGER_.print_error("windows_directory", Err, True)
End Function

Public Function program_files() As String
    On Error GoTo ErrHandler:
    program_files = LCase$(Environ("ProgramFiles"))
Done:
    Exit Function
ErrHandler:
    Call LOGGER_.print_error("program_files", Err, True)
End Function

Public Function common_program_files() As String
    On Error GoTo ErrHandler:
    common_program_files = LCase$(Environ("CommonProgramFiles"))
Done:
    Exit Function
ErrHandler:
    Call LOGGER_.print_error("common_program_files", Err, True)
End Function
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Private Functions
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' UserForm Builder Subroutines / Functions
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
