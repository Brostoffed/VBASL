VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ENV_USER_INFO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ==========================================================================
' Class Name        : ENV_USER_INFO
' Module Type       : Class
' Module Path		:
' Author            : Kevin Brostoff
' Site:             : https://github.com/Brostoffed
' --------------------------------------------------------------------------
' Description       : This class is a container for methods dealing
'                     with Environments => User Specific Information
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

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Private Functions
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Function username() As String
    On Error GoTo ErrHandler:
    username = LCase$(Environ("username"))
Done:
    Exit Function
ErrHandler:
    Call LOGGER_.print_error("username", Err, True)
End Function


Public Function computer_name() As String
    On Error GoTo ErrHandler:
    computer_name = LCase$(Environ("COMPUTERNAME"))
Done:
    Exit Function
ErrHandler:
    Call LOGGER_.print_error("computer_name", Err, True)
End Function

Public Function user_domain() As String
    On Error GoTo ErrHandler:
    user_domain = LCase$(Environ("USERDOMAIN"))
Done:
    Exit Function
ErrHandler:
    Call LOGGER_.print_error("user_domain", Err, True)
End Function

Public Function terminal_client_name() As String
    On Error GoTo ErrHandler:
    terminal_client_name = LCase$(Environ("ClientName"))
Done:
    Exit Function
ErrHandler:
    Call LOGGER_.print_error("terminal_client_name", Err, True)
End Function

Public Function terminal_session_name() As String
    On Error GoTo ErrHandler:
    terminal_session_name = LCase$(Environ("SESSIONNAME"))
Done:
    Exit Function
ErrHandler:
    Call LOGGER_.print_error("terminal_session_name", Err, True)
End Function

Public Function logon_server() As String
    On Error GoTo ErrHandler:
    logon_server = LCase$(Environ("LOGONSERVER"))
Done:
    Exit Function
ErrHandler:
    Call LOGGER_.print_error("logon_server", Err, True)
End Function

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' UserForm Builder Subroutines / Functions
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
