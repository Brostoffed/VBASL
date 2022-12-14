VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ENV_COMP_INFO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' ==========================================================================
' Class Name        : ENV_COMP_INFO
' Module Type       : Class
' Module Path		:
' Author            : Kevin Brostoff
' Site:             : https://github.com/Brostoffed
' --------------------------------------------------------------------------
' Description       : This class is a container for methods dealing
'                     with Environments => Core Computer Info
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
Public Function operating_system() As String
    On Error GoTo ErrHandler:
    computer_os = LCase$(Environ("OS"))
Done:
    Exit Function
ErrHandler:
    Call LOGGER_.print_error("operating_system_os", Err, True)
End Function

Public Function processor_architecture() As String
    On Error GoTo ErrHandler:
    processor_architecture = LCase$(Environ("PROCESSOR_ARCHITECTURE"))
Done:
    Exit Function
ErrHandler:
    Call LOGGER_.print_error("processor_architecture", Err, True)
End Function

Public Function processor_identifier() As String
    On Error GoTo ErrHandler:
    processor_identifier = LCase$(Environ("PROCESSOR_IDENTIFIER"))
Done:
    Exit Function
ErrHandler:
    Call LOGGER_.print_error("processor_identifier", Err, True)
End Function

Public Function processor_level() As String
    On Error GoTo ErrHandler:
    processor_level = LCase$(Environ("PROCESSOR_LEVEL"))
Done:
    Exit Function
ErrHandler:
    Call LOGGER_.print_error("processor_level", Err, True)
End Function

Public Function processor_revision() As String
    On Error GoTo ErrHandler:
    processor_revision = LCase$(Environ("PROCESSOR_REVISION"))
Done:
    Exit Function
ErrHandler:
    Call LOGGER_.print_error("processor_revision", Err, True)
End Function

Public Function number_of_processors() As String
    On Error GoTo ErrHandler:
    number_of_processors = LCase$(Environ("NUMBER_OF_PROCESSORS"))
Done:
    Exit Function
ErrHandler:
    Call LOGGER_.print_error("number_of_processors", Err, True)
End Function
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Private Functions
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' UserForm Builder Subroutines / Functions
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
