VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LOGGER_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
' ==========================================================================
' Class Name        : LOGGER_
' Module Type       : Class
' Module Path		:
' Author            : Kevin Brostoff
' Site:             : https://github.com/Brostoffed
' --------------------------------------------------------------------------
' Description       : This class is a container for methods dealing with logging
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
Public Sub print_error(sMethodName As String, vError As Object, LogToDB As Boolean)
    Dim oErrorObj As Object: Set oErrorObj = vError
    
    Debug.Print ("================================================")
    Debug.Print ("[Method]          => " & sMethodName)
    Debug.Print ("[Err Number]      => " & vError.Number)
    Debug.Print ("[Err Source]      => " & vError.Source)
    Debug.Print ("[Err Description] => " & vError.Description)
    Debug.Print ("================================================")
End Sub
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Private Subroutines
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Public Functions
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Private Functions
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' UserForm Builder Subroutines / Functions
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
