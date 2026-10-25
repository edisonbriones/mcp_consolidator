VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Distributor_Code As String
Public Customer_Code As String
Public Customer_Name As String
Public Province As String
Public Town As String
Public Barangay As String
Public Street As String
Public PSGC As String
Public Customer_Class As String
Public Credit_Term As String
Public salesman_code As String
Public Salesman_Name As String
Public service_model As String
Public Supervisor_Code As String
Public Supervisor_Name As String
Public frequency As String
Public Mon As Byte
Public Tue As Byte
Public Wed As Byte
Public Thu As Byte
Public Fri As Byte
Public Sat As Byte
Public Sun As Byte
Public WK1 As Byte
Public WK2 As Byte
Public WK3 As Byte
Public WK4 As Byte
Public Sequence As String
Public Day As String
'Public Remarks As String

' Member variable
Private m_remarks As String

' Properties
Property Get Remarks() As String
    Remarks = m_remarks
End Property

Property Let Remarks(value As String)
    m_remarks = value
End Property

' Event - triggered when class created
Private Sub Class_Initialize()
    m_remarks = Empty
End Sub

