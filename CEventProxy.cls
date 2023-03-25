VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CEventProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("HEMS")
Option Compare Database
Option Explicit

Private Const EventsExplained As String = "https://www.youtube.com/watch?v=kfIabBuUtmA"
Private Const EventsExpample As String = "https://www.youtube.com/watch?v=tVbUwdfm57Q"
Private Const CustomEvents As String = "https://sourcedaddy.com/ms-access/creating-custom-events.html"

Public Event ScheduleCustomerSelected(record As DAO.Fields)
Public Event SchedulesRowSelected(ID As Long)
Public Event CurrentPeriod(Year As Long, Month As Long)


Public Event AfterNewCustomer()
Public Event BeforeNewCustomer()




Public Sub SendSceduleCustomerSelected(record As DAO.Recordset)
  Debug.Print "Before Sending Event ScheduleCustomerSelected"
  RaiseEvent ScheduleCustomerSelected(record.Fields)
  DoEvents
End Sub
Public Sub SendSchedulesRowSelected(ID As Long)
  Debug.Print "Before Sending Event SchedulesRowSelected"
  RaiseEvent SchedulesRowSelected(ID)
  DoEvents
End Sub


Public Sub SendPeriod(Year As Long, Month As Long)
  RaiseEvent CurrentPeriod(Year, Month)
  DoEvents
End Sub


Public Sub evBeforeNewCustomer()
  RaiseEvent BeforeNewCustomer
  DoEvents
End Sub
Public Sub evAfterNewCustomer()
  RaiseEvent AfterNewCustomer
  DoEvents
End Sub

