VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLongScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' *************************************************************************
'  Copyright ©1994-2003 Karl E. Peterson
'  All Rights Reserved, http://www.mvps.org/vb
' *************************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code, non-compiled, without prior written consent.
' *************************************************************************
Option Explicit

' Object variables used to reference client controls.
Private WithEvents m_ClientH As HScrollBar
Attribute m_ClientH.VB_VarHelpID = -1
Private WithEvents m_ClientV As VScrollBar
Attribute m_ClientV.VB_VarHelpID = -1
Private m_Client As Object

' Variables to track virtual scrollbar properties.
Private m_Min As Long
Private m_Max As Long
Private m_SmallChange As Long
Private m_LargeChange As Long
Private m_Value As Long

' Variables to track real scrollbar properties.
Private m_vbValue As Integer
Private m_vbMin As Integer
Private m_vbMax As Integer
Private m_vbSmallChange As Integer
Private m_vbLargeChange As Integer

' Flag property to warn of possible recursion into
' real scrollbar's Change event.
Private m_Recursing As Boolean

' Default property values
Private Const defValue = 0
Private Const defMin = 0
Private Const defMax = 32767
Private Const defSmallChange = 1
Private Const defLargeChange = 1
Private Const defLargeChangeClient = 10

' Notification events
Public Event Change()
Public Event Scroll()


' **************************************************************
'  Initialization
' **************************************************************
Private Sub Class_Initialize()
   ' Set default startup property values same as VB
   m_Value = defValue
   m_Min = defMin
   m_Max = defMax
   m_SmallChange = defSmallChange
   m_LargeChange = defLargeChange
End Sub

' **************************************************************
'  Public Properties
' **************************************************************
Public Property Set Client(obj As Object)
   ' Allow for proper clean-up
   If obj Is Nothing Then
      Set m_Client = Nothing
      Set m_ClientH = Nothing
      Set m_ClientV = Nothing
   Else
      ' Assign passed object to m_Client if appropriate type
      If TypeOf obj Is HScrollBar Then
         Set m_ClientH = obj
         Set m_ClientV = Nothing
         Set m_Client = obj
      ElseIf TypeOf obj Is VScrollBar Then
         Set m_ClientH = Nothing
         Set m_ClientV = obj
         Set m_Client = obj
      Else
         Err.Clear
         Err.Raise Number:=vbObjectError + 513, _
                   Source:="CLongScroll", _
                   Description:="Client object must be a scrollbar."
      End If
   End If
   
   ' Assign new value to intrinsic properties
   If Not (m_Client Is Nothing) Then
      m_Client.Min = defMin
      m_Client.Max = defMax
      m_Client.SmallChange = defSmallChange
      m_Client.LargeChange = defLargeChangeClient
      m_Recursing = True
         m_Client.Value = defValue
      m_Recursing = False
   End If
End Property

'Public Property Get Client() As Object
   ' Return m_Client object
'   Set Client = m_Client
'End Property

Public Property Let Max(ByVal NewVal As Long)
   ' Assign Virtual Max property
   m_Max = NewVal
   
   ' Make sure m_Value is in legal range
   If m_Max > m_Min Then
      If m_Value > m_Max Then
         Me.Value = m_Max
      End If
   ElseIf m_Min > m_Max Then
      If m_Value > m_Min Then
         Me.Value = m_Min
      End If
   Else 'm_Min = m_Max
      If m_Max <> m_Value Then
         Me.Value = m_Max
      End If
   End If
End Property

'Public Property Get Max() As Long
   ' Return Virtual Max property
'   Max = m_Max
'End Property

Public Property Let Min(ByVal NewVal As Long)
   ' Assign Virtual Min property
   m_Min = NewVal
   
   ' Make sure m_Value is in legal range
   If m_Max > m_Min Then
      If m_Value < m_Min Then
         Me.Value = m_Min
      End If
   ElseIf m_Min > m_Max Then
      If m_Value < m_Max Then
         Me.Value = m_Min
      End If
   Else 'm_Min = m_Max
      If m_Min <> m_Value Then
         Me.Value = m_Min
      End If
   End If
End Property

'Public Property Get Min() As Long
   ' Return Virtual Min property
'   Min = m_Min
'End Property

Public Property Let LargeChange(ByVal NewVal As Long)
   ' Assign Virtual LargeChange property
   m_LargeChange = NewVal
End Property

'Public Property Get LargeChange() As Long
   ' Return Virtual LargeChange property
'   LargeChange = m_LargeChange
'End Property

Public Property Let SmallChange(ByVal NewVal As Long)
   ' Assign Virtual SmallChange property
   m_SmallChange = NewVal
End Property

'Public Property Get SmallChange() As Long
   ' Return Virtual SmallChange property
'   SmallChange = m_SmallChange
'End Property

Public Property Let Value(ByVal NewVal As Long)
   Dim VirtualRange As Long
   Dim RealRange As Long
   Dim Percent As Double
   Dim tmpMin As Long
   Dim tmpMax As Long
   
   ' Get current values from real scrollbar
   Call ReadRealValues
   
   ' Cases where Virtual(Min>Max) need to be handled specially.
   ' Some calculations require swapped values.
   If m_Min > m_Max Then
      tmpMin = m_Max
      tmpMax = m_Min
   Else
      tmpMin = m_Min
      tmpMax = m_Max
   End If
   
   ' Rather than raise an error, correct out-of-range values
   If NewVal < tmpMin Then
      NewVal = tmpMin
   ElseIf NewVal > tmpMax Then
      NewVal = tmpMax
   End If
   
   ' Set Virtual value
   m_Value = NewVal
   
   ' Calculate Real value of scrollbar
   VirtualRange = Abs(m_Max - m_Min)
   RealRange = Abs(m_vbMax - m_vbMin)
   If VirtualRange Then
      Percent = Abs(m_Value - tmpMin) / VirtualRange
   Else
      Percent = 0
   End If
   
   ' If Virtual(Min>Max) then flip value
   If m_Min <= m_Max Then
      m_vbValue = m_vbMin + (Percent * RealRange)
   Else
      m_vbValue = m_vbMax - (Percent * RealRange)
   End If
   
   ' Update real scrollbar and notify client
   Call UpdateRealValue
   RaiseEvent Change
End Property

Public Property Get Value() As Long
   ' Return Virtual value for scrollbar
   Value = m_Value
End Property

' **************************************************************
'  Sunken Client Events
' **************************************************************
Private Sub m_ClientH_Change()
   ' In this, and the other change/scroll events,
   ' pass execution to a generic calc routine that
   ' resets the scrollbar's values and updates
   ' internal tracking variables.
   Call ClientChange
End Sub

Private Sub m_ClientH_Scroll()
   Call ClientScroll
End Sub

Private Sub m_ClientV_Change()
   Call ClientChange
End Sub

Private Sub m_ClientV_Scroll()
   Call ClientScroll
End Sub

' **************************************************************
'  Private Methods
' **************************************************************
Private Sub ClientChange()
   Dim Delta As Long
   Dim VirtualRange As Long
   Dim RealRange As Long
   Dim Percent As Double
   Dim tmpMin As Long
   Dim tmpMax As Long
   
   ' Bail if recursing
   If m_Recursing Then Exit Sub
   
   ' Calculate real change
   Delta = m_Client.Value - m_vbValue
   
   ' Get current values from real scrollbar
   Call ReadRealValues
   
   ' Cases where Virtual(Min>Max) need to be handled specially.
   ' Most calculations can use swapped values.
   If m_Min > m_Max Then
      tmpMin = m_Max
      tmpMax = m_Min
      Delta = -1 * Delta
   Else
      tmpMin = m_Min
      tmpMax = m_Max
   End If
   
   ' See if Large or Small Change
   If Abs(Delta) = m_vbLargeChange Or _
      Abs(Delta) = m_vbSmallChange Then
      
      ' Adjust change to match virtual scaling
      If Abs(Delta) = m_vbLargeChange Then
         Delta = Sgn(Delta) * m_LargeChange
      ElseIf Abs(Delta) = m_vbSmallChange Then
         Delta = Sgn(Delta) * m_SmallChange
      End If
      
      ' Set virtual scale
      m_Value = m_Value + Delta
      
      ' Check if out of bounds
      If m_Value < tmpMin Then
         m_Value = tmpMin
      ElseIf m_Value > tmpMax Then
         m_Value = tmpMax
      End If
      
      ' Calculate Real value of scrollbar
      VirtualRange = Abs(m_Max - m_Min)
      RealRange = Abs(m_vbMax - m_vbMin)
      If VirtualRange Then
         Percent = Abs(m_Value - tmpMin) / VirtualRange
      Else
         Percent = 0
      End If
      
      ' If Virtual(Min>Max) then flip value
      If m_Min <= m_Max Then
         m_vbValue = m_vbMin + (Percent * RealRange)
      Else
         m_vbValue = m_vbMax - (Percent * RealRange)
      End If

   Else
      ' User moved thumb on scrollbar.
      ' Calculate Virtual value of scrollbar
      VirtualRange = Abs(m_Max - m_Min)
      RealRange = Abs(m_vbMax - m_vbMin)
      If RealRange Then
         Percent = Abs(m_vbValue - m_vbMin) / RealRange
      Else
         Percent = 0
      End If
      
      ' If Virtual(Min>Max) then flip value
      If m_Min <= m_Max Then
         m_Value = tmpMin + (Percent * VirtualRange)
      Else
         m_Value = tmpMax - (Percent * VirtualRange)
      End If
   End If
   
   ' Update real scrollbar
   Call UpdateRealValue
   
   ' Raise notification event
   RaiseEvent Change
End Sub

Private Sub ClientScroll()
   Dim VirtualRange As Long
   Dim RealRange As Long
   Dim Percent As Double
   Dim tmpMin As Long
   Dim tmpMax As Long
   
   ' Get current values from real scrollbar
   Call ReadRealValues
   
   ' Cases where Virtual(Min>Max) need to be handled specially.
   ' Most calculations can use swapped values.
   If m_Min > m_Max Then
      tmpMin = m_Max
      tmpMax = m_Min
   Else
      tmpMin = m_Min
      tmpMax = m_Max
   End If
   
   ' User moved thumb on scrollbar.
   ' Calculate Virtual value of scrollbar.
   VirtualRange = Abs(m_Max - m_Min)
   RealRange = Abs(m_vbMax - m_vbMin)
   If RealRange Then
      Percent = Abs(m_vbValue - m_vbMin) / RealRange
   Else
      Percent = 0
   End If
   
   ' If Virtual(Min>Max) then flip value
   If m_Min <= m_Max Then
      m_Value = tmpMin + (Percent * VirtualRange)
   Else
      m_Value = tmpMax - (Percent * VirtualRange)
   End If
   
   ' Raise notification event
   RaiseEvent Scroll
End Sub

Private Sub ReadRealValues()
   ' Read current values from scrollbar
   m_vbValue = m_Client.Value
   m_vbMin = m_Client.Min
   m_vbMax = m_Client.Max
   m_vbSmallChange = m_Client.SmallChange
   m_vbLargeChange = m_Client.LargeChange
End Sub

Private Sub UpdateRealValue()
   ' This assures that if the virtual value is not quite
   ' to either the Min or Max that there's still room to
   ' adjust the slider.
   If m_vbValue = m_vbMin Then
     If m_Value > m_Min Then
       m_vbValue = m_vbMin + 1
     End If
   ElseIf m_vbValue = m_vbMax Then
     If m_Value < m_Max Then
       m_vbValue = m_vbMax - 1
     End If
   End If
   
   ' Update display.  Note possible recursion!
   m_Recursing = True
      m_Client.Value = m_vbValue
   m_Recursing = False
End Sub




