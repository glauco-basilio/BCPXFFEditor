VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--- clsLayout ---

Option Explicit

Public Enum enLayoutStyle
   lsNone = 0     ' fixed position and size
   lsLeft = 1     ' fixed indentation from left of form
   lsRight = 2    ' fixed indentation from right of form
   lsHorizontal = 3 ' fixed left and right indentation (resize horizontally)
   lsTop = 4      ' fixed indentation from top of form
   lsBottom = 8   ' fixed indentation from bottom of form
   lsVertical = 12 'fixed top and bottom indentation (resize vertically)
   lsFull = 15    ' resize horizontally and vertically
   lsXProp = 16   ' position is dependent on form width
   lsYProp = 32   ' position is dependent on form height
   lsWprop = 64   ' width is proportional to form width
   lsDprop = 128  ' width is proportional to form depth (height)
End Enum

Private m_Controls As Collection ' of clsControl objects
Private m_minWidth As Long
Private m_minHeight As Long

Private Function AddControl(ctrl As Control, style As enLayoutStyle)
On Error GoTo errortrap
   Dim udt As New clsControl
   udt.style = style
   Set udt.ctrl = ctrl
   udt.LeftIndent = ctrl.Left
   udt.TopIndent = ctrl.Top
   udt.RightIndent = ctrl.Parent.Width - (ctrl.Width + ctrl.Left)
   udt.BottomIndent = ctrl.Parent.Height - (ctrl.Height + ctrl.Top)
   udt.Width = ctrl.Width
   udt.Height = ctrl.Height
   m_Controls.Add udt
errortrap:
   ' there are some controls that have no visual element in run time (ie timer, image list etc)
End Function

Private Function NotNeg(x As Long)
   NotNeg = IIf(x < 0, 0, x)
End Function

Private Function GetLayoutStyle(s As String) As enLayoutStyle
   Dim ls As enLayoutStyle
   ls = lsNone
   Dim i As Integer
   For i = 1 To Len(s)
      Select Case Mid$(s, i, 1)
      Case "L"
         ls = ls Or lsLeft
      Case "R"
         ls = ls Or lsRight
      Case "T"
         ls = ls Or lsTop
      Case "B"
         ls = ls Or lsBottom
      Case "V"
         ls = ls Or lsVertical
      Case "H"
         ls = ls Or lsHorizontal
      Case "F"
         ls = ls Or lsFull
      Case "X"
         ls = ls Or lsXProp
      Case "Y"
         ls = ls Or lsYProp
      Case "W"
         ls = ls Or lsWprop
      Case "D"
         ls = ls Or lsDprop
      Case "N"
         ls = lsNone
      End Select
   Next
   GetLayoutStyle = ls
End Function

Public Function RedrawLayout()
   Dim Width As Long
   Dim Height As Long
   Dim ctrl As clsControl

   For Each ctrl In m_Controls
      With ctrl
         If Width = 0 Then
            Width = .ctrl.Parent.Width
            If Width < m_minWidth Then Width = m_minWidth
         End If
         If Height = 0 Then
            Height = .ctrl.Parent.Height
            If Height < m_minHeight Then Height = m_minHeight
         End If
         
         
         If (.style And lsLeft) = lsLeft Then
            .ctrl.Left = .LeftIndent
         ElseIf (.style And lsXProp) = lsXProp Then
            .ctrl.Left = (.LeftIndent * Width) / m_minWidth
         End If
         If (.style And lsHorizontal) = lsHorizontal Then
            .ctrl.Width = NotNeg((Width - .RightIndent) - .LeftIndent)
         Else
            If (.style And lsWprop) = lsWprop Then
               .ctrl.Width = (.Width * Width) / m_minWidth
            End If
            If (.style And lsRight) = lsRight Then
               .ctrl.Left = NotNeg((Width - .RightIndent) - .ctrl.Width)
            End If
         End If
         
         If (.style And lsTop) = lsTop Then
            .ctrl.Top = .TopIndent
         ElseIf (.style And lsYProp) = lsYProp Then
            .ctrl.Top = (.TopIndent * Height) / m_minHeight
         End If
         If (.style And lsVertical) = lsVertical Then
            .ctrl.Height = NotNeg((Height - .BottomIndent) - .TopIndent)
         Else
            If (.style And lsDprop) = lsDprop Then
               .ctrl.Height = (.Height * Height) / m_minHeight
            End If
            If (.style And lsBottom) = lsBottom Then
               .ctrl.Top = NotNeg((Height - .BottomIndent) - .ctrl.Height)
            End If
         End If
      End With
   Next

End Function

Public Function SetLayout(frm As Form)
      m_minWidth = frm.Width
      m_minHeight = frm.Height
       
      Set m_Controls = New Collection ' of clsControl objects
      Dim con As Control
      Dim ls As enLayoutStyle
      For Each con In frm.Controls
         ls = GetLayoutStyle(con.Tag)
         If ls <> lsNone Then
            AddControl con, ls 'GetLayoutStyle(con.Tag)
         End If
      Next
End Function
