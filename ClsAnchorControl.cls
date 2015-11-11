VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsAnchorControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private p_Controls As New Scripting.Dictionary
Private p_Spaces As New Scripting.Dictionary


Private WithEvents m_form As Form
Attribute m_form.VB_VarHelpID = -1

Dim WithEvents b  As Frame
Attribute b.VB_VarHelpID = -1

Public AnchorTop As Boolean
Public AnchorBottom As Boolean
Public AnchorLeft As Boolean
Public AnchorRight As Boolean

'Private SpaceTop As Integer
'Private SpaceLeft As Integer
'Private SpaceBottom As Integer
'Private SpaceRight As Integer

Public Sub AddControl(ctrl As Control)
    Set p_Controls(ctrl.Name) = ctrl
    SetInitialState ctrl
End Sub

Public Property Set MyForm(frm As Form)
    Set m_form = frm
    Dim key As Variant
    For Each key In p_Controls.keys
        SetInitialState p_Controls.Item(key)
    Next
End Property

Public Property Get MyForm() As Form
    Set MyForm = m_form
    
End Property



Private Sub m_form_Resize()
    
    Dim key As Variant
    For Each key In p_Controls.keys
        Dim ctrl As Control: Set ctrl = p_Controls.Item(key)
        
        Dim sp As New ClsSpaces
        
        Set sp = p_Spaces(ctrl.Name)

        Dim intAux As Integer
        If AnchorRight Then
            If AnchorLeft Then
                'Cresce
                intAux = MyForm.Width - (sp.SpaceRight + sp.SpaceLeft)
                ctrl.Width = IIf(intAux > 0, intAux, 0)
                
            ElseIf Not AnchorLeft Then
                'Desliza para direita
                intAux = MyForm.Width - (sp.SpaceRight + ctrl.Width)
                ctrl.Left = IIf(intAux > 0, intAux, 0)
            End If
        End If
        
        If AnchorBottom Then
            If AnchorTop Then
                'Cresce
                intAux = MyForm.Height - (sp.SpaceTop + sp.SpaceBottom)
                ctrl.Height = IIf(intAux > 0, intAux, 0)
            ElseIf Not AnchorTop Then
                'Desliza para baixo
                intAux = MyForm.Height - (sp.SpaceBottom + ctrl.Height)
                ctrl.Top = IIf(intAux > 0, intAux, 0)
            End If
        End If
    Next
End Sub

Private Sub SetInitialState(ctrl As Control)
    Dim spAux As New ClsSpaces
    
    If Not MyForm Is Nothing _
       And Not ctrl Is Nothing Then
       spAux.SpaceTop = ctrl.Top
       spAux.SpaceLeft = ctrl.Left
       spAux.SpaceBottom = MyForm.Height - (ctrl.Top + ctrl.Height)
       spAux.SpaceRight = MyForm.Width - (ctrl.Left + ctrl.Width)
       Set p_Spaces(ctrl.Name) = spAux
    End If
End Sub
