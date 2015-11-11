VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEditor 
   Caption         =   "Form1"
   ClientHeight    =   8460
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   12150
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   12150
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11520
      Top             =   7890
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtLog 
      Height          =   1215
      Left            =   150
      MultiLine       =   -1  'True
      TabIndex        =   6
      Tag             =   "varWidth varVertical"
      Top             =   7140
      Width           =   11835
   End
   Begin VB.CommandButton cmdAddRow 
      Caption         =   "+"
      Height          =   225
      Left            =   870
      TabIndex        =   5
      Top             =   3600
      Width           =   555
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      Height          =   225
      Left            =   870
      TabIndex        =   3
      Top             =   90
      Width           =   555
   End
   Begin VB.Frame frameRow 
      Caption         =   "Row"
      Height          =   3375
      Left            =   120
      TabIndex        =   1
      Tag             =   "varWidth"
      Top             =   3630
      Width           =   11835
      Begin SHDocVwCtl.WebBrowser wbRow 
         Height          =   3045
         Left            =   60
         TabIndex        =   4
         Tag             =   "varWidth"
         Top             =   210
         Width           =   11685
         ExtentX         =   20611
         ExtentY         =   5371
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
   End
   Begin VB.Frame frameRecord 
      Caption         =   "Record"
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Tag             =   "varWidth varHeight"
      Top             =   90
      Width           =   11835
      Begin SHDocVwCtl.WebBrowser wbRecord 
         Height          =   3045
         Left            =   60
         TabIndex        =   2
         Tag             =   "varWidth"
         Top             =   240
         Width           =   11685
         ExtentX         =   20611
         ExtentY         =   5371
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
   End
   Begin VB.Menu mnArquivo 
      Caption         =   "Arquivo"
      Begin VB.Menu mnAbrir 
         Caption         =   "Abrir"
      End
      Begin VB.Menu mnSalvar 
         Caption         =   "Salvar"
      End
      Begin VB.Menu mnSalvarComo 
         Caption         =   "Salvar como..."
      End
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private fileLoaded As Boolean
Private filepath As String
Private documentChanged As Boolean


Private Sub cmdAddRow_Click()
    wbRow.Document.parentWindow.Add
    documentChanged = True
End Sub

Private Sub Command1_Click()
    wbRecord.Document.parentWindow.Add
    documentChanged = True
End Sub

Public Function GetFields() As Scripting.Dictionary
    'Vários dicinários dentro de um dicionário cuja a chave é o attributo ID
    wbRecord.Document.parentWindow.execScript "getFields()", "javascript"
    If Not IsNull(wbRecord.Document.parentWindow.recordDic) Then
        Set GetFields = wbRecord.Document.parentWindow.recordDic
    End If
End Function

Public Function GetColumns() As Scripting.Dictionary
    'Vários dicinários dentro de um dicionário cuja a chave é o attributo ID
    wbRow.Document.parentWindow.execScript "getColumns()", "javascript"
    If Not IsNull(wbRow.Document.parentWindow.rowDic) Then
        Set GetColumns = wbRow.Document.parentWindow.rowDic
    End If
End Function

Private Sub Form_Load()
    
    CommonDialog1.CancelError = True
    wbRecord.Navigate App.Path & "\" & "Record.htm"
    wbRow.Navigate App.Path & "\" & "Row.htm"
    
    
    AncoraControles
    
End Sub

Private Sub AncoraControles()
    Static anchorLeftRight As New ClsAnchorControl
    Set anchorLeftRight.MyForm = Me
    With anchorLeftRight
        .AnchorLeft = True
        .AnchorRight = True
        .AddControl frameRecord
        .AddControl wbRecord
        .AddControl frameRow
        .AddControl wbRow
        .AddControl txtLog
    End With
    
    Static anchorTopBottom As New ClsAnchorControl
    Set anchorTopBottom.MyForm = Me
    With anchorTopBottom
        .AnchorTop = True
        .AnchorBottom = True
        .AddControl frameRecord
        .AddControl wbRecord
        
    End With
    
    Static AnchorBottom As New ClsAnchorControl
    Set AnchorBottom.MyForm = Me
    With AnchorBottom
        .AnchorBottom = True
        .AddControl cmdAddRow
        .AddControl frameRow
        '.AddControl wbRow
        .AddControl txtLog
    End With
    
    
End Sub



Private Sub mnAbrir_Click()
    CommonDialog1.Filter = "Arquivo de BCP (*.xml)|*.xml|Todos os Arquivos (*.*)|*.*"
    If fileLoaded Or documentChanged Then
        Dim result As VbMsgBoxResult
        result = MsgBox("Deseja salvar as aterações feitas no arquivo ?", vbExclamation + vbYesNoCancel, "Salvar")
        Select Case result
            Case VbMsgBoxResult.vbCancel
                Exit Sub
            Case VbMsgBoxResult.vbNo
                
            Case VbMsgBoxResult.vbOK
                If filepath <> "" Then
                    SaveIn filepath
                Else
                    CommonDialog1.ShowSave
                    If CommonDialog1.FileName <> "" Then
                        SaveIn CommonDialog1.FileName
                    Else
                        
                    End If
                End If
        End Select
        
    End If
On Error GoTo CancelError
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
        LoadFile CommonDialog1.FileName
    End If
CancelError:
End Sub


Public Sub LoadFile(bcpFilePath As String)
    Dim doc As New MSXML2.DOMDocument
    
    If doc.Load(bcpFilePath) Then
        fileLoaded = True
        filepath = bcpFilePath
        ParseDocument doc
    Else
        MsgBox "Não foi possível abrir o arquivo de origem", vbCritical
    End If
End Sub

Public Sub ParseDocument(doc As MSXML2.DOMDocument)
    Dim filedList As IXMLDOMNodeList
    Dim rowList As IXMLDOMNodeList
    Set filedList = doc.getElementsByTagName("FIELD")
    Dim node As IXMLDOMNode
    wbRecord.Document.parentWindow.Reset
    wbRow.Document.parentWindow.Reset
    'Atributos de Field
    Dim attrID As String
    Dim length As String
    Dim Prefix_Length As String
    Dim Terminator As String
    Dim Max_Length As String
    Dim Collation As String
    Dim xsiType As String
    
    'Atributos de column
    Dim Source As String
    Dim name As String
    Dim precision As String
    Dim attrScale As String
    Dim nullable As String
    
    For Each node In filedList
       length = Empty
       Prefix_Length = Empty
       Terminator = Empty
       Max_Length = Empty
       Collation = Empty
       attrID = Empty
       xsiType = Empty
       
       If Not node.Attributes.getQualifiedItem(ATTR_TYPE, XSI_SCHEMA_URI) Is Nothing Then
            xsiType = node.Attributes.getQualifiedItem(ATTR_TYPE, XSI_SCHEMA_URI).nodeValue
       End If

       If InStr(xsiType, "Fixed") > 0 And Not (node.Attributes.getNamedItem(ATTR_LENGTH) Is Nothing) Then
          'Obrigatorio
          length = node.Attributes.getNamedItem(ATTR_LENGTH).nodeValue
       Else
        
       End If
       
       If InStr(xsiType, "Prefix") > 0 And Not (node.Attributes.getNamedItem(ATTR_PREFIX_LENGTH) Is Nothing) Then
          'Obrigatorio
          Prefix_Length = node.Attributes.getNamedItem(ATTR_PREFIX_LENGTH).nodeValue
          'Opcional
          If Not (node.Attributes.getNamedItem(ATTR_MAX_LENGTH) Is Nothing) Then
            Max_Length = node.Attributes.getNamedItem(ATTR_MAX_LENGTH).nodeValue
          End If
       Else
          
       End If
       
       If InStr(xsiType, "Term") > 0 And Not (node.Attributes.getNamedItem(ATTR_TERMINATOR) Is Nothing) Then
          'Obrigatorio
          Terminator = node.Attributes.getNamedItem(ATTR_TERMINATOR).nodeValue
       Else
        
       End If
       
       If InStr(xsiType, "Char") > 0 And Not (node.Attributes.getNamedItem(ATTR_COLLATION) Is Nothing) Then
            'opcional
            Collation = node.Attributes.getNamedItem(ATTR_COLLATION).nodeValue
          If Not (node.Attributes.getNamedItem(ATTR_MAX_LENGTH) Is Nothing) Then
            Max_Length = node.Attributes.getNamedItem(ATTR_MAX_LENGTH).nodeValue
          End If
       End If
       
       If Not node.Attributes.getNamedItem(ATTR_ID) Is Nothing Then
        attrID = node.Attributes.getNamedItem(ATTR_ID).nodeValue
       End If
       wbRecord.Document.parentWindow.AddWith attrID, _
            xsiType, length, Prefix_Length, Terminator, Max_Length, LCase(Collation)
            
    Next
    
    Set rowList = doc.getElementsByTagName("COLUMN")
    
    For Each node In rowList
        Source = Empty
        name = Empty
        length = Empty
        precision = Empty
        attrScale = Empty
        nullable = Empty
        xsiType = Empty
        If Not (node.Attributes.getNamedItem(ATTR_SOURCE) Is Nothing) Then
            Source = node.Attributes.getNamedItem(ATTR_SOURCE).nodeValue
        Else
        
        End If
        
        If Not (node.Attributes.getNamedItem(ATTR_NAME) Is Nothing) Then
            name = node.Attributes.getNamedItem(ATTR_NAME).nodeValue
        Else
        
        End If
        
        If Not (node.Attributes.getNamedItem(ATTR_LENGTH) Is Nothing) Then
            length = node.Attributes.getNamedItem(ATTR_LENGTH).nodeValue
        Else
        
        End If
        
        If Not (node.Attributes.getNamedItem(ATTR_PRECISION) Is Nothing) Then
            precision = node.Attributes.getNamedItem(ATTR_PRECISION).nodeValue
        Else
        
        End If
        
        If Not (node.Attributes.getNamedItem(ATTR_SCALE) Is Nothing) Then
            attrScale = node.Attributes.getNamedItem(ATTR_SCALE).nodeValue
        Else
        
        End If
        
        If Not (node.Attributes.getNamedItem(ATTR_NULLABLE) Is Nothing) Then
            nullable = node.Attributes.getNamedItem(ATTR_NULLABLE).nodeValue
        Else
        
        End If
        
        If Not (node.Attributes.getQualifiedItem(ATTR_TYPE, XSI_SCHEMA_URI) Is Nothing) Then
            xsiType = node.Attributes.getQualifiedItem(ATTR_TYPE, XSI_SCHEMA_URI).nodeValue
        Else
        
        End If
        
        wbRow.Document.parentWindow.AddWith Source, name, xsiType, length, precision, attrScale, nullable
        
    Next
End Sub

Public Function SaveIn(filepath As String) As Boolean
    Dim doc As New MSXML2.DOMDocument
    Dim rowNode As IXMLDOMNode
    Dim recordNode As IXMLDOMNode
    Dim oDic As Scripting.Dictionary
    Dim oDicRecord As Scripting.Dictionary
    Dim key As Variant
    Dim oNode As IXMLDOMNode
    Dim oAttr As IXMLDOMAttribute
    Dim sType As String
    Dim cont As Integer
    Dim sAttr As String
    
    cont = 1
    
    doc.Load App.Path & "\" & BCP_XML_BASE_FILE
    
    
    
    Dim i As Long
    For i = 0 To doc.documentElement.childNodes.length - 1
        If doc.documentElement.childNodes(i).nodeType = NODE_ELEMENT Then
            If doc.documentElement.childNodes(i).baseName = TAG_ROW Then
                Set rowNode = doc.documentElement.childNodes(i)
            ElseIf doc.documentElement.childNodes(i).baseName = TAG_RECORD Then
                Set recordNode = doc.documentElement.childNodes(i)
            End If
        End If
    Next
    
    Set oDic = GetFields()
    Set oDicRecord = oDic
    If oDic Is Nothing Then
        MensagemValidacao
        ValMessage GetLastErrorsWbRecord()
        Exit Function
    End If
    
    For Each key In oDic.Keys
         
        Set oNode = doc.createNode(MSXML2.NODE_ELEMENT, TAG_FIELD, DEFAULT_SCHEMA)
        Set oAttr = doc.createAttribute(ATTR_ID)
        oAttr.nodeValue = key
        oNode.Attributes.setNamedItem oAttr
        Set oAttr = doc.createAttribute(ATTR_ID)
        oAttr.nodeValue = key
        sType = oDic(key)(ATTR_TYPE)
        If sType = "" Or sType = EMPTY_CBO_STRING Then
            MensagemValidacao
            ValMessage "O atributo xsi:type é requerido."
            Exit Function
        End If
        Set oAttr = doc.createAttribute("xsi:" & ATTR_TYPE)
        oAttr.nodeValue = sType
        oNode.Attributes.setNamedItem oAttr
        
        If InStr(sType, "Char") > 0 Then
            If oDic(key)(ATTR_COLLATION) <> EMPTY_CBO_STRING Then
               Set oAttr = doc.createAttribute(ATTR_COLLATION)
               oAttr.nodeValue = oDic(key)(ATTR_COLLATION)
               oNode.Attributes.setNamedItem oAttr
            End If
            
            sAttr = oDic(key)(ATTR_MAX_LENGTH)
            If IsNumeric(sAttr) Then
                Set oAttr = doc.createAttribute(ATTR_MAX_LENGTH)
                oAttr.nodeValue = sAttr
                oNode.Attributes.setNamedItem oAttr
            End If
        End If
        If InStr(sType, "Fixed") > 0 Then
            sAttr = oDic(key)(ATTR_LENGTH)
            If IsNumeric(sAttr) Then
                Set oAttr = doc.createAttribute(ATTR_LENGTH)
                oAttr.nodeValue = sAttr
                oNode.Attributes.setNamedItem oAttr
            Else
                ValMessage "Field xsi:type " & sType & " requer um atributo " & ATTR_LENGTH & " numérico"
            End If
        End If
        If InStr(sType, "Prefix") > 0 Then
            
            sAttr = oDic(key)(ATTR_PREFIX_LENGTH)
            If IsNumeric(sAttr) Then
                Set oAttr = doc.createAttribute(ATTR_PREFIX_LENGTH)
                oAttr.nodeValue = sAttr
                oNode.Attributes.setNamedItem oAttr
            Else
                ValMessage "Field xsi:type " & sType & " requer um atributo " & ATTR_PREFIX_LENGTH & " numérico"
            End If
            
            sAttr = oDic(key)(ATTR_MAX_LENGTH)
            If IsNumeric(sAttr) Then
                Set oAttr = doc.createAttribute(ATTR_MAX_LENGTH)
                oAttr.nodeValue = sAttr
                oNode.Attributes.setNamedItem oAttr
            End If
            
        End If
        If InStr(sType, "Term") > 0 Then
            sAttr = oDic(key)(ATTR_TERMINATOR)
            If sAttr <> "" Then
                Set oAttr = doc.createAttribute(ATTR_TERMINATOR)
                oAttr.nodeValue = sAttr
                oNode.Attributes.setNamedItem oAttr
            Else
                ValMessage "Field xsi:type " & sType & " requer um atributo " & ATTR_TERMINATOR & " não vazio"
            End If
        End If

        recordNode.appendChild oNode
        If cont < oDic.Count Then
            recordNode.appendChild doc.createTextNode(vbNewLine & String(2, vbTab))
        Else
            recordNode.appendChild doc.createTextNode(vbNewLine & String(1, vbTab))
        End If
        cont = cont + 1
    Next
    cont = 1
    Set oDic = GetColumns()
    
    If oDic Is Nothing Then
        MensagemValidacao
        ValMessage GetLastErrorsWbRecord()
        Exit Function
    End If
    
    For Each key In oDic.Keys
        Set oNode = doc.createNode(MSXML2.NODE_ELEMENT, TAG_COLUMN, DEFAULT_SCHEMA)
        sAttr = key
        If sAttr = "" Then
            MensagemValidacao
            ValMessage "O atributo SOURCE é requerido"
            Exit Function
        Else
            Set oAttr = doc.createAttribute(ATTR_SOURCE)
            oAttr.nodeValue = sAttr
            oNode.Attributes.setNamedItem oAttr
        End If
        
        If Not oDicRecord.Exists(key) Then
            ValMessage "Column: Source: " & key & " O atributo SOURCE não está mapeado para um ID existente."
        End If
        
        sAttr = Trim(oDic(key)(ATTR_NAME))
        
        If sAttr <> "" Then
            Set oAttr = doc.createAttribute(ATTR_NAME)
            oAttr.nodeValue = sAttr
            oNode.Attributes.setNamedItem oAttr
        Else
            MensagemValidacao
            ValMessage "O atributo NAME é requerido"
            Exit Function
        End If
        
        sType = Trim(oDic(key)(ATTR_TYPE))
        If sType <> "" And sType <> EMPTY_CBO_STRING Then
            Set oAttr = doc.createAttribute("xsi:" & ATTR_TYPE)
            oAttr.nodeValue = sType
            oNode.Attributes.setNamedItem oAttr
        End If
        
        sAttr = Trim(oDic(key)(ATTR_LENGTH))
        If IsNumeric(sAttr) Then
            Set oAttr = doc.createAttribute(ATTR_LENGTH)
            oAttr.nodeValue = sAttr
            oNode.Attributes.setNamedItem oAttr
        ElseIf sAttr <> "" Then
            ValMessage "Columns: atributo LENGTH deveria ser um inteiro. Não adicionado"
        End If
        
        sAttr = Trim(oDic(key)(ATTR_LENGTH))
        If IsNumeric(sAttr) Then
            Set oAttr = doc.createAttribute(ATTR_LENGTH)
            oAttr.nodeValue = sAttr
            oNode.Attributes.setNamedItem oAttr
        ElseIf sAttr <> "" Then
            ValMessage "Columns: atributo LENGTH deveria ser um inteiro. Não adicionado."
        End If
        
        sAttr = Trim(oDic(key)(ATTR_PRECISION))
        If IsNumeric(sAttr) Then
            Set oAttr = doc.createAttribute(ATTR_PRECISION)
            oAttr.nodeValue = sAttr
            oNode.Attributes.setNamedItem oAttr
        ElseIf sAttr <> "" Then
            ValMessage "Columns: atributo PRECISION deveria ser um inteiro. Não adicionado."
        End If
        
        sAttr = Trim(oDic(key)(ATTR_SCALE))
        If IsNumeric(sAttr) Then
            Set oAttr = doc.createAttribute(ATTR_SCALE)
            oAttr.nodeValue = sAttr
            oNode.Attributes.setNamedItem oAttr
        ElseIf sAttr <> "" Then
            ValMessage "Columns: atributo SCALE deveria ser um inteiro. Não adicionado."
        End If
        
        sAttr = Trim(oDic(key)(ATTR_NULLABLE))
        If sAttr <> EMPTY_CBO_STRING Then
            Set oAttr = doc.createAttribute(ATTR_NULLABLE)
            oAttr.nodeValue = sAttr
            oNode.Attributes.setNamedItem oAttr
        End If
        
        rowNode.appendChild oNode
        If cont < oDic.Count Then
            rowNode.appendChild doc.createTextNode(vbNewLine & String(2, vbTab))
        Else
            rowNode.appendChild doc.createTextNode(vbNewLine & String(1, vbTab))
        End If
        cont = cont + 1
    Next
    
    doc.save filepath
    SaveIn = True
End Function

Public Sub ValMessage(message As String)
    txtLog.Text = txtLog.Text & message & vbNewLine
End Sub

Public Function GetLastErrorsWbRecord() As String
    GetLastErrorsWbRecord = wbRecord.Document.parentWindow.lastErrorMessage
End Function

Public Function GetLastErrorsWbRow() As String
    GetLastErrorsWbRow = wbRow.Document.parentWindow.lastErrorMessage
End Function

Public Sub MensagemValidacao()
    
    MsgBox "O documento apresenta erros que impedem o arquivo de ser salvo. Confira o log de validação." & _
            vbNewLine & "Corrija os erros e tente salvar novamente."
End Sub


Private Sub mnSalvar_Click()
    Dim fso As New Scripting.FileSystemObject
    If fileLoaded And Len(filepath) <> 0 And fso.FolderExists(Left(filepath, Len(filepath) - InStrRev(filepath, "\"))) Then
        If SaveIn(filepath) Then
            MsgBox "Salvo com sucesso."
        End If
    Else
        SalvarComo
    End If
End Sub

Public Sub SalvarComo()
On Error GoTo CancelError
        CommonDialog1.ShowSave
        If CommonDialog1.FileName <> "" Then
            If SaveIn(CommonDialog1.FileName) Then
                fileLoaded = True
                filepath = CommonDialog1.FileName
                MsgBox "Salvo com sucesso."
            End If
        End If
CancelError:
End Sub

Private Sub mnSalvarComo_Click()
    SalvarComo
End Sub
