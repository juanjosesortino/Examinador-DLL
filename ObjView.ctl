VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ObjView 
   ClientHeight    =   3735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6810
   ScaleHeight     =   3735
   ScaleWidth      =   6810
   Begin MSComctlLib.TreeView tvObjs 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   5953
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      Appearance      =   1
   End
End
Attribute VB_Name = "ObjView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Default Property Values:
Const m_def_Count = 0
'Property Variables:
Dim m_Col   As Collection
Dim m_Count As Long
Dim nRoot   As Node

Private ix  As Integer
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
    Set m_Col = New Collection
    ix = 0
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Add(Obj As Object, objName As String)
    Dim objData As cObjudt
    ix = ix + 1
    Set objData = New cObjudt
    Set objData.Obj = Obj
    objData.ID = objName & ix
    m_Col.Add objData, objName & ix
    CreateNode objData.Obj, objName & ix
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Remove(objName As String)
     m_Col.Remove objName
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,1,2,0
Public Property Get Count() As Long
Attribute Count.VB_MemberFlags = "400"
    Count = m_Count
End Property

Public Property Let Count(ByVal New_Count As Long)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_Count = New_Count
    PropertyChanged "Count"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Count = m_def_Count
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_Count = PropBag.ReadProperty("Count", m_def_Count)
    Set tvObjs.Font = PropBag.ReadProperty("Font", Ambient.Font)
End Sub

Private Sub UserControl_Resize()
    Const cW! = 3000
    Const cH! = 1500
    Static oldH!, oldW!
    Dim diff!
    
    With UserControl
        If oldH = 0 Then oldH = cH
        If oldW = 0 Then oldW = cW
        If .Height < cH Then
            .Height = cH
        End If
            diff = .Height - oldH - 800
            tvObjs.Height = tvObjs.Height + diff

            
        
        
        If .Width < cW Then
            .Width = cW
        End If
            diff = .Width - oldW - 800
            tvObjs.Width = tvObjs.Width + diff
        
        oldH = .Height
        oldW = .Width
    End With
    
End Sub

Private Sub UserControl_Terminate()
    Dim i  As Long
    For i = 1 To m_Col.Count
        m_Col.Remove 1
    Next
    Set m_Col = Nothing
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Count", m_Count, m_def_Count)
    Call PropBag.WriteProperty("Font", tvObjs.Font, Ambient.Font)
End Sub

Private Function CreateNode(Obj As Object, objName As String) As Node
    Dim nCur As Node
    If nRoot Is Nothing Then
        Set nRoot = tvObjs.Nodes.Add(, , objName, objName)
        Set nCur = nRoot
    Else
        Set nCur = tvObjs.Nodes.Add(nRoot, tvwNext, objName, objName)
    End If
    
    AddNode2 Obj, nCur
    

End Function
Private Function ReadFlag(vFlag As Variant) As String
    Dim sTmp As String
    Dim lType As Long, lSubType As Long
    lType = VarType(vFlag)
    If (lType And vbArray) Then
        Dim ub As Long, lb As Long, i As Long
        ub = UBound(vFlag)
        lb = LBound(vFlag)
        For i = lb To ub
            sTmp = sTmp + ReadFlag(vFlag(i)) + "$$"
        Next
    Else
        Select Case lType
        Case vbVariant
            sTmp = ReadFlag(vFlag)
        Case vbObject
            sTmp = "Object (" + TypeName(vFlag) + ")"
        Case vbVariant
        Case vbBoolean, vbLong, vbInteger, vbByte, vbCurrency, vbDate
            sTmp = CStr(vFlag)
        Case vbString
            sTmp = CStr(vFlag)
        End Select

    End If
            ReadFlag = sTmp
End Function
Public Sub Clear()
    tvObjs.Nodes.Clear
    Dim i  As Long
    For i = 1 To m_Col.Count
        m_Col.Remove 1
    Next
    Set nRoot = Nothing
    ix = 0
End Sub
Public Sub Refresh()
    tvObjs.Nodes.Clear
    Set nRoot = Nothing
    Dim Obj As cObjudt
        
    For Each Obj In m_Col
        CreateNode Obj.Obj, Obj.ID
    Next
    
End Sub
Public Sub RefreshOBJ(Obj As Object, objName As String)
    Dim objData As cObjudt
    
    Set objData = New cObjudt
    m_Col.Remove objName
    tvObjs.Nodes.Remove objName

    If tvObjs.Nodes.Count = 0 Then
        Set nRoot = Nothing
    End If
    
    Set objData.Obj = Obj
    objData.ID = objName
    m_Col.Add objData, objName
    
    

End Sub
Private Sub AddNodeText(nCurNode As Node, v As Variant)
    Dim i As Long
             
    i = VarType(v)
    If i And vbArray Then
        nCurNode.Text = nCurNode.Text + ": (Array)"
        
        Dim lb As Long, ub As Long, j As Long
        Dim n As Node
        lb = LBound(v)
        ub = UBound(v)
        Set n = tvObjs.Nodes.Add(nCurNode, tvwChild, , "Lower Bound: " + CStr(lb))
        Set n = tvObjs.Nodes.Add(nCurNode, tvwChild, , "Upper Bound: " + CStr(ub))
        
        For j = lb To ub
            Set n = tvObjs.Nodes.Add(nCurNode, tvwChild, , "Item(" + CStr(j) + ") : ")
            AddNodeText n, v(j)
        Next
        
    
    Else
        
        Select Case i
        Case vbSingle, vbDouble, vbInteger, vbByte, vbLong, vbCurrency, vbBoolean, vbString
            nCurNode.Text = nCurNode.Text + ": " + CStr(v)
        
        Case vbObject
            Dim sType As String
            sType = TypeName(v)
            nCurNode.Text = nCurNode.Text + " (Object: " + sType + ")"
            If v Is Nothing Then
                Exit Sub
            End If
''           If IsCollection(v) Then
''                For j = 1 To v.Count
''                    Set n = tvObjs.Nodes.Add(nCurNode, tvwChild, , "Item (" + CStr(j) + ") : ")
''                    ADDNODE v(
'                    AddNodeText n, v.Item(j)
'
'                Next
'            End If

        Case vbVariant
            AddNodeText nCurNode, v
        End Select
    End If
End Sub
Private Function FindEnumString(lVal As Long, msi As TLI.Members) As String
    Dim l As Long
    Dim mi As TLI.MemberInfo
    For Each mi In msi
        If lVal = mi.Value Then
            FindEnumString = mi.Name
        End If
    Next
End Function
Private Function IsCollection(v As Variant) As Boolean
    Dim tlApp As TLI.TLIApplication
    Dim tlIrf As TLI.InterfaceInfo
    Dim tlmi As TLI.MemberInfo
    Dim Obj As Object
    If IsObject(v) Then
        If TypeName(v) = "Collection" Then
            IsCollection = True
        End If
        Set Obj = v
        Set tlApp = New TLI.TLIApplication
        Set tlIrf = tlApp.InterfaceInfoFromObject(Obj)
        
        For Each tlmi In tlIrf.Members
            If LCase$(tlmi.Name) = "item" Then
                IsCollection = True
                Exit For
            End If
        Next
    End If
    
    Set tlmi = Nothing
    Set tlIrf = Nothing
    Set tlApp = Nothing
End Function
Private Sub AddNode2(Obj As Object, nRelNode As Node)
On Error Resume Next
    Dim tliApp As TLI.TLIApplication
    Dim tliIrf As TLI.InterfaceInfo
    Dim n As Node
    Dim v As Variant
    If Obj Is Nothing Then
        Exit Sub
    End If
    Set tliApp = New TLI.TLIApplication
    Set tliIrf = tliApp.InterfaceInfoFromObject(Obj)
    
    If IsCollection(Obj) Then
        Dim i As Long, nn As Node
        Set n = tvObjs.Nodes.Add(nRelNode, tvwChild, , tliIrf.Name + "(Collection)")
        Set nn = tvObjs.Nodes.Add(n, tvwChild, , "Count: " + CStr(Obj.Count))
        For i = 1 To Obj.Count
            Set nn = tvObjs.Nodes.Add(n, tvwChild, , "Item(" + CStr(i) + "): ")
            If IsObject(Obj.Item(i)) Then
                AddNode2 Obj.Item(i), nn
            End If
                AddNodeText nn, Obj.Item(i)
        Next
    Else
        Dim tliMi As TLI.MemberInfo '', v As Variant
        For Each tliMi In tliIrf.Members
            If (tliMi.InvokeKind = INVOKE_PROPERTYGET) And (tliMi.Parameters.Count = 0) Then
                Set n = tvObjs.Nodes.Add(nRelNode, tvwChild, , tliMi.Name)
                If IsObject(CallByName(Obj, tliMi.Name, VbGet)) Then
                    Dim oo As Object
                    Set v = CallByName(Obj, tliMi.Name, VbGet)
                    Set oo = v
                    AddNode2 oo, n
                    Set oo = Nothing
                  
                Else
                     v = CallByName(Obj, tliMi.Name, VbGet)
                     If Not (tliMi.ReturnType.TypeInfo Is Nothing) Then
                        v = FindEnumString(CLng(v), tliMi.ReturnType.TypeInfo.Members)
                     End If
                End If
                AddNodeText n, v
''                If (IsObject(v)) Then
''                    Dim oo As Object
''                    Set oo = v
''                End If
            End If
        Next
        
        Set tliMi = Nothing
    End If
    
    Set tliIrf = Nothing
    Set tliApp = Nothing
    
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=tvObjs,tvObjs,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = tvObjs.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set tvObjs.Font = New_Font
    PropertyChanged "Font"
End Property

