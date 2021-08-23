VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmExaminadorDLL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Examinador DLL"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12585
   Icon            =   "frmExaminadorDLL.frx":0000
   LinkTopic       =   "frmExaminadorDLL"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   12585
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab 
      Height          =   7785
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   12585
      _ExtentX        =   22199
      _ExtentY        =   13732
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   5
      TabHeight       =   529
      TabCaption(0)   =   "Clases"
      TabPicture(0)   =   "frmExaminadorDLL.frx":054A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ObjView"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Buscador"
      TabPicture(1)   =   "frmExaminadorDLL.frx":0566
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1(1)"
      Tab(1).Control(1)=   "ListViewFind"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame1 
         Caption         =   "Buscar"
         Height          =   885
         Index           =   1
         Left            =   -74940
         TabIndex        =   17
         Top             =   420
         Width           =   12450
         Begin VB.CommandButton cmdBuscar 
            Caption         =   "Buscar"
            Height          =   525
            Left            =   10860
            TabIndex        =   18
            Top             =   240
            Width           =   1245
         End
         Begin VB.TextBox txtTexto 
            Height          =   405
            Left            =   120
            TabIndex        =   1
            Top             =   300
            Width           =   10425
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Seleccionador de DLL y Clase"
         Height          =   885
         Index           =   0
         Left            =   60
         TabIndex        =   13
         Top             =   360
         Width           =   5580
         Begin VB.ComboBox cmbDLL 
            Height          =   315
            Left            =   210
            Sorted          =   -1  'True
            TabIndex        =   0
            Text            =   "DLL"
            Top             =   330
            Width           =   2500
         End
         Begin VB.ComboBox cmbClase 
            Height          =   315
            Left            =   2940
            Sorted          =   -1  'True
            TabIndex        =   14
            Text            =   "Clase"
            Top             =   330
            Width           =   2500
         End
      End
      Begin VB.CommandButton Command 
         Height          =   645
         Left            =   11820
         Picture         =   "frmExaminadorDLL.frx":0582
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   480
         Width           =   615
      End
      Begin VB.OptionButton Option4 
         Height          =   195
         Left            =   -72210
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   810
         Width           =   255
      End
      Begin VB.OptionButton Option5 
         Height          =   195
         Left            =   -71190
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   810
         Width           =   255
      End
      Begin VB.ComboBox cmbUsuarios 
         Height          =   315
         Left            =   -74940
         TabIndex        =   3
         Top             =   390
         Width           =   3435
      End
      Begin MSComctlLib.ListView ListViewTablas 
         Height          =   6150
         Left            =   -74940
         TabIndex        =   6
         Top             =   750
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   10848
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "TABLA"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "MB      "
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "EXTENTS"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "INITIAL_EXT"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "NEXT_EXT"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "MAX_EXT"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView ListViewConexion 
         Height          =   6600
         Left            =   -74970
         TabIndex        =   7
         Top             =   330
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   11642
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "OSuser"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Username"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Machine"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Program"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "SID"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Serial"
            Object.Width           =   1411
         EndProperty
      End
      Begin MSComctlLib.ListView ListViewTablespaces 
         Height          =   6600
         Left            =   -74970
         TabIndex        =   8
         Top             =   330
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   11642
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Tablespace"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "MB Tamaño"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "MB Usados"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "MB Libres"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Fichero de datos"
            Object.Width           =   5468
         EndProperty
      End
      Begin MSComctlLib.ListView ListViewSQL 
         Height          =   6600
         Left            =   -74970
         TabIndex        =   9
         Top             =   330
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   11642
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Programa"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Fecha/Hora"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Usuario"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "SQL"
            Object.Width           =   7056
         EndProperty
      End
      Begin ExaminadorDLL.ObjView ObjView 
         Height          =   7335
         Left            =   5640
         TabIndex        =   12
         Top             =   330
         Width           =   6825
         _ExtentX        =   12039
         _ExtentY        =   12938
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ListView ListView 
         Height          =   6420
         Left            =   60
         TabIndex        =   15
         Top             =   1260
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   11324
         SortKey         =   2
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Member"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Type"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "VarType"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Valor"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSComctlLib.ListView ListViewFind 
         Height          =   6390
         Left            =   -74940
         TabIndex        =   16
         Top             =   1320
         Width           =   12465
         _ExtentX        =   21987
         _ExtentY        =   11271
         SortKey         =   2
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Member"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Type"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "VarType"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Valor"
            Object.Width           =   8819
         EndProperty
      End
      Begin VB.Label lblUsuario 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -71310
         TabIndex        =   10
         Top             =   420
         Width           =   4935
      End
   End
End
Attribute VB_Name = "frmExaminadorDLL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmExaminadorDLL
' DateTime  : 08/2013
' Author    : Juan José Sortino
' Purpose   : Examinar propiedades de objetos BO
'---------------------------------------------------------------------------------------

'***********************************************************************
' Constantes Propias
'***********************************************************************
Private Const Si                  As String = "Sí"
Private Const No                  As String = "No"
Private Const NullString          As String = ""
Private Const UNKNOWN_ERRORSOURCE As String = "[Fuente de Error Desconocida]"
Private Const KNOWN_ERRORSOURCE   As String = "[Fuente de Error Conocida]"
'***********************************************************************

Option Explicit

'OLE Automation INVOKEKIND values
Public Enum InvokeKinds

  'OLE Automation INVOKEKIND values
  INVOKE_UNKNOWN = 0  '&H0

  'OLE Automation INVOKEKIND values
  INVOKE_FUNC = 1  '&H1

  'OLE Automation INVOKEKIND values
  INVOKE_PROPERTYGET = 2  '&H2

  'OLE Automation INVOKEKIND values
  INVOKE_PROPERTYPUT = 4  '&H4

  'OLE Automation INVOKEKIND values
  INVOKE_PROPERTYPUTREF = 8  '&H8

  'Special value for TLI
  INVOKE_EVENTFUNC = 16  '&H10

  'Special value for TLI
  INVOKE_CONST = 32  '&H20
  VT_VOID = 24
End Enum

Private ErrorLog            As ErrType

Private mvarControlData     As DataShare.udtControlData         'información de control
Private strState            As String * 50000

Private tliApp              As Object
Private tliApp2             As Object
Private objObjeto           As Object
Private bCargando           As Boolean
Private bBuscando           As Boolean
Private sngCoordenadaX      As Single
Private sngCoordenadaY      As Single
Private itmX                As ListItem
Private bStop               As Boolean

Private Sub Command_Click()
   ObjView.Clear
End Sub

Private Sub Form_Load()

10       On Error GoTo gesterr
         
20       bCargando = True
         
30       Cargar_cmbDLL
40       Cargar_cmbClase
         
50       bCargando = False
         
60       Exit Sub

gesterr:
70       Me.MousePointer = vbNormal
80       MsgBox "[Form_Load]" & vbCrLf & Err.Description & Erl
End Sub

Private Sub Cargar_cmbDLL()
   
   cmbDLL.AddItem "AlgStdFunc"
   
   cmbDLL.AddItem "BOFiscal"
   cmbDLL.AddItem "BOGeneral"
   cmbDLL.AddItem "BOContabilidad"
   cmbDLL.AddItem "BOSeguridad"
   cmbDLL.AddItem "BOProduccion"
   cmbDLL.AddItem "BOGesCom"
   cmbDLL.AddItem "BOCereales"
   
   cmbDLL.AddItem "DSCereales"
   cmbDLL.AddItem "DSGescom"
   cmbDLL.AddItem "DSProduccion"
   cmbDLL.AddItem "DSFiscal"
   cmbDLL.AddItem "DSContabilidad"
   cmbDLL.AddItem "DSGeneral"
   cmbDLL.AddItem "DSSeguridad"
   cmbDLL.AddItem "SPCereales"

   cmbDLL.ListIndex = 0
End Sub
Private Sub cmbDLL_Click()
   If bCargando Or bBuscando Then Exit Sub
   Cargar_cmbClase
End Sub
Private Sub Cargar_cmbClase()
         Dim tlibi     As Object
         Dim ti        As Object
         
10       On Error GoTo gesterr
         
20       Set tliApp = CreateObject("TLI.TLIApplication")
            
30       If InStr(cmbDLL.Text, "DS") Or InStr(cmbDLL.Text, "SP") Then
40          Set tlibi = tliApp.TypeLibInfoFromFile("C:\Archivos de programa\Algoritmo\server\" & cmbDLL.Text & ".dll")
50       Else
60          Set tlibi = tliApp.TypeLibInfoFromFile("C:\Archivos de programa\Algoritmo\" & cmbDLL.Text & ".dll")
70       End If

80       cmbClase.Clear
90       For Each ti In tlibi.TypeInfos
'            Debug.Print ti.AttributeMask
'            Debug.Print ti.Name
100         If ti.AttributeMask = 2 Or ti.AttributeMask = 11 Then
110            If Len(ti.Name) > 0 Then
120               cmbClase.AddItem ti.Name
130               DoEvents
140            End If
150         End If
160      Next
         
170      cmbClase.ListIndex = 0
            
180      Set tliApp = Nothing
190      Set tlibi = Nothing
200      Set ti = Nothing

210      Exit Sub

gesterr:
220      Set tliApp = Nothing
230      Set tlibi = Nothing
240      Set ti = Nothing
         
250      Me.MousePointer = vbNormal
260      MsgBox "[Cargar_cmbClase]" & vbCrLf & Err.Description & Erl
End Sub

Private Sub cmbClase_Click()
         Dim m_clsInterface As Variant 'InterfaceInfo
         Dim lMember        As Variant 'MemberInfo
         Dim itmX           As Variant
         
10       On Error GoTo gesterr
         
20       Set objObjeto = CreateObject(cmbDLL.Text & "." & cmbClase.Text)
30       ObjView.Add objObjeto, cmbDLL.Text & "." & cmbClase.Text
         
40       Set tliApp = CreateObject("TLI.TLIApplication")
50       Set m_clsInterface = tliApp.InterfaceInfoFromObject(objObjeto)
         
60       ListView.ListItems.Clear
70       For Each lMember In m_clsInterface.Members
80          Set itmX = ListView.ListItems.Add
                  
90          itmX.SubItems(1) = lMember.Name
100         itmX.SubItems(2) = WhatIsIt(lMember)
110         itmX.SubItems(3) = WhatIsIt2(lMember.ReturnType.VarType)
            
120         If itmX.SubItems(2) = "Property Get" Or itmX.SubItems(2) = "Property Let" Then
130            If itmX.SubItems(1) <> "LocalVar" And itmX.SubItems(1) <> "ObjectIsLoaded" And itmX.SubItems(1) <> "ControlData" Then
140               itmX.ListSubItems(1).ForeColor = vbBlue
150               itmX.ListSubItems(2).ForeColor = vbBlue
160               itmX.ListSubItems(3).ForeColor = vbBlue
170            End If
180         End If
            
190         If itmX.SubItems(2) = "Method" Then
200            If itmX.SubItems(1) <> "QueryInterface" And itmX.SubItems(1) <> "GetTypeInfoCount" And _
                  itmX.SubItems(1) <> "GetTypeInfo" And itmX.SubItems(1) <> "GetIDsOfNames" And _
                  itmX.SubItems(1) <> "Invoke" _
                  Then
210               itmX.ListSubItems(1).ForeColor = vbRed
220               itmX.ListSubItems(2).ForeColor = vbRed
230               itmX.ListSubItems(3).ForeColor = vbRed
240            End If
250         End If
            
260         If itmX.SubItems(2) = "Function" Then
270            If itmX.SubItems(1) <> "AddRef" And itmX.SubItems(1) <> "Release" And InStr(itmX.SubItems(1), "CallMetod") = 0 _
                  Then
280               itmX.ListSubItems(1).ForeColor = vbMagenta
290               itmX.ListSubItems(2).ForeColor = vbMagenta
300               itmX.ListSubItems(3).ForeColor = vbMagenta
310            End If
320         End If
330         If itmX.SubItems(2) = "Property Set" Then
340            itmX.ListSubItems(1).ForeColor = QBColor(3)
350            itmX.ListSubItems(2).ForeColor = QBColor(3)
360            itmX.ListSubItems(3).ForeColor = QBColor(3)
370         End If
380      Next

390      Set objObjeto = Nothing
400      Set tliApp = Nothing
410      Set m_clsInterface = Nothing
420      Set itmX = Nothing
         
430      Exit Sub

gesterr:
440      Set objObjeto = Nothing
450      Set tliApp = Nothing
460      Set m_clsInterface = Nothing
470      Set itmX = Nothing
         
480      Me.MousePointer = vbNormal
490      MsgBox "[Cargar_cmbClase]" & vbCrLf & Err.Description & Erl
End Sub
Private Function WhatIsIt2(ByVal vCodigo As Variant) As String
    Select Case vCodigo
        Case 0
            WhatIsIt2 = "Class"
        Case 2
            WhatIsIt2 = "Integer"
        Case 3
            WhatIsIt2 = "Long"
        Case 4
            WhatIsIt2 = "Single"
        Case 5
            WhatIsIt2 = "Double"
        Case 7
            WhatIsIt2 = "Date"
        Case 8
            WhatIsIt2 = "String"
        Case 9
            WhatIsIt2 = "Object"
        Case 11
            WhatIsIt2 = "Boolean"
        Case 12
            WhatIsIt2 = "Variant"
        Case 24
            WhatIsIt2 = "Sub"
        Case Else
            WhatIsIt2 = vCodigo & " (Unknown)"
    End Select
End Function
Private Function WhatIsIt(ByVal lMember As Object) As String
    Select Case lMember.InvokeKind
        Case INVOKE_FUNC
            If lMember.ReturnType.VarType <> VT_VOID Then
                WhatIsIt = "Function"
            Else
                WhatIsIt = "Method"
            End If
        Case INVOKE_PROPERTYGET
            WhatIsIt = "Property Get"
        Case INVOKE_PROPERTYPUT
            WhatIsIt = "Property Let"
        Case INVOKE_PROPERTYPUTREF
            WhatIsIt = "Property Set"
        Case INVOKE_CONST
            WhatIsIt = "Const"
        Case INVOKE_EVENTFUNC
            WhatIsIt = "Event"
        Case Else
            WhatIsIt = lMember.InvokeKind & " (Unknown)"
    End Select
End Function

Private Sub ListView_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   sngCoordenadaX = X
   sngCoordenadaY = Y
End Sub

Private Sub ListView_Click()
   On Error Resume Next
   
   Set itmX = ListView.HitTest(sngCoordenadaX, sngCoordenadaY)
   If itmX Is Nothing Then Exit Sub
   
   Clipboard.Clear
   Clipboard.SetText itmX.SubItems(1) & " " & itmX.SubItems(2) & " " & itmX.SubItems(3)
End Sub

Private Sub cmdBuscar_Click()

10       On Error GoTo gesterr
         
20       txtTexto = Trim(txtTexto)
         
30       If Len(txtTexto.Text) = 0 Then Exit Sub
         
40       If cmdBuscar.Caption = "Buscar" Then
50          MousePointer = vbArrowHourglass
60          cmdBuscar.Caption = "Parar"
            
70          bStop = False
80          BuscarDLL
            
90          cmdBuscar.Caption = "Buscar"
100         Me.MousePointer = vbNormal
110      Else
120         bStop = True
130      End If
         
140      Exit Sub

gesterr:
150      Me.MousePointer = vbNormal
160      cmdBuscar.Caption = "Buscar"
170      MsgBox "[cmdBuscar_Click]" & vbCrLf & Err.Description & Erl
End Sub
Private Sub BuscarDLL()

         Dim m_clsInterface As Variant 'InterfaceInfo
         Dim lMember        As Variant 'MemberInfo
         Dim itmX           As Variant
         Dim ix             As Integer
         Dim tlibi          As Object
         Dim ti             As Object
            
10       On Error GoTo gesterr
         
20       Set tliApp = CreateObject("TLI.TLIApplication")
         
30       ListViewFind.ListItems.Clear
40       bBuscando = True
                        
50       For ix = 0 To cmbDLL.ListCount - 1
60          cmbDLL.ListIndex = ix
            
70          If InStr(cmbDLL.Text, "DS") Or InStr(cmbDLL.Text, "SP") Then
80             Set tlibi = tliApp.TypeLibInfoFromFile("C:\Archivos de programa\Algoritmo\server\" & cmbDLL.Text & ".dll")
90          Else
100            Set tlibi = tliApp.TypeLibInfoFromFile("C:\Archivos de programa\Algoritmo\" & cmbDLL.Text & ".dll")
110         End If

120         For Each ti In tlibi.TypeInfos
130            If ti.AttributeMask = 2 Then
140               If Len(ti.Name) > 0 Then
                     
150                  Set objObjeto = CreateObject(cmbDLL.Text & "." & ti.Name)
                     
160                  Set tliApp2 = CreateObject("TLI.TLIApplication")
170                  Set m_clsInterface = tliApp2.InterfaceInfoFromObject(objObjeto)
                     
180                  For Each lMember In m_clsInterface.Members
190                     If InStr(UCase(lMember.Name), UCase(txtTexto.Text)) > 0 Then
                        'If UCase(lMember.Name) = UCase(txtTexto.Text) Then
200                        Set itmX = ListViewFind.ListItems.Add
210                        itmX.SubItems(1) = lMember.Name
220                        itmX.SubItems(2) = WhatIsIt(lMember)
230                        itmX.SubItems(3) = WhatIsIt2(lMember.ReturnType.VarType)
240                        itmX.SubItems(4) = cmbDLL.Text & "." & ti.Name
                           
250                        If itmX.SubItems(2) = "Property Get" Or itmX.SubItems(2) = "Property Let" Then
260                           If itmX.SubItems(1) <> "LocalVar" And itmX.SubItems(1) <> "ObjectIsLoaded" And itmX.SubItems(1) <> "ControlData" Then
270                              itmX.ListSubItems(1).ForeColor = vbBlue
280                              itmX.ListSubItems(2).ForeColor = vbBlue
290                              itmX.ListSubItems(3).ForeColor = vbBlue
300                              itmX.ListSubItems(4).ForeColor = vbBlue
310                           End If
320                        End If
                           
330                        If itmX.SubItems(2) = "Method" Then
340                           If itmX.SubItems(1) <> "QueryInterface" And itmX.SubItems(1) <> "GetTypeInfoCount" And _
                                 itmX.SubItems(1) <> "GetTypeInfo" And itmX.SubItems(1) <> "GetIDsOfNames" And _
                                 itmX.SubItems(1) <> "Invoke" _
                                 Then
350                              itmX.ListSubItems(1).ForeColor = vbRed
360                              itmX.ListSubItems(2).ForeColor = vbRed
370                              itmX.ListSubItems(3).ForeColor = vbRed
380                              itmX.ListSubItems(4).ForeColor = vbRed
390                           End If
400                        End If
                           
410                        If itmX.SubItems(2) = "Function" Then
420                           If itmX.SubItems(1) <> "AddRef" And itmX.SubItems(1) <> "Release" And InStr(itmX.SubItems(1), "CallMetod") = 0 _
                                 Then
430                              itmX.ListSubItems(1).ForeColor = vbMagenta
440                              itmX.ListSubItems(2).ForeColor = vbMagenta
450                              itmX.ListSubItems(3).ForeColor = vbMagenta
460                              itmX.ListSubItems(4).ForeColor = vbMagenta
470                           End If
480                        End If
490                        If itmX.SubItems(2) = "Property Set" Then
500                           itmX.ListSubItems(1).ForeColor = QBColor(3)
510                           itmX.ListSubItems(2).ForeColor = QBColor(3)
520                           itmX.ListSubItems(3).ForeColor = QBColor(3)
530                           itmX.ListSubItems(4).ForeColor = QBColor(3)
540                        End If
550                     End If
560                     DoEvents
570                     If bStop Then Exit Sub
580                  Next
                     
590                  DoEvents
600                  Set objObjeto = Nothing
610                  Set tliApp2 = Nothing
620                  Set m_clsInterface = Nothing
630               End If
640            End If
650         Next
            
660         DoEvents
670      Next ix
         
680      bBuscando = False
         
690      Set tliApp = Nothing
700      Set tliApp2 = Nothing
710      Set tlibi = Nothing
720      Set ti = Nothing
730      Set objObjeto = Nothing
740      Set tliApp = Nothing
750      Set m_clsInterface = Nothing
760      Set itmX = Nothing
         
770      Exit Sub

gesterr:
780      Set objObjeto = Nothing
790      Set tliApp = Nothing
800      Set m_clsInterface = Nothing
810      Set itmX = Nothing
         
820      Me.MousePointer = vbNormal
830      MsgBox "[BuscarDLL]" & vbCrLf & Err.Description & Erl
End Sub

'esta seria la sub que te muetra el dialogo para buscar el archivo y abrirlo, en la arte inferior esta el comentario de como se deberia guardar el archivo:
'
'
'
'Private Sub cmdExaminar_Click()
'
'   On Error GoTo gesterr
'
'   Dim varTemp As Variant
'
'   CDialog.ShowOpen
'
'   If Len(CDialog.FileName) = 0 Then Exit Sub
'   txtArchivo.Text = CDialog.FileName
'
'   lngfil = FreeFile()
'
'   Open txtArchivo.Text For Binary As lngfil
'   Get #lngfil, , varTemp
'   Close #lngfil
'
'   aByte = varTemp
'
'   PropBag.Contents = aByte
'
'   Exit Sub
'
'gesterr:
'
'   MsgBox "Error: " & Err.Number & vbCrLf & "Descripcion: " & Err.Description
''Forma que se debe guardar*******************************
''varTemp = pb.Contents
''Lo guarda en un archivo de texto.
''Open App.Path & "\" & sFic & ".txt" For Binary As #1
''Put #1, , varTemp
''Close #1
'End Sub
