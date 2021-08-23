Attribute VB_Name = "modInicio"
Option Explicit


'***********************************************************************
' Constantes Propias
'***********************************************************************
Public Const Si                  As String = "Sí"
Public Const No                  As String = "No"
Public Const NullString          As String = ""
Public Const UNKNOWN_ERRORSOURCE As String = "[Fuente de Error Desconocida]"
Public Const KNOWN_ERRORSOURCE   As String = "[Fuente de Error Conocida]"
'***********************************************************************
'Private Const MODULE_NAME        As String = "[ModInicio]"
Private ErrorLog                 As ErrType

Public Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Const SPI_GETWORKAREA = 48

'Private objTextBox As New AlgStdFunc.clsTextBoxEdit

'Public CUsuario                  As BOSeguridad.clsUsuario
Public CSysEnvironment           As AlgStdFunc.clsSysEnvironment
Public rstContextMenu            As ADODB.Recordset
Public rstMenu                   As ADODB.Recordset
Public rstVistasPersonalizadas   As ADODB.Recordset
Public rstVistasExportacion      As ADODB.Recordset
Public rstEmpresas               As ADODB.Recordset
Public RegistrySubKeys()         As Variant
Public SystemOptions()           As Variant
Public strSucursalElegida        As String

'/ deben ser visibles por frmMenu y frmMDIInicio
Public objSeguridad              As Object
Public objContabilidad           As Object
Public objFiscal                 As Object
Public objGesCom                 As Object
Public objGeneral                As Object
Public objCereales               As Object
Public objProduccion             As Object

'/ deben ser visibles por frmLlogin
Public bUserLogued               As Long

Public Enum EnumRegistrySubKeys
   Environment = 0
   DataBaseSettings = 1
   KeyMRUForms = 2
   MRUEmpresas = 3
   GridQueries = 4
   NavigationQueries = 5
   PrintQueries = 6
   QueryDBQueries = 7
   DataComboQueries = 8
   [_MAX_Value] = 8
End Enum

Public Enum EnumSystemOptions
   iCacheSize = 0             'valor del parámetro cachesize (registro del sistema)
   iZoom = 1                  'valor del Zoom por defecto en Vista Previa
   iFetchMode = 2             'indica el modo en el que vendran capturados los registros del server
   lngFetchLimit = 3          'si alFetchMode = 2, es el limite de registros recuperados en manera sincronica
   iFetchModeSearch = 4       'indica el modo en el que vendran capturados los registros del server (para la busqueda)
   lngFetchLimitSearch = 5    'si alFetchMode = 2, es el limite de registros recuperados en manera sincronica (para la busqueda)
   UseLocalCopy = 6           'Sí=Usa copias locales; No=Usa copias locales (Vista-Lista, Navegación e Impresión)
   UseLocalCopySearch = 7     'Sí=Usa copias locales; No=Usa copias locales (para la búsqueda)
   AskOldLocalCopy = 8        'Sí=Pregunta si usa copias locales desactualizadas;(Vista-Lista, Navegación e Impresión)
   UseMRUEnterprise = 9       'Si=recuerda las ultimas empresas;No=No recuerda
   MaxMRUForms = 10           'Dimension de la colecion MRUForms
   [_MAX_Value] = 10
End Enum

Public Enum ContextMenuEnum
   mnxNombre = 0
   mnxOrden = 1
   mnxForms = 2
   mnxCaption = 3
   mnxTarea = 4
   mnxClave = 5
End Enum


'Mensajes enviados por la Filter
Public Const FILTER_CALL_ADMIN = &H1
Public Const FILTER_QUERY_USER   As Long = &H2
Public Const FILTER_QUERY_CONTROLDATA   As Long = &H3

'  constantes para identificar los mensajes devueltos por Filter
Public Const MSG_CANCEL  As String = "CANCELFILTRO"
Public Const MSG_CONFIRM As String = "CONFIRMAFILTRO"
Public Const MSG_APPLY   As String = "APLICARFILTRO"

'  constantes para identificar los paneles del Status Bar
Public Const STB_PANEL1              As Integer = 1
Public Const STB_PANEL2              As Integer = 2
Public Const STB_PANEL3              As Integer = 3
Public Const STB_PANEL4              As Integer = 4

Public Enum alFetchMode
   alAsync = 1
   alSync = 2
   alTable = 3
End Enum

Public Const IX_CAMBIO_PWD          As Integer = 0
Public Const IX_ESTABLECER_EJERCI   As Integer = 1
Public Const IX_CAMBIO_LOGIN        As Integer = 2
Public Const IX_EMPRESAS            As Integer = 3
Public Const IX_SEPARA1             As Integer = 4
Public Const IX_LOG_ERRORES         As Integer = 5
Public Const IX_EDITOR_SQL          As Integer = 6
Public Const IX_SEPARA2             As Integer = 7
Public Const IX_BUSCAR_MENU         As Integer = 8
Public Const IX_BUSCAR_SIGUIENTE    As Integer = 9
Public Const IX_SEPARA3             As Integer = 10
Public Const IX_PROPIEDADES         As Integer = 11
Public Const IX_ARCHIVOS            As Integer = 12
Public Const IX_SEPARA4             As Integer = 13
Public Const IX_VER_FAVORITOS       As Integer = 14
Public Const IX_ORGANIZAR_FAVORITOS As Integer = 15
Public Const IX_SEPARA5             As Integer = 16
Public Const IX_OPCIONES            As Integer = 17 ' ---> ver aMenuTools(17)

Public aMenuTools(17)            As String                                                 'matriz elementos del menu Herramientas
Public aMenuWindow(2)            As String
Public aMenuHelp(4)              As String                                                'matriz elementos del menu Ayuda

Public mvarMDIForm               As MDIForm
Public MRUForms                  As New Collection

Public bRestart                  As Boolean    'sirve para saber si me estoy logeando como otro usuario

Public colInfoEmpresas As clsInfoEmpresas
Public objInfoEmpresa As clsInfoEmpresa

Private bStartUp  As Boolean

'
'  Definición de APIs
'
Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * 260
End Type
'-------------------------------------------------------
Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, _
ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long

Declare Function ProcessFirst Lib "kernel32.dll" Alias "Process32First" (ByVal hSnapshot As Long, _
uProcess As PROCESSENTRY32) As Long

Declare Function ProcessNext Lib "kernel32.dll" Alias "Process32Next" (ByVal hSnapshot As Long, _
uProcess As PROCESSENTRY32) As Long

Declare Function CreateToolhelpSnapshot Lib "kernel32.dll" Alias "CreateToolhelp32Snapshot" ( _
ByVal lFlags As Long, lProcessID As Long) As Long

Declare Function TerminateProcess Lib "kernel32.dll" (ByVal ApphProcess As Long, _
ByVal uExitCode As Long) As Long

Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
'
'
'

Sub Main()
      'Dim vValue As Variant
      'Dim s As String
      Dim objSPM  As DataShare.SPM
      Dim strUsaSubCas2009 As String
      'Dim iLinea  As Integer

10       On Error GoTo GestErr
   
   
         '
         ' Me fijo si es necesario actualizar AlgStar.exe
         '
         Dim strServerRoot As String
         Dim ServerDLLFolder As String
         Dim strServerFile As String
         Dim StrActualizarAlgStart As String
         '
         '  intenta establecer una conexón con el servicio
         '
20       Load frmTestServicio

30       Do While Not frmTestServicio.Connected
40          DoEvents
50       Loop
   
60       Unload frmTestServicio
  
      '
      '        Esto cambia el mensaje "Cambiar - Reintentar" por éste un poco más amigable
      '
70       App.OleRequestPendingMsgText = "El Servidor está ocupado procesando su requerimiento." & vbCrLf & vbCrLf & "Por favor, aguarde unos instantes."
80       App.OleRequestPendingMsgTitle = "Aplicaciones Algoritmo"
  
         'obtengo la ubicación en el server de la version del producto
90       strServerRoot = GetRegistryValue(HKEY_LOCAL_MACHINE, "Software\Algoritmo\Environment\Application Server", "Server Update Root", REG_SZ, "C:", True)
            '"Opciones\Usa Subclassing 2009;No"
91       strUsaSubCas2009 = GetRegistryValue(HKEY_LOCAL_MACHINE, "Software\Algoritmo", "Usa Subclassing 2009 (N/S)", REG_SZ, "N", True)
'92       UsaSubclassing2009 = IIf(UCase(strUsaSubCas2009) = "S", Si, No)
         
         
         'elimino los : del drive
100      strServerRoot = Replace(strServerRoot, ":", "")
   
         'directorio del server donde estan las DLL del Client"
110      ServerDLLFolder = AddBackslash(strServerRoot) & "Algoritmo\Componentes Client"
   
120      strServerFile = Dir(AddBackslash(ServerDLLFolder) & "AlgStart.exe", vbNormal)
130      StrActualizarAlgStart = GetSetting("Algoritmo", "AlgStart", "Actualizar", "No")
   
140      If Len(strServerFile) > 0 And StrActualizarAlgStart = Si Then
            '
            '  existe una nueva versión disponible de AlgStart.exe.
            '  Me traigo la nueva versión desde el server
            '
150         FileCopy AddBackslash(ServerDLLFolder) & "AlgStart.exe", AddBackslash(App.Path) & "AlgStart.exe"
            '
            ' pongo no para saltear evitar nuevas preguntas
            '
160         SaveSetting "Algoritmo", "AlgStart", "Actualizar", "No"

170         MsgBox "Se ha detectado una nueva versión de AlgStart.exe y la misma ha sido copiada en su puesto de trabajo." & vbCrLf & _
                   "Para que los cambios tengan efecto es necesario Reiniciar el Sistema Algoritmo"

180         End
190      End If
   
   
200      bStartUp = True
   
210      Set colInfoEmpresas = New clsInfoEmpresas
   
220      Set CUsuario = New BOSeguridad.clsUsuario
   
230      Set CSysEnvironment = New AlgStdFunc.clsSysEnvironment
   
240      SetRegistryEntries
   
250      ReadSystemOptions

         ' defino el application path para la clase clsEnvironment
260      SetAppPath App.Path
   
270      With ErrorLog
280         .Empresa = GetSPMProperty(DBSEmpresaPrimaria)
290         .Maquina = CSysEnvironment.Machine
300         .Aplicacion = App.EXEName
310      End With
   
   
         ' Leo el registro de la empresa primaria para ver si usa o no cache local
         '  ReadDefault
320      ReDim aKeys(1, 1)
330      aKeys(0, 0) = "Opciones\Utiliza cache local;" & No
340      aKeys(1, 0) = "Opciones\Permite Múltiples Instancias de la Aplicacion;" & No
350      Set objSPM = GetMyObject("DataShare.SPM")
360      objSPM.GetKeyValues objSPM.GetSPMProperty(DBSEmpresaPrimaria), aKeys
         'Control de Múltiples Instancias
370      If App.PrevInstance Then
380         Select Case aKeys(1, 1)
               Case "Avisar", "AVISAR", "avisar"
                  'Pregunta al usuario si quiere eliminar eventuales procesos colgados
                  Dim strCartel As String
390               strCartel = "Se ha detectado más de una Instancia de la Aplicación en Ejecución." & vbCrLf & vbCrLf & _
                              "Esto puede deberse a que algún proceso abortado debido a un error de la Aplicación." & vbCrLf & vbCrLf & _
                              "Si no está ejecutando otra Instancia de SoftCereal conteste Sí e ingresará normalmente." & vbCrLf & _
                              "En caso de tener otra Instancia válida, conteste No para permitir ejecutar más de una vez la Aplicación."
                        
400               If MsgBox(strCartel, vbYesNo, App.ProductName) = vbYes Then KillProcess "Inicio.exe", False
            
410            Case No
                  'Elimina automáticamente los Inicio.EXE que se estén ejecutando
420               KillProcess "Inicio.exe", False
430            Case Si
                  'No hace nada
440         End Select
450      End If
   
460      If IIf(IsNull(aKeys(0, 1)), No, aKeys(0, 1)) = No Then
            ' NO Usa cache
470         Set rstMenu = objSPM.GetSPMProperty(MNURecordset)
480         Set rstContextMenu = objSPM.GetSPMProperty(CMURecordset)

490         Set objSPM = Nothing
500      Else
   
510         Set objSPM = Nothing
      
            ' SI Usa cache
            'obtengo los recordset que van a usar las aplicaciones
520         If Dir(AddBackslash(App.Path) & "Cache", vbDirectory) = "" Then
               '
               '  creo la carpeta y guardo los archivos
               '
               Dim fso As Object
               Dim strRecursosFolder As String
   
530            Set fso = CreateObject("Scripting.FileSystemObject")
   
540            fso.CreateFolder AddBackslash(App.Path) & "Cache"

550            Set rstMenu = GetSPMProperty(MNURecordset)
560            rstMenu.Save AddBackslash(App.Path) & "Cache\Menu"

570            Set rstContextMenu = GetSPMProperty(CMURecordset)
580            rstContextMenu.Save AddBackslash(App.Path) & "Cache\MenuContextual"
               '
               ' recursos
               '
590            strRecursosFolder = GetRegistryValue(HKEY_LOCAL_MACHINE, "Software\Algoritmo\Environment", "Recursos")
               '
               '  creo la carpeta local
               '
600            fso.CopyFolder strRecursosFolder, AddBackslash(App.Path)
               '
               '  cambio el registro para que apunte al local
               '
610            SetRegistryValue HKEY_LOCAL_MACHINE, "Software\Algoritmo\Environment", "Recursos", REG_SZ, AddBackslash(App.Path) & "Recursos"
   
620            Set fso = Nothing
   
630         Else
               '
               '  los levanto
               '
640            Set rstMenu = New ADODB.Recordset
650            Set rstContextMenu = New ADODB.Recordset
   
660            rstMenu.Open AddBackslash(App.Path) & "Cache\Menu"
670            rstContextMenu.Open AddBackslash(App.Path) & "Cache\MenuContextual"
680         End If
   
690      End If
   
700      Set objSPM = Nothing
   
'710      LoadVistas
   
         'muestro el form MDI
720      frmMDIInicio.Show
         
730      Set mvarMDIForm = frmMDIInicio
   
         '  login usuario
740      FrmLogin.Show vbModal
         
745      LoadVistas
   
750      SetRegistryEntries CUsuario.Usuario
   
760      frmMDIInicio.stbMain.Panels(1).Text = "Cargando Opciones del Sistema ..."
   
         ' mostrar menú
770      frmMenu.Show
780      frmMenu.TreeView1(frmMenu.sst1.Tab).SetFocus
   
790      frmMDIInicio.stbMain.Panels(1).Text = CUsuario.Usuario
   
800      bStartUp = False
   
   
         ' Con esto registro el logoneo en Trans.Auditoria
         Dim ControlBuffer As String
810      ControlBuffer = AddAuditComment(ControlBuffer, "Login")
   
         Dim objAuditoria As DataAccess.clsAuditoriaDS
820      Set objAuditoria = GetMyObject("DataAccess.clsAuditoriaDS")
   
830      objAuditoria.WriteAudit "SIEMPRE_GRABA", 0, ControlBuffer
840      Set objAuditoria = Nothing
         ' FIN Con esto registro el logoneo en Trans.Auditoria
   
         ' Coloco el nombre de usuario utilizado en el servicio
         'Licencia_Transaccion_Auditoria = GetLicenciaTA(GetSPMProperty(DBSEmpresaPrimaria), CUsuario.Usuario, Maquina)
         'mvarMDIForm.tcpClient.SendData "Set Usuario;" & GetNombreMaquina(GetSPMProperty(DBSEmpresaPrimaria), CUsuario.Usuario) & ";" & Licencia_Transaccion_Auditoria
  
850      Exit Sub
   
GestErr:
860      Set objSPM = Nothing
870      Set objAuditoria = Nothing
   
880      LoadError ErrorLog, "Main" & Erl
   
890      If bStartUp Then
   
900         Select Case True

               Case CUsuario Is Nothing
910                  MsgBox "Se produjo un error en la linea " & Erl & vbCrLf & _
                            "Error en la creación del objeto Usuario durante la fase de Inicio del sistema. " & _
                            "Controle que todos sus componentes este correctamente registrados", vbExclamation, App.ProductName
920            Case Else
   
930                  MsgBox "Se produjo un error en la linea " & Erl & vbCrLf & _
                            "Error durante la inicialización del Sistema. Las posibles causas de este inconveniente podrían ser: " & vbCrLf & vbCrLf & _
                            "    - uno o mas Componentes no estan correctamente registrados en su PC Local" & vbCrLf & _
                            "    - el equipo Servidor no esta activo" & vbCrLf & _
                            "    - se produjo un error al intentar establecer un conexíon con el Paquete MTS" & vbCrLf & vbCrLf & _
                            "Intente Reiniciar el Paquete MTS en el servidor eliminando previamente todos los procesos mtx.exe visibles en el Administrador de Tareas del Servidor." & _
                            "Si el problema persiste, retroceda a la versión precedente. Si aún no ha podido iniciar el sistema contacte a su proveedor", vbExclamation, App.ProductName

940         End Select

950         End
960      Else
970         ShowErrMsg ErrorLog
980      End If
   
End Sub

Public Sub AutoIncr(ByVal strMenuKey As String)
'Dim iValue As Integer
'
'   'incremento el parametro en el registro de configuración para la estadistica de forms mas usados
'
'   iValue = GetRegistryValue(HKEY_LOCAL_MACHINE, RegistrySubKeys(EnumRegistrySubKeys.KeyMRUForms), strMenuKey, REG_DWORD, 0, False)
'   If iValue = 0 Then
'      'aún no ha sido grabado
'      SetRegistryValue HKEY_LOCAL_MACHINE, RegistrySubKeys(EnumRegistrySubKeys.KeyMRUForms), strMenuKey, REG_DWORD, 1
'   Else
'      SetRegistryValue HKEY_LOCAL_MACHINE, RegistrySubKeys(EnumRegistrySubKeys.KeyMRUForms), strMenuKey, REG_DWORD, iValue + 1
'   End If

End Sub

Private Sub LoadVistas()
      Dim i As Integer
      Dim sql  As String
      Dim rst  As ADODB.Recordset
      Dim mvarControlData  As DataShare.udtControlData

10       On Error GoTo GestErr

         ' Exportaciones a Excel
   
20       Set rstVistasExportacion = New ADODB.Recordset
   
30       With rstVistasExportacion

40          .Fields.Append "User", adVarChar, 20
50          .Fields.Append "MenuKey", adVarChar, 200
60          .Fields.Append "CodigoEstructura", adVarChar, 200
70          .Fields.Append "Column", adVarChar, 200
80          .Fields.Append "Key", adVarChar, 200
90          .Fields.Append "Orden", adNumeric, , adFldIsNullable
100         .Fields(rstVistasExportacion.Fields.Count - 1).Precision = 3
110         .Fields(rstVistasExportacion.Fields.Count - 1).NumericScale = 0
   
120         .Fields.Append "Visible", adVarChar, 2
   
130         .Open
      
140         mvarControlData = CUsuario.ControlData
      
            'Recupero las definiciones de exportacion del usuario
150         sql = " SELECT "
            sql = sql & "    NVL(TAB_CLAVE3, ' ') AS TAB_CLAVE3,"
            sql = sql & "    NVL(TAB_CLAVE4, 0) AS TAB_CLAVE4,"
            sql = sql & "    NVL(TAB_CLAVE5, ' ') AS TAB_CLAVE5,"
            sql = sql & "    NVL(TAB_VALOR, ' ') AS TAB_VALOR"
            sql = sql & "    FROM TABLAS  "
160         sql = sql & "    WHERE TABLAS.TAB_EMPRESA = '" & GetSPMProperty(DBSEmpresaPrimaria) & "' AND "
170         sql = sql & "          TABLAS.TAB_CLAVE1 = 'Exportar_A_Office' AND "
180         sql = sql & "          TABLAS.TAB_CLAVE2 = '" & mvarControlData.Usuario & "'"
190         sql = sql & "    ORDER BY TAB_CLAVE3, TO_NUMBER(TAB_CLAVE4) "
      
200         Set rst = Fetch(GetSPMProperty(DBSEmpresaPrimaria), sql)
      
210         Do While Not rst.EOF
220            rstVistasExportacion.AddNew
230            rstVistasExportacion.Fields("User").Value = mvarControlData.Usuario
240            rstVistasExportacion.Fields("MenuKey").Value = rst("TAB_CLAVE3").Value
250            rstVistasExportacion.Fields("Orden").Value = rst("TAB_CLAVE4").Value
         
260            Do While rstVistasExportacion.Fields("Orden").Value = CInt(rst("TAB_CLAVE4").Value)

270               If rst("TAB_CLAVE5").Value <> "Orden" Then
280                  rstVistasExportacion.Fields(rst("TAB_CLAVE5").Value) = rst("TAB_VALOR").Value
290               End If
            
300               rst.MoveNext
            
310               If rst.EOF Then Exit Do
320            Loop
         
330            rstVistasExportacion.Update
340         Loop

350         If Not rst Is Nothing Then
360            If rst.State <> adStateClosed Then rst.Close
370         End If
380         Set rst = Nothing
      
390         rstVistasExportacion.Sort = "MenuKey, CodigoEstructura, Orden, User "
400      End With

         ' Comienza el tema de las Vistas Personalizadas
   
410      On Error Resume Next
420      i = Len(Dir(LocalPath, vbDirectory))
430      Select Case Err.Number
            Case 52
440            MsgBox "No tiene permisos de lectura sobre la carpeta: " & LocalPath & vbCrLf & _
               "El sistema necesita de una carpeta temporal para almacenar las vistas personalizadas por usuario, " & vbCrLf & _
               "usted puede cambiar de carpeta desde el menú descolgable Herramientas/Opciones, en caso contrario " & vbCrLf & _
               "podrá trabajar normalmente pero sin hacer uso de dichas vistas personales."
450            Exit Sub
460      End Select
470      Err.Clear
   
480      On Error GoTo GestErr

490      Set rstVistasPersonalizadas = New ADODB.Recordset
   
500      If Len(Dir(LocalPath, vbDirectory)) = 0 Then
510         If MsgBox("La Carpeta Local " & LocalPath & " no existe. Desea configuar un nueva Carpeta para los archivos locales ?", vbQuestion + vbYesNo) = vbYes Then
520            frmOpciones.Show vbModal
530         End If
540      End If
   
550      With rstVistasPersonalizadas
560         If Len(Dir(LocalPath & "VistasPersonalizadas")) > 0 Then
570            .Open LocalPath & "VistasPersonalizadas", , , , adCmdFile
580         Else
590            .Fields.Append "User", adVarChar, 20
600            .Fields.Append "Application", adVarChar, 50
610            .Fields.Append "MenuKey", adVarChar, 50
620            .Fields.Append "Column", adVarChar, 50
630            .Fields.Append "Visible", adVarChar, 2
   
640            .Open
650            .Save LocalPath & "VistasPersonalizadas", adPersistADTG
660         End If
670      End With

         '
         ' cargo las vistas definidas para la exportacíon
         '
   
680      Exit Sub

GestErr:
690      LoadError ErrorLog, "LoadVistas " & Erl
700      ShowErrMsg ErrorLog
   
End Sub

'/
' Todas estas Subs y Funciones fueron tomadas del ModShare
'/

Public Function CallAdmin(ByVal ControlInfo As Variant, ByVal ControlData As Variant) As Long
'Dim hWndAdmin As Long

   '-- llamo al Form para la administracion del control activo
   '-- ControlInfo es del tipo ControlType
   
   On Error GoTo GestErr
   
   CallAdmin = mvarMDIForm.CallAdmin(ControlInfo.MenuKeyAdmin, ControlData)
   
   Exit Function
   
GestErr:
   LoadError ErrorLog, "CallAdmin"
   ShowErrMsg ErrorLog
End Function

Public Sub LoadError(ByRef ErrLog As ErrType, ByVal strSource As String)
Dim PropBag As PropertyBag

   ' carga la información del error en la variable ErrorLog
   SetError ErrLog, App.ProductName, strSource
   
   Set PropBag = New PropertyBag
   
   With PropBag
      .WriteProperty "ERR_EMPRESA", ErrLog.Empresa
      .WriteProperty "ERR_APLICACION", ErrLog.Aplicacion
      .WriteProperty "ERR_COMENTARIO", ErrLog.Comentario
      .WriteProperty "ERR_DESCRIPCION", ErrLog.Descripcion
      .WriteProperty "ERR_ERRORNATIVO", ErrLog.ErrorNativo
      .WriteProperty "ERR_FORM", ErrLog.Form
      .WriteProperty "ERR_MAQUINA", ErrLog.Maquina
      .WriteProperty "ERR_MODULO", ErrLog.Modulo
      .WriteProperty "ERR_NUMERROR", ErrLog.NumError
      .WriteProperty "ERR_SOURCE", ErrLog.Source
      .WriteProperty "ERR_USUARIO", ErrLog.Usuario
      .WriteProperty "WRITE_ERROR", ErrLog.WriteError
   End With
   
   TrapError PropBag.Contents
   Set PropBag = Nothing
   
   
End Sub
Public Sub SetError(ByRef ErrLog As ErrType, ByVal strModuleName As String, ByVal strSource As String)

   With ErrLog
   
      If Not CUsuario Is Nothing Then
         .Usuario = CUsuario.Usuario
      End If
      
      .Modulo = strSource
      .NumError = Err.Number
      .Source = Err.Source
      .Aplicacion = UCase(App.ProductName)
      .WriteError = Si
      
      If InStr(.Source, KNOWN_ERRORSOURCE) = 0 Then
         If InStr(.Source, UNKNOWN_ERRORSOURCE) = 0 Then
            .Source = UNKNOWN_ERRORSOURCE & vbCrLf & .Source
         End If
      Else
         .Source = Replace(.Source, KNOWN_ERRORSOURCE, NullString)
         .WriteError = False
      End If
      
      If InStr(.Source, strModuleName) > 0 Then
         .Source = .Source & vbCrLf & "[" & strSource & "]"
      Else
         .Source = .Source & vbCrLf & strModuleName & "[" & strSource & "]"
      End If
      
      .Descripcion = Err.Description

   End With

   
End Sub

Public Sub ShowErrMsg(ByRef ErrorLog As ErrType)
Dim iErrNumber         As Long                          ' numero de error (sin vbObjectError)
Dim bAlgError          As Boolean                       ' identifica un error de Algoritmo
'Dim ix                 As Integer
Dim strSource          As String
Dim n                  As Integer
Dim frmMsg             As frmMsgBox
Dim strMensaje         As String
Dim strDetalle         As String

   '  muestra en manera amigable un mensaje de error
   

   strSource = Trim(ErrorLog.Source)
   
   bAlgError = True
   n = InStr(strSource, UNKNOWN_ERRORSOURCE)
   If n > 0 Then
      ' es un error generado por alguna aplicacion de Algoritmo
      bAlgError = False
   End If
   
   strSource = Replace(strSource, UNKNOWN_ERRORSOURCE, NullString)
   
   If bAlgError Then
      ' errores de Algortimo
      iErrNumber = ErrorLog.NumError - vbObjectError
      Select Case iErrNumber
         Case Is < 10000
            ' warnings de Algortimo
            
            Set frmMsg = New frmMsgBox
            
            strMensaje = ErrorLog.Descripcion
            strDetalle = strSource
            
            frmMsg.ShowMsg App.ProductName, strMensaje, strDetalle, Warning
            
         Case 10000 To 20000
            'Errores Severos de Algoritmo
            
            Set frmMsg = New frmMsgBox
            
            strMensaje = ErrorLog.Descripcion
            frmMsg.ShowMsg App.ProductName, strMensaje, strDetalle, Error
            
         Case Else
         
            Set frmMsg = New frmMsgBox
            
            strMensaje = ErrorLog.Descripcion
            frmMsg.ShowMsg App.ProductName, strMensaje, strDetalle, Error
         
      End Select
   Else
      ' errores no generados por Algoritmo
             
         Set frmMsg = New frmMsgBox
          
         ErrorLog.Descripcion = Replace(ErrorLog.Descripcion, vbCr, NullString)
         ErrorLog.Descripcion = Replace(ErrorLog.Descripcion, vbLf, NullString)
          
         strMensaje = ErrorLog.Descripcion
         strDetalle = "Número     : " & ErrorLog.NumError & vbCrLf & strSource
         
         frmMsg.ShowMsg App.ProductName, strMensaje, strDetalle, Error
            
             
   End If
   
   ' una vez visualizado el mensaje de error, este viene limpiado
   With ErrorLog
      .Modulo = NullString
      .NumError = 0
      .Source = NullString
      .Descripcion = NullString
   End With
   
   Screen.MousePointer = vbDefault
   
End Sub

Public Sub CenterMDIActiveXChild(ByVal frmChild As Form)

   '--  centra el form MDIActiveX Child

   frmChild.Move (mvarMDIForm.ScaleWidth - frmChild.Width) / 2, (mvarMDIForm.ScaleHeight - frmChild.Height) / 2

End Sub

Public Sub CenterForm(ByRef frm As Form)
Dim r As RECT
Dim lRes As Long
Dim lw As Long
Dim lh As Long

   lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, r, 0)

   If lRes Then
      With r
         .Left = Screen.TwipsPerPixelX * .Left
         .Top = Screen.TwipsPerPixelY * .Top
         .Right = Screen.TwipsPerPixelX * .Right
         .Bottom = Screen.TwipsPerPixelY * .Bottom
         lw = .Right - .Left
         lh = .Bottom - .Top
         
         frm.Move .Left + (lw - frm.Width) \ 2, .Top + (lh - frm.Height) \ 2
      End With
   End If

End Sub

Public Sub SetRegistryEntries(Optional ByVal strUser As String)

         '  setea la ubicación de las claves del registro de windows
      
10       On Error GoTo GestErr
      
20       ReDim RegistrySubKeys(EnumRegistrySubKeys.[_MAX_Value])
   
30       If Len(strUser) = 0 Then
40          RegistrySubKeys(EnumRegistrySubKeys.DataBaseSettings) = "Software\Algoritmo\DataBaseSettings"
50          RegistrySubKeys(EnumRegistrySubKeys.Environment) = "Software\Algoritmo\Environment"
'60          RegistrySubKeys(EnumRegistrySubKeys.NavigationQueries) = "Software\Algoritmo\MRU Queries\NavigationStoredQueries"
'70          RegistrySubKeys(EnumRegistrySubKeys.GridQueries) = "Software\Algoritmo\MRU Queries\GridStoredQueries"
'80          RegistrySubKeys(EnumRegistrySubKeys.PrintQueries) = "Software\Algoritmo\MRU Queries\PrintStoredQueries"
'90          RegistrySubKeys(EnumRegistrySubKeys.QueryDBQueries) = "Software\Algoritmo\MRU Queries\QueryDBStoredQueries"
'100         RegistrySubKeys(EnumRegistrySubKeys.DataComboQueries) = "Software\Algoritmo\MRU Queries\DataComboStoredQueries"
110      Else
120         RegistrySubKeys(EnumRegistrySubKeys.MRUEmpresas) = "Software\Algoritmo\MRU Empresas\" & strUser
'130         RegistrySubKeys(EnumRegistrySubKeys.KeyMRUForms) = "Software\Algoritmo\MRU Forms\" & strUser
140      End If
      
150      Exit Sub
   
GestErr:
160      LoadError ErrorLog, "SetRegistryEntries" & Erl
170      ShowErrMsg ErrorLog
End Sub

Public Sub ReadSystemOptions()
      Dim vValue As Variant

         ' lectura de los parametros internos
   
10       On Error GoTo GestErr

20       ReDim SystemOptions(EnumSystemOptions.[_MAX_Value])
   
         '  CacheSize
30       vValue = GetKeyValuePI("ADO\CacheSize")
40       SystemOptions(EnumSystemOptions.iCacheSize) = IIf(IsNull(vValue), 1, vValue)
   
         '  Zoom
50       vValue = GetKeyValuePI("Opciones\Zoom Vista Previa\Valor Generico", 80)
60       SystemOptions(EnumSystemOptions.iZoom) = IIf(IsNull(vValue), 70, vValue)
   
         '  Fetch Mode
70       vValue = GetKeyValuePI("Performance\FetchMode")
80       SystemOptions(EnumSystemOptions.iFetchMode) = IIf(IsNull(vValue), 1, vValue)
90       If (SystemOptions(EnumSystemOptions.iFetchMode) <> alAsync) And SystemOptions(EnumSystemOptions.iFetchMode) <> alSync And (SystemOptions(EnumSystemOptions.iFetchMode) <> alTable) Then
100         MsgBox "El valor del parámetro 'Performance\FetchMode' admite los siguientes valores:" & vbCrLf & _
                   "    1 - Fetch Asincrónico" & vbCrLf & _
                   "    2 - Fetch Sincrónico" & vbCrLf & _
                   "    3 - Variable según 'Performance\Limite Fetch Sincrónico'" & vbCrLf & _
                   "En caso de omisión, asume la opción 2", vbInformation, App.ProductName
110      End If
   
120      If SystemOptions(EnumSystemOptions.iFetchMode) = alTable Then
130         vValue = GetKeyValuePI("Performance\Limite Fetch Sincronico")
140         SystemOptions(EnumSystemOptions.lngFetchLimit) = IIf(IsNull(vValue), 1000, vValue)
150      End If
   
   
         '  Fetch Mode Busqueda
160      vValue = GetKeyValuePI("Performance\FetchMode en Busqueda")
170      SystemOptions(EnumSystemOptions.iFetchModeSearch) = IIf(IsNull(vValue), 1, vValue)
180      If (SystemOptions(EnumSystemOptions.iFetchModeSearch) <> alAsync) And (SystemOptions(EnumSystemOptions.iFetchModeSearch) <> alSync) And (SystemOptions(EnumSystemOptions.iFetchModeSearch) <> alTable) Then
190         MsgBox "El valor del parámetro 'Performance\FetchMode' admite los siguientes valores:" & vbCrLf & _
                   "    1 - Fetch Asincrónico" & vbCrLf & _
                   "    2 - Fetch Sincrónico" & vbCrLf & _
                   "    3 - Variable según 'Performance\Limite Fetch Sincrónico'" & vbCrLf & _
                   "En caso de omisión, asume la opción 2", vbInformation, App.ProductName
200      End If
   
210      If SystemOptions(EnumSystemOptions.iFetchModeSearch) = alTable Then
220         vValue = GetKeyValuePI("Performance\Limite Fetch Sincronico en Busqueda")
230         SystemOptions(EnumSystemOptions.lngFetchLimitSearch) = IIf(IsNull(vValue), 300, vValue)
240      End If
   
   
         '  Usa copias locales
250      vValue = GetKeyValuePI("Performance\Usa Copias Locales")
260      SystemOptions(EnumSystemOptions.UseLocalCopy) = IIf(IsNull(vValue), Si, vValue)
270      If (SystemOptions(EnumSystemOptions.UseLocalCopy) <> Si) And (SystemOptions(EnumSystemOptions.UseLocalCopy) <> No) Then
280         MsgBox "El valor del parámetro 'Performance\Usa Copias Locales' puede ser Sí o No:" & vbCrLf & _
                   "En caso de omisión, asume la opción Sí", vbInformation, App.ProductName
290      End If

         '  Usa copias locales en Búsquedas
300      vValue = GetKeyValuePI("Performance\Usa Copias Locales en Busqueda")
310      SystemOptions(EnumSystemOptions.UseLocalCopySearch) = IIf(IsNull(vValue), Si, vValue)
320      If (SystemOptions(EnumSystemOptions.UseLocalCopySearch) <> Si) And (SystemOptions(EnumSystemOptions.UseLocalCopySearch) <> No) Then
330         MsgBox "El valor del parámetro 'Performance\Usa Copias Locales en Búsqueda' puede ser Sí o No:" & vbCrLf & _
                   "En caso de omisión, asume la opción Sí", vbInformation, App.ProductName
340      End If

         '  Pregunta si usa copias locales desactualizadas
350      vValue = GetKeyValuePI("Performance\Usa Copias Locales Desactualizadas")
360      SystemOptions(EnumSystemOptions.AskOldLocalCopy) = IIf(IsNull(vValue), Si, vValue)
370      If (SystemOptions(EnumSystemOptions.AskOldLocalCopy) <> Si) And (SystemOptions(EnumSystemOptions.AskOldLocalCopy) <> No) Then
380         MsgBox "El valor del parámetro 'Performance\Usa Copias Locales Desactualizadas' puede ser Sí o No:" & vbCrLf & _
                   "En caso de omisión, asume la opción Sí", vbInformation, App.ProductName
390      End If
   
         '  Pregunta si usa copias locales desactualizadas
400      vValue = GetKeyValuePI("Opciones\Empresas\Usa MRU de Empresas")
410      SystemOptions(EnumSystemOptions.UseMRUEnterprise) = IIf(IsNull(vValue), Si, vValue)
420      If (SystemOptions(EnumSystemOptions.UseMRUEnterprise) <> Si) And (SystemOptions(EnumSystemOptions.UseMRUEnterprise) <> No) Then
430         MsgBox "El valor del parámetro 'Opciones\Empresas\Usa MRU de Empresas' puede ser Sí o No:" & vbCrLf & _
                   "En caso de omisión, asume la opción Sí", vbInformation, App.ProductName
440      End If
   
         '  Dimension de MRUForms
450      vValue = GetKeyValuePI("Performance\Dimension MRUForms")
460      SystemOptions(EnumSystemOptions.MaxMRUForms) = IIf(IsNull(vValue), 0, vValue)
   
470      Exit Sub

GestErr:
480      LoadError ErrorLog, "ReadSystemOptions " & Erl
490      ShowErrMsg ErrorLog
   
End Sub

Public Function FindKey(strKey As String, TV As TreeView) As Boolean
Dim nodX As Node
  
   ' busco si existe la clave strKey en el treeview TV
   
   On Error Resume Next
   
   Set nodX = TV.Nodes(strKey)
   FindKey = (Not nodX Is Nothing)
   Err.Clear
   
End Function

Public Sub CopySubTree(SourceTV As TreeView, sourceND As Node, DestTV As TreeView, destND As Node)
Dim ix As Long, so As Node, de As Node
'Dim s As String

    ' rutina recursiva que copia o mueve todos los hijos de un nodo a otro nodo
    
    If sourceND.Children = 0 Then Exit Sub
    
    Set so = sourceND.Child
    For ix = 1 To sourceND.Children
'        s = so.key
'        so.key = ""
        ' agrega un nodo en el TreeView de destino
        Set de = DestTV.Nodes.Add(destND, tvwChild, so.Key, so.Text, so.Image, so.SelectedImage)
        de.ExpandedImage = so.ExpandedImage
        
        ' agrega todos los hijos de este nodo, en modo recursivo
        CopySubTree SourceTV, so, DestTV, de
        
        ' obtiene una referencia al proximo
        Set so = so.Next
    Next
    
End Sub

Public Function SetApplication(ByVal objApp As Object) As Object

   If objApp Is Nothing Then Exit Function

   Set objApp.CurrentUser = CUsuario
   Set objApp.FormMDI = frmMDIInicio
   Set objApp.Menus = rstMenu
   Set objApp.ContextMenus = rstContextMenu
   Set objApp.CustomViews = rstVistasPersonalizadas
   Set objApp.ExportViews = rstVistasExportacion
   Set objApp.SysEnvironment = CSysEnvironment
   Set objApp.FormsMRU = MRUForms
   objApp.SystemOptionsProperty = SystemOptions
   objApp.RegistrySubKeysProperty = RegistrySubKeys

   Set SetApplication = objApp
   
End Function

Public Function IsMRUForm(ByVal lngHndW As Long) As Boolean
Dim frm As Form

   '  determina si un forms esta cargado en la colección MRUForms
   
   For Each frm In MRUForms

      If frm.hWnd = lngHndW Then IsMRUForm = True: Exit For

   Next frm
   
End Function

Public Function GetEjercicioActivo(ByVal strEmpresa As String) As String
Dim rst As ADODB.Recordset
   
   On Error GoTo GestErr
   
   Set rst = Fetch(strEmpresa, "SELECT EJERCICIOS.EJE_CODIGO FROM EJERCICIOS WHERE EJERCICIOS.EJE_ESTADO = 'V'")
   If Not rst.EOF Then
      GetEjercicioActivo = rst("EJE_CODIGO").Value
   End If
   
   If Not rst Is Nothing Then
      If rst.State <> adStateClosed Then rst.Close
   End If
   Set rst = Nothing
   
   Exit Function

GestErr:

   MsgBox "No es posible establecer una conexion con la Empresa " & strEmpresa & "." & vbCrLf & vbCrLf & _
          "Verifique en el Servidor si el Servicio OracleService" & strEmpresa & " ha sido iniciado." & vbCrLf & _
          "Asegurese que el Servicio este iniciado o modifique localmente su Registro para evitar el uso de dicha empresa"
          

   Err.Raise vbObjectError + 100, "GetEjercicioActivo [modMain]" & KNOWN_ERRORSOURCE, "Ejercicio Activo de la Empresa " & strEmpresa & " no disponible."
                              

End Function

Public Sub CrearRstEmpresas()
Dim sql As String

   sql = " SELECT EMPRESAS.* "
   sql = sql & " FROM EMPRESAS, "
   sql = sql & "     USUARIOS"
   sql = sql & " WHERE"
   sql = sql & "     USUARIOS.USU_USUARIO = '" & CUsuario.Usuario & "'"
   sql = sql & "     AND"
   sql = sql & "     ("
   sql = sql & "       (   USUARIOS.USU_PERMISO_EMPRESA = 'P' AND"
   sql = sql & "           EMPRESAS.EMP_CODIGO_EMPRESA  IN (  SELECT USUARIOS_EMPRESAS.UEM_EMPRESA"
   sql = sql & "                                          FROM USUARIOS_EMPRESAS"
   sql = sql & "                                          WHERE USUARIOS_EMPRESAS.UEM_USUARIO = '" & CUsuario.Usuario & "'"
   sql = sql & "                                              AND USUARIOS_EMPRESAS.UEM_EMPRESA = EMPRESAS.EMP_CODIGO_EMPRESA"
   sql = sql & "                                  )"
   sql = sql & "       ) OR"
   sql = sql & "       (   USUARIOS.USU_PERMISO_EMPRESA = 'D' AND"
   sql = sql & "           EMPRESAS.EMP_CODIGO_EMPRESA NOT IN (  SELECT USUARIOS_EMPRESAS.UEM_EMPRESA"
   sql = sql & "                                          FROM USUARIOS_EMPRESAS"
   sql = sql & "                                          WHERE USUARIOS_EMPRESAS.UEM_USUARIO = '" & CUsuario.Usuario & "'"
   sql = sql & "                                              AND USUARIOS_EMPRESAS.UEM_EMPRESA = EMPRESAS.EMP_CODIGO_EMPRESA"
   sql = sql & "                                  )"
   sql = sql & "       ) OR"
   sql = sql & "       (   NVL(USUARIOS.USU_PERMISO_EMPRESA,'N') = 'N' "
   sql = sql & "          "
   sql = sql & "           "
   sql = sql & "       )                                      "
   sql = sql & "     )"
   sql = sql & " ORDER BY EMP_DESCRIPCION    "
   Set rstEmpresas = Fetch(GetSPMProperty(DBSEmpresaPrimaria), sql, adOpenStatic, adLockReadOnly, adUseClient)
End Sub

Public Function AddAuditComment(ByVal ControlBuffer As String, _
                                ByVal strComment As String) As String
Dim aByte()   As Byte
Dim pbControl As PropertyBag

   'obtengo el nombre del componente y de la clase
   Set pbControl = New PropertyBag
   aByte = ControlBuffer
   pbControl.Contents = aByte
   
   ' Grabar las propiedades minimas requeridas
   pbControl.WriteProperty "EMPRESA", GetSPMProperty(DBSEmpresaPrimaria)
'   pbControl.WriteProperty "SUCURSAL", GetSucursalActiva(GetSPMProperty(DBSEmpresaPrimaria)) 'CUsuario.Sucursal
   If GetRegistryValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\Winlogon", "DefaultDomainName") = "ALGORITMO" Then
      pbControl.WriteProperty "SUCURSAL", GetSucursalActiva(GetSPMProperty(DBSEmpresaPrimaria)) 'CUsuario.Sucursal
   Else
      pbControl.WriteProperty "SUCURSAL", strSucursalElegida   'vs116
   End If
   
   pbControl.WriteProperty "USUARIO", CUsuario.Usuario
   pbControl.WriteProperty "MAQUINA", CSysEnvironment.Machine
     
   ' Ahora agrego el Comentario pasado como argumento
   pbControl.WriteProperty "COMENTARIO_AUDITORIA", strComment
   
   AddAuditComment = pbControl.Contents
   
End Function


Public Sub KillProcess(ByVal NameProcess As String, Optional IncludeMe As Boolean = False)
Const PROCESS_ALL_ACCESS = &H1F0FFF
Const TH32CS_SNAPPROCESS As Long = 2&
Dim uProcess  As PROCESSENTRY32
Dim RProcessFound As Long
Dim hSnapshot As Long
Dim SzExename As String
Dim ExitCode As Long
Dim MyProcess As Long
Dim AppKill As Boolean
Dim AppCount As Integer
Dim i As Integer
Dim WinDirEnv As String
Dim mCol    As Collection

        
       If NameProcess <> "" Then
       
          Set mCol = New Collection
          AppCount = 0

          uProcess.dwSize = Len(uProcess)
          hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
          RProcessFound = ProcessFirst(hSnapshot, uProcess)
  
          Do
            i = InStr(1, uProcess.szexeFile, Chr(0))
            SzExename = LCase$(Left$(uProcess.szexeFile, i - 1))
            WinDirEnv = Environ("Windir") + "\"
            WinDirEnv = LCase$(WinDirEnv)
            
            If Right$(SzExename, Len(NameProcess)) = LCase$(NameProcess) Then
               AppCount = AppCount + 1
               MyProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
               mCol.Add MyProcess
            End If
            RProcessFound = ProcessNext(hSnapshot, uProcess)
          Loop While RProcessFound
          
          For i = 1 To mCol.Count - 1
            AppKill = TerminateProcess(mCol(i), ExitCode)
            Call CloseHandle(mCol(i))
          Next i
          
          If IncludeMe Then
             AppKill = TerminateProcess(mCol(mCol.Count), ExitCode)
             Call CloseHandle(mCol(mCol.Count))
          End If
          Call CloseHandle(hSnapshot)
       End If

End Sub

Public Function GetMyObject(ByVal strComponentClass As String, Optional ByVal strServerName As String = NullString) As Object
10       On Error GoTo GestErr

         ' Sin este Objeto Local (que termina en Nothing) se queda vivo el objeto en el servidor
         Dim objetoLocal   As Object
         Dim ix As Integer
   
20       ix = 0
30       Set objetoLocal = CreateObject(strComponentClass, strServerName)
40       Set GetMyObject = objetoLocal
   
50       Set objetoLocal = Nothing
   
60       Exit Function

GestErr:
70       ix = ix + 1
80       If ix < 3 Then
90          Resume
100      End If
   
110      Set objetoLocal = Nothing
120      Set GetMyObject = Nothing

130      LoadError ErrorLog, "Objeto: " & strComponentClass & vbCrLf & "Servidor: " & strServerName
140      ShowErrMsg ErrorLog

End Function

