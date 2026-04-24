Option Explicit On
Option Strict On


' El manejo de los "components": los detecta y los salta: si encuentra un "Component", no los pone en el oDictionary.
' Aunque salte el Component, entra a mirar qué tiene dentro, si adentro hay piezas reales, las trata normalmente.

' Links rotos: si el link de una pieza está roto, lo detecta porque al intentar acceder al documento de la referencia, lanza un error.
' Tiene un bloque try-catch para detectar si el link está roto. Si lo está, avisa por consola y omite ese elemento.



Public Class CatiaDataExtractor

    Public Function ExtractData(oRootProduct As ProductStructureTypeLib.Product,
                                folderPath As String,
                                takeSnaps As Boolean) As Dictionary(Of String, PwrProduct)

        Console.WriteLine("[" & DateTime.Now.ToString("HH:mm:ss") & "] Step 1/3: Extracting data from CATIA...")

        If takeSnaps AndAlso Not String.IsNullOrEmpty(folderPath) Then
            If Not IO.Directory.Exists(folderPath) Then IO.Directory.CreateDirectory(folderPath)
        End If

        Dim oDictionary As New Dictionary(Of String, PwrProduct)
        Dim PPRoot As New PwrProduct

        ' Documento del Root
        Dim rootDoc As INFITF.Document = CType(oRootProduct.ReferenceProduct.Parent, INFITF.Document)

        With PPRoot
            .Product = oRootProduct
            .Quantity = 1
            .ProductType = TypeName(rootDoc)
            .Source = oRootProduct.Source
            .Level = 0
            .FileName = rootDoc.Name
            .FullPath = GetJustDirectory(rootDoc.FullName) ' <--- MODIFICADO
            .ImageFilePath = If(takeSnaps, TakeSnapshot(oRootProduct, folderPath, True), "")
        End With

        oDictionary.Add(oRootProduct.PartNumber, PPRoot)

        ' Iniciamos recursividad pasando el documento del padre (en este caso el root)
        ProcesarHijosRecursivo(oRootProduct, oDictionary, 1, folderPath, takeSnaps, rootDoc)

        Return oDictionary
    End Function

    Private Sub ProcesarHijosRecursivo(oParent As ProductStructureTypeLib.Product,
                                  ByRef oDictionary As Dictionary(Of String, PwrProduct),
                                  ByVal currentLevel As Integer,
                                  folderPath As String,
                                  takeSnaps As Boolean,
                                  oParentDoc As INFITF.Document)

        For Each oChild As ProductStructureTypeLib.Product In oParent.Products

            Dim oChildDoc As INFITF.Document = Nothing

            'este Try bloque es el que detecta si el link está roto
            Try
                ' Intento obtener el documento de la referencia
                ' Si el link está roto, esta línea disparará el E_FAIL
                oChildDoc = CType(oChild.ReferenceProduct.Parent, INFITF.Document)
            Catch ex As Exception
                ' CASO: LINK ROTO
                Console.WriteLine(" ALERTA: Link roto detectado en '" & oChild.Name & "'. Se omitirá este elemento.")
                Continue For ' Salta al siguiente hijo del bucle
            End Try

            ' Si llegamos aquí, el documento existe y es accesible
            If oChildDoc.FullName = oParentDoc.FullName Then  ' ES UN COMPONENT (Internal)
                ProcesarHijosRecursivo(oChild, oDictionary, currentLevel, folderPath, takeSnaps, oParentDoc)
            Else
                ' ES UN ARCHIVO REAL (Part o Product)
                Dim pNumber As String = oChild.PartNumber

                If Not pNumber.StartsWith("Aux", StringComparison.OrdinalIgnoreCase) Then

                    If oDictionary.ContainsKey(pNumber) Then
                        oDictionary.Item(pNumber).Quantity += 1
                    Else
                        Dim PP As New PwrProduct
                        With PP
                            .Product = oChild
                            .Quantity = 1
                            .ProductType = TypeName(oChildDoc)
                            .Source = oChild.Source
                            .Level = currentLevel
                            .FileName = oChildDoc.Name
                            .FullPath = GetJustDirectory(oChildDoc.FullName)
                            ' Solo tomamos snapshot si el link no está roto (ya validado arriba)
                            .ImageFilePath = If(takeSnaps, TakeSnapshot(oChild, folderPath, False), "")
                        End With
                        oDictionary.Add(pNumber, PP)
                    End If

                End If

                ' Si es un ensamble real y no está roto, profundizamos
                If TypeOf oChildDoc Is ProductStructureTypeLib.ProductDocument Then
                    ProcesarHijosRecursivo(oChild, oDictionary, currentLevel + 1, folderPath, takeSnaps, oChildDoc)
                End If
            End If
        Next
    End Sub

    Private Function TakeSnapshot(oProd As ProductStructureTypeLib.Product, folder As String, isRoot As Boolean) As String

        ' Limpiar el nombre para evitar errores con caracteres como / o *
        Dim safePartNumber As String = CleanFileName(oProd.PartNumber)

        Dim finalFileName As String = IO.Path.Combine(folder, safePartNumber & ".jpg")

        Dim oApp As INFITF.Application = oProd.Application
        Dim docPrincipal As INFITF.Document = oApp.ActiveDocument

        ' --- 1. GESTIÓN DE VENTANAS ---
        If Not isRoot Then
            Dim oSelection As INFITF.Selection = docPrincipal.Selection
            oSelection.Clear()
            oSelection.Add(oProd)
            oApp.StartCommand("Open in New Window")
            oApp.RefreshDisplay = True

            ' Control de seguridad: si no cambió la ventana, abortar para no cerrar el principal
            If oApp.ActiveDocument Is docPrincipal Then
                oSelection.Clear()
                Return ""
            End If
        End If

        ' --- 2. CONFIGURACIÓN VISUAL ---
        Dim oCurrentWindow As INFITF.Window = oApp.ActiveWindow
        Dim oSpecsWin As INFITF.SpecsAndGeomWindow = CType(oCurrentWindow, INFITF.SpecsAndGeomWindow)
        Dim oViewer As INFITF.Viewer3D = CType(oSpecsWin.Viewers.Item(1), INFITF.Viewer3D)

        ' Fondo Blanco
        Dim oldColor(2), white(2) As Object
        white(0) = 1 : white(1) = 1 : white(2) = 1
        oViewer.GetBackgroundColor(oldColor)
        oViewer.PutBackgroundColor(white)

        ' Interfaz y Cámara
        oSpecsWin.Layout = INFITF.CatSpecsAndGeomWindowLayout.catWindowGeomOnly
        oApp.StartCommand("Compass") ' Ocultar

        oCurrentWindow.Height = 300
        oCurrentWindow.Width = 300


        ' 2. Cambiar el tipo de proyección a Paralela (Cylindric)
        ' En la API de CATIA, Parallel se identifica como catProjectionCylindric
        oViewer.Viewpoint3D.ProjectionMode = INFITF.CatProjectionMode.catProjectionCylindric

        ' Aplicar Vista Isométrica (Cámara 1)
        oViewer.Viewpoint3D = CType(oApp.ActiveDocument.Cameras.Item(1), INFITF.Camera3D).Viewpoint3D
        oViewer.Reframe()
        oViewer.Update()
        oApp.RefreshDisplay = True

        ' --- 3. CAPTURA ---
        oViewer.CaptureToFile(INFITF.CatCaptureFormat.catCaptureFormatJPEG, finalFileName)

        ' --- 4. RESTAURACIÓN ---
        oViewer.PutBackgroundColor(oldColor)
        oApp.StartCommand("Compass") ' Mostrar
        oSpecsWin.Layout = INFITF.CatSpecsAndGeomWindowLayout.catWindowSpecsAndGeom

        If Not isRoot Then
            oApp.ActiveDocument.Close()
            docPrincipal.Activate()
        Else
            oCurrentWindow.WindowState = INFITF.CatWindowState.catWindowStateMaximized
        End If
        Return finalFileName
    End Function


    Private Function CleanFileName(name As String) As String
        Dim invalidChars As New String(IO.Path.GetInvalidFileNameChars())
        Dim cleaned As String = name
        For Each c As Char In invalidChars
            cleaned = cleaned.Replace(c, "_"c)
        Next
        Return cleaned
    End Function


    Private Function GetJustDirectory(fullPath As String) As String
        If String.IsNullOrEmpty(fullPath) Then Return ""
        Dim lastSlash As Integer = Math.Max(fullPath.LastIndexOf("\"), fullPath.LastIndexOf("/"))
        If lastSlash > 0 Then
            Return fullPath.Substring(0, lastSlash)
        End If
        Return fullPath
    End Function


End Class