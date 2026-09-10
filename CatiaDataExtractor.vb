Option Explicit On
Option Strict On

' El manejo de los "components": los detecta y los salta: si encuentra un "Component", no los pone en el oDictionary.
' Aunque salte el Component, entra a mirar qué tiene dentro, si adentro hay piezas reales, las trata normalmente.

' Links rotos: si el link de una pieza está roto, lo detecta porque al intentar acceder al documento de la referencia, lanza un error.
' Tiene un bloque try-catch para detectar si el link está roto. Si lo está, avisa por consola y omite ese elemento.

' La clase trabaja con un diccionario que usa tipos nativos de VB.NET para almacenar la información.
' No usa "ProductStructureTypeLib.Product".Esto  evita problemas de compatibilidad.


Public Class CatiaDataExtractor

    Public Function ExtractData(oRootProduct As ProductStructureTypeLib.Product,
                            folderPath As String,
                            takeSnaps As Boolean) As Dictionary(Of String, (FullPath As String,
                                                                           FileName As String,
                                                                           ImageFilePath As String,
                                                                           PartNumber As String,
                                                                           DescriptionRef As String,
                                                                           Nomenclature As String,
                                                                           Definition As String,
                                                                           Quantity As Integer,
                                                                           Level As Integer,
                                                                           ProductType As String,
                                                                           Source As Integer))

        Console.WriteLine("[" & DateTime.Now.ToString("HH:mm:ss") & "] - Extracting data from CATIA...")

        If takeSnaps AndAlso Not String.IsNullOrEmpty(folderPath) Then
            If Not IO.Directory.Exists(folderPath) Then IO.Directory.CreateDirectory(folderPath)
        End If

        Dim oDictionary As New Dictionary(Of String, (FullPath As String,
                                                  FileName As String,
                                                  ImageFilePath As String,
                                                  PartNumber As String,
                                                  DescriptionRef As String,
                                                  Nomenclature As String,
                                                  Definition As String,
                                                  Quantity As Integer,
                                                  Level As Integer,
                                                  ProductType As String,
                                                  Source As Integer))


        Dim rootDoc As INFITF.Document = CType(oRootProduct.ReferenceProduct.Parent, INFITF.Document)


        oDictionary.Add(oRootProduct.PartNumber,
                        (FullPath:=GetJustDirectory(rootDoc.FullName),
                        FileName:=rootDoc.Name,
                        ImageFilePath:=If(takeSnaps, TakeSnapshot(oRootProduct, folderPath, True), ""),
                        oRootProduct.PartNumber,
                        oRootProduct.DescriptionRef,
                        oRootProduct.Nomenclature,
                        oRootProduct.Definition,
                        Quantity:=1,
                        Level:=0,
                        ProductType:=TypeName(rootDoc),
                        Source:=CInt(oRootProduct.Source)))

        ' seguir viendo el tema del source, porque 
        ' en el diccionario lo estamos guardando como Integer,
        ' pero en realidad es un enum CatProductSource.
        ' Source As ProductStructureTypeLib.CatProductSource))


        ProcesarHijosRecursivo(oRootProduct, oDictionary, 1, folderPath, takeSnaps, rootDoc)

        Return oDictionary

    End Function





    Private Sub ProcesarHijosRecursivo(oParent As ProductStructureTypeLib.Product,
                                       ByRef oDictionary As Dictionary(Of String,
                                       (FullPath As String,
                                       FileName As String,
                                       ImageFilePath As String,
                                       PartNumber As String,
                                       DescriptionRef As String,
                                       Nomenclature As String,
                                       Definition As String,
                                       Quantity As Integer,
                                       Level As Integer,
                                       ProductType As String,
                                       Source As Integer)), ByVal currentLevel As Integer,
                                       folderPath As String,
                                       takeSnaps As Boolean,
                                       oParentDoc As INFITF.Document)

        For Each oChild As ProductStructureTypeLib.Product In oParent.Products

            Dim oChildDoc As INFITF.Document = Nothing

            Try
                oChildDoc = CType(oChild.ReferenceProduct.Parent, INFITF.Document)
            Catch ex As Exception
                Console.WriteLine(" Broken Link '" & oChild.Name & "'. This element will be skipped.")
                Continue For
            End Try

            If oChildDoc.FullName = oParentDoc.FullName Then
                ProcesarHijosRecursivo(oChild, oDictionary, currentLevel, folderPath, takeSnaps, oParentDoc)
            Else
                Dim pNumber As String = oChild.PartNumber

                'aca hay un filtro hardcodeado:
                'si el PartNumber empieza con "Aux", lo ignoramos.
                'Esto hay que manejarlo de otra manera.
                If Not pNumber.StartsWith("Aux", StringComparison.OrdinalIgnoreCase) Then


                    If oDictionary.ContainsKey(pNumber) Then
                        Dim item = oDictionary(pNumber)
                        item.Quantity += 1
                        oDictionary(pNumber) = item
                    Else
                        oDictionary.Add(pNumber,
                                        (FullPath:=GetJustDirectory(oChildDoc.FullName),
                                        FileName:=oChildDoc.Name,
                                        ImageFilePath:=If(takeSnaps, TakeSnapshot(oChild, folderPath, False), ""),
                                        oChild.PartNumber,
                                        oChild.DescriptionRef,
                                        oChild.Nomenclature,
                                        oChild.Definition,
                                        Quantity:=1,
                                        Level:=currentLevel,
                                        ProductType:=TypeName(oChildDoc),
                                        Source:=CInt(oChild.Source)))
                    End If

                End If

                If TypeOf oChildDoc Is ProductStructureTypeLib.ProductDocument Then
                    ProcesarHijosRecursivo(oChild, oDictionary, currentLevel + 1, folderPath, takeSnaps, oChildDoc)
                End If

            End If

        Next

    End Sub



    Private Function TakeSnapshot(oProd As ProductStructureTypeLib.Product, folder As String, isRoot As Boolean) As String
        Dim safePartNumber As String = CleanFileName(oProd.PartNumber)
        Dim finalFileName As String = IO.Path.Combine(folder, safePartNumber & ".jpg")

        Dim oApp As INFITF.Application = oProd.Application
        Dim docPrincipal As INFITF.Document = oApp.ActiveDocument

        If Not isRoot Then
            Dim oSelection As INFITF.Selection = docPrincipal.Selection
            oSelection.Clear()
            oSelection.Add(oProd)
            oApp.StartCommand("Open in New Window")
            oApp.RefreshDisplay = True

            If oApp.ActiveDocument Is docPrincipal Then
                oSelection.Clear()
                Return ""
            End If
        End If

        Dim oCurrentWindow As INFITF.Window = oApp.ActiveWindow
        Dim oSpecsWin As INFITF.SpecsAndGeomWindow = CType(oCurrentWindow, INFITF.SpecsAndGeomWindow)
        Dim oViewer As INFITF.Viewer3D = CType(oSpecsWin.Viewers.Item(1), INFITF.Viewer3D)

        Dim oldColor(2), white(2) As Object
        white(0) = 1 : white(1) = 1 : white(2) = 1
        oViewer.GetBackgroundColor(oldColor)
        oViewer.PutBackgroundColor(white)

        oSpecsWin.Layout = INFITF.CatSpecsAndGeomWindowLayout.catWindowGeomOnly
        oApp.StartCommand("Compass")
        oCurrentWindow.Height = 300
        oCurrentWindow.Width = 300

        oViewer.Viewpoint3D.ProjectionMode = INFITF.CatProjectionMode.catProjectionCylindric
        oViewer.Viewpoint3D = CType(oApp.ActiveDocument.Cameras.Item(1), INFITF.Camera3D).Viewpoint3D
        oViewer.Reframe()
        oViewer.Update()
        oApp.RefreshDisplay = True

        oViewer.CaptureToFile(INFITF.CatCaptureFormat.catCaptureFormatJPEG, finalFileName)

        oViewer.PutBackgroundColor(oldColor)
        oApp.StartCommand("Compass")
        oSpecsWin.Layout = INFITF.CatSpecsAndGeomWindowLayout.catWindowSpecsAndGeom

        If Not isRoot Then
            oApp.ActiveDocument.Close()
            docPrincipal.Activate()
        Else
            oCurrentWindow.WindowState = INFITF.CatWindowState.catWindowStateMaximized
        End If

        Return finalFileName
    End Function





    Public Function ExtractPositions(oRootProduct As ProductStructureTypeLib.Product) As List(Of (InstanceName As String,
                                                                                                 PartNumber As String,
                                                                                                 Level As Integer,
                                                                                                 ProductType As String,
                                                                                                 RelX As Double, RelY As Double, RelZ As Double,
                                                                                                 AbsX As Double, AbsY As Double, AbsZ As Double))

        Dim positionsList As New List(Of (InstanceName As String,
                                          PartNumber As String,
                                          Level As Integer,
                                          ProductType As String,
                                          RelX As Double, RelY As Double, RelZ As Double,
                                          AbsX As Double, AbsY As Double, AbsZ As Double))

        Dim rootDoc As INFITF.Document = CType(oRootProduct.ReferenceProduct.Parent, INFITF.Document)

        ' Posición del Root (Origen absoluto)
        positionsList.Add((InstanceName:=oRootProduct.Name,
                           oRootProduct.PartNumber,
                           Level:=0,
                           ProductType:=TypeName(rootDoc),
                           RelX:=0.0, RelY:=0.0, RelZ:=0.0,
                           AbsX:=0.0, AbsY:=0.0, AbsZ:=0.0))

        ' Matriz identidad inicial (4x4 simplificada)
        Dim rootAbsMatrix(11) As Double
        rootAbsMatrix(0) = 1.0 : rootAbsMatrix(4) = 1.0 : rootAbsMatrix(8) = 1.0

        ProcesarPosicionesRecursivo(oRootProduct, positionsList, 1, rootAbsMatrix, rootDoc)

        Return positionsList
    End Function







    Private Sub ProcesarPosicionesRecursivo(oParent As ProductStructureTypeLib.Product,
                                            ByRef positionsList As List(Of (InstanceName As String,
                                                                            PartNumber As String,
                                                                            Level As Integer,
                                                                            ProductType As String,
                                                                            RelX As Double, RelY As Double, RelZ As Double,
                                                                            AbsX As Double, AbsY As Double, AbsZ As Double)),
                                            currentLevel As Integer,
                                            parentAbsMatrix As Double(),
                                            oParentDoc As INFITF.Document)

        For Each oChild As ProductStructureTypeLib.Product In oParent.Products

            Dim oChildDoc As INFITF.Document = Nothing

            Try
                oChildDoc = CType(oChild.ReferenceProduct.Parent, INFITF.Document)
            Catch ex As Exception
                Console.WriteLine(" Broken Link '" & oChild.Name & "'. Omitiendo cálculo de posición.")
                Continue For
            End Try

            ' 1. Obtener matriz relativa del hijo (COM SafeArray a Double())
            Dim posObj(11) As Object
            oChild.Position.GetComponents(posObj)

            Dim relMatrix(11) As Double
            For i As Integer = 0 To 11
                relMatrix(i) = CDbl(posObj(i))
            Next

            ' 2. Calcular matriz absoluta acumulada (Padre * Hijo)
            Dim absMatrix As Double() = MultiplyTransforms(parentAbsMatrix, relMatrix)

            ' 3. Manejo de componentes internos vs archivos reales
            Dim esComponentInterno As Boolean = (oChildDoc.FullName = oParentDoc.FullName)

            If Not esComponentInterno Then
                Dim pNumber As String = oChild.PartNumber

                If Not pNumber.StartsWith("Aux", StringComparison.OrdinalIgnoreCase) Then
                    positionsList.Add((InstanceName:=oChild.Name,
                                       PartNumber:=pNumber,
                                       Level:=currentLevel,
                                       ProductType:=TypeName(oChildDoc),
                                       RelX:=relMatrix(9), RelY:=relMatrix(10), RelZ:=relMatrix(11),
                                       AbsX:=absMatrix(9), AbsY:=absMatrix(10), AbsZ:=absMatrix(11)))
                End If
            End If

            ' 4. Si es subensamble o componente interno con hijos, profundizar
            If oChild.Products.Count > 0 Then
                Dim nextDoc As INFITF.Document = If(esComponentInterno, oParentDoc, oChildDoc)
                Dim nextLevel As Integer = If(esComponentInterno, currentLevel, currentLevel + 1)
                ProcesarPosicionesRecursivo(oChild, positionsList, nextLevel, absMatrix, nextDoc)
            End If

        Next

    End Sub

    Private Function MultiplyTransforms(A As Double(), B As Double()) As Double()
        Dim C(11) As Double

        ' Vector director eje X
        C(0) = A(0) * B(0) + A(3) * B(1) + A(6) * B(2)
        C(1) = A(1) * B(0) + A(4) * B(1) + A(7) * B(2)
        C(2) = A(2) * B(0) + A(5) * B(1) + A(8) * B(2)

        ' Vector director eje Y
        C(3) = A(0) * B(3) + A(3) * B(4) + A(6) * B(5)
        C(4) = A(1) * B(3) + A(4) * B(4) + A(7) * B(5)
        C(5) = A(2) * B(3) + A(5) * B(4) + A(8) * B(5)

        ' Vector director eje Z
        C(6) = A(0) * B(6) + A(3) * B(7) + A(6) * B(8)
        C(7) = A(1) * B(6) + A(4) * B(7) + A(7) * B(8)
        C(8) = A(2) * B(6) + A(5) * B(7) + A(8) * B(8)

        ' Traslación (Origen): T_res = Rot_A * T_B + T_A
        C(9) = A(0) * B(9) + A(3) * B(10) + A(6) * B(11) + A(9)
        C(10) = A(1) * B(9) + A(4) * B(10) + A(7) * B(11) + A(10)
        C(11) = A(2) * B(9) + A(5) * B(10) + A(8) * B(11) + A(11)

        Return C
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