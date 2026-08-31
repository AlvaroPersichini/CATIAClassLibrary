Option Explicit On
Option Strict On

Public Class CatiaDataInjector

    Public Sub InjectData(oRootProduct As ProductStructureTypeLib.Product,
                          dataToInject As Dictionary(Of String, (NewPartNumber As String, DescriptionRef As String, Quantity As Integer, Source As Integer, Nomenclature As String, Definition As String)))

        ' Obtenemos el documento raíz para la lógica de componentes
        Dim rootDoc As INFITF.Document = CType(oRootProduct.ReferenceProduct.Parent, INFITF.Document)

        Dim processedFiles As New HashSet(Of String)

        ' Procesar el Root
        ApplyToDocument(oRootProduct, rootDoc, dataToInject, processedFiles)

        ' Iniciar recorrido recursivo para los hijos
        ProcesarHijosRecursivo(oRootProduct, dataToInject, processedFiles, rootDoc)

        Console.WriteLine("[" & DateTime.Now.ToString("HH:mm:ss") & "] Injection complete.")

    End Sub

    Private Sub ProcesarHijosRecursivo(oParent As ProductStructureTypeLib.Product,
                                  dataToInject As Dictionary(Of String, (NewPartNumber As String, DescriptionRef As String, Quantity As Integer, Source As Integer, Nomenclature As String, Definition As String)),
                                  ByRef processedFiles As HashSet(Of String),
                                  oParentDoc As INFITF.Document)


        For Each oChild As ProductStructureTypeLib.Product In oParent.Products
            Dim oChildDoc As INFITF.Document = Nothing


            ' --- PROTECCIÓN CONTRA LINK ROTO ---
            ' Intentamos acceder al documento de la referencia
            ' Si falla, es un link roto. Consola y saltamos al siguiente.
            Try
                oChildDoc = CType(oChild.ReferenceProduct.Parent, INFITF.Document)
            Catch ex As Exception
                Console.WriteLine(" [SKIP] Link roto detectado en: " & oChild.Name & ". No se puede inyectar datos.")
                Continue For
            End Try


            ' Si el hijo vive en el mismo archivo que el padre entonces es Component.
            ' Entramos a sus hijos sin marcar el archivo como procesado
            If oChildDoc.FullName = oParentDoc.FullName Then
                ProcesarHijosRecursivo(oChild, dataToInject, processedFiles, oParentDoc)
            Else
                ' Si es un archivo real que no hemos modificado en esta sesión
                If Not processedFiles.Contains(oChildDoc.FullName) Then
                    ApplyToDocument(oChild, oChildDoc, dataToInject, processedFiles)
                End If
                ' Si el hijo es un ProductDocument, entramos recursivamente a sus hijos
                If TypeOf oChildDoc Is ProductStructureTypeLib.ProductDocument Then
                    ProcesarHijosRecursivo(oChild, dataToInject, processedFiles, oChildDoc)
                End If
            End If
        Next

    End Sub


    Private Sub ApplyToDocument(oInstancia As ProductStructureTypeLib.Product,
                            oDoc As INFITF.Document,
                            data As Dictionary(Of String, (NewPartNumber As String, DescriptionRef As String, Quantity As Integer, Source As Integer, Nomenclature As String, Definition As String)),
                            ByRef processed As HashSet(Of String))

        Dim currentPN As String = oInstancia.PartNumber

        If data.ContainsKey(currentPN) Then

            Dim info As (NewPartNumber As String,
                DescriptionRef As String,
                Quantity As Integer,
                Source As Integer,
                Nomenclature As String,
                Definition As String) = data(currentPN)

            Dim oRefProd As ProductStructureTypeLib.Product = oInstancia.ReferenceProduct

            ' 1. PartNumber
            If Not String.IsNullOrEmpty(info.NewPartNumber) AndAlso oRefProd.PartNumber <> info.NewPartNumber Then
                oRefProd.PartNumber = info.NewPartNumber & " "
                oRefProd.PartNumber = info.NewPartNumber.Trim()
            End If

            ' 2. Definition
            If oRefProd.Definition <> info.Definition Then
                oRefProd.Definition = info.Definition & " "
                oRefProd.Definition = info.Definition.Trim()
            End If

            ' 3. Nomenclature
            If oRefProd.Nomenclature <> info.Nomenclature Then
                oRefProd.Nomenclature = info.Nomenclature & " "
                oRefProd.Nomenclature = info.Nomenclature.Trim()
            End If

            ' 4. DescriptionRef
            If oRefProd.DescriptionRef <> info.DescriptionRef Then
                oRefProd.DescriptionRef = info.DescriptionRef & " "
                oRefProd.DescriptionRef = info.DescriptionRef.Trim()
            End If

            ' 5. Source
            Dim newSource As ProductStructureTypeLib.CatProductSource = CType(info.Source, ProductStructureTypeLib.CatProductSource)
            If oRefProd.Source <> newSource Then
                oRefProd.Source = newSource
            End If

        End If

        ' Marcamos como procesado
        If Not String.IsNullOrEmpty(oDoc.FullName) Then
            processed.Add(oDoc.FullName)
        End If

    End Sub

End Class
