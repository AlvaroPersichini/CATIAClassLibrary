Option Explicit On

Public Class PwrProduct

    Private oProduct As ProductStructureTypeLib.Product
    Private intQuantity As Integer
    Private intLevel As Integer
    Private strProductType As String
    Private strSource As ProductStructureTypeLib.CatProductSource
    Private strFileName As String
    Private strFullPath As String
    Private strImageFilePath As String


    Public Property FullPath As String
        Get
            Return strFullPath
        End Get
        Set(value As String)
            strFullPath = value
        End Set
    End Property

    Public Property FileName As String
        Get
            Return strFileName
        End Get
        Set(value As String)
            strFileName = value
        End Set
    End Property

    Public Property ImageFilePath As String
        Get
            Return strImageFilePath
        End Get
        Set(value As String)
            strImageFilePath = value
        End Set
    End Property

    Public Property Product As ProductStructureTypeLib.Product
        Get
            Return oProduct
        End Get
        Set(value As ProductStructureTypeLib.Product)
            oProduct = value
        End Set
    End Property

    Public Property Quantity As Integer
        Get
            Return intQuantity
        End Get
        Set(value As Integer)
            intQuantity = value
        End Set
    End Property

    Public Property Level As Integer
        Get
            Return intLevel
        End Get
        Set(value As Integer)
            intLevel = value
        End Set
    End Property

    Public Property ProductType As String
        Get
            Return strProductType
        End Get
        Set(value As String)
            strProductType = value
        End Set
    End Property

    Public Property Source As ProductStructureTypeLib.CatProductSource
        Get
            Return strSource
        End Get
        Set(value As ProductStructureTypeLib.CatProductSource)
            strSource = value
        End Set
    End Property

End Class