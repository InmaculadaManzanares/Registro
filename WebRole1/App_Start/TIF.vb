
Imports System.Data
Imports System.Drawing
Imports System.IO
Imports System.Drawing.Imaging
Imports System.Collections.Generic
Imports System.Linq
Imports System.Web

Namespace App.Files
    ''' <summary>
    ''' Summary description for TIF
    ''' </summary>

    Public Class TIF
        ' Fields
        Private m_Disposed As Boolean = False
        Private m_Img As Image = Nothing
        Private m_PageCount As Integer = -1
        Public m_FilePathName As [String]

#Region "CONSTRUCTOR/DESTRUCTOR"

        ''' <summary>
        ''' Creates an instance of this class
        ''' </summary>
        ''' <param name="FilePathName"></param>
        Public Sub New(ByVal FilePathName As [String])
            If FilePathName Is Nothing Then
                Throw New Exception("FilePathName cannot be NULL")
            End If
            If FilePathName = "" Then
                Throw New Exception("FilePathName cannot be NULL")
            End If
            m_FilePathName = FilePathName
        End Sub

        ''' <summary>
        ''' Disposes of any resources still available
        ''' </summary>
        Public Sub Dispose()
            If Me.m_Img IsNot Nothing Then
                Me.m_Img.Dispose()
            End If
            Me.m_Disposed = True
        End Sub


#End Region

#Region "PROPERTIES"

        ''' <summary>
        ''' Gets the Page Count of the TIF
        ''' </summary>
        Public ReadOnly Property PageCount() As Integer
            Get
                If m_PageCount = -1 Then
                    Me.m_PageCount = GetPageCount()
                End If

                Return m_PageCount
            End Get
        End Property


#End Region

#Region "METHODS"

        ''' <summary>
        ''' Returns the page count of the TIF
        ''' </summary>
        ''' <returns></returns>
        ''' 
        Public Shared Function SePuedeMostrar(ByVal ruta As String) As Boolean
            Dim Img As Image = Nothing
            Try
                Img = Image.FromFile(ruta)
                Return True
            Catch ex As Exception
                Return False
            Finally
                If Not Img Is Nothing Then
                    Img.Dispose()
                    GC.Collect()
                    GC.WaitForPendingFinalizers()
                End If
            End Try
        End Function
        Private Function GetPageCount() As Integer
            Dim Pgs As Integer = -1
            Dim Img As Image = Nothing
            Try
                Img = Image.FromFile(Me.m_FilePathName)
                Pgs = Img.GetFrameCount(FrameDimension.Page)

                Return Pgs
            Catch ex As Exception
                Throw ex
            Finally
                If Not Img Is Nothing Then
                    Img.Dispose()
                    GC.Collect()
                    GC.WaitForPendingFinalizers()
                End If
            End Try
        End Function

        ''' <summary>
        ''' Returns an Image of a TIF page
        ''' </summary>
        ''' <param name="PageNum"></param>
        ''' <returns></returns>
        Public Function GetTiffImage(ByVal PageNum As Integer) As Image
            If (PageNum < 1) Or (PageNum > Me.PageCount) Then
                Throw New InvalidOperationException("Page to be retrieved is outside the bounds of the total TIF file pages.  Please choose a page number that exists.")
            End If
            Dim ms As MemoryStream = Nothing
            Dim SrcImg As Image = Nothing
            Dim returnImage As Image = Nothing
            Try
                SrcImg = Image.FromFile(Me.m_FilePathName)
                ms = New MemoryStream()
                Dim FrDim As New FrameDimension(SrcImg.FrameDimensionsList(0))
                SrcImg.SelectActiveFrame(FrDim, PageNum - 1)
                SrcImg.Save(ms, ImageFormat.Tiff)
                returnImage = Image.FromStream(ms)
            Catch ex As Exception
                Throw ex
            Finally
                SrcImg.Dispose()
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try

            Return returnImage
        End Function

        ''' <summary>
        ''' Returns an Image of a TIF page, resized
        ''' </summary>
        ''' <param name="PageNum"></param>
        ''' <returns></returns>
        Public Function GetTiffImageThumb(ByVal PageNum As Integer, ByVal ImgWidth As Integer, ByVal ImgHeight As Integer) As Image
            If (PageNum < 1) Or (PageNum > Me.PageCount) Then
                PageNum = 1
                'Throw New InvalidOperationException("Page to be retrieved is outside the bounds of the total TIF file pages.  Please choose a page number that exists.")
            End If
            Dim ms As MemoryStream = Nothing
            Dim SrcImg As Image = Nothing
            Dim returnImage As Image = Nothing
            Try
                SrcImg = Image.FromFile(Me.m_FilePathName)
                ms = New MemoryStream()
                Dim FrDim As New FrameDimension(SrcImg.FrameDimensionsList(0))
                SrcImg.SelectActiveFrame(FrDim, PageNum - 1)
                SrcImg.Save(ms, ImageFormat.Tiff)
                ' Prevent using images internal thumbnail
                SrcImg.RotateFlip(System.Drawing.RotateFlipType.Rotate180FlipNone)
                SrcImg.RotateFlip(System.Drawing.RotateFlipType.Rotate180FlipNone)
                'Save Aspect Ratio
                If SrcImg.Width <= ImgWidth Then
                    ImgWidth = SrcImg.Width
                End If
                Dim NewHeight As Integer = SrcImg.Height * ImgWidth \ SrcImg.Width
                If NewHeight > ImgHeight Then
                    ' Resize with height instead
                    ImgWidth = SrcImg.Width * ImgHeight \ SrcImg.Height
                    NewHeight = ImgHeight
                End If
                'Return Image
                returnImage = Image.FromStream(ms).GetThumbnailImage(ImgWidth, NewHeight, Nothing, IntPtr.Zero)
            Catch ex As Exception
                Throw ex
            Finally
                SrcImg.Dispose()
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try

            Return returnImage
        End Function

        ''' <summary>
        ''' Returns an images beased on pages specified
        ''' </summary>
        ''' <param name="StartPageNum"></param>
        ''' <param name="EndPageNum"></param>
        ''' <returns></returns>
        Public Function GetTiffImages(ByVal StartPageNum As Integer, ByVal EndPageNum As Integer) As TIFPageCollection
            Dim Pgs As New TIFPageCollection()
            If ((StartPageNum < 1) Or (EndPageNum > Me.m_PageCount)) Or (EndPageNum > StartPageNum) Then
                Throw New InvalidOperationException("Page being retrieved is outside the bounds of the total TIF file pages.  Please choose a page number that exists.")
            End If
            Try
                Dim TotPgs As Integer = EndPageNum - StartPageNum
                For i As Integer = 0 To TotPgs
                    Pgs.Add(Me.GetTiffImage(StartPageNum + i))
                Next
            Catch ex As Exception
                Throw ex
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
            Return Pgs
        End Function

        ''' <summary>
        ''' Returns an images based on pages specified, resized
        ''' </summary>
        ''' <param name="StartPageNum"></param>
        ''' <param name="EndPageNum"></param>
        ''' <returns></returns>
        Public Function GetTiffImageThumbs(ByVal StartPageNum As Integer, ByVal EndPageNum As Integer, ByVal ImgWidth As Integer, ByVal ImgHeight As Integer) As System.Drawing.Image()
            Dim Pgs As New TIFPageCollection()
            If ((StartPageNum < 1) OrElse (EndPageNum > Me.m_PageCount)) OrElse (EndPageNum < StartPageNum) Then
                Throw New InvalidOperationException("Page being retrieved is outside the bounds of the total TIF file pages.  Please choose a page number that exists.")
            End If
            Dim returnImage As Image() = New Image((EndPageNum - StartPageNum)) {}
            Try
                Dim TotPgs As Integer = EndPageNum - StartPageNum
                For i As Integer = 0 To TotPgs
                    returnImage(i) = Me.GetTiffImageThumb(StartPageNum + i, ImgWidth, ImgHeight)
                Next
            Catch ex As Exception
                Throw ex
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
            Return returnImage
        End Function

        ''' <summary>
        ''' Returns an images based on pages specified, resized
        ''' </summary>
        ''' <param name="StartPageNum"></param>
        ''' <param name="EndPageNum"></param>
        ''' <returns></returns>
        Public Function GetTiffImageThumbsCollection(ByVal StartPageNum As Integer, ByVal EndPageNum As Integer, ByVal ImgWidth As Integer, ByVal ImgHeight As Integer) As TIFPageCollection
            Dim Pgs As New TIFPageCollection()
            If ((StartPageNum < 1) OrElse (EndPageNum > Me.m_PageCount)) OrElse (EndPageNum < StartPageNum) Then
                Throw New InvalidOperationException("Page being retrieved is outside the bounds of the total TIF file pages.  Please choose a page number that exists.")
            End If
            Try
                Dim TotPgs As Integer = EndPageNum - StartPageNum
                For i As Integer = 0 To TotPgs
                    Pgs.Add(Me.GetTiffImageThumb(StartPageNum + i, ImgWidth, ImgHeight))
                Next
            Catch ex As Exception
                Throw ex
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
            Return Pgs
        End Function

        ''' <summary>
        ''' Returns a Image of a specific page
        ''' </summary>
        ''' <param name="PageNum"></param>
        ''' <returns></returns>
        Default Public ReadOnly Property Item(ByVal PageNum As Integer) As Image
            Get
                Dim TiffPage As Image
                Try
                    Me.m_Img = Me.GetTiffImage(PageNum)
                    TiffPage = Me.m_Img
                Catch ex As Exception
                    Throw ex
                End Try
                Return TiffPage
            End Get
        End Property

#End Region

    End Class

    ''' <summary>
    ''' Collection of objects
    ''' </summary>
    <Serializable()> _
    Public Class TIFPageCollection
        Inherits System.Collections.CollectionBase
        Private m_Disposed As Boolean = False

#Region "CONSTURCTORS"

        ''' <summary>
        ''' Default Constructor
        ''' </summary>
        Public Sub New()
        End Sub

        ''' <summary>
        ''' Disposes of the object by clearing the collection
        ''' </summary>
        Public Sub Dispose()
            'Make sure each image is disposed of properly
            For Each Img As System.Drawing.Image In Me
                Img.Dispose()
                GC.Collect()
                GC.WaitForPendingFinalizers()
            Next
            Me.m_Disposed = True
            Me.Clear()
        End Sub

#End Region

#Region "METHODS"

        ''' <summary>
        ''' Adds an item to the collection
        ''' </summary>
        ''' <param name="Obj"></param>
        Public Sub Add(ByVal Obj As System.Drawing.Image)
            Me.List.Add(Obj)
        End Sub

        ''' <summary>
        ''' Indicates if the object exists in the collection
        ''' </summary>
        ''' <param name="Obj"></param>
        ''' <returns></returns>
        Public Function Contains(ByVal Obj As System.Drawing.Image) As Boolean
            Return Me.List.Contains(Obj)
        End Function

        ''' <summary>
        ''' Returns the indexOf the object
        ''' </summary>
        ''' <param name="Obj"></param>
        ''' <returns></returns>
        Public Function IndexOf(ByVal Obj As System.Drawing.Image) As Integer
            Return Me.List.IndexOf(Obj)
        End Function

        ''' <summary>
        ''' Inserts an object into the collection at the specifed index
        ''' </summary>
        ''' <param name="index"></param>
        ''' <param name="Obj"></param>
        Public Sub Insert(ByVal index As Integer, ByVal Obj As System.Drawing.Image)
            Me.List.Insert(index, Obj)
        End Sub

        ''' <summary>
        ''' Removes the object from the collection
        ''' </summary>
        ''' <param name="Obj"></param>
        Public Sub Remove(ByVal Obj As System.Drawing.Image)
            Me.List.Remove(Obj)
        End Sub

#End Region

#Region "PROPERTIES"

        ''' <summary>
        ''' Returns a reference to the object at specified index
        ''' </summary>
        ''' <param name="index">Index of object</param>
        ''' <returns></returns>
        Default Public Property Item(ByVal index As Integer) As System.Drawing.Image
            Get
                Return DirectCast(Me.List(index), System.Drawing.Image)
            End Get
            Set(ByVal value As System.Drawing.Image)
                Me.List(index) = value
            End Set
        End Property

#End Region

    End Class
End Namespace
