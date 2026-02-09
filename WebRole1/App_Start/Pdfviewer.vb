
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Text
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls


Namespace PdfViewer
    <DefaultProperty("FilePath")> _
    <ToolboxData("<{0}:ShowPdf runat=server></{0}:ShowPdf>")> _
    Public Class ShowPdf
        Inherits WebControl

#Region "Declarations"

        Private mFilePath As String

#End Region



#Region "Properties"

        <Category("Source File")> _
        <Browsable(True)> _
        <Description("Set path to source file.")> _
        <Editor(GetType(System.Web.UI.DesignTimeParseData), GetType(System.Drawing.Design.UITypeEditor))> _
        Public Property FilePath() As String
            Get
                Return mFilePath
            End Get
            Set(value As String)
                If value = String.Empty Then
                    mFilePath = String.Empty
                Else
                    Dim tilde As Integer = -1
                    tilde = value.IndexOf("~"c)
                    If tilde <> -1 Then
                        mFilePath = value.Substring((tilde + 2)).Trim()
                    Else
                        mFilePath = value
                    End If
                End If
            End Set
        End Property
        ' end FilePath property

#End Region



#Region "Rendering"

        Protected Overrides Sub RenderContents(writer As HtmlTextWriter)
            Try
                Dim sb As New StringBuilder()
                sb.Append("<iframe src=" + FilePath.ToString() + " ")
                sb.Append("width=" + Width.ToString() + " height=" + Height.ToString() + " ")
                sb.Append("<View PDF: <a href=" + FilePath.ToString() + "</a></p> ")
                sb.Append("</iframe>")

                writer.RenderBeginTag(HtmlTextWriterTag.Div)
                writer.Write(sb.ToString())
                writer.RenderEndTag()
            Catch
                ' with no properties set, this will render "Display PDF Control" in a
                ' a box on the page
                writer.RenderBeginTag(HtmlTextWriterTag.Div)
                writer.Write("Display PDF Control")
                writer.RenderEndTag()
            End Try
            ' end try-catch
        End Sub
        ' end RenderContents

#End Region

    End Class
    ' end class
End Namespace
' end namespace

'=======================================================
'Service provided by Telerik (www.telerik.com)
'Conversion powered by NRefactory.
'Twitter: @telerik
'Facebook: facebook.com/telerik
'=======================================================

