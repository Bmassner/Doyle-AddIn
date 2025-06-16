Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Runtime.InteropServices
Imports Docnet.Core
Imports Docnet.Core.Models

Public Module PDFToImage
    ''' <summary>
    ''' Converts the first page of a PDF to a JPEG image using Docnet.Core, with a white background.
    ''' </summary>
    ''' <param name="pdfFilePath">The path to the input PDF file.</param>
    ''' <param name="imageFilePath">The path to the output image file.</param>
    Public Sub ExportFirstPageAsImage(pdfFilePath As String, imageFilePath As String)
        ' Set desired DPI or pixel dimensions
        Dim dpi As Integer = 3200

        Using docReader = DocLib.Instance.GetDocReader(pdfFilePath, New PageDimensions(dpi, dpi))
            Using pageReader = docReader.GetPageReader(0)
                Dim width = pageReader.GetPageWidth()
                Dim height = pageReader.GetPageHeight()
                Dim rawBytes = pageReader.GetImage()

                ' Create a bitmap from the raw BGRA bytes
                Using bmp As New Bitmap(width, height, PixelFormat.Format32bppArgb)
                    Dim bmpData = bmp.LockBits(New Rectangle(0, 0, width , height ), Imaging.ImageLockMode.WriteOnly, bmp.PixelFormat)
                    Marshal.Copy(rawBytes, 0, bmpData.Scan0, rawBytes.Length)
                    bmp.UnlockBits(bmpData)

                    ' Composite onto a white background
                    Using whiteBmp As New Bitmap(width, height, PixelFormat.Format24bppRgb)
                        Using g As Graphics = Graphics.FromImage(whiteBmp)
                            g.Clear(Color.White)
                            g.DrawImage(bmp, 0, 0)
                        End Using
                        whiteBmp.Save(imageFilePath, ImageFormat.Jpeg)
                    End Using
                End Using
            End Using
        End Using
    End Sub
End Module