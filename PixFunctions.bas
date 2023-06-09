Attribute VB_Name = "PixFunctions"
Function funcPixGetFileExt(PixImageObject As Object) As String

    Select Case PixImageObject.ScanPackaging
        Case &H0          ' No packaging
             PixImageObject.ScanFileExt = ""
        
        Case &H10000
             PixImageObject.ScanFileExt = ".PCX"
        
        Case &H20000
             PixImageObject.ScanFileExt = ".PDA"
        
        Case &H30000
            PixImageObject.ScanFileExt = ".TIF"
        
        Case &H50000
            PixImageObject.ScanFileExt = ".DCX"
        
        Case &H60000
            PixImageObject.ScanFileExt = ".BMP"
        
        Case &H80000
            PixImageObject.ScanFileExt = ".GIF"
         
        Case &HB0000
            PixImageObject.ScanFileExt = ".JPEG"
         
        Case &HC0000
            PixImageObject.ScanFileExt = ".CALS"
         
        Case &HE0000
            PixImageObject.ScanFileExt = ".DCA"
         
        Case &H100000
            PixImageObject.ScanFileExt = ".PDF"
         
        Case &H110000
            PixImageObject.ScanFileExt = ".TIF"
         
        Case &H120000
            PixImageObject.ScanFileExt = ".JBIG"
         
        Case &H130000
            PixImageObject.ScanFileExt = ".PNG"
         
        Case &H140000
            PixImageObject.ScanFileExt = ".JP2"
         
    End Select

    funcPixGetFileExt = PixImageObject.ScanFileExt
    
End Function

