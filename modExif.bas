Attribute VB_Name = "modExif"
Option Explicit

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Source As Any, ByVal Length As Long)

Dim bytByteFormat As Byte   ' 0=Intel 1=Motorola
Dim strChunkData As String  ' The whole header chunk of the 'App data'
Global ListInfo() As String ' 2,X   0=TagNum 1=TagType 2=TagLength, all in Hex

Function ReadExif(strFileName As String, strListBox As ListBox) As Boolean
    'On Error Resume Next
    Dim I             As Long ' Generic Counter
    Dim ArrCnt        As Long ' Array Counter
    Dim First2Bytes   As String
    Dim Last2Bytes    As String
    Dim tmpStr        As String ' Temp value used within a few lines
    Dim AppDataLength As Long   ' "APP1"
    Dim Offset2IFD    As Long   ' Main offset used for directory thingies
    
    strListBox.Clear
    Open strFileName For Binary Access Read As #1
    
    
    
    
    ' Check to make sures its a jpeg file
    First2Bytes = Input(2, #1)
    Seek #1, FileLen(strFileName) - 1 ' Jump to near end
    Last2Bytes = Input(2, #1)
    If Byte2Hex(First2Bytes) <> "FFD8" Or Byte2Hex(Last2Bytes) <> "FFD9" Then
        strListBox.AddItem "Not a JPeg file"
        Close #1
        Exit Function
    End If
    
    
    
    
    ' Look for the start of the exif data 'Application Marker' - should be straight away
    I = 3
    Seek #1, I
    Do
        tmpStr = Input(2, 1)
        Select Case Left$(Byte2Hex(tmpStr), 3)
        Case "FFE" ' FFE0 - FFEF, hopefully FFE1
            ' Found Marker so continue, remembering the file position
            Exit Do
        Case Else
            tmpStr = Input(2, 1)
            I = I + (Byte2Dec(tmpStr) - 2) ' Size includes the size bytes as well
            Seek #1, I ' Skip the chunk of data
        End Select
    Loop Until EOF(1)
    If EOF(1) Then strListBox.AddItem "File information not found": Close #1: Exit Function
    
    
    
    
    ' Get some information about the structure
    AppDataLength = Byte2Dec(Input(2, 1)) - 2 ' Motorola byte format
    'strListBox.AddItem "Exif Application Data Length = " & AppDataLength: ArrCnt = ArrCnt + 1
    ' Confirm its really Exif
    tmpStr = Input(6, 1)
    If tmpStr <> "Exif" & Chr$(0) & Chr$(0) Then strListBox.AddItem "Not 'Exif' data, Panic": Close #1: Exit Function
    ' Get whole 'App Data' chunk info, (49492A00 08000000 - Common TIFF Header)
    strChunkData = Input(AppDataLength, 1)
    Select Case Mid$(strChunkData, 1, 2)
    Case "II": bytByteFormat = 0 ': strListBox.AddItem "Intel Header Format": ArrCnt = ArrCnt + 1 ' Reverse bytes
    Case "MM": bytByteFormat = 1: strListBox.AddItem "Motarola Header Format - Might have prob's": ArrCnt = ArrCnt + 1
    Case Else: strListBox.AddItem "Unknown/Error Header Format": Close #1: Exit Function
    End Select
    ' Skip the next 2 bytes = "002A"
    Offset2IFD = Byte2Dec(Rev(Mid$(strChunkData, 5, 4)))
    
    ' Main offsets are relitive to the TIFF header (II or MM)
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Dim NumofDirEntries     As Long
    Dim DirEntryInfo        As String
    Dim DataFormat          As Long
    Dim SpecialField        As Boolean
    Dim TagName             As String
    Dim SizeMultiplier      As Long
    Dim LenOfTagData        As Double
    Dim MoreChunkOffsets(2) As Long ' for more 'Exif' data, 'GPS' and 'Interoperability'
    Dim ActiveOffset        As Byte ' for just above (0 to 2), so the correct TagName field is selected
    
    
    
    
    ' Loops though the differently placed directories in the header
    Do
        
        
        ' Start of the Image File Directory (IFD)
        NumofDirEntries = Byte2Dec(Rev(Mid$(strChunkData, Offset2IFD + 1, 2)))
        'strListBox.AddItem "No of Dir Entries = " & NumofDirEntries: ArrCnt = ArrCnt + 1
        
        
        ' Run though the Directory Entries
        For I = 0 To NumofDirEntries - 1
            DirEntryInfo = Mid$(strChunkData, Offset2IFD + 3 + (I * 12), 12) ' Grab 12 bytes which hold the !file! information
            
            
            ' Define most of the I entry
            Select Case ActiveOffset                                    ' Get Descriptive Name
            Case 0: TagName = GetTagName(Rev(Mid$(DirEntryInfo, 1, 2)), SpecialField)
            Case 1: TagName = GetTagNameGPS(Rev(Mid$(DirEntryInfo, 1, 2)), SpecialField)
            Case 2: TagName = GetTagNameInteroperability(Rev(Mid$(DirEntryInfo, 1, 2)), SpecialField)
            End Select
            DataFormat = Byte2Dec(Rev(Mid$(DirEntryInfo, 3, 2)))        ' E.G. Long (DWord)
            SizeMultiplier = Byte2Dec(Rev(Mid$(DirEntryInfo, 5, 4)))    ' 2 of them
            LenOfTagData = CLng(TypeOfTag(DataFormat)) * SizeMultiplier ' 2 X Long(4)
            'If TagName = "Unknown" Then Stop
            
            
            'Save some data, for changing values
            ReDim Preserve ListInfo(2, ArrCnt) As String
            ListInfo(0, ArrCnt) = Byte2Hex(Rev(Mid$(DirEntryInfo, 1, 2)))
            ListInfo(1, ArrCnt) = Byte2Hex(Rev(Mid$(DirEntryInfo, 3, 2)))
            ListInfo(2, ArrCnt) = Hex(LenOfTagData)
            ArrCnt = ArrCnt + 1
            
            
            ' Grab data from within the I entry or grab from defined offset, convert special entry if needed
            If LenOfTagData <= 4 Then ' No Offset < 5 Bytes
                tmpStr = ConvertData2Format(DataFormat, Rev(Right$(Mid$(DirEntryInfo, 9, 4), LenOfTagData)))
            Else                      ' Offset required > 4 Bytes
                ' Strings seem to have a Chr(0) at the end
                If TagName = "User comments" Then LenOfTagData = (2 ^ 16) ' bodge fix :)
                tmpStr = ConvertData2Format(DataFormat, Mid$(strChunkData, Byte2Dec(Rev(Mid$(DirEntryInfo, 9, 4))) + 1, LenOfTagData))
            End If
            If SpecialField = True Then tmpStr = DefineSpecialTag(Rev(Mid$(DirEntryInfo, 1, 2)), tmpStr)
            strListBox.AddItem TagName & " = " & tmpStr
            strListBox.ItemData(strListBox.NewIndex) = Byte2Dec(Rev(Mid$(DirEntryInfo, 1, 2))) ' Save the tag number
            
            
            ' Offset pointers to other data chunks, there should only be max 1 each, in the whole file
            If TagName = "Exif IFD" Then ' Normally at the end of the list, onto the next
                MoreChunkOffsets(0) = ConvertData2Format(DataFormat, Rev(Right$(Mid$(DirEntryInfo, 9, 4), LenOfTagData)))
                strListBox.RemoveItem (strListBox.NewIndex): ArrCnt = ArrCnt - 1 ' Remove the info
            ElseIf TagName = "GPS IFD" Then
                MoreChunkOffsets(1) = ConvertData2Format(DataFormat, Rev(Right$(Mid$(DirEntryInfo, 9, 4), LenOfTagData)))
                strListBox.RemoveItem (strListBox.NewIndex): ArrCnt = ArrCnt - 1 ' Remove the info
            ElseIf TagName = "Interoperability IFD" Then
                MoreChunkOffsets(2) = ConvertData2Format(DataFormat, Rev(Right$(Mid$(DirEntryInfo, 9, 4), LenOfTagData)))
                strListBox.RemoveItem (strListBox.NewIndex): ArrCnt = ArrCnt - 1 ' Remove the info
            End If
        Next I
        
        
        ' Offset of next Directory structure, kinda like FAT
        If MoreChunkOffsets(0) <> 0 Then
            Offset2IFD = MoreChunkOffsets(0): MoreChunkOffsets(0) = 0: ActiveOffset = 0
        ElseIf MoreChunkOffsets(1) <> 0 Then
            Offset2IFD = MoreChunkOffsets(1): MoreChunkOffsets(1) = 0: ActiveOffset = 1
        ElseIf MoreChunkOffsets(2) <> 0 Then
            Offset2IFD = MoreChunkOffsets(2): MoreChunkOffsets(2) = 0: ActiveOffset = 2
        Else
            Exit Do
        End If
        
    Loop
    
    
    Close #1
    ReadExif = True
End Function

Function Hex2Byte(InHex As String) As String ' Conv Hex to Bytes
    Dim I As Long
    
    For I = 1 To Len(InHex) Step 2
        Hex2Byte = Hex2Byte & Chr$(CLng("&H" & Mid$(InHex, I, 2)))
    Next I
End Function

Function Byte2Dec(InBytes As String) As Double ' Conv. Bytes to Decimal - Could be > 4 Billion
    Dim I As Long
    Dim tmp As String
    
    For I = 1 To Len(InBytes)
        tmp = tmp & Hex(Format$(Asc(Mid$(InBytes, I, 1)), "00"))
    Next I
    Byte2Dec = "&H" & tmp
End Function

Function Byte2Hex(InBytes As String) As String ' Conv. Bytes to Hex
    Dim I As Long
    Dim tmp As String
    
    For I = 1 To Len(InBytes)
        tmp = Hex(Asc(Mid$(InBytes, I, 1)))
        Byte2Hex = Byte2Hex & String$(2 - Len(tmp), "0") & tmp
    Next I
End Function

Function Rev(InBytes As String) As String ' Reverse the byte order
    If bytByteFormat = 1 Then Rev = InBytes: Exit Function ' Not needed for Motorola format
    
    Dim I As Long
    Dim tmp As String
    
    For I = Len(InBytes) To 1 Step -1
        tmp = tmp & Mid$(InBytes, I, 1)
    Next I
    Rev = tmp
End Function

Function GetTagName(TagNum As String, retSpecialField As Boolean) As String
    retSpecialField = False
    
    Select Case Byte2Hex(TagNum)
    Case "00FE": GetTagName = "New Subfile Type" ' ?
    Case "00FF": GetTagName = "SubfileType"      ' ?
    Case "0100": GetTagName = "Image width"
    Case "0101": GetTagName = "Image height"
    Case "0102": GetTagName = "Number of bits per component"
    Case "0103": GetTagName = "Compression scheme": retSpecialField = True
    Case "0106": GetTagName = "Pixel composition": retSpecialField = True
    Case "010E": GetTagName = "Image title"
    Case "010F": GetTagName = "Manufacturer of image input equipment"
    Case "0110": GetTagName = "Model of image input equipment"
    Case "0111": GetTagName = "Strip Offsets"
    Case "0112": GetTagName = "Orientation of image": retSpecialField = True ': Stop
    Case "0115": GetTagName = "Samples Per Pixel"
    Case "0116": GetTagName = "Number of rows per strip"
    Case "0117": GetTagName = "Bytes per compressed strip"
    Case "011A": GetTagName = "Image resolution in width"
    Case "011B": GetTagName = "Image resolution in height"
    Case "011C": GetTagName = "Image data arrangement, Planar Configuration": retSpecialField = True
    Case "0128": GetTagName = "Unit of X and Y resolution": retSpecialField = True
    Case "012D": GetTagName = "Transfer function"
    Case "0131": GetTagName = "Software used"
    Case "0132": GetTagName = "File change date and time"
    Case "013B": GetTagName = "Person who created the image"
    Case "013D": GetTagName = "Predictor"
    Case "013E": GetTagName = "White point chromaticity"
    Case "013F": GetTagName = "Chromaticities of primaries"
    Case "0142": GetTagName = "Tile Width"
    Case "0143": GetTagName = "Tile Length"
    Case "0144": GetTagName = "Tile Offsets"
    Case "0145": GetTagName = "Tile Byte Counts"
    Case "014A": GetTagName = "Sub IFDs"
    Case "015B": GetTagName = "JPEG Tables"
    Case "0200": GetTagName = "Special Mode"
    Case "0201": GetTagName = "Offset to JPEG SOI"
    Case "0202": GetTagName = "Bytes of JPEG data"
    Case "0204": GetTagName = "Digi Zoom"
    Case "0207": GetTagName = "Software Release"
    Case "0208": GetTagName = "Pict Info"
    Case "0209": GetTagName = "Camera ID"
    Case "0211": GetTagName = "Color space transformation matrix coefficients"
    Case "0212": GetTagName = "Subsampling ratio of Y to C": retSpecialField = True
    Case "0213": GetTagName = "Y and C positioning": retSpecialField = True
    Case "0214": GetTagName = "Pair of black and white reference values"
    Case "0F00": GetTagName = "Data Dump"
    Case "828D": GetTagName = "CFA Repeat Pattern Dim"
    Case "828E": GetTagName = "CFA Pattern"
    Case "828F": GetTagName = "Battery Level"
    Case "8298": GetTagName = "Person who created the image Copyright holder"
    Case "829A": GetTagName = "Exposure time"
    Case "829D": GetTagName = "F number"
    Case "83BB": GetTagName = "IPTC/NAA"
    Case "8769": GetTagName = "Exif IFD" ' ***
    Case "8773": GetTagName = "InterColorProfile"
    Case "8822": GetTagName = "Exposure program": retSpecialField = True
    Case "8824": GetTagName = "Spectral sensitivity"
    Case "8825": GetTagName = "GPS IFD" ' ***
    Case "8827": GetTagName = "ISO speed rating"
    Case "8828": GetTagName = "Optoelectric conversion factor"
    Case "8829": GetTagName = "Interlace"
    Case "882A": GetTagName = "Time Zone Offset"
    Case "882B": GetTagName = "Self Timer Mode"
    Case "9000": GetTagName = "Exif Version"
    Case "9003": GetTagName = "Date and time original image was generated"
    Case "9004": GetTagName = "Date and time image was made digital data"
    Case "9101": GetTagName = "Component Configuration": retSpecialField = True
    Case "9102": GetTagName = "Image compression mode"
    Case "9201": GetTagName = "Shutter Speed"
    Case "9202": GetTagName = "Aperture Value"
    Case "9203": GetTagName = "Brightness Value"
    Case "9204": GetTagName = "Exposure Bias Value"
    Case "9205": GetTagName = "Maximum lens aperture"
    Case "9206": GetTagName = "Subject distance"
    Case "9207": GetTagName = "Metering mode": retSpecialField = True
    Case "9208": GetTagName = "Light source": retSpecialField = True
    Case "9209": GetTagName = "Flash": retSpecialField = True
    Case "920A": GetTagName = "Lens focal length"
    Case "920B": GetTagName = "Flash Energy"
    Case "920C": GetTagName = "Spatial Frequency Response"
    Case "920D": GetTagName = "Noise"
    Case "9211": GetTagName = "Image Number"
    Case "9212": GetTagName = "Security Classification"
    Case "9213": GetTagName = "Image History"
    Case "9214": GetTagName = "Subject area": retSpecialField = True
    Case "9215": GetTagName = "Exposure Index"
    Case "9216": GetTagName = "TIFF/EP Standard ID"
    Case "927C": GetTagName = "Manufacturer notes": retSpecialField = True ': Stop
    Case "9286": GetTagName = "User comments": retSpecialField = True ': Stop
    Case "9290": GetTagName = "DateTime subseconds"
    Case "9291": GetTagName = "DateTimeOriginal subseconds"
    Case "9292": GetTagName = "DateTimeDigitized subseconds"
    Case "A000": GetTagName = "Supported Flashpix version"
    Case "A001": GetTagName = "Color space information": retSpecialField = True
    Case "A002": GetTagName = "Exif Image Width"
    Case "A003": GetTagName = "Exif Image Height"
    Case "A004": GetTagName = "Related audio file"
    Case "A005": GetTagName = "Interoperability IFD" ' ***
    Case "A20B": GetTagName = "Flash energy"
    Case "A20C": GetTagName = "Spatial frequency response"
    Case "A20E": GetTagName = "Focal plane X resolution"
    Case "A20F": GetTagName = "Focal plane Y resolution"
    Case "A210": GetTagName = "Focal plane resolution unit"
    Case "A214": GetTagName = "Subject location"
    Case "A215": GetTagName = "Exposure index"
    Case "A217": GetTagName = "Sensing method": retSpecialField = True
    Case "A300": GetTagName = "File source": retSpecialField = True
    Case "A301": GetTagName = "Scene type": retSpecialField = True
    Case "A302": GetTagName = "CFA pattern": retSpecialField = True
    Case "A401": GetTagName = "Custom image processing": retSpecialField = True
    Case "A402": GetTagName = "Exposure mode": retSpecialField = True
    Case "A403": GetTagName = "White balance": retSpecialField = True
    Case "A404": GetTagName = "Digital zoom ratio"
    Case "A405": GetTagName = "Focal length in 35 mm film"
    Case "A406": GetTagName = "Scene capture type": retSpecialField = True
    Case "A407": GetTagName = "Gain control": retSpecialField = True
    Case "A408": GetTagName = "Contrast": retSpecialField = True
    Case "A409": GetTagName = "Saturation": retSpecialField = True
    Case "A40A": GetTagName = "Sharpness": retSpecialField = True
    Case "A40B": GetTagName = "Device settings description"
    Case "A40C": GetTagName = "Subject distance range": retSpecialField = True
    Case "A420": GetTagName = "Unique image ID"
    'Case "": GetTagName = ""
    Case Else: GetTagName = "Unknown"
    End Select
End Function

Function GetTagNameGPS(TagNum As String, retSpecialField As Boolean) As String
    retSpecialField = False
    
    Select Case Byte2Hex(TagNum)
    Case "0000": GetTagNameGPS = "GPS tag version"
    Case "0001": GetTagNameGPS = "GPS North or South Latitude"
    Case "0002": GetTagNameGPS = "GPS Latitude"
    Case "0003": GetTagNameGPS = "GPS East or West Longitude"
    Case "0004": GetTagNameGPS = "GPS Longitude"
    Case "0005": GetTagNameGPS = "GPS Altitude Reference"
    Case "0006": GetTagNameGPS = "GPS Altitude"
    Case "0007": GetTagNameGPS = "GPS time (atomic clock)"
    Case "0008": GetTagNameGPS = "GPS satellites used for measurement"
    Case "0009": GetTagNameGPS = "GPS receiver status"
    Case "000A": GetTagNameGPS = "GPS measurement mode"
    Case "000B": GetTagNameGPS = "GPS Measurement precision"
    Case "000C": GetTagNameGPS = "GPS Speed unit"
    Case "000D": GetTagNameGPS = "GPS Speed of GPS receiver"
    Case "000E": GetTagNameGPS = "GPS Reference for direction of movement"
    Case "000F": GetTagNameGPS = "GPS Direction of movement"
    Case "0010": GetTagNameGPS = "GPS Reference for direction of image"
    Case "0011": GetTagNameGPS = "GPS Direction of image"
    Case "0012": GetTagNameGPS = "GPS Geodetic survey data used"
    Case "0013": GetTagNameGPS = "GPS Reference for latitude of destination"
    Case "0014": GetTagNameGPS = "GPS Latitude of destination"
    Case "0015": GetTagNameGPS = "GPS Reference for longitude of destination"
    Case "0016": GetTagNameGPS = "GPS Longitude of destination"
    Case "0017": GetTagNameGPS = "GPS Reference for bearing of destination"
    Case "0018": GetTagNameGPS = "GPS Bearing of destination"
    Case "0019": GetTagNameGPS = "GPS Reference for distance to destination"
    Case "001A": GetTagNameGPS = "GPS Distance to destination"
    Case "001B": GetTagNameGPS = "GPS Name of GPS processing method"
    Case "001C": GetTagNameGPS = "GPS Name of GPS area"
    Case "001D": GetTagNameGPS = "GPS date"
    Case "001E": GetTagNameGPS = "GPS differential correction"
    End Select
End Function

Function GetTagNameInteroperability(TagNum As String, retSpecialField As Boolean) As String
    retSpecialField = False
    
    Select Case Byte2Hex(TagNum)
    Case "0001": GetTagNameInteroperability = "Interoperability Identification": retSpecialField = True
    Case "0002": GetTagNameInteroperability = "Interoperability Version"
    Case "1000": GetTagNameInteroperability = "Related Image File Format"
    Case "1001": GetTagNameInteroperability = "Related Image Width\Length"
    End Select
End Function

Function DefineSpecialTag(TagNum As String, InStri As String) As String
    Select Case Byte2Hex(TagNum)
    Case "0001"                   ' Interoperability Identification
        Select Case Trim$(InStri)
        Case "R98": DefineSpecialTag = "Conforming to R98 file specification"
        Case "THM": DefineSpecialTag = "The file conforming to DCF thumbnail file"
        End Select
        
        
        
        
        
        
    Case "0103" ' Compression
        Select Case Val(InStri)
        Case 0: DefineSpecialTag = "Uncompressed"
        Case 6: DefineSpecialTag = "JPEG compression (thumbnails only)"
        End Select
        
    Case "0106" ' PhotometricInterpretation
        Select Case Val(InStri)
        Case 2: DefineSpecialTag = "RGB"
        Case 6: DefineSpecialTag = "YCbCr"
        End Select
        
    Case "0112" ' Orientation
        Select Case Val(InStri)
        Case 1: DefineSpecialTag = "1st Row is top, 1st Col is left"     'The 0th row is at the visual top of the image, and the 0th column is the visual left-hand side
        Case 2: DefineSpecialTag = "1st Row is top, 1st Col is right"    'The 0th row is at the visual top of the image, and the 0th column is the visual right-hand side
        Case 3: DefineSpecialTag = "1st Row is bottom, 1st Col is right" 'The 0th row is at the visual bottom of the image, and the 0th column is the visual right-hand side
        Case 4: DefineSpecialTag = "1st Row is bottom, 1st Col is left"  'The 0th row is at the visual bottom of the image, and the 0th column is the visual left-hand side
        Case 5: DefineSpecialTag = "1st Row is left, 1st Col is top"     'The 0th row is the visual left-hand side of the image, and the 0th column is the visual top
        Case 6: DefineSpecialTag = "1st Row is right, 1st Col is top"    'The 0th row is the visual right-hand side of the image, and the 0th column is the visual top
        Case 7: DefineSpecialTag = "1st Row is right, 1st Col is bottom" 'The 0th row is the visual right-hand side of the image, and the 0th column is the visual bottom
        Case 8: DefineSpecialTag = "1st Row is left, 1st Col is bottom"  'The 0th row is the visual left-hand side of the image, and the 0th column is the visual bottom
        End Select
        
    Case "011C" ' Image data arrangement, Planar Configuration
        Select Case Val(InStri)
        Case 0: DefineSpecialTag = "Chunky format"
        Case 1: DefineSpecialTag = "Planar format"
        End Select
        
    Case "0128" ' Unit of X and Y resolution
        Select Case Val(InStri)
        Case 2: DefineSpecialTag = "Inches"
        Case 3: DefineSpecialTag = "Centimeters"
        End Select
        
    Case "0212" ' Subsampling ratio of Y to C
        Select Case Val(InStri)
        Case 131073: DefineSpecialTag = "YCbCr4:2:2"
        Case 131074: DefineSpecialTag = "YCbCr4:2:0"
        End Select
        
    Case "0213" ' Y and C positioning
        Select Case Val(InStri)
        Case 1: DefineSpecialTag = "Centered"
        Case 2: DefineSpecialTag = "Co-sited"
        End Select
        
    Case "8822" ' Exposure program
        Select Case Val(InStri)
        Case 1: DefineSpecialTag = "Manual"
        Case 2: DefineSpecialTag = "Normal program"
        Case 3: DefineSpecialTag = "Aperture priority"
        Case 4: DefineSpecialTag = "Shutter priority"
        Case 5: DefineSpecialTag = "Creative program (biased toward depth of field)"
        Case 6: DefineSpecialTag = "Action program (biased toward fast shutter speed)"
        Case 7: DefineSpecialTag = "Portrait mode (for closeup photos with the background out of focus)"
        Case 8: DefineSpecialTag = "Landscape mode (for landscape photos with the background in focus)"
        End Select
        
    Case "9101" ' Component Configuration
        If InStr(1, InStri, "1") > 0 Then DefineSpecialTag = "Y "
        If InStr(1, InStri, "2") > 0 Then DefineSpecialTag = DefineSpecialTag & "Cb "
        If InStr(1, InStri, "3") > 0 Then DefineSpecialTag = DefineSpecialTag & "Cr "
        If InStr(1, InStri, "4") > 0 Then DefineSpecialTag = DefineSpecialTag & "R "
        If InStr(1, InStri, "5") > 0 Then DefineSpecialTag = DefineSpecialTag & "G "
        If InStr(1, InStri, "6") > 0 Then DefineSpecialTag = DefineSpecialTag & "B "
        
    Case "9207" ' Metering mode
        Select Case Val(InStri)
        Case 1: DefineSpecialTag = "Average"
        Case 2: DefineSpecialTag = "Center Weighted Average"
        Case 3: DefineSpecialTag = "Spot"
        Case 4: DefineSpecialTag = "MultiSpot"
        Case 5: DefineSpecialTag = "Pattern"
        Case 6: DefineSpecialTag = "Partial"
        End Select
        
    Case "9208" ' Light source
        Select Case Val(InStri)
        Case 0: DefineSpecialTag = "Unknown"
        Case 1: DefineSpecialTag = "Daylight"
        Case 2: DefineSpecialTag = "Fluorescent"
        Case 3: DefineSpecialTag = "Tungsten"
        Case 4: DefineSpecialTag = "Flash"
        Case 9: DefineSpecialTag = "Fine weather"
        Case 10: DefineSpecialTag = "Cloudy weather"
        Case 11: DefineSpecialTag = "Shade"
        Case 12: DefineSpecialTag = "Daylight fluorescent (D 5700 – 7100K)"
        Case 13: DefineSpecialTag = "Day white fluorescent (N 4600 – 5400K)"
        Case 14: DefineSpecialTag = "Cool white fluorescent (W 3900 – 4500K)"
        Case 15: DefineSpecialTag = "White fluorescent (WW 3200 – 3700K)"
        Case 17: DefineSpecialTag = "Standard light A"
        Case 18: DefineSpecialTag = "Standard light B"
        Case 19: DefineSpecialTag = "Standard light C"
        Case 20: DefineSpecialTag = "D55"
        Case 21: DefineSpecialTag = "D65"
        Case 22: DefineSpecialTag = "D75"
        Case 23: DefineSpecialTag = "D50"
        Case 24: DefineSpecialTag = "ISO studio tungsten"
        End Select
        
    Case "9214" ' Subject area
        Select Case Val(InStri)
        Case 2: DefineSpecialTag = "The main subject given as x y coordinates"
        Case 3: DefineSpecialTag = "The area of the main subject is given as a circle"
        Case 4: DefineSpecialTag = "The area of the main subject is given as a rectangle"
        End Select
        
    Case "927C" ' Maker Notes
        DefineSpecialTag = Mid$(InStri, 419, 54) ' Crude - Cannon
        
    Case "9286" ' User Comment
        DefineSpecialTag = Mid$(InStri, 9, InStr(10, InStri, Chr$(0)) - 9) ' Skip 8 byte start info
        
    Case "A001" ' Color space information
        Select Case Val(InStri)
        Case 1: DefineSpecialTag = "sRGB"
        Case 65535: DefineSpecialTag = "Uncalibrated"
        End Select
        
    Case "A217" ' Sensing method
        Select Case Val(InStri)
        Case 1: DefineSpecialTag = "Not defined"
        Case 2: DefineSpecialTag = "One-chip color area sensor"
        Case 3: DefineSpecialTag = "Two-chip color area sensor"
        Case 4: DefineSpecialTag = "Three-chip color area sensor"
        Case 5: DefineSpecialTag = "Color sequential area sensor"
        Case 6: DefineSpecialTag = "Trilinear sensor"
        Case 7: DefineSpecialTag = "Color sequential linear sensor"
        End Select
        
    Case "A300" ' File source
        If Val(InStri) = 3 Then DefineSpecialTag = "DSC"
        
    Case "A301" ' Scene type
        If Val(InStri) = 1 Then DefineSpecialTag = "A directly photographed image"
        
    Case "A302" ' CFA pattern
        DefineSpecialTag = "Sorry skipping this :(" ' Not borthered about this one
        
    Case "A401" ' Custom image processing
        Select Case Val(InStri)
        Case 0: DefineSpecialTag = "Normal process"
        Case 1: DefineSpecialTag = "Custom process"
        End Select
        
    Case "A402" ' Exposure mode
        Select Case Val(InStri)
        Case 0: DefineSpecialTag = "Auto exposure"
        Case 1: DefineSpecialTag = "Manual exposure"
        Case 2: DefineSpecialTag = "Auto bracket"
        End Select
        
    Case "A403" ' White balance
        Select Case Val(InStri)
        Case 0: DefineSpecialTag = "Auto white balance"
        Case 1: DefineSpecialTag = "Manual white balance"
        End Select
        
    Case "A406" ' Scene capture type
        Select Case Val(InStri)
        Case 0: DefineSpecialTag = "Standard"
        Case 1: DefineSpecialTag = "Landscape"
        Case 2: DefineSpecialTag = "Portrait"
        Case 3: DefineSpecialTag = "Night scene"
        End Select
        
    Case "A407" ' Gain control
        Select Case Val(InStri)
        Case 0: DefineSpecialTag = "None"
        Case 1: DefineSpecialTag = "Low gain up"
        Case 2: DefineSpecialTag = "High gain up"
        Case 3: DefineSpecialTag = "Low gain down"
        Case 4: DefineSpecialTag = "High gain down"
        End Select
        
    Case "A408" ' Contrast
        Select Case Val(InStri)
        Case 0: DefineSpecialTag = "Normal"
        Case 1: DefineSpecialTag = "Soft"
        Case 2: DefineSpecialTag = "Hard"
        End Select
        
    Case "A409" ' Saturation
        Select Case Val(InStri)
        Case 0: DefineSpecialTag = "Normal"
        Case 1: DefineSpecialTag = "Low saturation"
        Case 2: DefineSpecialTag = "High saturation"
        End Select
        
    Case "A40A" ' Sharpness
        Select Case Val(InStri)
        Case 0: DefineSpecialTag = "Normal"
        Case 1: DefineSpecialTag = "Soft"
        Case 2: DefineSpecialTag = "Hard"
        End Select
        
    Case "A40C" ' Subject distance range
        Select Case Val(InStri)
        Case 1: DefineSpecialTag = "Macro"
        Case 2: DefineSpecialTag = "Close View"
        Case 3: DefineSpecialTag = "Distant view"
        End Select
        
    End Select
    
    If DefineSpecialTag = "" Then DefineSpecialTag = "Undefined/Unknown"
End Function

Function TypeOfTag(InDec As Long) As Byte
    ' InDec, Bytes per component, Type
    '  1     1                    unsigned byte
    '  2     1                    ascii Strings
    '  3     2                    unsigned Short
    '  4     4                    unsigned long
    '  5     8                    unsigned rational
    '  6     1                    signed byte
    '  7     1                    undefined       ' Normally more tags within
    '  8     2                    signed Short
    '  9     4                    signed long
    ' 10     8                    signed rational
    ' 11     4                    single float    ' Not a standard
    ' 12     8                    double float    ' Not a standard
    
    Select Case InDec
    Case 1:  TypeOfTag = 1
    Case 2:  TypeOfTag = 1
    Case 3:  TypeOfTag = 2
    Case 4:  TypeOfTag = 4
    Case 5:  TypeOfTag = 8
    Case 6:  TypeOfTag = 1
    Case 7:  TypeOfTag = 1
    Case 8:  TypeOfTag = 2
    Case 9:  TypeOfTag = 4
    Case 10: TypeOfTag = 8
    Case 11: TypeOfTag = 4
    Case 12: TypeOfTag = 8
    End Select
End Function

Function ConvertData2Format(DataFormat As Long, InBytes As String) As String
    ' Read function aboves details
    ' Double check for Motorola format esp. CopyMemory
    Dim tmpInt As Integer
    Dim tmpLng As Long
    Dim tmpSng As Single
    Dim tmpDbl As Double
    'Dim tmpStr As String * 4
    
    Select Case DataFormat
    Case 1, 3, 4: ConvertData2Format = Byte2Dec(InBytes)
    Case 2, 7: ConvertData2Format = InBytes
    
    Case 5 ' Unsigned Sortof Fraction (Longs)
        ConvertData2Format = CDbl(Byte2Dec(Rev(Mid$(InBytes, 1, 4)))) / CDbl(Byte2Dec(Rev(Mid$(InBytes, 5, 4))))
        
    Case 6 ' Signed Byte
        tmpLng = Byte2Dec(InBytes)
        If tmpLng > 127 Then ConvertData2Format = -(tmpLng Xor 127) Else ConvertData2Format = tmpLng
        
    Case 8 ' Signed Short
        tmpLng = Byte2Dec(Rev(InBytes))
        If tmpLng > 32767 Then ConvertData2Format = -(tmpLng Xor 32767) Else ConvertData2Format = tmpLng
        'CopyMemory tmpInt, InBytes, 2
        'ConvertData2Format = tmpInt
        
    Case 9 ' Signed Long
        tmpDbl = Byte2Dec(Rev(InBytes))
        If tmpDbl > 2147483647 Then ConvertData2Format = -(tmpDbl Xor 2147483647) Else ConvertData2Format = tmpDbl
        'CopyMemory tmpLng, InBytes, 4
        'ConvertData2Format = tmpLng
        
    Case 10 ' Signed Sortof Fraction (Longs)
        tmpDbl = Byte2Dec(Rev(Mid$(InBytes, 1, 4)))
        If tmpDbl > 2147483647 Then ConvertData2Format = -(tmpDbl Xor 2147483647) Else ConvertData2Format = tmpDbl
        tmpDbl = Byte2Dec(Rev(Mid$(InBytes, 5, 4)))
        If tmpDbl > 2147483647 Then
            ConvertData2Format = ConvertData2Format / -(tmpDbl Xor 2147483647)
        Else
            ConvertData2Format = ConvertData2Format / tmpDbl
        End If
        '      Its like CopyMemory does not work!, it produces different outputs from constant inputs!
        'ZeroMemory tmpLng, 4
        'tmpStr = Mid$(InBytes, 1, 4)
        'frmMain.List1.AddItem Byte2Hex(Mid$(InBytes, 1, 4))
        'x = VarPtr(tmpLng)
        'y = StrPtr(tmpStr)
        'CopyMemory tmpLng, tmpStr, 4
        'ConvertData2Format = x
        'tmpStr = Mid$(InBytes, 5, 4)
        'CopyMemory tmpLng, tmpStr, 4
        'ConvertData2Format = ConvertData2Format / tmpLng
        
    Case 11 ' Single Float
        CopyMemory tmpSng, InBytes, 4 ' *** Probably will not work
        ConvertData2Format = tmpSng
        
    Case 12 ' Double Float
        CopyMemory tmpDbl, InBytes, 8 ' *** Probably will not work
        ConvertData2Format = tmpDbl
    End Select
End Function
























Function ChangeExif(strFileName As String, LstIndex As Long, NewVal1 As String, NewVal2 As String)
    ' This function is just a copy of the read function with a few commented changes
    
    'On Error Resume Next
    Dim I             As Long
    Dim tmpStr        As String
    Dim AppDataLength As Long
    Dim Offset2IFD    As Long
    
    Open strFileName For Binary Access Read Write As #5
    I = 3
    Seek #5, I
    Do
        tmpStr = Input(2, 5)
        Select Case Left$(Byte2Hex(tmpStr), 3)
        Case "FFE": Exit Do
        Case Else
            tmpStr = Input(2, 5)
            I = I + (Byte2Dec(tmpStr) - 2)
            Seek #5, I
        End Select
    Loop Until EOF(5)
    If EOF(5) Then Close #5: Exit Function
    
    AppDataLength = Byte2Dec(Input(2, 5)) - 2
    tmpStr = Input(6, 5) ' Waster
    strChunkData = Input(AppDataLength, 5)
    Offset2IFD = Byte2Dec(Rev(Mid$(strChunkData, 5, 4)))
    
    Dim NumofDirEntries     As Long
    Dim DirEntryInfo        As String
    Dim DataFormat          As Long
    Dim SpecialField        As Boolean
    Dim TagName             As String
    Dim SizeMultiplier      As Long
    Dim LenOfTagData        As Double
    Dim MoreChunkOffsets(2) As Long
    Dim ActiveOffset        As Byte
    
    Dim RawData             As String ' Temporary holds the converted format of a number/string
    
    Do
        NumofDirEntries = Byte2Dec(Rev(Mid$(strChunkData, Offset2IFD + 1, 2)))
        For I = 0 To NumofDirEntries - 1
            DirEntryInfo = Mid$(strChunkData, Offset2IFD + 3 + (I * 12), 12)
            
            Select Case ActiveOffset
            Case 0: TagName = GetTagName(Rev(Mid$(DirEntryInfo, 1, 2)), SpecialField)
            Case 1: TagName = GetTagNameGPS(Rev(Mid$(DirEntryInfo, 1, 2)), SpecialField)
            Case 2: TagName = GetTagNameInteroperability(Rev(Mid$(DirEntryInfo, 1, 2)), SpecialField)
            End Select
            DataFormat = Byte2Dec(Rev(Mid$(DirEntryInfo, 3, 2)))
            SizeMultiplier = Byte2Dec(Rev(Mid$(DirEntryInfo, 5, 4)))
            LenOfTagData = CLng(TypeOfTag(DataFormat)) * SizeMultiplier
            
            
            
            ' Check for a tag match
            If ListInfo(0, LstIndex) = Byte2Hex(Rev(Mid$(DirEntryInfo, 1, 2))) Then
                
                ' Convert the values into there correct formats to write
                RawData = ConvertData2FormatBackwards(DataFormat, NewVal1)
                Select Case DataFormat
                Case 5, 10: RawData = RawData & ConvertData2FormatBackwards(DataFormat, NewVal2) ' Change 2nd for rational
                End Select
                
                ' Write the new raw data
                If LenOfTagData <= 4 Then
                    Seek #5, Offset2IFD + 3 + (I * 12) + 20 + 2
                    Put #5, , RawData
                Else
                    tmpStr = ConvertData2Format(DataFormat, Mid$(strChunkData, Byte2Dec(Rev(Mid$(DirEntryInfo, 9, 4))) + 1, LenOfTagData))
                    Seek #5, Byte2Dec(Rev(Mid$(DirEntryInfo, 9, 4))) + 1 + 12 ' +12 for Tiff header
                    Put #5, , RawData
                End If
            End If
            
            
            
            If TagName = "Exif IFD" Then
                MoreChunkOffsets(0) = ConvertData2Format(DataFormat, Rev(Right$(Mid$(DirEntryInfo, 9, 4), LenOfTagData)))
            ElseIf TagName = "GPS IFD" Then
                MoreChunkOffsets(1) = ConvertData2Format(DataFormat, Rev(Right$(Mid$(DirEntryInfo, 9, 4), LenOfTagData)))
            ElseIf TagName = "Interoperability IFD" Then
                MoreChunkOffsets(2) = ConvertData2Format(DataFormat, Rev(Right$(Mid$(DirEntryInfo, 9, 4), LenOfTagData)))
            End If
        Next I
        
        If MoreChunkOffsets(0) <> 0 Then
            Offset2IFD = MoreChunkOffsets(0): MoreChunkOffsets(0) = 0: ActiveOffset = 0
        ElseIf MoreChunkOffsets(1) <> 0 Then
            Offset2IFD = MoreChunkOffsets(1): MoreChunkOffsets(1) = 0: ActiveOffset = 1
        ElseIf MoreChunkOffsets(2) <> 0 Then
            Offset2IFD = MoreChunkOffsets(2): MoreChunkOffsets(2) = 0: ActiveOffset = 2
        Else
            Exit Do
        End If
    Loop
    Close #5
End Function

Function ConvertData2FormatBackwards(DataFormat As Long, InValue As String) As String
    ' This function mostly converts the number to hex, then groups of 2 hex's to a character
    
    Dim I As Long ' Counter
    Dim tmpInt As Integer
    Dim tmpLng As Long
    Dim tmpSng As Single
    Dim tmpDbl As Double
    Dim tmpStr1 As String, tmpStr2 As String ' Converter helpers
    Dim tmpStr4 As String * 4
    Dim tmpStr8 As String * 8
    'Dim tmpCur1 As Currency, tmpCur2 As Currency ' holds a big slow whole value
    
    Select Case DataFormat
    Case 1 ' Unsigned Byte
        ConvertData2FormatBackwards = Chr$("&H" & Right$(Hex(Val(InValue)), 2))
        
    Case 2, 7: ConvertData2FormatBackwards = InValue
        
    Case 3 ' Unsigned Short
        tmpStr1 = Hex(Val(InValue))
        tmpStr1 = "000" & tmpStr1                            ' make sure the length is at least 2 bytes
        For I = Len(tmpStr1) - 1 To Len(tmpStr1) - 3 Step -2 ' get the last 2 bytes backwards
            ConvertData2FormatBackwards = ConvertData2FormatBackwards & Chr$("&H" & Mid$(tmpStr1, I, 2))
        Next I
        
    Case 4, 5 ' Unsigned Long
        tmpStr1 = Hex(Val(InValue))
        tmpStr1 = "0000000" & tmpStr1                        ' make sure the length is at least 4 bytes
        For I = Len(tmpStr1) - 1 To Len(tmpStr1) - 7 Step -2 ' get the last 4 bytes backwards
            ConvertData2FormatBackwards = ConvertData2FormatBackwards & Chr$("&H" & Mid$(tmpStr1, I, 2))
        Next I
        
    Case 6 ' Signed Byte
        tmpLng = Val(InValue)
        If tmpLng < 0 Then tmpStr2 = (tmpLng Xor -256) Else tmpStr2 = tmpLng ' use the correct bit arrangment for converting
        ConvertData2FormatBackwards = ConvertData2FormatBackwards(1, tmpStr2)
        
    Case 8 ' Signed Short
        tmpLng = Val(InValue)
        If tmpLng < 0 Then tmpStr2 = (tmpLng Xor -65536) Else tmpStr2 = tmpLng
        ConvertData2FormatBackwards = ConvertData2FormatBackwards(3, tmpStr2)
        
    Case 9, 10 ' Signed Long - this one is not properly tested
        'tmpCur2 = -4294967296#
        tmpLng = (2 ^ 30)
        If Val(InValue) < 0 Then
            tmpDbl = Val(InValue) Xor tmpLng       ' 30 bits   ' Stops overflow problems with xor
            If tmpDbl > tmpLng Then
                tmpDbl = tmpDbl Xor tmpLng         ' 31 bits
                If tmpDbl > tmpLng Then
                    tmpDbl = tmpDbl Xor tmpLng     ' 31.5 bits
                    If tmpDbl > tmpLng Then
                        tmpDbl = tmpDbl Xor tmpLng ' 32 bits
                    End If
                End If
            End If
        Else
            tmpStr2 = Val(InValue)
        End If
        tmpStr1 = Hex(Val(tmpStr2))
        tmpStr1 = "0000000" & tmpStr1                        ' make sure the length is at least 4 bytes
        For I = Len(tmpStr1) - 1 To Len(tmpStr1) - 7 Step -2 ' get the last 4 bytes backwards
            ConvertData2FormatBackwards = ConvertData2FormatBackwards & Chr$("&H" & Mid$(tmpStr1, I, 2))
        Next I
        'Stop
        
    Case 11 ' Single Float
        tmpSng = CSng(InValue)
        CopyMemory tmpStr4, tmpSng, 4 ' *** Probably will crash the program - need a bit of help here
        ConvertData2FormatBackwards = tmpStr4
        
    Case 12 ' Double Float
        tmpDbl = CDbl(InValue)
        CopyMemory tmpStr8, tmpDbl, 8 ' *** Probably will crash the program
        ConvertData2FormatBackwards = tmpStr8
    End Select
End Function
