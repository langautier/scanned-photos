# Issues with my collection of 30 years scanned photo
I have more than 3 000 photos scanned from negatives with two issues 
- A number of images were scanned with too high resolution which resulted in files being too large
- The images were classified in the scanning order very far from chronological classification I need for my albums

I wrote this utility for single use, so don't be angry with me if it's a little bit rustic. Converting and updating metadata is done through two proven programs to reduce the risk of image damage.

# Using Word to sort this out
The need is to display images with text next to them and to be able to easily move when sorting.

I had as constraint that it is simple and quick to do given my current skills and the single use.

So, Word and VBA will be the convenient solution for me.

As a result, two additional tools are necessary :
- [ImageMagick](https://imagemagick.org/) for resizing images files
- [ExifTool](https://exiftool.org/) for reading, writing and editing meta information in images files

# Main functionalities
## List and select images in you collection
Display to see this menu which allows you to:
- select one or more photo directories
- filter the selection according to the tags associated with the photo, the type of device or the size of the photo
- and use command button as:
    - __reset cache__ to delete the Exiftool cached files and force a new reading of metadata in the images files
    - __run__ to display in Word the selecting photos
    - __convert__ to resize all images over the targeted size

![Main screen](https://github.com/langautier/scanned-photos/blob/master/mainscreen.png)

## Sort photos by dates and edit meta tags
Once in word, you can select one or more photos and use buttons in the banner :
- __display__ to open the photo in the MS Photo editor
- __set date__ to set the DateTimeOriginal according to following rules 
    - '1971:04:01 00:00:01' first day of the month and one minute after midnight meaning that theses values are just set to order the photos
    - if you have several images in your selection, minutes will be incremented to avoid having two photos in the directory with the same name
    - if you have set another year, the photo will be moved in the right directory (including creation if useful)
    - after renaming, the images will be moved in the Word display as the right place according to your time
    - to order photos within a day, you can move using the Word cut/copy images then __set date__ on the right selection
- __set artist__ to update the meta tag of the same name
- __add a tag__ or __set all tags__ to update XPKeywords tag, even if I just have use it to clear previous value
- __GPS address__ to set the location tag which is a strong *XPKeywords avoider* and well adapted with photos galleries tools. Ok, this will not be as perfect as having a GPS in the camera, but at that time they were not yet invented ! GPS list is get on the fly from a text file you can edit in a separate window. GPS coordinates are the one in the Google map format.
    
![Test Image 4](https://github.com/langautier/scanned-photos/blob/master/setGPS.png)

# Install it
- First of all, you photos must be stored in directories as '1970' or '1970 summer holidays' and obviously in a test environment
- Create your working directory as shown
- Open the word file
![Installation](https://github.com/langautier/scanned-photos/blob/master/installation.png)

Knowing that :
- __caches__ will contain the ExifTool metadata output files to avoid doing too often these quite slow operation
- __originals__ to save files before the resize
- __journal__ of every actions done by the program

# Code sharing
Just for the pleasure to share some solutions I used in this program
## VBA resize a photo using ImageMagick
Convert is done using [ImageMagick](https://imagemagick.org/) targeting a photo properly printed on an A4 page which also means perfectly displayed on laptop or tablet.
```VBScript
xy = Me.imageSize
If Int(xy(0)) > Int(xy(1)) Then                         ' provide our target on the largest dimension
    size = IIf(xy(0) < 3508, xy(0), "3508") & "x"       ' avoid to enlarge image if already smaller
Else
    size = "x" & IIf(xy(1) < 3508, xy(1), "3508")       ' A4 as 300 dpi = 3508 x 2480 pixels/inches
End If
'   convert will create another file, so we restart Onedrive version historical from zero
'         this will allow later to easily free disk space just deleting _original
msg = img_.Convert(original, "+repage", "-resize", size, "-density", "300x300", "-units", "PixelsPerInch", destination)
```
## VBA Build a sorted list of strings in a table
As there is no simple solution to sort a table of strings in VBA, the best way is to insert new elements directly at the right place
```VBScript
Dim aFileSize() As String
Private Sub Class_Initialize()
    ReDim aFileSize(0) As String
End Sub
Property Let fileSize(fs As String)
Dim i As Integer, nbkey As Integer
If aFileSize(0) = "" Then
    aFileSize(0) = fs
Else
    If UBound(filter(aFileSize, fs, True)) = -1 Then
        nbkey = 1 + UBound(aFileSize)
        ReDim Preserve aFileSize(nbkey)

        For i = nbkey To 1 Step -1
            prev = aFileSize(i - 1)
            If LCase(fs) > LCase(prev) Then
                aFileSize(i) = fs
                Exit For
            Else
                aFileSize(i) = prev
            End If
        Next
        If i = 0 Then aFileSize(i) = fs
    End If
End If
End Property
```
## VBA Build a sorted list of objects in a collection
As there is no solution to sort items in a collection, the best way is to insert new elements directly at the right place
```VBScript
' create a sorted collection of the shapes of our document
Dim sh As Shape
Set aShapes = New Collection
For Each sh In ActiveDocument.Shapes
    If aShapes.count = 0 Then
        aShapes.add sh, sh.name
    Else
        If sh.name < aShapes(1).name Then
            aShapes.add sh, sh.name, aShapes(1).name
        Else
            For i = aShapes.count To 1 Step -1
                lastnam = aShapes(i).name
                If lastnam < sh.name Then Exit For
            Next i
            If i = aShapes.count Then
                aShapes.add sh, sh.name
            Else
                aShapes.add sh, sh.name, , lastnam
            End If
        End If
    End If
Next sh
```
## VBA Get an object by key in a collection
I imagine that the VBA designers did not want to write a potentially hateful performance routine. However, the possibility of finding an element from a key is often necessary and that sequential reading may not be too penalizing if the size of the list remains reasonable.
```VBScript
Property Get item(ByVal vID As Variant) As Shape
Dim sh As Shape

    Select Case VarType(vID)
        Case vbString
            For Each sh In aShapes
                If StrComp(sh.name, vID) = 0 Then
                    Set item = sh
                    Exit For
                End If
            Next
        Case vbLong, vbInteger, vbByte, vbDecimal
            Set item = aShapes.item(vID)
    End Select
End Property
```
# ExifTool Build a directory files metadata list in an XML file
```VBScript
out = ExifCache_filename(name)
'                                               -m (-ignoreMinorErrors)
'                                               -X (-xmlFormat)
'                                               -f (-forcePrint) Force printing of tags even if their values are not found
'                                               -s print tag names instead of descriptions
'                                               -L use Latin encoding for windows accentued characters é è... in keywords
cmd = exifToolExe & " -m -X -s -f -L -charset filename=latin " & _
        "-directory -filename -ExifIFD:DateTimeOriginal -ExifIFD:CreateDate -XPKeywords -Artist -IFD0:Model -IFD0:Make -ProfileCreator -About -FileSize -Location -GPSposition -ImageSize -JFIF:resolutionunit -JFIF:XResolution -JFIF:YResolution -IFD0:Orientation -ext JPG " & _
        """" & imagesFolderPath & "\" & name & """"

'   open and close the cmd command which is not very pleasant
'   s = Split(CreateObject("wscript.shell").Exec(cmd).StdOut.ReadAll, "<rdf:Description rdf:about=")
CreateObject("WScript.Shell").Run "cmd.exe /C " & cmd & " >""" & out & """", 0, True    ' bWaitOnReturn

' even with the waitonreturn it may happen the loadfromfile failed as coming to soon
Dim start As Single
start = Timer                   ' number of seconds elapsed since midnight
Do While Timer < start + 1
    DoEvents
Loop
```
