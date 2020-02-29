# Issues with my collection of 30 years scanned photo
I have more than 3 000 photos scanned from negatives with two issues 
- A number of images were scanned with too high resolution which resulted in files being too large
- The images were classified in the scanning order very far from chronological classification

# Using Word to sort this out
The need is to display images with text next to them and to be able to easily move when sorting. With the constraint that it is simple and quick to do given my current skills and the single use.
So, Word and VBA will be the convenitent solution.

As a result, two additional tools are necessary :
- [ImageMagick](https://imagemagick.org/) for resizing images files
- [ExifTool](https://exiftool.org/) for reading, writing and editing meta information in images files


# Main functionalities
![Test Image 4](https://github.com/langautier/scanned-photos/blob/master/mainscreen.png)


![Test Image 4](https://github.com/langautier/scanned-photos/blob/master/setGPS.png)

## Convert
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
