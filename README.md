# scanned-photos
I have more than 3 000 photos scanned from negatives with two issues 
- a lot of them have been scanned with an inapropriate setting and 


So, I d

![Test Image 4](https://github.com/langautier/scanned-photos/blob/master/mainscreen.png)


![Test Image 4](https://github.com/langautier/scanned-photos/blob/master/setGPS.png)

## Convert
Convert is done using [ImageMagick](https://imagemagick.org/index.php)
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
