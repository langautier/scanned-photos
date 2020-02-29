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
## List and select images in you collection
Display to see this menu which allows you to:
- select one or more photo directories
- filter the selection according to the tags associated with the photo, the type of device or the size of the photo
- and use command button as:
    - __reset cache__ to delete the Exiftool cached files and force a new reading of metadata in the images files
    - __run__ to display in Word the selecting photos
    - __convert__ to resize all images over the targeted size

![Test Image 4](https://github.com/langautier/scanned-photos/blob/master/mainscreen.png)

## Sort photos by dates and edit meta tags
Once in word, you can select one or more photos and use buttons in the banner :
- __display__ to open the photo in the MS Photo editor
- __set date__ to set the DateTimeOriginal according to following rules 
    - '1971:04:01 00:00:01' first day of the month and one minute after midnight meaning that theses values are just set to order the photos
    - if you have several images in your selection, minutes will be incremented to avoid having to photos in the directory with the same name
    - if you have set another year, the photo will be moved in the right directory (including creation if useful)
    - after renaming, the images will be moved in the display as the right place according to your time
    - to order photos within a day, you can move using the Word cut/copy images then __set date__ on the right selection
- __set artist__ to update the meta tag of the same name
- __add a tag__ or __set all tags__ to update XPKeywords tag, even if I just have use it to clear previous value
- __GPS address__ to set the location tag which is a strong *XPKeywords avoider* and well adapted with photos galleries tools.  Ok
    
![Test Image 4](https://github.com/langautier/scanned-photos/blob/master/setGPS.png)

# How to use it

# Document how it is written
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
