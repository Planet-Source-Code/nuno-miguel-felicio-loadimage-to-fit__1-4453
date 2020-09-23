<div align="center">

## LoadImage\(\) to Fit


</div>

### Description

It's a single function, without API's, that loads an image and puts it in a destination PictureBox. If the source image is bigger that the destination PictureBox, then it will resized to fit in (mantaining the ratio). In other words, the image loaded will nicely fit the destination, but will not be deformed.

If the source is smaller, then it will remain that size...

This function is an upgrade (in speed, error trapping and results) of Jason Monroe original post. Thanks Janson.
 
### More Info
 
FilePath$ -> the path of the file to be loaded

PicMain -> The destination picturebox of the image

ImgMain -> An image object created inside PicMain (this image will be the "container" of the final image

PicTemp -> A picture box used as temp during the process

You need to know how to call a function... :-)

Put the code as it is in a new module... and then call it...

Use the return code to know if the image was loaded...

PS

We Only See Well With Our Heart;

The Essential Things Are Invisible To Our Eyes!

Greetings from Portugal...

0 -> The image was loaded

## -> The file was not loaded - Returns the vbError code

NOP


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Nuno Miguel Felicio](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/nuno-miguel-felicio.md)
**Level**          |Unknown
**User Rating**    |2.0 (10 globes from 5 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/nuno-miguel-felicio-loadimage-to-fit__1-4453/archive/master.zip)

### API Declarations

NOP


### Source Code

```
Option Explicit
Public Function LoadImage(FilePath$, picTemp As PictureBox, picMain As PictureBox, imgMain As Image) As Integer
  Dim X As Long
  Dim xo As Long
  Dim Y As Long
  Dim yo As Long
'vars to save the user initial picture boxes and images settings
  Dim pMainSM As Integer
  Dim pTempSM As Integer
  Dim pMainAS As Boolean
  Dim pTempAS As Boolean
  Dim iMainST As Boolean
'saves the initial conditions of picture boxes and images, for future reposition
  pMainSM = picMain.ScaleMode
  pMainAS = picMain.AutoSize
  pTempSM = picTemp.ScaleMode
  pTempAS = picTemp.AutoSize
  iMainST = imgMain.Stretch
'set the necessary conditions to picture boxes and image
  picMain.ScaleMode = vbTwips
  picMain.AutoSize = False
  picTemp.ScaleMode = vbTwips
  picTemp.AutoSize = True
  imgMain.Stretch = True
  'while sizing, make destination image invisible
  imgMain.Visible = False
  On Error Resume Next
  picTemp.Picture = LoadPicture(FilePath)
  If Err Then 'the image was not loaded, so set the image to blank and exit sub
    imgMain.Picture = LoadPicture()
    LoadImage = Err 'return the error code
    Exit Function
  End If
  'obtain the loaded image size
  xo = picTemp.Width
  yo = picTemp.Height
  ' First shrink the image so the sides fit
  If xo > picMain.Width Then
    X = picMain.Width
    Y = yo - (xo - X)
  End If
  ' if the image is still too tall, shrink it some more
  yo = Y
  If Y > picMain.Height Then
    Y = picMain.Height
    X = X - (yo - Y)
  End If
  'Now we have the X and Y that have the best fit, so set the destination to that size
  imgMain.Width = X
  imgMain.Height = Y
  ' Center the image(imgmain) in the main picture box(picmain)
  imgMain.Top = (picMain.Height \ 2) - (imgMain.Height \ 2)
  imgMain.Left = (picMain.Width \ 2) - (imgMain.Width \ 2)
  ' Now copy the image from the start picbox(picstart) into the
  ' display image field (imgmain)
  imgMain.Picture = picTemp.Picture
  picTemp.Picture = LoadPicture() 'clar the temp picture, because it's not necessary
  imgMain.Visible = True 'make the destination visible
'restore the initial user settings
  picMain.ScaleMode = pMainSM
  picMain.AutoSize = pMainAS
  picTemp.ScaleMode = pTempSM
  picTemp.AutoSize = pTempAS
  imgMain.Stretch = iMainST
  LoadImage = 0 'and returns 0, the image was sucessfuly loaded
End Function
```

