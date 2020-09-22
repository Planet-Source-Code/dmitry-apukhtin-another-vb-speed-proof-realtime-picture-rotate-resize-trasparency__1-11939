VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000018&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ROT"
   ClientHeight    =   3000
   ClientLeft      =   1515
   ClientTop       =   885
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   366
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      ClipControls    =   0   'False
      Height          =   1065
      Left            =   180
      ScaleHeight     =   1005
      ScaleWidth      =   2370
      TabIndex        =   0
      Top             =   150
      Width           =   2430
      Begin VB.OptionButton Option1 
         Caption         =   "Double"
         Height          =   195
         Index           =   3
         Left            =   105
         TabIndex        =   4
         Top             =   705
         Width           =   1050
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Normal"
         Height          =   195
         Index           =   2
         Left            =   105
         TabIndex        =   3
         Top             =   495
         Value           =   -1  'True
         Width           =   1080
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Half"
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   2
         Top             =   285
         Width           =   1080
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Bounce"
         Height          =   195
         Index           =   0
         Left            =   1410
         TabIndex        =   1
         Top             =   495
         Width           =   1080
      End
      Begin VB.Label lFPS 
         Caption         =   "xx FPS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1605
         TabIndex        =   6
         Top             =   45
         Width           =   690
      End
      Begin VB.Label Label1 
         Caption         =   """Rotator"" Size..."
         Height          =   195
         Left            =   30
         TabIndex        =   5
         Top             =   30
         Width           =   1380
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const PI As Double = 3.1415296

' rotate, scale and transparent blit. all-in-one.

' (c) Damian, 2000
'     dmitrya@thewercs.com
'
'  hope these stupid comments will help you to break thru :)
'  drop me a note about it, let me know if I should stop wasting
'  my time writing all this stuff :)

'  and sorry for poor English. I did my best ;-)

' note: if your first Q is "whatta hell is DIB?", you better do some
'       research of MSDN or whatever. sorry, this sample won't help
'       you to understand it. it just uses it.
'       generally speaking, DIB is "that thing that allows you to work
'       with bitmap's pixels directly, without PutPixel/GetPixel"

' a bunch of DIB's
Private dib As CDIB     ' working one - all is got rendered on it
Private dibP As CDIB    ' picture holder - we'll rotate it
Private dibB1 As CDIB   ' two
Private dibB2 As CDIB   '    background pics

Private A As Double ' current rotating angle (in radians)
Private X As Long   ' X-position of second backgrounder (which one moves back'n'forth)
Private d As Long   ' delta value (or direction, is prefer)
                    ' gets -1 or 1, indicating moving (and rotating) direction

Private dcX As Long, dcY As Long ' rotated picture center pos at working dib

Private tm As Long, fps As Long ' for FPS measuring

' all pos coordinates are in pixels

Private Sub Form_Load()
    ' get our DIBs inited
    Set dibB1 = New CDIB
    dibB1.Clone LoadPicture(App.Path & "\202.jpg")
    Set dibB2 = New CDIB
    dibB2.Clone LoadPicture(App.Path & "\203.jpg")
    Set dibP = New CDIB
    dibP.Clone LoadPicture(App.Path & "\1.gif")
    
    ' resize form to backgrounders size
    ' working dib - will have same size
    Width = dibB1.Width * 15: Height = dibB1.Height * 15
    Set dib = New CDIB
    dib.Create ScaleWidth, ScaleHeight
    
    ' center point on form (as well as on working DIB)
    dcX = ScaleWidth - 150: dcY = ScaleHeight \ 2

    ' starting direction is positive
    d = 1
End Sub

Private Sub Form_Paint()
    A = A - PI / 180 * 2 * d ' change angle by 2 grads at a time
    X = X + 2 * d            ' ... and position by 2 pixels
    If X < 0 Or X > dib.Width Then d = -d ' change direction on at edge points
    
    ' scrolling backgrounders part
    ' blit them into working dib, erasing previous frame as well
    BitBlt dib.hDC, X, 0, dib.Width - X, dib.Height, dibB1.hDC, 0, 0, vbSrcCopy
    BitBlt dib.hDC, 0, 0, X, dib.Height, dibB2.hDC, 0, 0, vbSrcCopy
    
    ' shift X-position in current direction
    dcX = dcX - d
    
    ' see what ResiZe factor is active
    Dim rz As Long, Y As Long
    Select Case True
        Case Option1(0) ' boucing
            rz = Sin(A) * 180 ' makes value running from -200 to 200
        Case Option1(1)
            rz = 200    ' half size (result of 100/200=1/2)
        Case Option1(3)
            rz = 50     ' double size (100/50=2)
        Case Else
            rz = 100    ' normal (yes, 100/100=1 :)
    End Select
    
    ' let's make it flying by sinus graph
    Y = dcY + Sin(A) * (dcY - dibP.Height)
    
    ' ok, here we go to main job. nice sub' name, isnt it? ;-)
    doTransparentRotateAndResize dibP, dib, A, dcX, Y, rz
    
    ' finally, blit what we have to form's face
    dib.PaintTo Me.hDC, 0, 0

    fps = fps + 1
    If timeGetTime > tm Then
        lFPS = fps & " FPS": lFPS.Refresh
        fps = 0
        tm = timeGetTime + 1000
    End If

    ' nice trick to eliminate timer'ed or loop'ed calls to this sub
    ' windows will send WM_PAINT to us as soon as it will be able
    ' and we'll have this proc called again and again
    ' still having form's controls working as usual.
    ' nice'n' easy and it looks like we have additional tread in VB
    Dim rc As RECT
    rc.Right = 1: rc.Bottom = 1
    InvalidateRect Me.hWnd, rc, False
End Sub

' ok, here's general rotation formula
'     x' = x * Cos(a) - y * Sin(a)
'     y' = x * Sin(a) + y * Cos(a)
' where (x,y) is original point, (x',y') is rotated one and "a" is angle
'
' so, to rotate a bitmap, we will "map" (x,y) coordinates in dest bitmap
' to corresponding (x',y') in source. in other words, we will move color
' bytes (RGB values) from source(x',y') to dest(x,y).
'
' we can't do much to optimize formula, but we can optimize
' our rotation procedure. so we will calculate points mappings
' for only one quadrant of source-bitmap QUAD and use OX/OY axes mirror
' reflections of calculated point to move RGB bytes
'
' second step in optimisation is to work with linear (one-dimensioned)
' array of bytes representing bitmaps. knowing how many bytes are in
' one line of bitmap, we can get linear addres of (X,Y) point in bitmap
' as Y*(num_of_bytes_per_line)+X*3 (hope it's clear why *3 - R,G and B bytes)

Sub doTransparentRotateAndResize(s As CDIB, d As CDIB, _
                                 Angle As Double, _
                                 ByVal dcX As Long, ByVal dcY As Long, _
                                 ByVal rz As Long)
' s - source,
' d - destination DIBs
' Angle - guess what?
' dcX,dCY - point at dest DIB where result picture's center should be
' rz - resize factor. take it as percent value of 100/rz,
'      ie 50 is for double-size and 200 for half-size

    ' source DIB's center AS WELL as quadrant (1/4th of whole pic)
    ' side length
    Dim scX As Long, scY As Long
    scX = (s.Width - 1) \ 2: scY = (s.Height - 1) \ 2
    
    ' ok, we work only on QUAD bitmaps so let's get rid of bigger dimension
    If scX < scY Then scY = scX Else scX = scY
    
    ' stretch quadrant to calculate since rotated quad surrounding
    ' area will grow in size. max growth is at 45 degree (PI/4)
    ' - remember hypotenuse stuff? hehe, do your math better ;)
    scY = Sqr(scX * scX * 2)
    
    ' pre-calc these sin/cos to save computing time in loop
    Dim aSin As Double, aCos As Double
    aSin = Sin(Angle): aCos = Cos(Angle)

    ' ok, let's get our linear memory spaces for both DIB's
    ' we will map it to byte-arrays for...
    Dim sB() As Byte, sP As Long, sRB As Long ' ...source DIB
    Dim dB() As Byte, dP As Long, dRB As Long ' ...dest DIB, using...
    ' ...very useful function of my DIB-Helper class. MapArray fools VB
    ' making him think that his array (which is not bounded, in fact)
    ' is mapped to particular space in memory (DIB bits in our case).
    ' it returns byte-width of one line of pixels in DIB
    sRB = s.MapArray(sB): dRB = d.MapArray(dB)
    ' so, after this call count that sB(0),sB(1) and sB(2) are
    ' B,G and R components of first pixel in source DIB.
    ' don't forget that it's in fact LAST pixel because DIBs are upside-down
    
    ' get linear address of CENTERs in our DIBs
    sP = scX * sRB + scX * 3 ' sP stands for "source Pointer"
    dP = dcY * dRB + dcX * 3 ' -"- "dest" one. not really matching names
                             ' but I like shorties ;)

    ' transparency part - suppose first upper-left pixel color is transparent one
    ' for speed, use only one color component (blue)
    Dim TransB As Byte
    TransB = sB(UBound(sB) - 1)
    
    Dim X As Long, Y As Long
    Dim xx As Long, yy As Long
    ' ok, lets do real business
    ' loop all pixels in ONE quadrant of ENLARGED (see scY comment above)
    ' area of source dib. it's upper-right quadrant, counting that DIB bits are upside-down
    For Y = 0 To scY - 1
        For X = 0 To scY - 1
            ' here we get that "mapped" coordinates
            ' and apply tricky resize-part. yes, one extra MUL
            ' and we have resize functionality - so easy...
            xx = (X * aCos - Y * aSin) * rz / 100
            yy = (X * aSin + Y * aCos) * rz / 100

            ' it could fall off original DIB because we scan
            ' enlarged area, so be aware
            If Abs(xx) <= scX And Abs(yy) <= scX Then
                Dim i As Long, j As Long
                ' same way to get linear address in DIB bytes memspace
                ' this time it will be OFFSETS from CENTER pointers...
                j = Y * dRB + X * 3   ' ... in dest bitmap
                i = yy * sRB + xx * 3 ' ... in source
                
                ' is pixel is not transparent one,
                ' copy three bytes (R,G and B values) from source(xx,yx) to dest(X,Y)
                ' in calculated quadrant (upper-right)...
                If sB(sP + i) <> TransB Then dB(dP + j) = sB(sP + i): dB(dP + j + 1) = sB(sP + i + 1): dB(dP + j + 2) = sB(sP + i + 2)
                ' ... and it's opposite one (down-left), "mirrored" by OX and OY axes
                If sB(sP - i) <> TransB Then dB(dP - j) = sB(sP - i): dB(dP - j + 1) = sB(sP - i + 1): dB(dP - j + 2) = sB(sP - i + 2):
                
                ' now, we have to recalc linear offsets for
                ' mirrored (x,y) in upper-left quadrant...
                j = X * dRB - Y * 3
                i = xx * sRB - yy * 3
                ' copy 3 bytes in it...
                If sB(sP + i) <> TransB Then dB(dP + j) = sB(sP + i): dB(dP + j + 1) = sB(sP + i + 1): dB(dP + j + 2) = sB(sP + i + 2)
                ' ... and mirror it once more (down-right quadrant)
                If sB(sP - i) <> TransB Then dB(dP - j) = sB(sP - i): dB(dP - j + 1) = sB(sP - i + 1): dB(dP - j + 2) = sB(sP - i + 2):
            End If
        Next
    Next
    ' uff, seems like we're done
    ' have to "un-fool" VB back so it won't lost its mind and GPF :-)
    d.UnMapArray dB: s.UnMapArray sB
End Sub

Private Sub Form_Activate()
    Picture1.Refresh
End Sub


