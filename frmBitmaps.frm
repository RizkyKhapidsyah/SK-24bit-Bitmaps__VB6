VERSION 5.00
Begin VB.Form frmBitmaps 
   AutoRedraw      =   -1  'True
   Caption         =   "Working with Bitmaps"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   ScaleHeight     =   302
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   340
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBright 
      Height          =   285
      Left            =   4200
      TabIndex        =   9
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtRipple 
      Height          =   285
      Left            =   4200
      TabIndex        =   8
      Top             =   240
      Width           =   495
   End
   Begin VB.CommandButton cmdRipple 
      Caption         =   "&Ripple"
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdBright 
      Caption         =   "&Brightness"
      Height          =   495
      Left            =   3120
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdInvert 
      Caption         =   "&Invert"
      Height          =   495
      Left            =   4200
      TabIndex        =   5
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmdGreen 
      Caption         =   "Gr&een it"
      Height          =   495
      Left            =   4200
      TabIndex        =   4
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdRed 
      Caption         =   "Re&d it"
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdBlue 
      Caption         =   "&Blue it"
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "Rest&ore"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdGrey 
      Caption         =   "Gra&y it"
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   3960
      Width           =   975
   End
End
Attribute VB_Name = "frmBitmaps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'Chapter 1
'Image processing with 24-bit bitmaps
'

Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
'Constants for the GenerateDC function
'**LoadImage Constants**
Const IMAGE_BITMAP As Long = 0
Const LR_LOADFROMFILE As Long = &H10
Const LR_CREATEDIBSECTION As Long = &H2000
Const LR_DEFAULTCOLOR As Long = &H0
Const LR_COLOR As Long = &H2
'****************************************
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Private Type BITMAP
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type

Dim BitmapImage As Long 'Bitmap DC
Dim bm As BITMAP        'bitmap structure
Dim hbm As Long         'Bitmap handle

Dim OriginalBits() As Byte

Dim BitmapWidth As Long
Dim BitmapHeight As Long

'IN: FileName: The file name of the graphics
'    BitmapHandle: The receiver of the loaded bitmap handle
'OUT: The Generated DC
Public Function GenerateDC(FileName As String, ByRef BitmapHandle As Long) As Long
Dim DC As Long
Dim hBitmap As Long

'Create a Device Context, compatible with the screen
DC = CreateCompatibleDC(0)

If DC < 1 Then
    GenerateDC = 0
    'Raise error
    Err.Raise vbObjectError + 1
    Exit Function
End If

'Load the image....BIG NOTE: This function is not supported under NT, there you can not
'specify the LR_LOADFROMFILE flag
hBitmap = LoadImage(0, FileName, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE Or LR_CREATEDIBSECTION)

If hBitmap = 0 Then 'Failure in loading bitmap
    DeleteDC DC
    GenerateDC = 0
    'Raise error
    Err.Raise vbObjectError + 2
    Exit Function
End If

'Throw the Bitmap into the Device Context
SelectObject DC, hBitmap

'Return the device context and handle
BitmapHandle = hBitmap
GenerateDC = DC


End Function
'Deletes a generated DC
Private Function DeleteGeneratedDC(DC As Long) As Long

If DC > 0 Then
    DeleteGeneratedDC = DeleteDC(DC)
Else
    DeleteGeneratedDC = 0
End If

End Function

Private Sub cmdBlue_Click()
Dim BitmapWidthBytes As Long
Dim ByteArray() As Byte
Dim I As Long, J As Long

ReDim ByteArray(1 To bm.bmWidthBytes, 1 To bm.bmHeight)

For I = 1 To bm.bmWidthBytes Step 3
    For J = 1 To bm.bmHeight
        
        ByteArray(I, J) = 0
        ByteArray(I + 1, J) = 0
        ByteArray(I + 2, J) = OriginalBits(I + 2, J)
        
    Next J
Next I

SetBitmapBits hbm, bm.bmWidthBytes * bm.bmHeight, ByteArray(1, 1)

BitBlt Me.hdc, 0, 0, BitmapWidth, BitmapHeight, BitmapImage, 0, 0, vbSrcCopy
Me.Refresh

End Sub

Private Sub cmdBright_Click()
Dim BitmapWidthBytes As Long
Dim ByteArray() As Byte
Dim I As Long, J As Long
Dim TempValue As Long
'Brightness Table
Dim BrightTable(255) As Byte

'Build brightness lookup table
For I = 0 To 255
    TempValue = I * Val(txtBright.Text)
    
    If TempValue > 255 Then
        BrightTable(I) = 255
    Else
        BrightTable(I) = TempValue
    End If
Next I


ReDim ByteArray(1 To bm.bmWidthBytes, 1 To bm.bmHeight)

For I = 1 To bm.bmWidthBytes Step 3
    For J = 1 To bm.bmHeight
        
        ByteArray(I, J) = BrightTable(OriginalBits(I, J))
        ByteArray(I + 1, J) = BrightTable(OriginalBits(I + 1, J))
        ByteArray(I + 2, J) = BrightTable(OriginalBits(I + 2, J))
        
    Next J
Next I

SetBitmapBits hbm, bm.bmWidthBytes * bm.bmHeight, ByteArray(1, 1)

BitBlt Me.hdc, 0, 0, BitmapWidth, BitmapHeight, BitmapImage, 0, 0, vbSrcCopy
Me.Refresh
End Sub

Private Sub cmdGreen_Click()
Dim BitmapWidthBytes As Long
Dim ByteArray() As Byte
Dim I As Long, J As Long

ReDim ByteArray(1 To bm.bmWidthBytes, 1 To bm.bmHeight)

For I = 1 To bm.bmWidthBytes Step 3
    For J = 1 To bm.bmHeight
        
        ByteArray(I, J) = 0
        ByteArray(I + 1, J) = OriginalBits(I + 1, J)
        ByteArray(I + 2, J) = 0
        
    Next J
Next I

SetBitmapBits hbm, bm.bmWidthBytes * bm.bmHeight, ByteArray(1, 1)

BitBlt Me.hdc, 0, 0, BitmapWidth, BitmapHeight, BitmapImage, 0, 0, vbSrcCopy
Me.Refresh

End Sub

Private Sub cmdGrey_Click()
Dim BitmapWidthBytes As Long
Dim ByteArray() As Byte
Dim I As Long, J As Long
Dim TempColor As Long

ReDim ByteArray(1 To bm.bmWidthBytes, 1 To bm.bmHeight)

For I = 1 To bm.bmWidthBytes Step 3
    For J = 1 To bm.bmHeight
        
        TempColor = OriginalBits(I, J)
        TempColor = TempColor + OriginalBits(I + 1, J)
        TempColor = TempColor + OriginalBits(I + 2, J)
        TempColor = TempColor / 3
        
        ByteArray(I, J) = TempColor
        ByteArray(I + 1, J) = TempColor
        ByteArray(I + 2, J) = TempColor
            

    Next J
Next I

SetBitmapBits hbm, bm.bmWidthBytes * bm.bmHeight, ByteArray(1, 1)

BitBlt Me.hdc, 0, 0, BitmapWidth, BitmapHeight, BitmapImage, 0, 0, vbSrcCopy
Me.Refresh

End Sub

Private Sub cmdInvert_Click()
Dim BitmapWidthBytes As Long
Dim ByteArray() As Byte
Dim I As Long, J As Long

ReDim ByteArray(1 To bm.bmWidthBytes, 1 To bm.bmHeight)

For I = 1 To bm.bmWidthBytes Step 3
    For J = 1 To bm.bmHeight
        
        ByteArray(I, J) = 255 - OriginalBits(I, J)
        ByteArray(I + 1, J) = 255 - OriginalBits(I + 1, J)
        ByteArray(I + 2, J) = 255 - OriginalBits(I + 2, J)
        
    Next J
Next I

SetBitmapBits hbm, bm.bmWidthBytes * bm.bmHeight, ByteArray(1, 1)

BitBlt Me.hdc, 0, 0, BitmapWidth, BitmapHeight, BitmapImage, 0, 0, vbSrcCopy
Me.Refresh
End Sub

Private Sub cmdRed_Click()
Dim BitmapWidthBytes As Long
Dim ByteArray() As Byte
Dim I As Long, J As Long

ReDim ByteArray(1 To bm.bmWidthBytes, 1 To bm.bmHeight)

For I = 1 To bm.bmWidthBytes Step 3
    For J = 1 To bm.bmHeight
        
        ByteArray(I, J) = OriginalBits(I, J)
        ByteArray(I + 1, J) = 0
        ByteArray(I + 2, J) = 0
    Next J
Next I

SetBitmapBits hbm, bm.bmWidthBytes * bm.bmHeight, ByteArray(1, 1)

BitBlt Me.hdc, 0, 0, BitmapWidth, BitmapHeight, BitmapImage, 0, 0, vbSrcCopy
Me.Refresh
End Sub

Private Sub cmdRestore_Click()

SetBitmapBits hbm, bm.bmWidthBytes * bm.bmHeight, OriginalBits(1, 1)

BitBlt Me.hdc, 0, 0, BitmapWidth, BitmapHeight, BitmapImage, 0, 0, vbSrcCopy
Me.Refresh
End Sub

Private Sub cmdRipple_Click()

Dim ByteArray() As Byte
Dim I As Long, J As Long
Dim TempValue As Long
Dim RippleTable() As Byte

'Dimension the ripple lookup table
ReDim RippleTable(1 To BitmapWidth)

'Build ripple table
For I = 1 To BitmapWidth
    TempValue = I + Sin(I / 5) * Val(txtRipple.Text)
    If TempValue > BitmapWidth Then
        RippleTable(I) = BitmapWidth
    ElseIf TempValue < 1 Then
        RippleTable(I) = 1
    Else
        RippleTable(I) = TempValue
    End If
    
Next I

ReDim ByteArray(1 To bm.bmWidthBytes, 1 To bm.bmHeight)

For I = 1 To bm.bmWidthBytes Step 3
    For J = 1 To bm.bmHeight
        
        ByteArray(I, J) = OriginalBits(I, RippleTable(J))
        ByteArray(I + 1, J) = OriginalBits(I + 1, RippleTable(J))
        ByteArray(I + 2, J) = OriginalBits(I + 2, RippleTable(J))
    Next J
Next I

SetBitmapBits hbm, bm.bmWidthBytes * bm.bmHeight, ByteArray(1, 1)

BitBlt Me.hdc, 0, 0, BitmapWidth, BitmapHeight, BitmapImage, 0, 0, vbSrcCopy
Me.Refresh

End Sub

Private Sub Form_Load()

'Load the image
BitmapImage = GenerateDC(App.Path & "\bitmap.bmp", hbm)

'Get the bitmap structure
GetObjectAPI hbm, Len(bm), bm

'Preinitialize the byte array
ReDim OriginalBits(1 To bm.bmWidthBytes, 1 To bm.bmHeight)

BitmapWidth = bm.bmWidth
BitmapHeight = bm.bmHeight

'Get teh bits
GetBitmapBits hbm, bm.bmWidthBytes * bm.bmHeight, OriginalBits(1, 1)
   
'Draw the bitmap
BitBlt Me.hdc, 0, 0, BitmapWidth, BitmapWidth, BitmapImage, 0, 0, vbSrcCopy
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
DeleteGeneratedDC BitmapImage
DeleteObject hbm
End Sub
