Attribute VB_Name = "Module1"
Option Explicit

Private Const LF_FACESIZE = 32
Private Const SPI_GETNONCLIENTMETRICS = 41

Private Type LOGFONT
   lfHeight As Long
   lfWidth As Long
   lfEscapement As Long
   lfOrientation As Long
   lfWeight As Long
   lfItalic As Byte
   lfUnderline As Byte
   lfStrikeOut As Byte
   lfCharSet As Byte
   lfOutPrecision As Byte
   lfClipPrecision As Byte
   lfQuality As Byte
   lfPitchAndFamily As Byte
   lfFaceName As String * LF_FACESIZE
End Type

Private Type NONCLIENTMETRICS
   cbSize As Long
   iBorderWidth As Long
   iScrollWidth As Long
   iScrollHeight As Long
   iCaptionWidth As Long
   iCaptionHeight As Long
   lfCaptionFont As LOGFONT
   iSMCaptionWidth As Long
   iSMCaptionHeight As Long
   lfSMCaptionFont As LOGFONT
   iMenuWidth As Long
   iMenuHeight As Long
   lfMenuFont As LOGFONT
   lfStatusFont As LOGFONT
   lfMessageFont As LOGFONT
End Type

Private Declare Function SystemParametersInfo Lib "USER32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long

Global pL As Long
Global pT As Long

Public Function centerObject(frmParent As Form, frmChild As Form, ByVal objContainer As PictureBox, Optional Fill As Boolean = True)
    Dim nPrntMode As Long
    Dim nChldMode As Long
    Dim nContMode As Long
    Dim factor As Long
    Dim Lfactor As Long
    Dim Tfactor As Long
    Dim nc As NONCLIENTMETRICS
    
    'get the scale mode for all objects we need
    nPrntMode = frmParent.ScaleMode
    nChldMode = frmChild.ScaleMode
    nContMode = objContainer.ScaleMode
    
    'ensure that we are using twips for all scale modes
    frmParent.ScaleMode = vbTwips
    frmChild.ScaleMode = vbTwips
    objContainer.ScaleMode = vbTwips
    
    If Fill Then
        'set the child form to the size of the container
        frmChild.Width = objContainer.ScaleWidth
        frmChild.Height = objContainer.ScaleHeight
    End If
    
    'since non client metrics stores everything as pixels we will calculate it as pixels
    nc.cbSize = Len(nc)
    Call SystemParametersInfo(SPI_GETNONCLIENTMETRICS, 0, nc, 0)
    
    'set the factor based on picture box style
    If TypeOf objContainer Is PictureBox Then
        If objContainer.BorderStyle = vbFixedSingle Then
            If objContainer.Appearance = 1 Then
                'for picture box that has a 3D border - 2 pixel border
                factor = 2
            Else
                'for picture box that has a border but is flat - 1 pixel border
                factor = 1
            End If
        Else
            'for picture box that has no border
            factor = 0
        End If
    End If
    
    'adjust the factor for parent forms border style
    Select Case frmParent.BorderStyle
        Case vbFixedSingle, vbFixedDialog
            Lfactor = (factor + 1) * nc.iBorderWidth
            'if the title bar of the form is visible then we
            If Not frmParent.ControlBox And frmParent.Caption = vbNullString Then
                Tfactor = Lfactor
            Else
                Tfactor = (factor + 2) * nc.iBorderWidth + nc.iCaptionHeight
            End If
        Case vbSizable
            Lfactor = (factor + 2) * nc.iBorderWidth
            If Not frmParent.ControlBox And frmParent.Caption = vbNullString Then
                Tfactor = Lfactor
            Else
                Tfactor = (factor + 3) * nc.iBorderWidth + nc.iCaptionHeight
            End If
        Case vbFixedToolWindow
            Lfactor = (factor + 1) * nc.iBorderWidth
            If Not frmParent.ControlBox And frmParent.Caption = vbNullString Then
                Tfactor = Lfactor
            Else
                Tfactor = (factor + 2) * nc.iBorderWidth + nc.iSMCaptionHeight
            End If
        Case vbSizableToolWindow
            Lfactor = (factor + 2) * nc.iBorderWidth
            If Not frmParent.ControlBox And frmParent.Caption = vbNullString Then
                Tfactor = Lfactor
            Else
                Tfactor = (factor + 3) * nc.iBorderWidth + nc.iSMCaptionHeight
            End If
        Case vbBSNone
            Lfactor = factor - 2
            Tfactor = factor - 2
    End Select
    
    'convert the pixel factors to twips
    Lfactor = Lfactor * Screen.TwipsPerPixelX
    Tfactor = Tfactor * Screen.TwipsPerPixelY
    
    'Set the LEFT value for the Child form
    pL = frmParent.Left + objContainer.Left + Lfactor + objContainer.Width \ 2 - frmChild.Width \ 2
    'Set the TOP value for the Child form
    pT = frmParent.Top + objContainer.Top + Tfactor + objContainer.Height \ 2 - frmChild.Height \ 2
  
    'Move the Child form into the container on the main form.
    frmChild.Move pL, pT
    
    'set scale modes back to previous settings
    frmParent.ScaleMode = nPrntMode
    frmChild.ScaleMode = nChldMode
    objContainer.ScaleMode = nContMode
  
End Function
