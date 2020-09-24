VERSION 5.00
Begin VB.UserControl HTMLLabel 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1635
   ScaleHeight     =   675
   ScaleWidth      =   1635
   Begin VB.HScrollBar hscScroll 
      Height          =   285
      Left            =   0
      SmallChange     =   10
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   390
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox picViewPort 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   0
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   47
      TabIndex        =   2
      Top             =   0
      Width           =   705
   End
   Begin VB.Timer tmrHyperlinkClick 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   210
      Top             =   90
   End
   Begin VB.VScrollBar vscScroll 
      Height          =   525
      Left            =   1290
      SmallChange     =   10
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picHTML 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   0
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape shpCorner 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   315
      Left            =   720
      Top             =   30
      Width           =   315
   End
End
Attribute VB_Name = "HTMLLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'
' UserControl HTMLLabel.
'
' Version 0.3.0.
'
' A static HTML rendering control.
'
' Copyright © 2001-2002 Woodbury Associates.
'
'

Option Explicit

'
' Windows API declarations.
'
Private Const SRCCOPY = &HCC0020
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
                                             ByVal X As Long, ByVal Y As Long, _
                                             ByVal nWidth As Long, ByVal nHeight As Long, _
                                             ByVal hSrcDC As Long, _
                                             ByVal xSrc As Long, ByVal ySrc As Long, _
                                             ByVal dwRop As Long) As Long

'
' Private constants.
'
Private Const mcstrVersion          As String = "0.3.0"
Private Const mcstrDefaultFontName  As String = "Arial"
Private Const mcsngDefaultFontSize  As Single = 10
Private Const mclngDefaultBackColor As Long = vbButtonFace
Private Const mcstrResIDHandCursor  As String = "HAND_CURSOR"
Private Const mcintMaxTableCols     As Integer = 16
Private Const mcintMaxNestingLevel  As Integer = 16
Private Const mcintDefaultPadding   As Long = 4
Private Const mcintDefaultSpacing   As Long = 4

'
' Private enumerations.
'

'
' enumHTMLElementType
'
' HTML tag tokens.
'
Private Enum enumHTMLElementType
    hetContent
    hetUnknown
    hetHEADon
    hetHEADoff
    hetTITLEon
    hetTITLEoff
    hetBODYon
    hetBODYoff
    hetCOMMENTon
    hetCOMMENToff
    hetSTRONGon
    hetSTRONGoff
    hetEMon
    hetEMoff
    hetUon
    hetUoff
    hetPon
    hetPoff
    hetDIVon
    hetDIVoff
    hetBR
    hetHR
    hetULon
    hetULoff
    hetOLon
    hetOLoff
    hetLI
    hetTABLEon
    hetTABLEoff
    hetTHEADon
    hetTHEADoff
    hetTBODYon
    hetTBODYoff
    hetTFOOTon
    hetTFOOToff
    hetTRon
    hetTRoff
    hetTDon
    hetTDoff
    hetFONTon
    hetFONToff
    hetAon
    hetAoff
    hetIMG
    hetBLOCKQUOTEon
    hetBLOCKQUOTEoff
    hetHeaderon
    hetHeaderoff
    hetBIGon
    hetBIGoff
    hetSMALLon
    hetSMALLoff
    hetCENTERon
    hetCenterOff
    hetSUBon                                    ' Not implemented.
    hetSUBoff                                   ' Not implemented.
    hetSUPon                                    ' Not implemented.
    hetSUPoff                                   ' Not implemented.
    hetFORMon                                   ' Always ignored.
    hetFORMoff                                  ' Always ignored.
    hetSCRIPTon                                 ' Always ignored.
    hetSCRIPToff                                ' Always ignored.
    hetSTYLEon                                  ' Always ignored.
    hetSTYLEoff                                 ' Always ignored.
End Enum
'
' enumAlignmentConstants
'
Private Enum enumContentAlignmentStyle
    casHorizontalLeft
    casHorizontalCentre
    casHorizontalRight
    casVerticalTop
    casVerticalCentre
    casVerticalBottom
End Enum

'
' Private types.
'

'
' tHTMLElement
'
' Represents a single HTML element.
'
Private Type tHTMLElement
    ' General properties.
    strHTML         As String
    blnIsTag        As Boolean
    hetType         As enumHTMLElementType
    strID           As String
    blnUnSpaced     As Boolean
    blnBGColorSet   As Boolean

    ' Text words.
    astrWords()     As String

    ' Font attributes.
    strFontName     As String
    sngFontSize     As Single
    lngFontColor    As Long
    
    ' Anchor attributes.
    strAhref        As String
    strTitle        As String
    lngTop          As Long
    lngLeft         As Long
    lngBottom       As Long
    lngRight        As Long

    lngIndent       As Long
    blnCentre       As Boolean
    blnRight        As Boolean

    ' Image attributes.
    strImgSrc       As String
    strImgAlt       As String
    lngImgWidth     As Long
    lngImgHeight    As Long
    intHSpace       As Integer
    intVSpace       As Integer

    ' List attributes.
    blnListNumbered As Boolean
    intListNumber   As Integer

    ' Table attributes.
    sngWidth            As Single
    lngWidth            As Long
    sngHeight           As Single
    intBorderWidth      As Integer
    intCellPadding      As Integer
    intCellSpacing      As Integer
    intColSpan          As Integer
    lngContentHeight    As Long
    lngRowHeight        As Long
    casVAlign           As enumContentAlignmentStyle
    lngBgColour         As Long

    ' Document hierarchy attributes.
    intChildElements    As Integer
    aintChildElements() As Integer
    intParentElement    As Integer
    intChildIndex       As Integer
    intElementIndex     As Integer
End Type
'
' tColumn
'
' A single table column.
'
Private Type tColumn
    lngLeft     As Long
    lngRight    As Long
End Type
'
' tTable
'
' A table.
'
Private Type tTable
    blnCentre                   As Boolean
    intBorderWidth              As Integer
    lngTableLeft                As Long
    lngTableTop                 As Long
    lngTableWidth               As Long
    lngTableHeight              As Long
    lngRowTop                   As Long
    lngRowHeight                As Long
    lngCellLeft                 As Long
    lngMarginRight              As Long
    intCol                      As Integer
    audtCol(mcintMaxTableCols)  As tColumn
    intCellPadding              As Integer
    intCellSpacing              As Integer
    intElement                  As Integer
End Type

'
' Public events.
'
Public Event HyperlinkClick(Href As String)
Public Event LoadImage(Source As String, Image As Picture)

'
' Private member variables.
'
Private mstrDefaultFontName     As String
Private msngDefaultFontSize     As Single
Private mstrHTML                As String
Private mintElements            As Integer
Private maudtElement()          As tHTMLElement
Private mastrTagAttrName()      As String
Private mastrTagAttrValue()     As String
Private mblnEnableScroll        As Boolean
Private mblnEnableAnchors       As Boolean
Private mblnEnableTooltips      As Boolean
Private mintAnchors             As Integer
Private maintAnchor()           As Integer
Private mstrAhref               As String
Private mlngTextColor           As Long
Private mlngLinkColor           As Long
Private mstrBackground          As String
Private mlngScrollWidth         As Long
Private mlngScrollHeight        As Long
Private mblnUnderlineLinks      As Boolean
Private mlngBgColorDefault      As Long
Private mlngBgColorDocument     As Long
Private mintBODYElementIndex    As Integer
Private mintDefaultPadding      As Integer
Private mintDefaultSpacing      As Integer

'
' Public properties.
'

'
' Version
'
Public Property Get Version() As String
    Version = mcstrVersion
End Property
'
' DefaultFontName
'
Public Property Get DefaultFontName() As String
    DefaultFontName = mstrDefaultFontName
End Property
Public Property Let DefaultFontName(strNewVal As String)
    mstrDefaultFontName = strNewVal
    If Len(mstrHTML) > 0 Then
        DocumentHTML = mstrHTML
    End If
End Property
'
' DefaultFontSize
'
Public Property Get DefaultFontSize() As Single
    DefaultFontSize = msngDefaultFontSize
End Property
Public Property Let DefaultFontSize(sglNewVal As Single)
    msngDefaultFontSize = sglNewVal
    If Len(mstrHTML) > 0 Then
        DocumentHTML = mstrHTML
    End If
End Property
'
' BackColor
'
Public Property Get BackColor() As Long
    BackColor = mlngBgColorDefault
End Property
Public Property Let BackColor(lngNewVal As Long)
    mlngBgColorDefault = lngNewVal
    shpCorner.BackColor = mlngBgColorDefault
    UserControl.BackColor = mlngBgColorDefault

    If mintElements = 0 Then
        mlngBgColorDocument = mlngBgColorDefault
        picHTML.BackColor = mlngBgColorDefault
        picViewPort_Paint
    ElseIf Not maudtElement(mintBODYElementIndex).blnBGColorSet Then
        mlngBgColorDocument = mlngBgColorDefault
        picHTML.BackColor = mlngBgColorDefault
        DocumentHTML = mstrHTML
    End If
End Property
'
' Appearance
'
Public Property Get Appearance() As Integer
    Appearance = UserControl.Appearance
End Property
Public Property Let Appearance(lngNewVal As Integer)
    UserControl.Appearance = lngNewVal
End Property
'
' BorderStyle
'
Public Property Get BorderStyle() As Integer
    BorderStyle = UserControl.BorderStyle
End Property
Public Property Let BorderStyle(lngNewVal As Integer)
    UserControl.BorderStyle = lngNewVal
End Property
'
' EnableTooltips
'
Public Property Get EnableTooltips() As Boolean
    EnableTooltips = mblnEnableTooltips
End Property
Public Property Let EnableTooltips(blnNewVal As Boolean)
    mblnEnableTooltips = blnNewVal
End Property
'
' DocumentHTML
'
Public Property Get DocumentHTML() As String
    DocumentHTML = mstrHTML
End Property
Public Property Let DocumentHTML(strNewVal As String)
    picViewPort.MousePointer = vbHourglass
    DoEvents

    ' Add <HTML> and <BODY> tags if not present.
    If InStr(UCase(strNewVal), "<BODY") = 0 Then
        strNewVal = "<BODY>" & strNewVal & "</BODY>"
    End If
    If InStr(UCase(strNewVal), "<HTML>") = 0 Then
        strNewVal = "<HTML>" & strNewVal & "</HTML>"
    End If

    ' Convert CR adn LF characters to spaces.
    mstrHTML = Replace(Replace(strNewVal, Chr(10), " "), Chr(13), " ")

    ' Reset the colour.
    mlngBgColorDocument = mlngBgColorDefault
    mlngTextColor = vbBlack
    mlngLinkColor = vbBlue
    mSetDefaultStyle

    ' Replace some common character entities with their character literals.
    mstrHTML = Replace(mstrHTML, "&lt;", "&#" & Format(Asc("<"), "000") & ";")
    mstrHTML = Replace(mstrHTML, "&gt;", "&#" & Format(Asc(">"), "000") & ";")
    mstrHTML = Replace(mstrHTML, "&quot;", "&#" & Format(Asc(""""), "000") & ";")
    mstrHTML = Replace(mstrHTML, "&copy;", "&#169;")
    mstrHTML = Replace(mstrHTML, "&deg;", "&#176;")
    mstrHTML = Replace(mstrHTML, "&amp;", "&#" & Format(Asc("&"), "000") & ";")
    mstrHTML = Replace(mstrHTML, "&middot;", "&#183;")
    mstrHTML = Replace(mstrHTML, "&agrave;", "&#224;")
    mstrHTML = Replace(mstrHTML, "&aacute;", "&#225;")
    mstrHTML = Replace(mstrHTML, "&egrave;", "&#232;")
    mstrHTML = Replace(mstrHTML, "&eacute;", "&#233;")
    mstrHTML = Replace(mstrHTML, "&euml;", "&#235;")

    ' Strip whitespace.
    mstrHTML = Replace(mstrHTML, vbTab, " ")
    mstrHTML = Replace(mstrHTML, vbCrLf, " ")
    mstrHTML = Replace(Replace(Replace(mstrHTML, "  ", " "), "  ", " "), "  ", " ")

    ' Split the HTML into its constituent elements.
    mElementSplit
    mintBODYElementIndex = 0

    ' Parse the elements.
    mstrBackground = ""
    mParseHTMLElements
    mFixUnpairedTags
    mBuildAnchorList
    mBuildHierarchy

    ' Refresh the display if we are already visible.
    If UserControl.Parent.Visible Then
        Refresh False
    End If

    picViewPort.MousePointer = vbDefault
    DoEvents
End Property
'
' EnableScroll
'
Public Property Get EnableScroll() As Boolean
    EnableScroll = mblnEnableScroll
End Property
Public Property Let EnableScroll(blnNewVal As Boolean)
    mblnEnableScroll = blnNewVal
End Property
'
' EnableAnchors
'
Public Property Get EnableAnchors() As Boolean
    EnableAnchors = mblnEnableAnchors
End Property
Public Property Let EnableAnchors(blnNewVal As Boolean)
    mblnEnableAnchors = blnNewVal
End Property
'
' DocumentTitle
'
Public Property Get DocumentTitle() As String
    Dim intElem As Integer

    DocumentTitle = "Unknown"

    ' Locate the <TITLE></TITLE> tag within our list of HTML elements.
    If mintElements > 0 Then
        For intElem = 0 To UBound(maudtElement) - 1
            If maudtElement(intElem).hetType = hetTITLEon Then
                DocumentTitle = mstrDecodeText(maudtElement(intElem + 1).strHTML)
                Exit For
            End If
        Next intElem
    End If
End Property
'
' UnderlineLinks
'
Public Property Get UnderlineLinks() As Boolean
    UnderlineLinks = mblnUnderlineLinks
End Property
Public Property Let UnderlineLinks(blnNewVal As Boolean)
    mblnUnderlineLinks = blnNewVal
End Property
'
' DefaultPadding
'
Public Property Get DefaultPadding() As Integer
    DefaultPadding = mintDefaultPadding
End Property
Public Property Let DefaultPadding(intNewVal As Integer)
    mintDefaultPadding = intNewVal
End Property
'
' DefaultSpacing
'
Public Property Get DefaultSpacing() As Integer
    DefaultSpacing = mintDefaultSpacing
End Property
Public Property Let DefaultSpacing(intNewVal As Integer)
    mintDefaultSpacing = intNewVal
End Property
'
' picViewPort_Paint()
'
' Repaint the viewing window.
'
Private Sub picViewPort_Paint()
    On Error GoTo ErrorHandler

    ' Copy the visible portion of the document into the viewport.
    BitBlt picViewPort.hDC, 0, 0, picViewPort.ScaleWidth, picViewPort.ScaleHeight, _
                        picHTML.hDC, hscScroll.Value, vscScroll.Value, SRCCOPY

ExitPoint:
    Exit Sub

ErrorHandler:
    Resume ExitPoint
End Sub
'
' tmrHyperlinkClick_Timer()
'
' Fire the "hyperlink clicked" event after a delay which allows the control to complete processing before the event is fired.
'
Private Sub tmrHyperlinkClick_Timer()
    Dim strMethod   As String
    Dim varArgs     As Variant

    tmrHyperlinkClick.Enabled = False

    If Len(mstrAhref) > 0 Then
        If Left(UCase(mstrAhref), 3) = "VB:" Then
            mParseVBURL Mid(mstrAhref, 4), strMethod, varArgs
            mCallByName strMethod, varArgs
        Else
            ' Inform the container that an external target has been requested.
            RaiseEvent HyperlinkClick(mstrAhref)
        End If
        mstrAhref = ""
    End If
End Sub

'
' Private methods.
'

'
' UserControl_Initialize()
'
' Perform default initialisation.
'
Private Sub UserControl_Initialize()
    mstrDefaultFontName = mcstrDefaultFontName
    msngDefaultFontSize = mcsngDefaultFontSize
    mlngBgColorDefault = mclngDefaultBackColor
    mlngBgColorDocument = mlngBgColorDefault
    UserControl.BackColor = mlngBgColorDefault
    mlngTextColor = vbBlack
    mlngLinkColor = vbBlue
    picViewPort.MouseIcon = LoadResPicture(mcstrResIDHandCursor, vbResCursor)
    mblnUnderlineLinks = True
    mintDefaultPadding = mcintDefaultPadding
    mintDefaultSpacing = mcintDefaultSpacing
End Sub
'
' UserControl_ReadProperties()
'
' Load the properties set at design time for this instance of the control.
'
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Appearance = PropBag.ReadProperty("Appearance", 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    mlngBgColorDefault = PropBag.ReadProperty("BackColor", vbButtonFace)
    mlngBgColorDocument = mlngBgColorDefault
    UserControl.BackColor = mlngBgColorDefault
    picHTML.BackColor = mlngBgColorDefault
    mblnEnableAnchors = PropBag.ReadProperty("EnableAnchors", False)
    mblnEnableScroll = PropBag.ReadProperty("EnableScroll", False)
    mblnEnableTooltips = PropBag.ReadProperty("EnableTooltips", True)
    mstrDefaultFontName = PropBag.ReadProperty("DefaultFontName", "MS Sans Serif")
    msngDefaultFontSize = PropBag.ReadProperty("DefaultFontSize", 10)
    mblnUnderlineLinks = PropBag.ReadProperty("UnderlineLinks", True)
    mintDefaultPadding = PropBag.ReadProperty("DefaultPadding", mcintDefaultPadding)
    mintDefaultSpacing = PropBag.ReadProperty("DefaultSpacing", mcintDefaultSpacing)
End Sub
'
' UserControl_WriteProperties()
'
' Store the properties set at design time for this instance of the control.
'
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Appearance", UserControl.Appearance
    PropBag.WriteProperty "BorderStyle", UserControl.BorderStyle
    PropBag.WriteProperty "BackColor", mlngBgColorDefault
    PropBag.WriteProperty "EnableAnchors", mblnEnableAnchors
    PropBag.WriteProperty "EnableScroll", mblnEnableScroll
    PropBag.WriteProperty "EnableTooltips", mblnEnableTooltips
    PropBag.WriteProperty "DefaultFontName", mstrDefaultFontName
    PropBag.WriteProperty "DefaultFontSize", msngDefaultFontSize
    PropBag.WriteProperty "UnderlineLinks", mblnUnderlineLinks
    PropBag.WriteProperty "DefaultPadding", mintDefaultPadding
    PropBag.WriteProperty "DefaultSpacing", mintDefaultSpacing
End Sub
'
' UserControl_Resize()
'
' Resize event handler.
'
Private Sub UserControl_Resize()
    If UserControl.Parent.WindowState <> vbMinimized And Height > 360 Then
        ' Position our controls.
        If mblnEnableScroll Then
            vscScroll.Left = Width - vscScroll.Width - IIf(UserControl.Appearance = 1, 60, 30)
            vscScroll.Height = Height - vscScroll.Top - IIf(UserControl.Appearance = 1, 60, 30) - hscScroll.Height
            hscScroll.Width = vscScroll.Left
            hscScroll.Top = vscScroll.Height
            picHTML.Width = vscScroll.Left - picHTML.Left
            picHTML.Height = vscScroll.Height
            picViewPort.Width = picHTML.Width
            picViewPort.Height = picHTML.Height
            vscScroll.Value = 0
            hscScroll.Value = 0
            shpCorner.Top = hscScroll.Top
            shpCorner.Left = vscScroll.Left
            shpCorner.Visible = True
            shpCorner.ZOrder
        Else
            picHTML.Width = Width - IIf(UserControl.Appearance = 1, 30, 0)
            picHTML.Height = Height - IIf(UserControl.Appearance = 1, 30, 0)
            picViewPort.Width = picHTML.Width
            picViewPort.Height = picHTML.Height
            shpCorner.Visible = False
        End If
    End If
End Sub
'
' picViewPort_MouseMove()
'
' Show the "hand" cursor if the mouse pointer moves across an anchor.
'
Private Sub picViewPort_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim blnHit      As Boolean
    Dim intAnchor   As Integer

    On Error Resume Next

    If mblnEnableAnchors Then
        ' Is the mouse pointer curretly over a hyperlink ?
        For intAnchor = 0 To mintAnchors - 1
            If Len(maudtElement(maintAnchor(intAnchor)).strAhref) > 0 Then
                If maudtElement(maintAnchor(intAnchor)).lngLeft - hscScroll.Value < X And _
                   maudtElement(maintAnchor(intAnchor)).lngRight - hscScroll.Value > X And _
                   maudtElement(maintAnchor(intAnchor)).lngBottom - vscScroll.Value > Y And _
                   maudtElement(maintAnchor(intAnchor)).lngTop - vscScroll.Value < Y Then
                    blnHit = True
                    Exit For
                End If
            End If
        Next intAnchor

        ' Set the cursor depending on whether or not the pointer is over a hyperlink.
        If blnHit Then
            picViewPort.MousePointer = vbCustom
            If mblnEnableTooltips Then
                If Len(maudtElement(maintAnchor(intAnchor)).strTitle) > 0 Then
                    picViewPort.ToolTipText = maudtElement(maintAnchor(intAnchor)).strTitle
                Else
                    picViewPort.ToolTipText = maudtElement(maintAnchor(intAnchor)).strAhref
                End If
            End If
        Else
            picViewPort.MousePointer = vbArrow 'vbDefault
            picViewPort.ToolTipText = ""
        End If
    End If
End Sub
'
' picViewPort_MouseUp()
'
' Fire the "hyperlink clicked" event if the mouse is clicked on an anchor.
'
Private Sub picViewPort_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim blnHit      As Boolean
    Dim intAnchor   As Integer
    Dim intTarget   As Integer

    On Error Resume Next

    If mblnEnableAnchors Then
        ' Is the mouse pointer curretly over a hyperlink ?
        For intAnchor = 0 To mintAnchors - 1
            If maudtElement(maintAnchor(intAnchor)).lngLeft - hscScroll.Value <= X And _
               maudtElement(maintAnchor(intAnchor)).lngRight - hscScroll.Value >= X And _
               maudtElement(maintAnchor(intAnchor)).lngBottom - vscScroll.Value >= Y And _
               maudtElement(maintAnchor(intAnchor)).lngTop - vscScroll.Value <= Y Then
                blnHit = (Len(maudtElement(maintAnchor(intAnchor)).strAhref) > 0)
                Exit For
            End If
        Next intAnchor

        If blnHit Then
            ' Scroll to the referenced anchor if the clicked hyperlink refers to an internal
            ' destination anchor.
            If mblnEnableScroll And Left(maudtElement(maintAnchor(intAnchor)).strAhref, 1) = "#" Then
                For intTarget = 0 To UBound(maudtElement)
                    If maudtElement(intTarget).strID = Mid(maudtElement(maintAnchor(intAnchor)).strAhref, 2) Then
                        If maudtElement(intTarget).lngTop <= vscScroll.Max Then
                            vscScroll.Value = maudtElement(intTarget).lngTop
                        Else
                            vscScroll.Value = vscScroll.Max
                        End If
                    End If
                Next intTarget
            Else
                ' Prepare to fire the HyperlinkClick event.
                mstrAhref = maudtElement(maintAnchor(intAnchor)).strAhref
                tmrHyperlinkClick.Enabled = True
            End If
        End If
    End If
End Sub
'
' picViewPort_KeyDown()
'
' Provide keyboard-only scrolling.
'
Private Sub picViewPort_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next

    If mblnEnableScroll Then
        Select Case KeyCode
            Case vbKeyUp
                If vscScroll.Value > vscScroll.Min Then
                    vscScroll.Value = vscScroll.Value - vscScroll.SmallChange
                End If
            Case vbKeyDown
                If vscScroll.Value < vscScroll.Max Then
                    vscScroll.Value = vscScroll.Value + vscScroll.SmallChange
                End If
            Case vbKeyPageUp
                If vscScroll.Value > vscScroll.Min Then
                    If vscScroll.Value - vscScroll.LargeChange >= vscScroll.Min Then
                        vscScroll.Value = vscScroll.Value - vscScroll.LargeChange
                    Else
                        vscScroll.Value = vscScroll.Min
                    End If
                End If
            Case vbKeyPageDown
                If vscScroll.Value < vscScroll.Max Then
                    If vscScroll.Value + vscScroll.LargeChange <= vscScroll.Max Then
                        vscScroll.Value = vscScroll.Value + vscScroll.LargeChange
                    Else
                        vscScroll.Value = vscScroll.Max
                    End If
                End If
            Case vbKeyHome
                If (Shift And vbCtrlMask) > 0 Then
                    vscScroll.Value = vscScroll.Min
                End If
            Case vbKeyEnd
                If (Shift And vbCtrlMask) > 0 Then
                    vscScroll.Value = vscScroll.Max
                End If
            Case Else
        End Select
    End If
End Sub
'
' vscScroll_Change()
'
' Update the display after a scrollbar change.
'
Private Sub vscScroll_Change()
    If UserControl.Ambient.UserMode Then
        If mintElements > 0 Then
            picViewPort_Paint
        End If
    End If
End Sub
'
' vscScroll_Scroll()
'
' Update the display during drag-and-drop scrolling.
'
Private Sub vscScroll_Scroll()
    If UserControl.Ambient.UserMode Then
        If mintElements > 0 Then
            picViewPort_Paint
        End If
    End If
End Sub
Private Sub hscScroll_Change()
    If UserControl.Ambient.UserMode Then
        If mintElements > 0 Then
            picViewPort_Paint
        End If
    End If
End Sub
Private Sub hscScroll_Scroll()
    If UserControl.Ambient.UserMode Then
        If mintElements > 0 Then
            picViewPort_Paint
        End If
    End If
End Sub
'
' Refresh()
'
' Refresh the display.
'
' PaintOnly :   When True, indicates that the entire document should be redrawn, otherwise only the current
'               viewable region should be drawn.
'
Public Sub Refresh(Optional PaintOnly As Boolean = True)
    If UserControl.Ambient.UserMode Then
        picViewPort.MousePointer = vbHourglass
        UserControl_Resize

        ' Refresh the display.
        If mintElements > 0 Then
            picHTML.AutoRedraw = False
            mRenderElements False
            If mlngScrollHeight > picViewPort.ScaleHeight Then
                picHTML.Height = mlngScrollHeight * Screen.TwipsPerPixelX
            End If
            If mlngScrollWidth > picViewPort.ScaleWidth Then
                picHTML.Width = mlngScrollWidth * Screen.TwipsPerPixelX
            End If
            picHTML.AutoRedraw = True
            mRenderElements True
        End If
    
        ' Re-initialise the scroll bars.
        If mblnEnableScroll Then
            If mintElements > 0 Then
                vscScroll.Max = mlngScrollHeight - picViewPort.ScaleHeight
                hscScroll.Max = mlngScrollWidth - picViewPort.ScaleWidth '- 10
            Else
                vscScroll.Max = 0
                hscScroll.Max = 0
            End If

            vscScroll.LargeChange = picViewPort.ScaleHeight
            hscScroll.LargeChange = picViewPort.ScaleWidth
        
            If vscScroll.Max > 0 Then
                vscScroll.Enabled = True
                vscScroll.Value = 0
                vscScroll.Enabled = True
            Else
                vscScroll.Max = 0
                vscScroll.Value = 0
                vscScroll.Enabled = False
            End If
            vscScroll.Visible = True

            If hscScroll.Max > 0 Then
                hscScroll.Enabled = True
                hscScroll.Value = 0
                hscScroll.Enabled = True
            Else
                hscScroll.Max = 0
                hscScroll.Value = 0
                hscScroll.Enabled = False
            End If
            hscScroll.Visible = True
        Else
            vscScroll.Visible = False
            hscScroll.Visible = False
        End If

        ' Refresh the display.
        picViewPort_Paint
        picViewPort.MousePointer = vbDefault
    End If
End Sub
'
' mElementSplit()
'
' Split the current HTML into its constituent HTML elements.
'
Private Sub mElementSplit()
    Dim blnUnSpaced As Boolean
    Dim intStart    As Integer
    Dim intEnd      As Integer

    On Error Resume Next
    mintElements = 0
    Erase maudtElement

    On Error GoTo ErrorHandler

    intStart = 1
    intEnd = 0

    While intEnd < Len(mstrHTML)
        ' Locate the start of the next tag.
        intStart = InStr(intStart, mstrHTML, "<")

        If intStart > 0 Then
            If Mid(mstrHTML, intStart, 4) = "<!--" Then
                ' Grab everything within the comment.
                intEnd = InStr(intStart, mstrHTML, "-->") + 2
            Else
                ' Extract the tag (if one is found).
                intEnd = InStr(intStart, mstrHTML, ">")
            End If

            If intEnd > 0 Then
                If Len(Trim(Mid(mstrHTML, intStart, intEnd - intStart + 1))) > 0 Then
                    ReDim Preserve maudtElement(mintElements)
                    maudtElement(mintElements).strHTML = Mid(mstrHTML, intStart, intEnd - intStart + 1)
                    maudtElement(mintElements).intElementIndex = mintElements
                    If intStart > 1 Then
                        blnUnSpaced = (Mid(mstrHTML, intStart - 1, 1) <> " ")
                    End If
                    mintElements = mintElements + 1
                    intEnd = intEnd + 1
                End If
                intStart = intEnd
            Else
                intStart = intStart + 1
            End If

            ' Extract the content which follows the tag (if there is any).
            intEnd = InStr(intStart, mstrHTML, "<")
            If intEnd > 0 And intEnd - intStart > 0 Then
                If Len(Trim(Mid(mstrHTML, intStart, intEnd - intStart))) Then
                    ReDim Preserve maudtElement(mintElements)
                    maudtElement(mintElements).strHTML = Mid(mstrHTML, intStart, intEnd - intStart)
                    maudtElement(mintElements).intElementIndex = mintElements
                    maudtElement(mintElements).blnUnSpaced = blnUnSpaced
                    mintElements = mintElements + 1
                End If
                intStart = intEnd
            ElseIf intEnd = 0 Then
                ' Pass 1 complete.
                intEnd = Len(mstrHTML)
            End If
        End If
    Wend

ExitPoint:
    Exit Sub

ErrorHandler:
    Resume ExitPoint
End Sub
'
' mblnIsTag()
'
' Return True if the specified text is an HTML tag.
'
Private Function mblnIsTag(strText As String) As Boolean
    mblnIsTag = (Left(strText, 1) = "<" And Right(strText, 1) = ">")
End Function
'
' mstrTagID()
'
' Extract and return the HTML tag identifier from the specified string.
'
Public Function mstrTagID(strTag As String) As String
    Dim intEnd      As Integer
    Dim strRetVal   As String

    intEnd = InStr(strTag, " ")

    If intEnd > 0 Then
        strRetVal = Mid(strTag, 2, intEnd - 1)
    Else
        strRetVal = Mid(Trim(strTag), 2, Len(Trim(strTag)) - 2)
    End If

    mstrTagID = UCase(Trim(strRetVal))
End Function
'
' mintExtractTagAttributes()
'
' Extract the attribute names and values from the tag contained in the specified string.
'
Public Function mintExtractTagAttributes(strTag As String) As Integer
    Dim intRetVal   As Integer
    Dim intStart    As Integer
    Dim intEnd      As Integer
    Dim strDelim    As String

    Erase mastrTagAttrName
    Erase mastrTagAttrValue
    intStart = InStr(strTag, " ")

    If intStart > 0 Then
    While InStr(intStart, strTag, "=") > 0
        ' Extract the next attribute name.
        intEnd = InStr(intStart + 1, strTag, "=")

        ReDim Preserve mastrTagAttrName(intRetVal)
        mastrTagAttrName(intRetVal) = Replace(Trim(UCase(Mid(strTag, intStart, intEnd - intStart))), vbTab, "")

        ' Ascertain the value delimiter ("'", """ or " ").
        strDelim = " "
        intStart = intEnd + 1
        While Mid(strTag, intStart, 1) = " "
            intStart = intStart + 1
        Wend
        If Mid(strTag, intStart, 1) = "'" Or Mid(strTag, intStart, 1) = """" Or Mid(strTag, intStart, 1) = " " Then
            strDelim = Mid(strTag, intStart, 1)
        End If

        ' Locate the end delimiter.
        If InStr(intStart + 1, strTag, strDelim) > 0 Then
            intEnd = InStr(intStart + 1, strTag, strDelim)
        Else
            intEnd = Len(strTag)
        End If

        ' Extract the attribute value.
        ReDim Preserve mastrTagAttrValue(intRetVal)
        mastrTagAttrValue(intRetVal) = Trim(Mid(strTag, intStart, intEnd - intStart))
        If Left(mastrTagAttrValue(intRetVal), 1) = strDelim Then
            mastrTagAttrValue(intRetVal) = Mid(mastrTagAttrValue(intRetVal), 2)
        End If

        intStart = intEnd + 1

        intRetVal = intRetVal + 1
    Wend
    End If

    mintExtractTagAttributes = intRetVal
End Function
'
' mSetDefaultStyle()
'
' Reset the PictureBox's style using the current defaults.
'
Private Sub mSetDefaultStyle()
    picHTML.Font.Name = mstrDefaultFontName
    picHTML.Font.Size = msngDefaultFontSize
    picHTML.ForeColor = mlngTextColor
    picHTML.Font.Bold = False
    picHTML.Font.Italic = False
    picHTML.Font.Underline = False
End Sub
'
' mstrDecodeText()
'
' Decode the specified HTML-encoded text.
'
Private Function mstrDecodeText(strText) As String
    Dim intPos      As Integer
    Dim intChar     As Integer
    Dim strRetVal   As String

    If InStr(strText, "&#") > 0 Then
        intPos = 1
        While intPos <= Len(strText)
            If Mid(strText, intPos, 2) = "&#" And InStr(intPos, strText, ";") > 0 Then
                ' Translate the character literal.
                intPos = intPos + 2
                intChar = 0
                While IsNumeric(Mid(strText, intPos, 1))
                    intChar = (intChar * 10) + CInt(Mid(strText, intPos, 1))
                    intPos = intPos + 1
                Wend
                If Len(CStr(intChar)) < 4 Then
                    strRetVal = strRetVal & Chr(intChar) 'ChrW(intChar)
                End If
                intPos = intPos + 1
            Else
                strRetVal = strRetVal & Mid(strText, intPos, 1)
                intPos = intPos + 1
            End If
        Wend
    Else
        strRetVal = strText
    End If

    mstrDecodeText = Replace(Replace(Replace(strRetVal, vbCrLf, " "), Chr(10), " "), vbTab, " ")
End Function
'
' mParseHTMLElement()
'
' Parse the HTML element contained in the specified tHTMLElement structure.
'
Private Sub mParseHTMLElement(udtElem As tHTMLElement)
    On Error GoTo ErrorHandler

    If mblnIsTag(udtElem.strHTML) Then
        ' Store the tag's token and attributes.
        udtElem.blnIsTag = True

        Select Case mstrTagID(udtElem.strHTML)
            Case "P"
                udtElem.hetType = hetPon
            Case "/P"
                udtElem.hetType = hetPoff
            Case "DIV"
                udtElem.hetType = hetDIVon
            Case "/DIV"
                udtElem.hetType = hetDIVoff
            Case "HEAD"
                udtElem.hetType = hetHEADon
            Case "/HEAD"
                udtElem.hetType = hetHEADoff
            Case "TITLE"
                udtElem.hetType = hetTITLEon
            Case "/TITLE"
                udtElem.hetType = hetTITLEoff
            Case "BODY", "NOFRAMES"
                udtElem.hetType = hetBODYon
                mintBODYElementIndex = udtElem.intElementIndex
            Case "/BODY", "/NOFRAMES"
                udtElem.hetType = hetBODYoff
            Case "!--"
                udtElem.hetType = hetCOMMENTon
            Case "--"
                udtElem.hetType = hetCOMMENToff
            Case "STRONG", "B"
                udtElem.hetType = hetSTRONGon
            Case "/STRONG", "/B"
                udtElem.hetType = hetSTRONGoff
            Case "EM", "I"
                udtElem.hetType = hetEMon
            Case "/EM", "/I"
                udtElem.hetType = hetEMoff
            Case "U"
                udtElem.hetType = hetUon
            Case "/U"
                udtElem.hetType = hetUoff
            Case "BR"
                udtElem.hetType = hetBR
            Case "HR"
                udtElem.hetType = hetHR
            Case "UL"
                udtElem.hetType = hetULon
            Case "/UL"
                udtElem.hetType = hetULoff
            Case "OL"
                udtElem.hetType = hetOLon
            Case "/OL"
                udtElem.hetType = hetOLoff
            Case "LI"
                udtElem.hetType = hetLI
            Case "BLOCKQUOTE"
                udtElem.hetType = hetBLOCKQUOTEon
            Case "/BLOCKQUOTE"
                udtElem.hetType = hetBLOCKQUOTEoff
            Case "TABLE"
                udtElem.hetType = hetTABLEon
                udtElem.intCellSpacing = 2
                udtElem.intCellPadding = 2
                udtElem.sngWidth = 1
            Case "/TABLE"
                udtElem.hetType = hetTABLEoff
            Case "THEAD"
                udtElem.hetType = hetTHEADon
            Case "/THEAD"
                udtElem.hetType = hetTHEADoff
            Case "TBODY"
                udtElem.hetType = hetTBODYon
            Case "/TBODY"
                udtElem.hetType = hetTBODYoff
            Case "TFOOT"
                udtElem.hetType = hetTFOOTon
            Case "/TFOOT"
                udtElem.hetType = hetTFOOToff
            Case "TR"
                udtElem.hetType = hetTRon
            Case "/TR"
                udtElem.hetType = hetTRoff
            Case "TD", "TH"
                udtElem.hetType = hetTDon
                udtElem.intColSpan = 1
                udtElem.sngWidth = 1
                udtElem.casVAlign = casVerticalTop
                udtElem.lngBgColour = BackColor
            Case "/TD", "/TH"
                udtElem.hetType = hetTDoff
            Case "FONT"
                udtElem.hetType = hetFONTon
                udtElem.strFontName = mstrDefaultFontName
                udtElem.lngFontColor = mlngTextColor
                udtElem.sngFontSize = msngDefaultFontSize
            Case "/FONT"
                udtElem.hetType = hetFONToff
            Case "H1", "H2", "H3", "H4", "H5", "H6"
                udtElem.hetType = hetHeaderon
                udtElem.sngFontSize = msngDefaultFontSize + (1.2 * (7 - CSng(Mid(mstrTagID(udtElem.strHTML), 2, 1))))
            Case "/H1", "/H2", "/H3", "/H4", "/H5", "/H6"
                udtElem.hetType = hetHeaderoff
            Case "BIG"
                udtElem.hetType = hetBIGon
                udtElem.sngFontSize = msngDefaultFontSize + 1
            Case "/BIG"
                udtElem.hetType = hetBIGoff
            Case "SMALL"
                udtElem.hetType = hetSMALLon
                udtElem.sngFontSize = msngDefaultFontSize - 1
            Case "/SMALL"
                udtElem.hetType = hetSMALLoff
            Case "SUP"
                udtElem.hetType = hetSUPon
            Case "/SUP"
                udtElem.hetType = hetSUPoff
            Case "SUB"
                udtElem.hetType = hetSUBon
            Case "/SUB"
                udtElem.hetType = hetSUBoff
            Case "A"
                udtElem.hetType = hetAon
                udtElem.lngTop = -1
                udtElem.lngLeft = -1
                udtElem.lngBottom = -1
                udtElem.lngRight = -1
            Case "/A"
                udtElem.hetType = hetAoff
            Case "IMG"
                udtElem.hetType = hetIMG
                udtElem.intHSpace = 2
                udtElem.intVSpace = 2
            Case "CENTER"
                udtElem.hetType = hetCENTERon
            Case "/CENTER"
                udtElem.hetType = hetCenterOff
            Case "FORM"
                udtElem.hetType = hetFORMon
            Case "/FORM"
                udtElem.hetType = hetFORMoff
            Case "SCRIPT"
                udtElem.hetType = hetSCRIPTon
            Case "/SCRIPT"
                udtElem.hetType = hetSCRIPToff
            Case "STYLE"
                udtElem.hetType = hetSTYLEon
            Case "/STYLE"
                udtElem.hetType = hetSTYLEoff
            Case Else
                udtElem.hetType = hetUnknown
        End Select

        ' Set the tag's attributes.
        mSetElementAttributes udtElem
    Else
        udtElem.hetType = hetContent
        ' Split the text content into individual words.
        If InStr(mstrDecodeText(udtElem.strHTML), " ") > 0 Then
            udtElem.astrWords = Split(mstrDecodeText(udtElem.strHTML), " ")
        Else
            ReDim udtElem.astrWords(0)
            udtElem.astrWords(0) = mstrDecodeText(udtElem.strHTML)
        End If
    End If

ExitPoint:
    Exit Sub

ErrorHandler:
    Debug.Print "Error (" & Err.Number & ") " & Err.Description
    Resume ExitPoint
End Sub
'
' mParseHTMLElements()
'
' Parse the entire set of HTML elements.
'
Private Sub mParseHTMLElements()
    Dim intElem As Integer

    For intElem = 0 To mintElements - 1
        ' Parse the element.
        mParseHTMLElement maudtElement(intElem)
    Next intElem
End Sub
'
' mBuildAnchorList()
'
' Build a list of indexes which poin to the anchors in the document.
'
Private Sub mBuildAnchorList()
    Dim intElem As Integer

    Erase maintAnchor
    mintAnchors = 0

    For intElem = 0 To mintElements - 1
        ' Add any anchors to the anchors array.
        If mblnEnableAnchors And maudtElement(intElem).hetType = hetAon Then
            ReDim Preserve maintAnchor(mintAnchors)
            maintAnchor(mintAnchors) = intElem
            mintAnchors = mintAnchors + 1
        End If
    Next intElem
End Sub
'
' mSetElementAttributes()
'
' Set the attributes for the specified HTML element from its source tag.
'
Private Sub mSetElementAttributes(udtElem As tHTMLElement)
    Dim intAttr     As Integer
    Dim strValue    As String

    If mintExtractTagAttributes(udtElem.strHTML) > 0 Then
        For intAttr = 0 To UBound(mastrTagAttrName)
            Select Case mastrTagAttrName(intAttr)
                Case "ALIGN"
                    Select Case UCase(mastrTagAttrValue(intAttr))
                        Case "CENTER"
                            udtElem.blnCentre = True
                        Case "RIGHT"
                            udtElem.blnRight = True
                        Case Else
                    End Select
                Case "ALT"
                    udtElem.strImgAlt = mstrDecodeText(mastrTagAttrValue(intAttr))
                Case "BACKGROUND"
                    mstrBackground = mastrTagAttrValue(intAttr)
                Case "BGCOLOR"
                    strValue = Replace(mastrTagAttrValue(intAttr), "#", "")
                    If IsNumeric("&H" & strValue) Then
                        If udtElem.hetType = hetBODYon Then
                            mlngBgColorDocument = RGB(CLng("&H" & Left(strValue, 2)), _
                                                      CLng("&H" & Mid(strValue, 3, 2)), _
                                                      CLng("&H" & Right(strValue, 2)))
                        Else
                            udtElem.lngBgColour = RGB(CLng("&H" & Left(strValue, 2)), _
                                              CLng("&H" & Mid(strValue, 3, 2)), _
                                              CLng("&H" & Right(strValue, 2)))
                            udtElem.blnBGColorSet = True
                        End If
                    Else
                        If udtElem.hetType = hetBODYon Then
                            mlngBgColorDocument = mlngTranslateHTMLColour(mastrTagAttrValue(intAttr))
                        Else
                            udtElem.lngBgColour = mlngTranslateHTMLColour(mastrTagAttrValue(intAttr))
                            udtElem.blnBGColorSet = True
                        End If
                    End If
                Case "BORDER"
                    udtElem.intBorderWidth = mastrTagAttrValue(intAttr)
                Case "CELLPADDING"
                    udtElem.intCellPadding = mastrTagAttrValue(intAttr)
                Case "CELLSPACING"
                    udtElem.intCellSpacing = mastrTagAttrValue(intAttr)
                Case "COLOR"
                    strValue = Replace(mastrTagAttrValue(intAttr), "#", "")
                    If IsNumeric("&H" & strValue) Then
                        udtElem.lngFontColor = RGB(CLng("&H" & Left(strValue, 2)), _
                                               CLng("&H" & Mid(strValue, 3, 2)), _
                                               CLng("&H" & Right(strValue, 2)))
                    Else
                        udtElem.lngFontColor = mlngTranslateHTMLColour(mastrTagAttrValue(intAttr))
                    End If
                Case "COLSPAN"
                    udtElem.intColSpan = mastrTagAttrValue(intAttr)
                Case "FACE"
                    If InStr(mastrTagAttrValue(intAttr), ",") > 1 Then
                        udtElem.strFontName = Left(mastrTagAttrValue(intAttr), InStr(mastrTagAttrValue(intAttr), ",") - 1)
                    Else
                        udtElem.strFontName = mastrTagAttrValue(intAttr)
                    End If
                Case "HEIGHT"
                    If udtElem.hetType = hetIMG Then
                        udtElem.lngImgHeight = mastrTagAttrValue(intAttr)
                    Else
                        If InStr(mastrTagAttrValue(intAttr), "%") > 0 Then
                            udtElem.sngHeight = Replace(mastrTagAttrValue(intAttr), "%", "") / 100
                        Else
                            udtElem.sngHeight = mastrTagAttrValue(intAttr)
                        End If
                    End If
                Case "HREF"
                    udtElem.strAhref = mstrDecodeText(mastrTagAttrValue(intAttr))
                Case "HSPACE"
                    udtElem.intHSpace = mastrTagAttrValue(intAttr)
                Case "ID", "NAME"
                    udtElem.strID = mastrTagAttrValue(intAttr)
                Case "LINK"
                    strValue = Replace(mastrTagAttrValue(intAttr), "#", "")
                    If IsNumeric("&H" & strValue) Then
                        mlngLinkColor = RGB(CLng("&H" & Left(strValue, 2)), _
                                               CLng("&H" & Mid(strValue, 3, 2)), _
                                               CLng("&H" & Right(strValue, 2)))
                    Else
                        mlngLinkColor = mlngTranslateHTMLColour(mastrTagAttrValue(intAttr))
                    End If
                Case "SIZE"
                    If IsNumeric(mastrTagAttrValue(intAttr)) Then
                        If Left(mastrTagAttrValue(intAttr), 1) = "+" Or _
                           Left(mastrTagAttrValue(intAttr), 1) = "-" Then
                            udtElem.sngFontSize = msngDefaultFontSize + (1.2 * CSng(mastrTagAttrValue(intAttr)))
                        Else
                            udtElem.sngFontSize = msngDefaultFontSize + (1.2 * (CSng(mastrTagAttrValue(intAttr) - 3)))
                        End If
                    End If
                Case "SRC"
                    udtElem.strImgSrc = mstrDecodeText(mastrTagAttrValue(intAttr))
                Case "TEXT"
                    strValue = Replace(mastrTagAttrValue(intAttr), "#", "")
                    If IsNumeric("&H" & strValue) Then
                        mlngTextColor = RGB(CLng("&H" & Left(strValue, 2)), _
                                               CLng("&H" & Mid(strValue, 3, 2)), _
                                               CLng("&H" & Right(strValue, 2)))
                    Else
                        mlngTextColor = mlngTranslateHTMLColour(mastrTagAttrValue(intAttr))
                    End If
                Case "TITLE"
                    udtElem.strTitle = mastrTagAttrValue(intAttr)
                Case "VALIGN"
                    If UCase(mastrTagAttrValue(intAttr)) = "CENTER" Then
                        udtElem.casVAlign = casVerticalCentre
                    End If
                    If UCase(mastrTagAttrValue(intAttr)) = "BOTTOM" Then
                        udtElem.casVAlign = casVerticalBottom
                    End If
                Case "VSPACE"
                    udtElem.intVSpace = mastrTagAttrValue(intAttr)
                Case "WIDTH"
                    If udtElem.hetType = hetIMG Then
                        udtElem.lngImgWidth = mastrTagAttrValue(intAttr)
                    Else
                        If InStr(mastrTagAttrValue(intAttr), "%") > 0 Then
                            udtElem.sngWidth = Replace(mastrTagAttrValue(intAttr), "%", "") / 100
                        Else
                            udtElem.sngWidth = Replace(mastrTagAttrValue(intAttr), "px", "")
                        End If
                    End If
                Case Else
            End Select
        Next intAttr
    End If
End Sub
'
' mRenderElements()
'
' Render the entire set of current HTML elements into our PictureBox.
'
' blnFinalLayout    :   Flag indicating that this is the final call before the document is displayed.
'
' WA    19.11.2001                              Content rendering code removed to method mRenderElementContent().
'
Private Sub mRenderElements(blnFinalLayout As Boolean)
    Const clngListIndent        As Long = 20

    Dim blnCentre                           As Boolean
    Dim blnRight                            As Boolean
    Dim blnIgnore                           As Boolean
    Dim blnStartUnderline                   As Boolean
    Dim blnSpacerInserted                   As Boolean
    Dim blnInTable                          As Boolean
    Dim intElem                             As Integer
    Dim intWord                             As Integer
    Dim intNestingLevel                     As Integer
    Dim aintNumber(mcintMaxNestingLevel, 1) As Integer
    Dim intLinkElement                      As Integer
    Dim intTableNestLevel                   As Integer
    Dim lngX                                As Long
    Dim lngY                                As Long
    Dim lngIndent                           As Long
    Dim lngLastIndent                       As Long
    Dim lngScrollOffset                     As Long
    Dim lngLineHeight                       As Long
    Dim lngIndentStep                       As Long
    Dim lngXExtent                          As Long
    Dim lngMarginLeft                       As Long
    Dim lngMarginRight                      As Long
    Dim lngHeight                           As Long
    Dim audtTable(mcintMaxNestingLevel - 1) As tTable
    Dim sngLastFontSize                     As Single
    Dim strValue                            As String
    Dim sngCellWidth                        As Single
    Dim objImg                              As Picture
    Dim udtParent                           As tHTMLElement

    On Error GoTo ErrorHandler

    ' Initialise.
    picHTML.Cls
    picHTML.BackColor = mlngBgColorDocument
    If Not blnFinalLayout Then
        mlngScrollWidth = picViewPort.ScaleWidth
    End If
    mlngScrollHeight = picViewPort.ScaleHeight

    mSetDefaultStyle
    sngLastFontSize = msngDefaultFontSize
    lngLineHeight = picHTML.TextHeight("X") + mintDefaultSpacing
    lngIndentStep = picHTML.TextWidth("W") * 2

    lngMarginLeft = mintDefaultPadding
    lngMarginRight = picHTML.ScaleWidth
    lngX = lngMarginLeft
    lngY = mintDefaultPadding - lngScrollOffset
    lngIndent = 0
    lngLastIndent = 0

    If Len(mstrBackground) > 0 Then
        mRenderBackground lngScrollOffset
    End If

    ' Ignore everything up to the <BODY> tag.
    Do
        intElem = intElem + 1
        If intElem = mintElements Then
            Exit Do
        End If
    Loop While maudtElement(intElem).hetType <> hetBODYon

    ' Render the HTML elements.
    Do While intElem < mintElements
        maudtElement(intElem).lngTop = lngY
        maudtElement(intElem).lngIndent = 0
        maudtElement(intElem).lngLeft = lngX
        If maudtElement(intElem).hetType <> hetTDon Then
            maudtElement(intElem).blnCentre = blnCentre Or maudtElement(intElem).blnCentre
        End If

        If maudtElement(intElem).blnIsTag Then
            ' Update the prevailing mark-up style.
            Select Case maudtElement(intElem).hetType
                Case hetCENTERon
                    blnCentre = True
                Case hetCenterOff
                    blnCentre = False
                Case hetCOMMENTon
                    ' Ignore comments.
                Case hetFORMon
'                    blnIgnore = True
                Case hetSCRIPTon, hetSTYLEon
                    blnIgnore = True
                Case hetFORMoff, hetSCRIPToff, hetSTYLEoff
                    blnIgnore = False
                Case hetSTRONGon
                    picHTML.Font.Bold = True
                Case hetSTRONGoff
                    picHTML.Font.Bold = False
                Case hetEMon
                    picHTML.Font.Italic = True
                Case hetEMoff
                    picHTML.Font.Italic = False
                Case hetUon
                    picHTML.Font.Underline = True
                Case hetUoff
                    picHTML.Font.Underline = False
                Case hetPon, hetDIVon
                    lngX = lngMarginLeft + lngIndent
                    lngY = lngY + lngLineHeight
                    lngLastIndent = lngX - lngMarginLeft
                    maudtElement(intElem).lngIndent = lngX - lngMarginLeft
                    lngLineHeight = picHTML.TextHeight("X") + mintDefaultSpacing
                    blnRight = maudtElement(intElem).blnRight
                    blnCentre = maudtElement(intElem).blnCentre
                    If lngMarginRight = picHTML.ScaleWidth Then
                        lngMarginRight = picHTML.ScaleWidth - mintDefaultPadding
                    End If
                Case hetPoff, hetDIVoff
                    blnRight = False
                    blnCentre = False
                    If mblnGetTypedParent(maudtElement(intElem), udtParent, hetDIVon) Then
                        blnRight = udtParent.blnRight
                        blnCentre = udtParent.blnCentre
                    ElseIf mblnGetTypedParent(maudtElement(intElem), udtParent, hetTDon) Then
                        blnRight = udtParent.blnRight
                        blnCentre = udtParent.blnCentre
                    ElseIf mblnGetTypedParent(maudtElement(intElem), udtParent, hetCENTERon) Then
                        blnCentre = True
                    End If
                    lngX = lngMarginLeft + lngIndent
                    lngLastIndent = lngX - lngMarginLeft
                    maudtElement(intElem).lngIndent = lngX - lngMarginLeft
                    If Not blnSpacerInserted Then
                        lngY = lngY + lngLineHeight
                        lngLineHeight = picHTML.TextHeight("X") + mintDefaultSpacing
                        blnSpacerInserted = True
                    Else
                        blnSpacerInserted = False
                    End If
                    If lngMarginRight = picHTML.ScaleWidth - mintDefaultPadding Then
                        lngMarginRight = picHTML.ScaleWidth
                    End If
                Case hetBR
                    lngX = lngMarginLeft + lngIndent
                    lngY = lngY + lngLineHeight
                    lngLineHeight = picHTML.TextHeight("X") + mintDefaultSpacing
                    lngLastIndent = lngX - lngMarginLeft
                    maudtElement(intElem).lngIndent = lngX - lngMarginLeft
                    blnSpacerInserted = True 'False
                Case hetHR
                    lngX = lngMarginLeft
                    maudtElement(intElem).lngTop = lngY
                    picHTML.Line (lngMarginLeft, lngY)-(lngMarginRight - mintDefaultPadding, lngY)
                    lngY = lngY + mintDefaultSpacing + 1 'lngLineHeight
                Case hetULon
                    intNestingLevel = intNestingLevel + 1
                    aintNumber(intNestingLevel, 0) = False
                    lngIndent = lngIndent + lngIndentStep
                    lngLastIndent = lngIndent + lngIndentStep
                    maudtElement(intElem).lngIndent = lngLastIndent
                    If intNestingLevel = 1 And Not blnSpacerInserted Then
                        lngY = lngY + lngLineHeight
                        lngLineHeight = picHTML.TextHeight("X") + mintDefaultSpacing
                        blnSpacerInserted = True
                    Else
                        blnSpacerInserted = False
                    End If
                Case hetULoff
                    aintNumber(intNestingLevel, 0) = False
                    intNestingLevel = intNestingLevel - 1
                    lngIndent = IIf(lngIndent - lngIndentStep < 0, 0, lngIndent - lngIndentStep)
                    lngLastIndent = lngIndent
                    maudtElement(intElem).lngIndent = lngIndent
                    If intNestingLevel = 0 And Not blnSpacerInserted Then
                        lngY = lngY + lngLineHeight
                        lngLineHeight = picHTML.TextHeight("X") + mintDefaultSpacing
                        blnSpacerInserted = True
                    Else
                        blnSpacerInserted = False
                    End If
                Case hetOLon
                    intNestingLevel = intNestingLevel + 1
                    aintNumber(intNestingLevel, 0) = True
                    aintNumber(intNestingLevel, 1) = 0
                    lngIndent = lngIndent + lngIndentStep
                    lngLastIndent = lngIndent + lngIndentStep
                    maudtElement(intElem).lngIndent = lngLastIndent
                    If intNestingLevel = 1 And Not blnSpacerInserted Then
                        lngY = lngY + lngLineHeight
                        lngLineHeight = picHTML.TextHeight("X") + mintDefaultSpacing
                        blnSpacerInserted = True
                    Else
                        blnSpacerInserted = False
                    End If
                Case hetOLoff
                    aintNumber(intNestingLevel, 0) = False
                    intNestingLevel = intNestingLevel - 1
                    lngIndent = IIf(lngIndent - lngIndentStep < 0, 0, lngIndent - lngIndentStep)
                    lngLastIndent = lngIndent
                    maudtElement(intElem).lngIndent = lngIndent
                    If intNestingLevel = 0 And Not blnSpacerInserted Then
                        lngY = lngY + lngLineHeight
                        lngLineHeight = picHTML.TextHeight("X") + mintDefaultSpacing
                        blnSpacerInserted = True
                    Else
                        blnSpacerInserted = False
                    End If
                Case hetLI
                    lngX = lngMarginLeft + lngIndent
                    lngY = lngY + lngLineHeight
                    lngLineHeight = picHTML.TextHeight("X") + mintDefaultSpacing
                    maudtElement(intElem).lngTop = lngY
                    maudtElement(intElem).lngIndent = lngIndent
                    If aintNumber(intNestingLevel, 0) Then
                        aintNumber(intNestingLevel, 1) = aintNumber(intNestingLevel, 1) + 1
                        picHTML.CurrentX = lngX
                        picHTML.CurrentY = lngY
                        maudtElement(intElem).blnListNumbered = True
                        maudtElement(intElem).intListNumber = aintNumber(intNestingLevel, 1)
                    End If
                    picHTML.CurrentY = lngY
                    If maudtElement(intElem).blnListNumbered Then
                        ' Insert the list element's number.
                        picHTML.CurrentX = lngMarginLeft + maudtElement(intElem).lngIndent
                        picHTML.Print maudtElement(intElem).intListNumber & ". ";
                        lngX = lngX + picHTML.TextWidth("W" & ". ")
                    Else
                        ' Insert the list element's bullet.
                        picHTML.CurrentX = lngMarginLeft + maudtElement(intElem).lngIndent
                        picHTML.Print Chr(149) & "  ";
                        lngX = lngX + picHTML.TextWidth(Chr(149) & "  ")
                    End If
                    lngLastIndent = lngX
                Case hetTABLEon
                    If intTableNestLevel < 0 Then
                        intTableNestLevel = 0
                    End If

                    ' Move to a new line if we're not already inside a table.
                    If (Not blnInTable) Then
                        If lngY > mintDefaultPadding Then
                            lngY = lngY + lngLineHeight
                            lngLineHeight = picHTML.TextHeight("X") + mintDefaultSpacing
                            maudtElement(intElem).lngTop = lngY
                        ElseIf lngY = mintDefaultPadding And picHTML.CurrentX > lngMarginLeft Then
                            lngY = lngY + lngLineHeight
                            lngLineHeight = picHTML.TextHeight("X") + mintDefaultSpacing
                            maudtElement(intElem).lngTop = lngY
                        End If
                    End If
                    blnInTable = True

                    ' Calculate the table's width
                    If maudtElement(intElem).sngWidth <= 1 Then
                        audtTable(intTableNestLevel).lngTableWidth = _
                                (lngMarginRight - mintDefaultPadding * 2 - lngMarginLeft) * _
                                maudtElement(intElem).sngWidth
                    Else
                        audtTable(intTableNestLevel).lngTableWidth = maudtElement(intElem).sngWidth
                    End If

                    ' Layout the table.
                    maudtElement(intElem).lngWidth = audtTable(intTableNestLevel).lngTableWidth
                    mLayoutTable maudtElement(intElem)

                    ' Initialise the table.
                    audtTable(intTableNestLevel).lngTableTop = lngY
                    audtTable(intTableNestLevel).lngRowTop = lngY + _
                                                             maudtElement(intElem).intBorderWidth + _
                                                             maudtElement(intElem).intCellSpacing
                    audtTable(intTableNestLevel).lngRowHeight = 0
                    audtTable(intTableNestLevel).intBorderWidth = maudtElement(intElem).intBorderWidth
                    If Not blnFinalLayout Then
                        If maudtElement(intElem).sngHeight > 1 Then
                            audtTable(intTableNestLevel).lngTableHeight = maudtElement(intElem).sngHeight
                        Else
                            audtTable(intTableNestLevel).lngTableHeight = _
                                maudtElement(intElem).intBorderWidth * 2
                        End If
                    Else
                        audtTable(intTableNestLevel).lngTableHeight = _
                            maudtElement(intElem).lngContentHeight
                    End If
                    audtTable(intTableNestLevel).intCellSpacing = maudtElement(intElem).intCellSpacing
                    audtTable(intTableNestLevel).intCellPadding = maudtElement(intElem).intCellPadding
                    audtTable(intTableNestLevel).intElement = intElem

                    ' Set the table's left edge.
                    If maudtElement(intElem).blnCentre Then
                        audtTable(intTableNestLevel).lngTableLeft = _
                            ((lngMarginRight - lngMarginLeft) - _
                            audtTable(intTableNestLevel).lngTableWidth) \ 2 + lngMarginLeft
                        If audtTable(intTableNestLevel).lngTableLeft < lngMarginLeft Then
                            If intTableNestLevel = 0 Then
                                audtTable(intTableNestLevel).lngTableLeft = lngMarginLeft
                            Else
                                audtTable(intTableNestLevel).lngTableLeft = lngMarginLeft + audtTable(intTableNestLevel).intCellPadding
                            End If
                        End If
                    Else
                        If intTableNestLevel = 0 Then
                            audtTable(intTableNestLevel).lngTableLeft = lngMarginLeft
                        Else
                            audtTable(intTableNestLevel).lngTableLeft = lngMarginLeft + audtTable(intTableNestLevel).intCellPadding
                        End If
                    End If

                    If blnFinalLayout Then
                        If maudtElement(intElem).blnBGColorSet Then
                            ' Render the table's background.
                            picHTML.Line (audtTable(intTableNestLevel).lngTableLeft, _
                                          audtTable(intTableNestLevel).lngTableTop)- _
                                          (audtTable(intTableNestLevel).lngTableLeft + _
                                          audtTable(intTableNestLevel).lngTableWidth - 1, _
                                          audtTable(intTableNestLevel).lngTableTop + _
                                          audtTable(intTableNestLevel).lngTableHeight + _
                                          audtTable(intTableNestLevel).intCellSpacing - 1), _
                                          maudtElement(intElem).lngBgColour, BF

                        End If

                        ' Render the table's border.
                        If audtTable(intTableNestLevel).intBorderWidth > 0 Then
                            mRender3DBorder False, _
                                            audtTable(intTableNestLevel).lngTableLeft, _
                                            audtTable(intTableNestLevel).lngTableTop, _
                                            audtTable(intTableNestLevel).lngTableLeft + _
                                            audtTable(intTableNestLevel).lngTableWidth + 1, _
                                            audtTable(intTableNestLevel).lngTableTop + _
                                            audtTable(intTableNestLevel).lngTableHeight + _
                                            audtTable(intTableNestLevel).intCellSpacing - 1
                        End If
                    End If

                    ' Store the current centreing state.
                    audtTable(intTableNestLevel).blnCentre = maudtElement(intElem).blnCentre And blnCentre
                    blnCentre = False

                    ' Allow tables to be nested.
                    intTableNestLevel = intTableNestLevel + 1

                    If mblnEnableScroll And Not blnFinalLayout Then
                        If mlngScrollWidth < maudtElement(intElem).lngWidth + mintDefaultPadding Then
                            mlngScrollWidth = maudtElement(intElem).lngWidth + mintDefaultPadding
                        End If
                    End If
                Case hetTABLEoff
                    maudtElement(audtTable(intTableNestLevel - 1).intElement).lngBottom = _
                        audtTable(intTableNestLevel - 1).lngTableTop + _
                        audtTable(intTableNestLevel - 1).lngTableHeight - 1 + _
                    lngLineHeight + mintDefaultSpacing

                    If Not blnFinalLayout Then
                        If mblnGetParent(maudtElement(intElem), udtParent) Then
                            maudtElement(udtParent.intElementIndex).lngContentHeight = _
                                audtTable(intTableNestLevel - 1).lngTableHeight
                        End If
                    End If

                    ' Allow tables to be nested.
                    intTableNestLevel = intTableNestLevel - 1

                    ' Insert vertical spacing after the table.
                    lngY = lngY + lngLineHeight
                    lngLineHeight = picHTML.TextHeight("X") + mintDefaultSpacing

                    ' Reset the left and right margins.
                    If intTableNestLevel = 0 Then
                        lngMarginLeft = mintDefaultPadding
                        lngX = lngMarginLeft
                        lngMarginRight = picHTML.ScaleWidth
                        blnInTable = False
                    Else
                        lngMarginLeft = audtTable(intTableNestLevel - 1).lngTableLeft + _
                                        audtTable(intTableNestLevel - 1).lngCellLeft + _
                                        audtTable(intTableNestLevel - 1).intCellPadding
                        lngX = lngMarginLeft + lngIndent
                        ' Restore the previous centreing state.
                        blnCentre = audtTable(intTableNestLevel - 1).blnCentre
                    End If

                    ' Set the vertical rendering position below the table.
                    lngY = audtTable(intTableNestLevel).lngTableTop + _
                           audtTable(intTableNestLevel).lngTableHeight _
                           - audtTable(intTableNestLevel).intCellSpacing \ 2 '_
                           '- mintDefaultPadding

                    lngLineHeight = picHTML.TextHeight("X") + mintDefaultSpacing
                Case hetTRon
                    lngX = lngMarginLeft + lngIndent
                    maudtElement(intElem).lngIndent = lngX - lngMarginLeft

                    ' Set the row's height.
                    If maudtElement(intElem).sngHeight > 1 Then
                        audtTable(intTableNestLevel - 1).lngRowHeight = maudtElement(intElem).sngHeight
                    Else
                        audtTable(intTableNestLevel - 1).lngRowHeight = _
                                    audtTable(intTableNestLevel - 1).intCellSpacing / 1
                    End If

                    ' Set the first cell's left edge to the row's left edge.
                    audtTable(intTableNestLevel - 1).lngCellLeft = audtTable(intTableNestLevel - 1).intBorderWidth + _
                                                                    audtTable(intTableNestLevel - 1).intCellSpacing
                    audtTable(intTableNestLevel - 1).intCol = 0
                Case hetTRoff
                    If lngY <= audtTable(intTableNestLevel - 1).lngRowTop Then
                        lngY = lngY + lngLineHeight
                    End If
                    lngX = lngMarginLeft + lngIndent

                    ' Set the containing table's height.
                    If Not blnFinalLayout Then
                        If audtTable(intTableNestLevel - 1).lngTableHeight + _
                            audtTable(intTableNestLevel - 1).lngRowHeight + _
                            audtTable(intTableNestLevel - 1).intCellSpacing + _
                            audtTable(intTableNestLevel - 1).intBorderWidth > _
                            audtTable(intTableNestLevel - 1).lngTableHeight Then
                            audtTable(intTableNestLevel - 1).lngTableHeight = _
                                audtTable(intTableNestLevel - 1).lngTableHeight + _
                                audtTable(intTableNestLevel - 1).lngRowHeight + _
                                audtTable(intTableNestLevel - 1).intCellSpacing + _
                                audtTable(intTableNestLevel - 1).intBorderWidth
                        End If
                    End If

                    ' Draw borders around the cells in the row.
                    If audtTable(intTableNestLevel - 1).intBorderWidth > 0 Then
                        Dim idx As Integer
                        For idx = 0 To audtTable(intTableNestLevel - 1).intCol - 1
                            mRender3DBorder True, _
                                            audtTable(intTableNestLevel - 1).audtCol(idx).lngLeft, _
                                            audtTable(intTableNestLevel - 1).lngRowTop, _
                                            audtTable(intTableNestLevel - 1).audtCol(idx).lngRight, _
                                            audtTable(intTableNestLevel - 1).lngRowTop + _
                                            audtTable(intTableNestLevel - 1).lngRowHeight
                        Next idx
                    End If

                    If Not blnFinalLayout Then
                        If mblnGetParent(maudtElement(intElem), udtParent) Then
                            maudtElement(udtParent.intElementIndex).lngRowHeight = audtTable(intTableNestLevel - 1).lngRowHeight
                        End If
                    End If

                    ' Adjust the row's height.
                    audtTable(intTableNestLevel - 1).lngRowTop = _
                                audtTable(intTableNestLevel - 1).lngRowTop + _
                                audtTable(intTableNestLevel - 1).lngRowHeight + _
                                audtTable(intTableNestLevel - 1).intCellSpacing + _
                                audtTable(intTableNestLevel - 1).intBorderWidth '+ 1
                 Case hetTDon
                    sngCellWidth = maudtElement(intElem).lngWidth
                    blnCentre = maudtElement(intElem).blnCentre
                    blnRight = maudtElement(intElem).blnRight

                    ' Set the left and right margins to the cell's left and right edges.
                    lngMarginLeft = audtTable(intTableNestLevel - 1).lngTableLeft + _
                                    audtTable(intTableNestLevel - 1).lngCellLeft + _
                                    audtTable(intTableNestLevel - 1).intBorderWidth + _
                                    audtTable(intTableNestLevel - 1).intCellPadding
                    lngMarginRight = audtTable(intTableNestLevel - 1).lngTableLeft + _
                                    audtTable(intTableNestLevel - 1).lngCellLeft + _
                                    sngCellWidth - _
                                    audtTable(intTableNestLevel - 1).intBorderWidth * 2 - _
                                    audtTable(intTableNestLevel - 1).intCellPadding
                    audtTable(intTableNestLevel - 1).lngMarginRight = lngMarginRight
                    lngX = lngMarginLeft

                    ' Store the cell's left and right margins.
                    audtTable(intTableNestLevel - 1).audtCol(audtTable(intTableNestLevel - 1).intCol).lngLeft = _
                        audtTable(intTableNestLevel - 1).lngTableLeft + _
                        audtTable(intTableNestLevel - 1).lngCellLeft
                    audtTable(intTableNestLevel - 1).audtCol(audtTable(intTableNestLevel - 1).intCol).lngRight = _
                        lngMarginRight + _
                        audtTable(intTableNestLevel - 1).intCellPadding + _
                        audtTable(intTableNestLevel - 1).intBorderWidth

                    ' Stretch the containing table to fit the cell.
                    If audtTable(intTableNestLevel - 1). _
                        audtCol(audtTable(intTableNestLevel - 1).intCol).lngRight - _
                        audtTable(intTableNestLevel - 1).lngTableLeft > _
                        audtTable(intTableNestLevel - 1).lngTableWidth Then
                        audtTable(intTableNestLevel - 1).lngTableWidth = _
                        audtTable(intTableNestLevel - 1). _
                        audtCol(audtTable(intTableNestLevel - 1).intCol).lngRight - _
                        audtTable(intTableNestLevel - 1).lngTableLeft
                    End If
                    audtTable(intTableNestLevel - 1).intCol = audtTable(intTableNestLevel - 1).intCol + 1

                    If mblnEnableScroll And Not blnFinalLayout Then
                        If mlngScrollWidth < audtTable(intTableNestLevel - 1).lngTableWidth Then
                            mlngScrollWidth = audtTable(intTableNestLevel - 1).lngTableWidth
                        End If
                    End If

                    If Not blnFinalLayout Then
                        ' Set y to the containing row's top edge.
                        lngY = audtTable(intTableNestLevel - 1).lngRowTop + _
                                audtTable(intTableNestLevel - 1).intBorderWidth + _
                                audtTable(intTableNestLevel - 1).intCellPadding
                        ' Allow the cell's content height to be calculated on the first pass.
                        maudtElement(intElem).lngContentHeight = lngY '- lngLineHeight
                    Else
                        If mblnGetParent(maudtElement(intElem), udtParent) Then
                            lngHeight = udtParent.lngRowHeight - IIf(audtTable(intTableNestLevel - 1).intBorderWidth = 0, 1, 0) '- mintDefaultPadding / 1
                        Else
                            lngHeight = audtTable(intTableNestLevel - 1).lngRowHeight
                        End If
                        If maudtElement(intElem).blnBGColorSet Then
                            ' Render the cell's background.
                            picHTML.Line (audtTable(intTableNestLevel - 1).lngTableLeft + _
                                          audtTable(intTableNestLevel - 1).lngCellLeft + _
                                          audtTable(intTableNestLevel - 1).intBorderWidth, _
                                          audtTable(intTableNestLevel - 1).lngRowTop + _
                                          audtTable(intTableNestLevel - 1).intBorderWidth)- _
                                          (lngMarginRight + _
                                          audtTable(intTableNestLevel - 1).intCellPadding - _
                                          IIf(audtTable(intTableNestLevel - 1).intBorderWidth = 0, 1, 0), _
                                          audtTable(intTableNestLevel - 1).lngRowTop + lngHeight - _
                                          audtTable(intTableNestLevel - 1).intBorderWidth), _
                                          maudtElement(intElem).lngBgColour, BF
                        End If

                        ' Vertically align the cell's content.
                        If maudtElement(intElem).casVAlign = casVerticalTop Then
                            lngY = audtTable(intTableNestLevel - 1).lngRowTop + _
                                    audtTable(intTableNestLevel - 1).intCellSpacing / 2 + _
                                    (audtTable(intTableNestLevel - 1).intBorderWidth * 0) + _
                                    audtTable(intTableNestLevel - 1).intCellPadding
                        ElseIf maudtElement(intElem).casVAlign = casVerticalCentre Then
                            lngY = audtTable(intTableNestLevel - 1).lngRowTop + _
                                    (lngHeight - _
                                    maudtElement(intElem).lngContentHeight) \ 2
                        Else
                            lngY = audtTable(intTableNestLevel - 1).lngRowTop + _
                                    lngHeight - maudtElement(intElem).lngContentHeight - _
                                    audtTable(intTableNestLevel - 1).intCellSpacing / 2 - _
                                    audtTable(intTableNestLevel - 1).intCellPadding * 2
                        End If
                    End If
                Case hetTDoff
                    If Not blnFinalLayout Then
                        ' Store the cell's content height on the first pass.
                        If mblnGetParent(maudtElement(intElem), udtParent) Then
                            maudtElement(udtParent.intElementIndex).lngContentHeight = _
                                lngY - udtParent.lngContentHeight + lngLineHeight - mintDefaultSpacing
                        End If
                    End If

                    If lngLineHeight <> picHTML.TextHeight("X") + mintDefaultSpacing Then
                        If picHTML.CurrentY < lngY + lngLineHeight Then
                            picHTML.CurrentY = lngY + lngLineHeight
                        End If
                        lngLineHeight = picHTML.TextHeight("X") + mintDefaultSpacing
                    End If
                    If blnFinalLayout Then
                        If mblnGetParent(maudtElement(intElem), udtParent) Then
                            If mblnGetParent(udtParent, udtParent) Then
                                audtTable(intTableNestLevel - 1).lngRowHeight = maudtElement(udtParent.intElementIndex).lngRowHeight
                            End If
                        End If
                    Else
                        ' Set the containing row's height to the highest cell in the row.
                        If picHTML.CurrentY - _
                            audtTable(intTableNestLevel - 1).lngRowTop > _
                            audtTable(intTableNestLevel - 1).lngRowHeight Then
                            audtTable(intTableNestLevel - 1).lngRowHeight = _
                                picHTML.CurrentY - _
                                audtTable(intTableNestLevel - 1).lngRowTop
                        End If
                    End If

                    ' Set the next cell's left egde.
                    audtTable(intTableNestLevel - 1).lngCellLeft = _
                        audtTable(intTableNestLevel - 1).audtCol(audtTable(intTableNestLevel - 1).intCol - 1).lngRight + _
                        audtTable(intTableNestLevel - 1).intCellSpacing + audtTable(intTableNestLevel - 1).intBorderWidth - _
                        audtTable(intTableNestLevel - 1).lngTableLeft
                    blnCentre = False
                    blnRight = False
                    If mblnGetTypedParent(maudtElement(intElem), udtParent, hetDIVon) Then
                        blnRight = udtParent.blnRight
                        blnCentre = udtParent.blnCentre
                    ElseIf mblnGetParent(maudtElement(intElem), udtParent) Then
                        If mblnGetTypedParent(udtParent, udtParent, hetTDon) Then
                            blnRight = udtParent.blnRight
                            blnCentre = udtParent.blnCentre
                        ElseIf mblnGetTypedParent(maudtElement(intElem), udtParent, hetCENTERon) Then
                            blnCentre = True
                        End If
                    ElseIf mblnGetTypedParent(maudtElement(intElem), udtParent, hetCENTERon) Then
                        blnCentre = True
                    End If
                Case hetFONTon
                    On Error Resume Next
                    picHTML.FontName = maudtElement(intElem).strFontName
                    picHTML.FontSize = maudtElement(intElem).sngFontSize
                    picHTML.ForeColor = maudtElement(intElem).lngFontColor
                    sngLastFontSize = maudtElement(intElem).sngFontSize
                    On Error GoTo ErrorHandler
                Case hetFONToff
                    mSetDefaultStyle
                    sngLastFontSize = msngDefaultFontSize
                    lngLineHeight = picHTML.TextHeight("X") + mintDefaultSpacing
                    lngIndentStep = picHTML.TextWidth("W") * 2
                Case hetBLOCKQUOTEon
                    lngY = lngY + lngLineHeight
                    If Not blnSpacerInserted Then
                        lngY = lngY + lngLineHeight
                        lngLineHeight = picHTML.TextHeight("X") + mintDefaultSpacing
                    End If
                    lngX = lngMarginLeft + lngIndentStep
                    lngIndent = lngIndent + lngIndentStep
                    lngLastIndent = lngIndent
                    maudtElement(intElem).lngIndent = lngLastIndent
                Case hetBLOCKQUOTEoff
                    lngY = lngY + lngLineHeight
                    lngLineHeight = picHTML.TextHeight("X") + mintDefaultSpacing
                    lngX = lngX - lngIndentStep
                    lngIndent = lngIndent - lngIndentStep
                    lngLastIndent = lngIndent
                    maudtElement(intElem).lngIndent = lngLastIndent
                Case hetHeaderon
                    lngX = lngMarginLeft + lngIndent
                    picHTML.FontSize = maudtElement(intElem).sngFontSize
                    picHTML.Font.Bold = True
                    lngLineHeight = picHTML.TextHeight("X") + mintDefaultSpacing
                    lngIndentStep = picHTML.TextWidth("W") * 2
                    If (Not blnInTable) And picHTML.CurrentY > mintDefaultPadding Then
                        lngY = lngY + lngLineHeight  '+ mintDefaultPadding
                        lngLineHeight = picHTML.TextHeight("X") + mintDefaultSpacing
                    End If
                Case hetBIGon, hetSMALLon
                    picHTML.FontSize = maudtElement(intElem).sngFontSize
                Case hetHeaderoff, hetBIGoff, hetSMALLoff
                    lngX = lngMarginLeft + lngIndent
                    lngY = lngY + lngLineHeight  '- mintDefaultPadding
                    lngLineHeight = picHTML.TextHeight("X") + mintDefaultSpacing
                    picHTML.FontSize = sngLastFontSize
                    picHTML.Font.Bold = False
                    lngLineHeight = picHTML.TextHeight("X") + mintDefaultSpacing
                    lngIndentStep = picHTML.TextWidth("W") * 2
                Case hetAon
                    If mblnEnableAnchors Then
                        If Len(maudtElement(intElem).strAhref) > 0 Then
                            picHTML.ForeColor = mlngLinkColor
                            blnStartUnderline = True
                        End If
                        intLinkElement = intElem
                        lngXExtent = lngX
                    End If
                Case hetAoff
                    If mblnEnableAnchors Then
                        picHTML.Font.Underline = False
                        blnStartUnderline = False
                        picHTML.ForeColor = mlngTextColor
                        If intLinkElement > -1 Then
                            maudtElement(intLinkElement).lngBottom = lngY - mintDefaultSpacing + lngLineHeight
                            maudtElement(intLinkElement).lngRight = lngXExtent
                            intLinkElement = -1
                        End If
                    End If
                Case hetIMG
                    ' Load the referenced image.
                    RaiseEvent LoadImage(maudtElement(intElem).strImgSrc, objImg)

                    If (Not blnInTable) And lngY > mintDefaultPadding And Not blnSpacerInserted Then
                        lngY = lngY + lngLineHeight
                        maudtElement(intElem).lngTop = lngY
                        lngLineHeight = picHTML.TextHeight("X") + mintDefaultSpacing
                    End If

                    ' Store the image's size (in pixels) if no explicit size was given.
                    If Not (objImg Is Nothing) Then
                        If maudtElement(intElem).lngImgWidth = 0 Then
                            maudtElement(intElem).lngImgWidth = picHTML.ScaleX(objImg.Width, vbHimetric, vbPixels)
                        End If
                        If maudtElement(intElem).lngImgHeight = 0 Then
                            maudtElement(intElem).lngImgHeight = picHTML.ScaleY(objImg.Height, vbHimetric, vbPixels)
                        End If
                    Else
                        If maudtElement(intElem).lngImgWidth = 0 Then
                            maudtElement(intElem).lngImgWidth = 24 'lngLineHeight
                        End If
                        If maudtElement(intElem).lngImgHeight = 0 Then
                            maudtElement(intElem).lngImgHeight = 24 'lngLineHeight
                        End If
                        maudtElement(intElem).intBorderWidth = 1
                    End If

                    ' Centre the image, if necessary.
                    If blnCentre Or maudtElement(intElem).blnCentre Then
                        maudtElement(intElem).blnCentre = True
                        lngX = ((lngMarginRight - lngMarginLeft) - maudtElement(intElem).lngImgWidth) \ 2 + lngMarginLeft

                        If lngX < lngMarginLeft Then
                            lngX = lngMarginLeft
                        End If
                    End If

                    lngX = lngX + maudtElement(intElem).intHSpace

                    ' Render the image's border, if it has one.
                    If maudtElement(intElem).intBorderWidth > 0 Then
                        mRender3DBorder True, lngX, lngY + maudtElement(intElem).intVSpace, _
                                        lngX + maudtElement(intElem).lngImgWidth + _
                                        maudtElement(intElem).intBorderWidth, _
                                        lngY + maudtElement(intElem).intVSpace + _
                                        maudtElement(intElem).lngImgHeight + _
                                        maudtElement(intElem).intBorderWidth
                    End If
                    lngX = lngX + maudtElement(intElem).intBorderWidth
                    'lngY = lngY + maudtElement(intElem).intBorderWidth

                    If Not (objImg Is Nothing) Then
                        ' Render the image.
                        picHTML.PaintPicture objImg, lngX, _
                                             lngY + maudtElement(intElem).intVSpace, _
                                             maudtElement(intElem).lngImgWidth, _
                                             maudtElement(intElem).lngImgHeight
                        Set objImg = Nothing
                    End If

                    lngX = lngX + maudtElement(intElem).lngImgWidth + _
                           maudtElement(intElem).intBorderWidth + _
                           maudtElement(intElem).intHSpace
                    If lngLineHeight < maudtElement(intElem).lngImgHeight + _
                                    (maudtElement(intElem).intBorderWidth * 2) + _
                                    (maudtElement(intElem).intVSpace * 2) Then '+ mintDefaultPadding Then
                        lngLineHeight = maudtElement(intElem).lngImgHeight + _
                                        (maudtElement(intElem).intBorderWidth * 2) + _
                                        (maudtElement(intElem).intVSpace * 2)
                    End If
                    lngXExtent = lngX

                    If mblnEnableScroll And Not blnFinalLayout Then
                        If mlngScrollWidth < lngX + mintDefaultPadding Then
                            mlngScrollWidth = lngX + mintDefaultPadding
                        End If
                    End If
                Case hetCENTERon
                    blnCentre = True
                    blnRight = False
                    lngX = lngMarginLeft + lngIndent
                    maudtElement(intElem).blnCentre = True
                Case hetCenterOff
                    blnCentre = False
                    lngX = lngMarginLeft + lngIndent
                    maudtElement(intElem).blnCentre = False
                Case Else
            End Select

            If maudtElement(intElem).hetType <> hetCOMMENTon Then
                maudtElement(intElem).lngBottom = lngY + lngLineHeight + mintDefaultSpacing
                maudtElement(intElem).lngRight = lngXExtent
            End If

            intElem = intElem + 1
        ElseIf Not blnIgnore Then
            maudtElement(intElem).lngIndent = lngLastIndent

            If blnInTable Then
                mRenderElementContent intElem, lngX, lngY, blnRight, blnCentre, _
                                      lngMarginRight, lngMarginLeft + (audtTable(intTableNestLevel - 1).intCellSpacing \ 2), _
                                      lngXExtent, lngLineHeight, _
                                      intLinkElement, blnSpacerInserted, blnStartUnderline, lngLastIndent, _
                                      blnIgnore, sngLastFontSize, lngIndentStep, _
                                      audtTable(intTableNestLevel - 1).intCellSpacing \ 2
            Else
                mRenderElementContent intElem, lngX, lngY, blnRight, blnCentre, _
                                      lngMarginRight, lngMarginLeft, _
                                      lngXExtent, lngLineHeight, _
                                      intLinkElement, blnSpacerInserted, blnStartUnderline, lngLastIndent, _
                                      blnIgnore, sngLastFontSize, lngIndentStep, _
                                      mintDefaultSpacing
            End If
        Else
            intElem = intElem + 1
        End If

        If lngY > mlngScrollHeight - 10 Then
            mlngScrollHeight = lngY + lngLineHeight
        End If
    Loop

ExitPoint:
    Exit Sub

ErrorHandler:
    Resume ExitPoint
End Sub
'
' mRenderElementContent()
'
' Render the text content contained in the elements starting from intElem until a non-content
' element is encountered.
'
Private Sub mRenderElementContent(ByRef intElem As Integer, _
                                  ByRef lngX As Long, ByRef lngY As Long, _
                                  ByVal blnRight As Boolean, _
                                  ByVal blnCentre As Boolean, _
                                  ByVal lngMarginRight As Long, ByVal lngMarginLeft As Long, _
                                  ByRef lngXExtent As Long, _
                                  ByRef lngLineHeight As Long, _
                                  ByRef intLinkElement As Integer, _
                                  ByRef blnSpacerInserted As Boolean, _
                                  ByRef blnStartUnderline As Boolean, _
                                  ByRef lngLastIndent As Long, _
                                  ByRef blnIgnore As Boolean, _
                                  ByRef sngLastFontSize As Single, _
                                  ByRef lngIndentStep As Long, _
                                  ByVal lngPadding As Long)
    Dim blnContentExhausted As Boolean
    Dim blnSpaceInserted    As Boolean
    Dim blnUnderlineStart   As Boolean
    Dim intWord             As Integer
    Dim intElemLoop         As Integer
    Dim intIdxLine          As Integer
    Dim aintLineStartElem() As Integer
    Dim aintLineEndElem()   As Integer
    Dim aintLineStartWord() As Integer
    Dim aintLineEndWord()   As Integer
    Dim alngLineX()         As Long
    Dim lngIndent           As Long

    Dim strFontName     As String
    Dim sngFontSize     As Single
    Dim lngFontColor    As Long
    Dim blnBold         As Boolean
    Dim blnItalic       As Boolean
    Dim blnUnderline    As Boolean

    Dim astrLineText()      As String

    On Error GoTo ErrorHandler

    blnUnderlineStart = blnStartUnderline
    strFontName = picHTML.FontName
    sngFontSize = picHTML.FontSize
    lngFontColor = picHTML.ForeColor
    blnBold = picHTML.FontBold
    blnItalic = picHTML.FontItalic
    blnUnderline = picHTML.FontUnderline

    intElemLoop = intElem
    intWord = 0
    lngIndent = maudtElement(intElem).lngIndent

    ' Build a set of arrays for each line of content which hold the start and end elements and start and end words,
    ' amongst other things.
    While Not blnContentExhausted
        ' Initialise storage for the new line.
        ReDim Preserve aintLineStartElem(intIdxLine)
        ReDim Preserve aintLineEndElem(intIdxLine)
        ReDim Preserve aintLineStartWord(intIdxLine)
        ReDim Preserve aintLineEndWord(intIdxLine)
        ReDim Preserve alngLineX(intIdxLine)
        ReDim Preserve astrLineText(intIdxLine)

        aintLineStartElem(intIdxLine) = intElemLoop
        aintLineStartWord(intIdxLine) = intWord
        aintLineEndElem(intIdxLine) = intElemLoop
        aintLineEndWord(intIdxLine) = intWord
        astrLineText(intIdxLine) = ""

        ' Set the X position to the start of the line.
        If intIdxLine = 0 Then
            alngLineX(intIdxLine) = lngX
            alngLineX(intIdxLine) = alngLineX(intIdxLine)
        Else
            alngLineX(intIdxLine) = lngMarginLeft + lngIndent
        End If

        ' Iterate through the following elements' content until the end of the line is reached.
        Do While (alngLineX(intIdxLine) + _
            picHTML.TextWidth(IIf(Len(astrLineText(intIdxLine)) > 0 And Right(astrLineText(intIdxLine), 1) <> " " And _
            (Not maudtElement(intElemLoop).blnUnSpaced), " ", "") + _
            maudtElement(intElemLoop).astrWords(intWord)) + lngPadding < lngMarginRight) Or _
            Len(astrLineText(intIdxLine)) = 0

            aintLineEndElem(intIdxLine) = intElemLoop
            aintLineEndWord(intIdxLine) = intWord

            ' Add the current word to the overall line width.
            If Len(astrLineText(intIdxLine)) > 0 And Right(astrLineText(intIdxLine), 1) <> " " And _
                (Not maudtElement(intElemLoop).blnUnSpaced) Or intWord > 0 Then
                alngLineX(intIdxLine) = alngLineX(intIdxLine) + picHTML.TextWidth(" ")
                astrLineText(intIdxLine) = astrLineText(intIdxLine) & " "
            End If
            alngLineX(intIdxLine) = alngLineX(intIdxLine) + _
                                    picHTML.TextWidth(Replace(maudtElement(intElemLoop).astrWords(intWord), "&nbsp;", " "))
            astrLineText(intIdxLine) = astrLineText(intIdxLine) & maudtElement(intElemLoop).astrWords(intWord)

            If intWord = UBound(maudtElement(intElemLoop).astrWords) Then
                ' All of the words in the current element have been exhausted.
                intElemLoop = intElemLoop + 1
                intWord = 0

                If intElemLoop > UBound(maudtElement) Then
                    blnContentExhausted = True
                    Exit Do
                End If

                ' At this point we need to look ahead to the next element.
                ' If the next element constitutes a text style change, keep going, otherwise stop.
                While maudtElement(intElemLoop).blnIsTag
                    If mblnIsStyleTag(maudtElement(intElemLoop).hetType) Or maudtElement(intElemLoop).hetType = hetUnknown Then
                        mSetStyleFromTag intElemLoop, sngLastFontSize, lngLineHeight, lngIndentStep, blnStartUnderline
                        aintLineEndElem(intIdxLine) = intElemLoop
                        intElemLoop = intElemLoop + 1
                        If intElemLoop > UBound(maudtElement) Then
                            blnContentExhausted = True
                            Exit Do
                        End If
                    ElseIf maudtElement(intElemLoop).hetType = hetSCRIPTon Then
                        Do
                            intElemLoop = intElemLoop + 1
                            If maudtElement(intElemLoop).blnIsTag Then
                                If maudtElement(intElemLoop).hetType = hetSCRIPToff Then
                                    intElemLoop = intElemLoop + 1
                                    Exit Do
                                End If
                            End If
                        Loop
                    Else
                        blnContentExhausted = True
                        Exit Do
                    End If
                Wend
            Else
                ' Move on to the next word in the current element.
                intWord = intWord + 1
            End If
        Loop

        ' Remove any trailing spaces from the end of the line.
        If Right(astrLineText(intIdxLine), 1) = " " Then
            alngLineX(intIdxLine) = alngLineX(intIdxLine) - picHTML.TextWidth(" ")
        End If

        intIdxLine = intIdxLine + 1
    Wend

    ' Return to the calling code with intElem set to the next (non-content) element.
    intElem = intElemLoop

    ' Reset the current rendering style to the first element's style.
    blnStartUnderline = blnUnderlineStart
    picHTML.FontName = strFontName
    picHTML.FontSize = sngFontSize
    picHTML.ForeColor = lngFontColor
    picHTML.FontBold = blnBold
    picHTML.FontItalic = blnItalic
    picHTML.FontUnderline = blnUnderline

    ' Render the lines.
    For intIdxLine = 0 To UBound(aintLineStartElem)
        ' Align the current line.
        intElemLoop = aintLineStartElem(intIdxLine)

        If intIdxLine = 0 And lngX <> lngMarginLeft And lngX + lngPadding <> lngMarginLeft Then
            lngIndent = maudtElement(intElemLoop).lngIndent - lngPadding
        Else
            If blnRight Then
                lngX = lngMarginRight - (alngLineX(intIdxLine) - lngMarginLeft) - lngPadding
            ElseIf blnCentre Then
                lngX = (((lngMarginRight - lngMarginLeft) - (alngLineX(intIdxLine) - lngMarginLeft)) / 2) + lngMarginLeft
            Else
                lngX = lngMarginLeft + lngIndent
            End If
        End If

        ' Render the current line.
        Do While intElemLoop <= aintLineEndElem(intIdxLine)
            Do While maudtElement(intElemLoop).blnIsTag
                maudtElement(intElemLoop).lngLeft = lngX
                maudtElement(intElemLoop).lngTop = lngY
                maudtElement(intElemLoop).lngIndent = lngLastIndent
                maudtElement(intElemLoop).lngBottom = lngY + lngLineHeight + mintDefaultSpacing
                maudtElement(intElemLoop).lngRight = lngX
                If mblnIsStyleTag(maudtElement(intElemLoop).hetType) Then
                    ' Set the style for the next element.
                    mSetStyleFromTag intElemLoop, sngLastFontSize, lngLineHeight, lngIndentStep, blnStartUnderline

                    If mblnEnableAnchors Then
                        Select Case maudtElement(intElemLoop).hetType
                            Case hetAon
                                intLinkElement = intElemLoop
                                lngXExtent = lngX
                                maudtElement(intLinkElement).lngLeft = lngX
                            Case hetAoff
                                If intLinkElement > -1 Then
                                    maudtElement(intLinkElement).lngBottom = lngY + lngLineHeight + mintDefaultSpacing
                                    maudtElement(intLinkElement).lngRight = lngXExtent
                                    intLinkElement = -1
                                End If
                        End Select
                    End If

                    intElemLoop = intElemLoop + 1
                ElseIf maudtElement(intElemLoop).hetType = hetSCRIPTon Then
                    Do
                        intElemLoop = intElemLoop + 1
                        If maudtElement(intElemLoop).blnIsTag Then
                            If maudtElement(intElemLoop).hetType = hetSCRIPToff Then
                                intElemLoop = intElemLoop + 1
                                Exit Do
                            End If
                        End If
                    Loop
                ElseIf maudtElement(intElemLoop).hetType = hetUnknown Then
                    intElemLoop = intElemLoop + 1
                Else
                    Exit Do
                End If

                If intElemLoop > aintLineEndElem(intIdxLine) Then
                    Exit Do
                End If
            Loop

            ' Set the word iterator to the first word in the line.
            If intElemLoop = aintLineStartElem(intIdxLine) Then
                intWord = aintLineStartWord(intIdxLine)
            Else
                intWord = 0
            End If

            Do While intElemLoop <= aintLineEndElem(intIdxLine)
                ' Render the next word.
                picHTML.CurrentY = lngY
                picHTML.CurrentX = lngX

                ' Insert a space if necessary.
                If ((Not maudtElement(intElemLoop).blnUnSpaced) Or intWord > 0) And lngX > lngMarginLeft + lngIndent Then
                    If blnStartUnderline Then
                        If Not blnSpaceInserted Then
                            picHTML.CurrentX = picHTML.CurrentX + picHTML.TextWidth(" ")
                            blnSpaceInserted = True
                        End If
                    Else
                        If Len(Trim(maudtElement(intElemLoop).astrWords(intWord))) > 0 Then
                            If Not blnSpaceInserted Then
                                If (blnRight And _
                                    lngX - 1 > lngMarginRight - (alngLineX(intIdxLine) - lngMarginLeft) - lngPadding) _
                                    Or (blnCentre And _
                                    lngX - 1 > (((lngMarginRight - lngMarginLeft) - (alngLineX(intIdxLine) - lngMarginLeft)) / 2) + lngMarginLeft) _
                                    Or (Not blnRight And Not blnCentre) Then
                                    picHTML.Print " ";
                                    blnSpaceInserted = True
                                End If
                            End If
                        End If
                    End If
                End If

                ' Start underlining for anchors.
                If blnStartUnderline And mblnEnableAnchors And intLinkElement > -1 Then
                    maudtElement(intLinkElement).lngLeft = lngX
                End If
                If Len(Trim(maudtElement(intElemLoop).astrWords(intWord))) > 0 Then
                    If blnStartUnderline Then
                        picHTML.Font.Underline = mblnUnderlineLinks 'True
                        blnStartUnderline = False
                    End If
                    picHTML.Print Replace(maudtElement(intElemLoop).astrWords(intWord), "&nbsp;", " ");
                    blnSpaceInserted = False
                End If
                lngX = picHTML.CurrentX
                If intLinkElement > -1 Then
                    lngXExtent = lngX
                End If

                ' Move onto the next element if the current element's content is exhausted.
                If intElemLoop = aintLineEndElem(intIdxLine) And intWord = aintLineEndWord(intIdxLine) Then
                    Exit Do
                ElseIf intElemLoop < aintLineEndElem(intIdxLine) And intWord = UBound(maudtElement(intElemLoop).astrWords) Then
                    Exit Do
                Else
                    intWord = intWord + 1
                    If maudtElement(intElemLoop).hetType = hetContent Then
                        If intWord > UBound(maudtElement(intElemLoop).astrWords) Then
                            Exit Do
                        End If
                    End If
                End If
            Loop

            If intElemLoop <= UBound(maudtElement) Then
                If maudtElement(intElemLoop).hetType <> hetCOMMENTon Then
                    maudtElement(intElemLoop).lngBottom = lngY + (lngLineHeight * 2) + mintDefaultSpacing
                    maudtElement(intElemLoop).lngRight = lngX
                End If
            End If

            intElemLoop = intElemLoop + 1
        Loop

        ' Wrap to the next line.
        lngY = lngY + lngLineHeight
        picHTML.CurrentY = lngY
        lngX = lngMarginLeft + lngIndent
    Next intIdxLine

ExitPoint:
    lngY = lngY - lngLineHeight
    blnSpacerInserted = False
    Exit Sub

ErrorHandler:
    Resume ExitPoint
End Sub
'
' mblnIsStyleTag()
'
' Returns True if the specified tag type is a style change tag.
'
Private Function mblnIsStyleTag(hetTagType As enumHTMLElementType) As Boolean
    Dim blnRetVal   As Boolean

    Select Case hetTagType
        Case hetFONTon, hetFONToff, _
             hetSTRONGon, hetSTRONGoff, _
             hetEMon, hetEMoff, _
             hetUon, hetUoff, _
             hetBIGon, hetBIGoff, _
             hetSMALLon, hetSMALLoff, _
             hetAon, hetAoff, _
             hetSUPon, hetSUPoff, _
             hetSUBon, hetSUBoff
            blnRetVal = True
        Case Else
            blnRetVal = False
    End Select

    mblnIsStyleTag = blnRetVal
End Function
'
' mSetStyleFromTag()
'
Private Sub mSetStyleFromTag(ByVal intElem As Integer, _
                             ByRef sngLastFontSize As Single, ByRef lngLineHeight As Long, ByRef lngIndentStep As Long, _
                             ByRef blnStartUnderline As Boolean)
    Select Case maudtElement(intElem).hetType
        Case hetFONTon
            On Error Resume Next
            picHTML.FontName = maudtElement(intElem).strFontName
            picHTML.FontSize = maudtElement(intElem).sngFontSize
            picHTML.ForeColor = maudtElement(intElem).lngFontColor
            sngLastFontSize = maudtElement(intElem).sngFontSize
            lngLineHeight = picHTML.TextHeight("X") + mintDefaultSpacing
            lngIndentStep = picHTML.TextWidth("W") * 2
        Case hetFONToff
            mSetDefaultStyle
            sngLastFontSize = msngDefaultFontSize
        Case hetSTRONGon
            picHTML.Font.Bold = True
        Case hetSTRONGoff
            picHTML.Font.Bold = False
        Case hetEMon
            picHTML.Font.Italic = True
        Case hetEMoff
            picHTML.Font.Italic = False
        Case hetUon
            picHTML.Font.Underline = True
        Case hetUoff
            picHTML.Font.Underline = False
        Case hetBIGon, hetSMALLon
            picHTML.FontSize = maudtElement(intElem).sngFontSize
        Case hetBIGoff, hetSMALLoff
            lngLineHeight = picHTML.TextHeight("X") + mintDefaultSpacing
            picHTML.FontSize = sngLastFontSize
            picHTML.Font.Bold = False
            lngLineHeight = picHTML.TextHeight("X") + mintDefaultSpacing
            lngIndentStep = picHTML.TextWidth("W") * 2
        Case hetAon
            picHTML.ForeColor = mlngLinkColor
            blnStartUnderline = True
        Case hetAoff
            picHTML.Font.Underline = False
            picHTML.ForeColor = mlngTextColor
            blnStartUnderline = False
        Case Else
    End Select
End Sub
'
' mlngTranslateHTMLColour()
'
' Translate the specified HTML colour name into a suitable RGB colour value.
'
' strColourName :   The name of the colour to be translated.
'
Private Function mlngTranslateHTMLColour(strColourName As String) As Long
    Dim strRGB      As String

    Select Case LCase(strColourName)
        Case "black"
            strRGB = "000000"
        Case "green"
            strRGB = "008000"
        Case "silver"
            strRGB = "C0C0C0"
        Case "lime"
            strRGB = "00FF00"
        Case "gray"
            strRGB = "808080"
        Case "olive"
            strRGB = "808000"
        Case "white"
            strRGB = "FFFFFF"
        Case "yellow"
            strRGB = "FFFF00"
        Case "maroon"
            strRGB = "800000"
        Case "navy"
            strRGB = "000080"
        Case "red"
            strRGB = "FF0000"
        Case "blue"
            strRGB = "0000FF"
        Case "purple"
            strRGB = "800080"
        Case "teal"
            strRGB = "008080"
        Case "fuchsia"
            strRGB = "FF00FF"
        Case "aqua"
            strRGB = "00FFFF"
        Case Else
            strRGB = "000000"
    End Select

    mlngTranslateHTMLColour = RGB(CLng("&H" & Left(strRGB, 2)), _
                              CLng("&H" & Mid(strRGB, 3, 2)), _
                              CLng("&H" & Right(strRGB, 2)))
End Function
'
' mRender3DBorder()
'
' Draw a 3D border around the specified rectangle.
'
Private Sub mRender3DBorder(blnInset As Boolean, lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long)
    Dim lngCol  As Long

    lngCol = picHTML.ForeColor

    If blnInset Then
        picHTML.ForeColor = vbButtonShadow
    Else
        If picHTML.BackColor = vbWhite Then
            picHTML.ForeColor = vbButtonFace
        Else
            picHTML.ForeColor = vbWhite
        End If
    End If
    picHTML.Line (lngLeft, lngTop)-(lngRight, lngTop)
    picHTML.Line (lngLeft, lngTop)-(lngLeft, lngBottom)

    If blnInset Then
        If picHTML.BackColor = vbWhite Then
            picHTML.ForeColor = vbButtonFace
        Else
            picHTML.ForeColor = vbWhite
        End If
    Else
        picHTML.ForeColor = vbButtonShadow
    End If
    picHTML.Line (lngRight, lngTop)-(lngRight, lngBottom)
    picHTML.Line (lngLeft, lngBottom)-(lngRight + 1, lngBottom)

    picHTML.ForeColor = lngCol
End Sub
'
' mBuildHierarchy()
'
' Structure the elements array as a hierarchy,.
'
Public Sub mBuildHierarchy()
    If mintElements > 0 Then
        maudtElement(0).intChildElements = mBuildElementHierarchy(0, maudtElement(0))
    End If
End Sub
'
' mBuildElementHierarchy()
'
' Structure the specified HTML element as a hierarchy.
'
Private Function mBuildElementHierarchy(ByRef intElem As Integer, ByRef udtElem As tHTMLElement) As Integer
    Dim intChildElem    As Integer

    Do While intElem < mintElements
        intElem = intElem + 1
        If intElem >= mintElements Then
            Exit Do
        End If

        ReDim Preserve udtElem.aintChildElements(intChildElem)
        udtElem.aintChildElements(intChildElem) = intElem
        maudtElement(intElem).intParentElement = udtElem.intElementIndex
        maudtElement(intElem).intChildIndex = intChildElem

        If maudtElement(intElem).blnIsTag Then
            Select Case maudtElement(intElem).hetType
                Case hetHEADon, hetTITLEon, hetBODYon, hetSTRONGon, hetEMon, hetUon, hetPon, hetDIVon, _
                        hetULon, hetOLon, hetTABLEon, hetTRon, hetTDon, hetFONTon, hetAon, hetBLOCKQUOTEon, _
                        hetHeaderon, hetBIGon, hetSMALLon, hetCENTERon
                    maudtElement(udtElem.aintChildElements(intChildElem)).intChildElements = _
                        mBuildElementHierarchy(intElem, maudtElement(intElem))
                Case hetHEADoff, hetTITLEoff, hetBODYoff, hetSTRONGoff, hetEMoff, hetUoff, hetPoff, hetDIVoff, _
                        hetULoff, hetOLoff, hetTABLEoff, hetTDoff, hetFONToff, hetAoff, hetBLOCKQUOTEoff, _
                        hetHeaderoff, hetBIGoff, hetSMALLoff, hetCenterOff
                    Exit Do
                Case hetTRoff
                    Exit Do
            End Select
        End If

        intChildElem = intChildElem + 1
    Loop

    On Error Resume Next
    mBuildElementHierarchy = intChildElem
End Function
'
' mblnGetParent()
'
' Return the immediate parent of udtIn as udtOut.
' Returns True if successful, False otherwise.
'
Private Function mblnGetParent(ByRef udtIn As tHTMLElement, ByRef udtOut As tHTMLElement) As Boolean
    If udtIn.intParentElement > 0 Then
        udtOut = maudtElement(udtIn.intParentElement)
        mblnGetParent = True
    Else
        mblnGetParent = False
    End If
End Function
'
' mblnGetTypedParent()
'
' Return the element of type hetType which contains elemnt udtIn as udtOut.
' Returns True if successful, False otherwise.
'
Private Function mblnGetTypedParent(ByRef udtIn As tHTMLElement, ByRef udtOut As tHTMLElement, _
                                    hetType As enumHTMLElementType) As Boolean

    Dim blnRetVal   As Boolean

    If mblnGetParent(udtIn, udtOut) Then
        Do
            If udtOut.hetType = hetType Then
                blnRetVal = True
                Exit Do
            End If
        Loop While mblnGetParent(udtOut, udtOut)
    End If

    mblnGetTypedParent = blnRetVal
End Function
'
' mblnGetFirstChild()
'
' Return the first child of udtIn as udtOut.
' Returns True if successful, False otherwise.
'
Private Function mblnGetFirstChild(ByRef udtIn As tHTMLElement, ByRef udtOut As tHTMLElement) As Boolean
    If udtIn.intChildElements > 0 Then
        udtOut = maudtElement(udtIn.aintChildElements(0))
        mblnGetFirstChild = True
    Else
        mblnGetFirstChild = False
    End If
End Function
'
' mblnGetNextSibling()
'
' Return the next sibling of udtIn as udtOut.
' Returns True if successful, False otherwise.
'
Private Function mblnGetNextSibling(ByRef udtIn As tHTMLElement, ByRef udtOut As tHTMLElement) As Boolean
    Dim udtTemp As tHTMLElement

    If mblnGetParent(udtIn, udtTemp) Then
        If udtIn.intChildIndex + 1 < udtTemp.intChildElements Then
            udtOut = maudtElement(udtTemp.aintChildElements(udtIn.intChildIndex + 1))
            mblnGetNextSibling = True
        Else
            mblnGetNextSibling = False
        End If
    Else
        mblnGetNextSibling = False
    End If
End Function
'
' mLayoutTable()
'
' Calculate the width of the specified TABLE element and its contained TD elements.
'
Private Sub mLayoutTable(ByRef udtTable As tHTMLElement)
    Dim blnFound                        As Boolean
    Dim intTableCols                    As Integer
    Dim intColIndex                     As Integer
    Dim intColSpan                      As Integer
    Dim intUnsizedCols                  As Integer
    Dim lngAvailWidth                   As Long
    Dim sngTotalWidth                   As Single
    Dim asngColWidth(mcintMaxTableCols) As Single
    Dim udtRow(1)                       As tHTMLElement
    Dim udtCell(1)                      As tHTMLElement
    Dim udtSizingRow                    As tHTMLElement

    ' Count the number of columns in the table.
    If mblnGetFirstChild(udtTable, udtRow(0)) Then
        While udtRow(0).hetType <> hetTRon
            If Not mblnGetNextSibling(udtRow(0), udtRow(0)) Then
                Exit Sub
            End If
        Wend
        If mblnGetFirstChild(udtRow(0), udtCell(0)) Then
            intTableCols = udtCell(0).intColSpan
            While mblnGetNextSibling(udtCell(0), udtCell(1))
                intTableCols = intTableCols + udtCell(1).intColSpan
                udtCell(0) = udtCell(1)
            Wend
        End If
    End If

    ' Calculate the actual width available to cells.
    lngAvailWidth = udtTable.lngWidth - _
                    (intTableCols + 1) * udtTable.intCellSpacing

    ' Locate the first row with where no cell has its COLSPAN attribute set.
    If mblnGetFirstChild(udtTable, udtRow(1)) Then
        Do
            udtRow(0) = udtRow(1)
            If udtRow(0).hetType = hetTRon Then
                udtSizingRow = udtRow(0)
                blnFound = True

                If mblnGetFirstChild(udtRow(0), udtCell(1)) Then
                    Do
                        udtCell(0) = udtCell(1)
                        If udtCell(0).intColSpan > 1 Then
                            blnFound = False
                            Exit Do
                        End If
                    Loop While mblnGetNextSibling(udtCell(0), udtCell(1))
                End If
            End If
        Loop Until blnFound Or Not mblnGetNextSibling(udtRow(0), udtRow(1))
    End If

    If blnFound Then
        ' Collect the size-dictating row's cell widths.
        If mblnGetFirstChild(udtSizingRow, udtCell(1)) Then
            Do
                udtCell(0) = udtCell(1)
                If udtCell(0).hetType = hetTDon Then
                    If udtCell(0).sngWidth < 1 Then
                        asngColWidth(intColIndex) = udtCell(0).sngWidth * lngAvailWidth
                        sngTotalWidth = sngTotalWidth + asngColWidth(intColIndex)
                    ElseIf udtCell(0).sngWidth > 1 Then
                        asngColWidth(intColIndex) = udtCell(0).sngWidth
                        sngTotalWidth = sngTotalWidth + asngColWidth(intColIndex)
                    Else
                        asngColWidth(intColIndex) = 1
                        intUnsizedCols = intUnsizedCols + 1
                    End If
                    intColIndex = intColIndex + 1
                End If
            Loop While mblnGetNextSibling(udtCell(0), udtCell(1))
        End If
    Else
        ' No sizing row was found, assign proportional widths to the columns.
        For intColIndex = 0 To intTableCols - 1
            asngColWidth(intColIndex) = (1 / intTableCols) * CSng(lngAvailWidth)
        Next intColIndex
    End If

    ' Proportionally size any remaining unsized columns.
    For intColIndex = 0 To intTableCols - 1
        If asngColWidth(intColIndex) = 1 Then
            asngColWidth(intColIndex) = (CSng(lngAvailWidth) - sngTotalWidth) \ CSng(intUnsizedCols)
        End If
    Next intColIndex

    ' Calculate the cumulative width of all the columns.
    sngTotalWidth = 0
    For intColIndex = 0 To intTableCols - 1
        sngTotalWidth = sngTotalWidth + asngColWidth(intColIndex)
    Next intColIndex
    If sngTotalWidth > 1 And udtTable.sngWidth > 1 And sngTotalWidth > udtTable.lngWidth Then
        ' Increase the table width if the cumulative width of all the columns is greater than the
        ' specified table width and the table and columns are not proportionally sized.
        maudtElement(udtTable.intElementIndex).lngWidth = sngTotalWidth
    ElseIf sngTotalWidth < udtTable.lngWidth Then
        ' Proportionally increase the width of each column if the cumulative width of all the columns
        ' is less than the table width.
        For intColIndex = 0 To intTableCols - 1
            asngColWidth(intColIndex) = lngAvailWidth * _
                                        (asngColWidth(intColIndex) / sngTotalWidth)
        Next intColIndex
    End If

    ' Set the width of every cell in the table.
    If mblnGetFirstChild(udtTable, udtRow(1)) Then
        Do
            udtRow(0) = udtRow(1)
            If udtRow(0).hetType = hetTRon Then
                intColIndex = 0

                If mblnGetFirstChild(udtRow(0), udtCell(1)) Then
                    Do
                        udtCell(0) = udtCell(1)
                        If udtCell(0).hetType = hetTDon Then
                            intColSpan = udtCell(0).intColSpan
'                            maudtElement(udtCell(0).intElementIndex).lngWidth = _
                                (intColSpan - 1) * (udtTable.intBorderWidth * 2 + udtTable.intCellSpacing)
                            maudtElement(udtCell(0).intElementIndex).lngWidth = 0
                            While intColSpan > 0
                                maudtElement(udtCell(0).intElementIndex).lngWidth = _
                                    maudtElement(udtCell(0).intElementIndex).lngWidth + _
                                    asngColWidth(intColIndex)
                                intColIndex = intColIndex + 1
                                intColSpan = intColSpan - 1
                            Wend
                        End If
                    Loop While mblnGetNextSibling(udtCell(0), udtCell(1))
                End If
            End If
        Loop While mblnGetNextSibling(udtRow(0), udtRow(1))
    End If
End Sub
'
' mRenderBackground()
'
' Tile the background with the current background image.
'
' lngScrollOffset   :   The current vertical scrolling offset.
'
Private Sub mRenderBackground(lngScrollOffset As Long)
    Dim intTileV        As Integer
    Dim intTileH        As Integer
    Dim lngImgWidth     As Long
    Dim lngImgHeight    As Long
    Dim objImage        As Picture

    ' Load the image.
    RaiseEvent LoadImage(mstrBackground, objImage)

    If Not (objImage Is Nothing) Then
        ' Get the image's dimensions.
        lngImgWidth = picHTML.ScaleX(objImage.Width, vbHimetric, vbPixels)
        lngImgHeight = picHTML.ScaleY(objImage.Height, vbHimetric, vbPixels)

        ' Tile the image across the background.
        For intTileV = 0 To picHTML.ScaleHeight / lngImgHeight + 1
            For intTileH = 0 To picHTML.ScaleWidth / lngImgWidth
                picHTML.PaintPicture objImage, _
                                     intTileH * lngImgWidth, _
                                     intTileV * lngImgHeight - lngScrollOffset Mod lngImgHeight, _
                                     lngImgWidth, _
                                     lngImgHeight
            Next intTileH
        Next intTileV
        Set objImage = Nothing
    End If
End Sub
'
' mParseVBURL()
'
' Parse the specified "VB URL" and return the method name and argument list from it.
'
' strVBURL  :   The VB URL to be parsed, e.g. "MyFunc()".
' strMethod :   [out] The name of the method to be called.
' varArgs   :   [out] Variant array containing the argument list.
'
Private Sub mParseVBURL(ByVal strVBURL As String, ByRef strMethod As String, ByRef varArgs As Variant)
    Dim intCh   As Integer
    Dim intArg  As Integer
    Dim strCh   As String * 1
    Dim strTemp As String

    On Error GoTo ErrorHandler

    intCh = 1
    While intCh <= Len(strVBURL)
        strCh = Mid(strVBURL, intCh, 1)

        If Len(strMethod) = 0 Then
            If strCh = " " Or strCh = "(" Then
                ' Return everything up to the first space or left bracket as the method name.
                strMethod = Trim(strTemp)
                strTemp = ""
            Else
                ' Append the next character to the method name.
                strTemp = strTemp & strCh
            End If
        Else
            If strCh = "," Or strCh = ")" Then
                If (Left(strTemp, 1) = """" And Right(strTemp, 1) <> """") Or (Left(strTemp, 1) = "'" And Right(strTemp, 1) <> "'") Then
                    ' Append the next character to the current (string constant) argument.
                    strTemp = strTemp & strCh
                Else
                    ' Append the next argument to the argument list.
                    strTemp = Trim(strTemp)
                    If (Left(strTemp, 1) = """" And Right(strTemp, 1) = """") Or (Left(strTemp, 1) = "'" And Right(strTemp, 1) = "'") Then
                        strTemp = Mid(strTemp, 2, Len(strTemp) - 2)
                    End If
                    If Len(strTemp) > 0 Then
                        If intArg > 0 Then
                            ReDim Preserve varArgs(intArg) As Variant
                        Else
                            ReDim varArgs(0) As Variant
                        End If
                        varArgs(intArg) = strTemp
                        intArg = intArg + 1
                        strTemp = ""
                    End If
                End If
            Else
                ' Append the next character to the current argument.
                strTemp = strTemp & strCh
            End If
        End If

        intCh = intCh + 1
    Wend

    If Len(Trim(strTemp)) > 0 Then
        If Len(strMethod) = 0 Then
            ' The VB URL contains only a method name.
            strMethod = Trim(strTemp)
        Else
            ' Append the final argument to the argument list.
            If intArg > 0 Then
                ReDim Preserve varArgs(intArg) As Variant
            Else
                ReDim varArgs(0) As Variant
            End If
            strTemp = Trim(strTemp)
            If (Left(strTemp, 1) = """" And Right(strTemp, 1) = """") Or (Left(strTemp, 1) = "'" And Right(strTemp, 1) = "'") Then
                strTemp = Mid(strTemp, 2, Len(strTemp) - 2)
            End If
            varArgs(intArg) = strTemp
        End If
    End If

ExitPoint:
    Exit Sub

ErrorHandler:
    Resume ExitPoint
End Sub
'
' mCallByName()
'
' Call the specified method on our container object with the specified arguments.
'
' strMethod :   Name of the method to be called.
' varArgs   :   Variant containing the argument list (this implementation supports a maximum of eight parameters).
'
Private Sub mCallByName(strMethod As String, varArgs As Variant)
    On Error GoTo ErrorHandler

    If IsArray(varArgs) Then
        Select Case UBound(varArgs) + 1
            Case 1
                CallByName Extender.Parent, strMethod, VbMethod, _
                           CVar(varArgs(0))
            Case 2
                CallByName Extender.Parent, strMethod, VbMethod, _
                           CVar(varArgs(0)), CVar(varArgs(1))
            Case 3
                CallByName Extender.Parent, strMethod, VbMethod, _
                           CVar(varArgs(0)), CVar(varArgs(1)), CVar(varArgs(2))
            Case 4
                CallByName Extender.Parent, strMethod, VbMethod, _
                           CVar(varArgs(0)), CVar(varArgs(1)), CVar(varArgs(2)), CVar(varArgs(3))
            Case 5
                CallByName Extender.Parent, strMethod, VbMethod, _
                           CVar(varArgs(0)), CVar(varArgs(1)), CVar(varArgs(2)), CVar(varArgs(3)), _
                           CVar(varArgs(4))
            Case 6
                CallByName Extender.Parent, strMethod, VbMethod, _
                           CVar(varArgs(0)), CVar(varArgs(1)), CVar(varArgs(2)), CVar(varArgs(3)), _
                           CVar(varArgs(4)), CVar(varArgs(5))
            Case 7
                CallByName Extender.Parent, strMethod, VbMethod, _
                           CVar(varArgs(0)), CVar(varArgs(1)), CVar(varArgs(2)), CVar(varArgs(3)), _
                           CVar(varArgs(4)), CVar(varArgs(5)), CVar(varArgs(6))
            Case 8
                CallByName Extender.Parent, strMethod, VbMethod, _
                           CVar(varArgs(0)), CVar(varArgs(1)), CVar(varArgs(2)), CVar(varArgs(3)), _
                           CVar(varArgs(4)), CVar(varArgs(5)), CVar(varArgs(6)), CVar(varArgs(7))
            Case Else
                CallByName Extender.Parent, strMethod, VbMethod
        End Select
    Else
        CallByName Extender.Parent, strMethod, VbMethod
    End If

ExitPoint:
    Exit Sub

ErrorHandler:
    Resume ExitPoint
End Sub
'
' mInsertElement()
'
' Insert a new element at the specified point.
'
' intInsertAt   :   Array position at which to insert a the new element.
'
Private Sub mInsertElement(intInsertAt As Integer)
    Dim intIndex    As Integer

    ' Reserve space for the new element.
    ReDim Preserve maudtElement(UBound(maudtElement) + 1)

    ' Shift everything forwards by one element.
    For intIndex = UBound(maudtElement) - 1 To intInsertAt Step -1
        maudtElement(intIndex + 1) = maudtElement(intIndex)
        maudtElement(intIndex + 1).intElementIndex = intIndex + 1
    Next intIndex
End Sub
'
' mFixUnpairedTags()
'
' Fix any unpaired tag by inserting a closing tag at the appropriate position in the array of elements.
' This version supports unpaired <P>, <TR> and <TD> tags only.
'
Private Sub mFixUnpairedTags()
    Dim intStart        As Integer
    Dim intEnd          As Integer
    Dim intNext         As Integer
    Dim intNestLevel    As Integer

    On Error GoTo ErrorHandler

    ' Fix unpaired <P> tags.
    While intStart <= UBound(maudtElement)
        ' Find the next <P> element.
        If maudtElement(intStart).hetType = hetPon Then
            intEnd = intStart + 1

            Do While intEnd <= UBound(maudtElement)
                If maudtElement(intEnd).hetType = hetPoff Then
                    ' Carry on searching if a </P> tag is found.
                    intStart = intEnd + 1
                    Exit Do
                ElseIf maudtElement(intEnd).hetType = hetPon Or _
                    maudtElement(intEnd).hetType = hetTABLEon Or _
                    maudtElement(intEnd).hetType = hetTDoff Then
                        ' Insert a </P> tag if a <P> or <TABLE> tag is encountered.
                        mInsertElement intEnd
                        maudtElement(intEnd).hetType = hetPoff
                        maudtElement(intEnd).strHTML = "</P>"
                        intStart = intEnd
                        Exit Do
                Else
                    intEnd = intEnd + 1
                End If
            Loop
        End If

        intStart = intStart + 1
    Wend

    ' Fix unpaired <TR> tags.
    intStart = 0
    While intStart <= UBound(maudtElement)
        ' Find the next <TR> element.
        If maudtElement(intStart).hetType = hetTRon Then
            intEnd = intStart + 1
            intNext = intStart + 1
            intNestLevel = 0

            Do While intEnd <= UBound(maudtElement)
                If intNestLevel = 0 Then
                    If maudtElement(intEnd).hetType = hetTRoff Then
                        ' Carry on searching if a </TR> tag is found.
                        intStart = intNext
                        Exit Do
                    ElseIf maudtElement(intEnd).hetType = hetTABLEon Then
                        intNestLevel = intNestLevel + 1
                        intEnd = intEnd + 1
                    ElseIf maudtElement(intEnd).hetType = hetTRon Or _
                        maudtElement(intEnd).hetType = hetTABLEoff Then
                        ' Insert a </TR> tag if a <TR> or </TABLE> tag is encountered.
                        mInsertElement intEnd
                        maudtElement(intEnd).hetType = hetTRoff
                        maudtElement(intEnd).strHTML = "</TR>"
                        intStart = intNext
                        Exit Do
                    Else
                        intEnd = intEnd + 1
                    End If
                Else
                    If maudtElement(intEnd).hetType = hetTABLEon Then
                        intNestLevel = intNestLevel + 1
                    ElseIf maudtElement(intEnd).hetType = hetTABLEoff Then
                        intNestLevel = intNestLevel - 1
                    End If
                    intEnd = intEnd + 1
                End If
            Loop
        End If

        intStart = intStart + 1
    Wend

    ' Fix unpaired <TD> tags.
    intStart = 0
    While intStart <= UBound(maudtElement)
        ' Find the next <TD> element.
        If maudtElement(intStart).hetType = hetTDon Then
            intEnd = intStart + 1
            intNext = intStart + 1
            intNestLevel = 0

            Do While intEnd <= UBound(maudtElement)
                If intNestLevel = 0 Then
                    If maudtElement(intEnd).hetType = hetTDoff Then
                        ' Carry on searching if a </TD> tag is found.
                        intStart = intNext
                        Exit Do
                    ElseIf maudtElement(intEnd).hetType = hetTABLEon Then
                        intNestLevel = intNestLevel + 1
                        intEnd = intEnd + 1
                    ElseIf maudtElement(intEnd).hetType = hetTDon Or _
                        maudtElement(intEnd).hetType = hetTRoff Or _
                        maudtElement(intEnd).hetType = hetTABLEoff Then
                        ' Insert a </TD> tag if a <TD>, </TR> or </TABLE> tag is encountered.
                        mInsertElement intEnd
                        maudtElement(intEnd).hetType = hetTRoff
                        maudtElement(intEnd).strHTML = "</TD>"
                        intStart = intNext
                        Exit Do
                    Else
                        intEnd = intEnd + 1
                    End If
                Else
                    If maudtElement(intEnd).hetType = hetTABLEon Then
                        intNestLevel = intNestLevel + 1
                    ElseIf maudtElement(intEnd).hetType = hetTABLEoff Then
                        intNestLevel = intNestLevel - 1
                    End If
                    intEnd = intEnd + 1
                End If
            Loop
        End If

        intStart = intStart + 1
    Wend

ExitPoint:
    Exit Sub

ErrorHandler:
    Resume ExitPoint
End Sub
