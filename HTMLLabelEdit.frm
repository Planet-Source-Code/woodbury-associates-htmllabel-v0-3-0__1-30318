VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FHTMLLabelEdit 
   Caption         =   "HTMLLabel Editor"
   ClientHeight    =   7950
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   11880
   Icon            =   "HTMLLabelEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog ctlCommonDialog 
      Left            =   30
      Top             =   810
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtHTMLSource 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4050
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   780
      Width           =   1245
   End
   Begin HTMLLabelEdit.HTMLLabel ctlPanel 
      Height          =   705
      Left            =   30
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   30
      Width           =   3960
      _ExtentX        =   6985
      _ExtentY        =   1244
      Appearance      =   1
      BorderStyle     =   1
      BackColor       =   -2147483633
      EnableAnchors   =   -1  'True
      EnableScroll    =   0   'False
      EnableTooltips  =   -1  'True
      DefaultFontName =   "Tahoma"
      DefaultFontSize =   9
      UnderlineLinks  =   0   'False
      DefaultPadding  =   0
      DefaultSpacing  =   4
   End
   Begin HTMLLabelEdit.HTMLLabel ctlHTMLView 
      Height          =   705
      Left            =   4050
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   30
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   1244
      Appearance      =   1
      BorderStyle     =   1
      BackColor       =   -2147483633
      EnableAnchors   =   -1  'True
      EnableScroll    =   -1  'True
      EnableTooltips  =   -1  'True
      DefaultFontName =   "Tahoma"
      DefaultFontSize =   9
      UnderlineLinks  =   -1  'True
      DefaultPadding  =   4
      DefaultSpacing  =   4
   End
End
Attribute VB_Name = "FHTMLLabelEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'
' Form FHTMLLabelEdit.
'
' HTMLLabel demo edit/view application.
'
' Copyright © 2001 Woodbury Associates.
'

'
' Windows API declarations.
'
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                                                               ByVal lpOperation As String, _
                                                                               ByVal lpFile As String, _
                                                                               ByVal lpParameters As String, _
                                                                               ByVal lpDirectory As String, _
                                                                               ByVal nShowCmd As Long) As Long

'
' Private member variables.
'
Private mblnSizing      As Boolean
Private mstrCurrentView As String

'
' ctlHTMLView_LoadImage()
'
' HTMLLabel callback which is fired to obtain the specified image.
'
' Source    :   The SRC attribute from the HTML <IMG> tag.
' Image     :   A Picture object reference to be set to the loaded image.
'
Private Sub ctlHTMLView_LoadImage(Source As String, Image As stdole.Picture)
    On Error Resume Next
    If Mid(Trim(Source), 2, 1) = ":" Or Left(Trim(Source), 2) = "\\" Then
        ' Treat Source as an absolute file path.
        Set Image = LoadPicture(Source)
    Else
        ' Treat Source as a path relative to the current directory.
        Set Image = LoadPicture(App.Path & "\" & Source)
    End If
End Sub
'
' ctlPanel_HyperlinkClick()
'
' Respond to any clicked links in the instruction HTMLLabel.
'
Private Sub ctlPanel_HyperlinkClick(Href As String)
    Select Case Href
        Case "Show my HTML"
            ShowMyHTML
        Case "Hello, World!"
            txtHTMLSource.Text = _
                                "<html>" & vbCrLf & _
                                "     <body>" & vbCrLf & _
                                "        <br>" & vbCrLf & _
                                "         <table align=center border=1 width=200 bgcolor=white>" & vbCrLf & _
                                "             <tr height=100>" & vbCrLf & _
                                "                 <td align=center valign=center bgcolor=#6070b0>" & vbCrLf & _
                                "                     <font size=+1 color=white>" & vbCrLf & _
                                "                         <b>Hello, World !</b>" & vbCrLf & _
                                "                     </font>" & vbCrLf & _
                                "                 </td>" & vbCrLf & _
                                "             </tr>" & vbCrLf & _
                                "         </table>" & vbCrLf & _
                                "     </body>" & vbCrLf & _
                                "</html>"
            txtHTMLSource.SelLength = Len(txtHTMLSource.Text)
            txtHTMLSource.SetFocus
        Case "Help"
            HelpContents
        Case Else
    End Select
End Sub
'
' ctlPanel_LoadImage()
'
' Load a standard Picture object with the specified image file.
'
Private Sub ctlPanel_LoadImage(Source As String, Image As stdole.Picture)
    On Error Resume Next
    If Mid(Trim(Source), 2, 1) = ":" Or Left(Trim(Source), 2) = "\\" Then
        ' Treat Source as an absolute file path.
        Set Image = LoadPicture(Source)
    Else
        ' Treat Source as a path relative to the current directory.
        Set Image = LoadPicture(App.Path & "\" & Source)
    End If
End Sub
'
' Form_KeyDown()
'
' F5 accelerator for [Show my HTML] and F1 accelerator for Help | Contents.
'
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            HelpContents
            KeyCode = 0
        Case vbKeyF5
            ShowMyHTML
            KeyCode = 0
        Case Else
    End Select
End Sub
'
' Form_Load()
'
' Form initialisation.
'
Private Sub Form_Load()
    mstrCurrentView = "Welcome"
    ctlPanel.DocumentHTML = mstrPanelHTML()
End Sub
'
' Form_MouseDown()
'
' Allow the user to size the HTML view and source controls.
'
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MousePointer = vbSizeNS Then
        mblnSizing = True
    Else
        mblnSizing = False
    End If
End Sub
'
' Form_MouseMove()
'
' Allow the user to size the HTML view and source controls.
'
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not mblnSizing Then
        If X >= ctlHTMLView.Left And X <= ctlHTMLView.Left + ctlHTMLView.Width And _
            Y > ctlHTMLView.Top + ctlHTMLView.Height And Y <= txtHTMLSource.Top Then
            MousePointer = vbSizeNS
        Else
            MousePointer = vbDefault
        End If
    Else
        If Y > 500 And Y < Height - 1200 Then
            ' Position and size our controls.
            ctlHTMLView.Height = Y - ctlHTMLView.Top
            txtHTMLSource.Top = ctlHTMLView.Top + ctlHTMLView.Height + 60
            txtHTMLSource.Height = Height - txtHTMLSource.Top - 730
        End If
    End If
End Sub
'
' Form_MouseUp()
'
' Allow the user to size the HTML view and source controls.
'
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MousePointer = vbDefault

    If mblnSizing Then
        ' Refresh our HTMLLabel controls.
        ctlPanel.Refresh False
        ctlHTMLView.Refresh False

        mblnSizing = False
    End If
End Sub
'
' txtHTMLSource_MouseMove()
'
Private Sub txtHTMLSource_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not mblnSizing Then
        MousePointer = vbDefault
    End If
End Sub
'
' Form_Resize()
'
' Layout the user interface.
'
Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        ' Fix a minimum size for the form.
        If Width < 5000 Then
            Width = 5000
        End If
        If Height < 4000 Then
            Height = 4000
        End If

        ' Position and size our controls.
        ctlPanel.Height = Height - 460

        ctlHTMLView.Height = (Height - 450) / 2
        ctlHTMLView.Width = Width - ctlHTMLView.Left - 120

        txtHTMLSource.Top = ctlHTMLView.Top + ctlHTMLView.Height + 60
        txtHTMLSource.Height = ctlHTMLView.Height - 60
        txtHTMLSource.Width = ctlHTMLView.Width

        ctlPanel.DocumentHTML = mstrPanelHTML()
    End If

    ' Refresh our HTMLLabel controls.
    ctlPanel.Refresh False
    ctlHTMLView.Refresh False
End Sub
'
' ctlHTMLView_HyperlinkClick()
'
' Follow any clicked hyperlinks.
'
Private Sub ctlHTMLView_HyperlinkClick(Href As String)
    ShellExecute 0, "open", Href, "", "", 0
End Sub
'
' mstrPanelHTML()
'
' Build the HTML document which provides the main control panel user interface.
'
Private Function mstrPanelHTML() As String
    Dim lngPanelHeight  As Long
    Dim strRetVal       As String
    
    lngPanelHeight = (ctlPanel.Height \ Screen.TwipsPerPixelY) - 124

    strRetVal = "<html><body link=white bgcolor='#FFFFE1'>" & _
                "<table height=" & lngPanelHeight & " width=255 border=0 cellspacing=0 cellpadding=0 bgcolor=#FFFFE1>"

    strRetVal = strRetVal & "<tr height=30><td>" & _
                "<a href=vb:SetView(""Welcome"") title='Show the welcome message'>" & _
                "<img src=welcome.gif hspace=0 vspace=0 border=0></a>" & _
                "</td></tr>"

    Select Case mstrCurrentView
        Case "Welcome"
            strRetVal = strRetVal & "<tr height=" & lngPanelHeight & "><td>"
            strRetVal = strRetVal & "<table border=0><tr><td>"
            strRetVal = strRetVal & "<p>Welcome to the HTMLLabel Editor," & _
                        "a simple application which demonstrates some of the capabilities of the HTMLLabel control and allows you to test it with your own HTML.</p>" & _
                        "<p>To try out HTMLLabel, first<a href='Hello, World!'><font color=blue> type</font></a> your HTML source into the textbox on the right.</p>" & _
                        "<p>Then press the &quot;Show my HTML&quot; button (below) or click<a href='Show my HTML'><font color=blue> here</font></a> to see how it looks it in HTMLLabel.</p>" & _
                        "<p>For more information, including full details on the HTML tags supported by HTMLLabel and using HTMLLabel in your own software, view the<a href='Help'>" & _
                        "<font color=blue> readme file</font></a>.</p>"
            strRetVal = strRetVal & "</td></tr></table>"
            strRetVal = strRetVal & "</td></tr>"
            strRetVal = strRetVal & "<tr height=30><td>" & _
                        "<a href=vb:Setview(""Menu"") title='Show the menu'>" & _
                        "<img src=menu.gif hspace=0 vspace=0 border=0></a>" & _
                        "</td></tr>"
            strRetVal = strRetVal & "<tr height=30><td>" & _
                        "<a href=vb:Setview(""Options"") title='Show options'>" & _
                        "<img src=options.gif hspace=0 vspace=0 border=0></a>" & _
                        "</td></tr>"
        Case "Menu"
            strRetVal = strRetVal & "<tr height=30><td>" & _
                        "<a href=vb:Setview(""Menu"") title='Show the menu'>" & _
                        "<img src=menu.gif hspace=0 vspace=0 border=0></a>" & _
                        "</td></tr>"
            strRetVal = strRetVal & "<tr height=" & lngPanelHeight & "><td>"
            strRetVal = strRetVal & "<table border=0><tr><td>"
            strRetVal = strRetVal & "<br>What do you want to do ?" & _
                        "<blockquote>" & _
                        "<a href='vb:FileOpen()' title='Open a file'><font color=blue>Open a file</font></a><br>" & _
                        "<a href='vb:FileSave()' title='Save my HTML'><font color=blue>Save my HTML</font></a><br>" & _
                        "<a href='vb:FileCopyAsVBCode()' title='Copy my HTML as VB code'><font color=blue>Copy my HTML as VB code</font></a><br>" & _
                        "<a href='vb:FileExit()' title='Close the application'><font color=blue>Close the application</font></a><br>" & _
                        "<br>" & _
                        "<a href='vb:HelpContents()' title='View the help file'><font color=blue>View the help file</font></a><br>" & _
                        "<a href='vb:HelpAbout()' title='View the About... box'><font color=blue>View the About... box</font></a><br>" & _
                        "</blockquote>"
            strRetVal = strRetVal & "</td></tr></table>"
            strRetVal = strRetVal & "</td></tr>"
            strRetVal = strRetVal & "<tr height=30><td>" & _
                        "<a href=vb:Setview(""Options"") title='Show options'>" & _
                        "<img src=options.gif hspace=0 vspace=0 border=0></a>" & _
                        "</td></tr>"
        Case "Options"
            strRetVal = strRetVal & "<tr height=30><td>" & _
                        "<a href=vb:Setview(""Menu"") title='Show the menu'>" & _
                        "<img src=menu.gif hspace=0 vspace=0 border=0></a>" & _
                        "</td></tr>"
            strRetVal = strRetVal & "<tr height=30><td>" & _
                        "<a href=vb:Setview(""Options"") title='Show options'>" & _
                        "<img src=options.gif hspace=0 vspace=0 border=0></a>" & _
                        "</td></tr>"
            strRetVal = strRetVal & "<tr height=" & lngPanelHeight & "><td>"
                        strRetVal = strRetVal & "<table border=0>" & _
                        "<tr><td width=50% valign=center align=right><br>Underline links:</td><td><br><a href='vb:SetOption(""UnderlineLinks"", " & IIf(ctlHTMLView.UnderlineLinks, 0, 1) & ")' title='Underline links'><img src='" & IIf(ctlHTMLView.UnderlineLinks, "tick.gif", "untick.gif") & "' width=16 height=16></a></td></tr>" & _
                        "<tr><td valign=center align=right>Enable scrollbars:</td><td><a href='vb:SetOption(""EnableScroll"", " & IIf(ctlHTMLView.EnableScroll, 0, 1) & ")' title='Enable scrollbars'><img src=" & IIf(ctlHTMLView.EnableScroll, "tick.gif", "untick.gif") & " width=16 height=16></a></td></tr>" & _
                        "<tr><td valign=center align=right>Enable anchors:</td><td><a href='vb:SetOption(""EnableAnchors"", " & IIf(ctlHTMLView.EnableAnchors, 0, 1) & ")' title='Enable anchors'><img src=" & IIf(ctlHTMLView.EnableAnchors, "tick.gif", "untick.gif") & " width=16 height=16></a></td></tr>" & _
                        "<tr><td valign=center align=right>Enable tooltips:</td><td><a href='vb:SetOption(""EnableTooltips"", " & IIf(ctlHTMLView.EnableTooltips, 0, 1) & ")' title='Enable tooltips'><img src=" & IIf(ctlHTMLView.EnableTooltips, "tick.gif", "untick.gif") & " width=16 height=16></a></td></tr>" & _
                        "<tr><td valign=center align=right>Default padding:</td>" & _
                        "<td><table width=56 border=0 cellspacing=0 cellpadding=0><tr><td width=16 align=center><a href='vb:IncDefaultPadding(-1)' title='Decrease the default padding'>" & _
                        "<img src='dec.gif' width=16 height=16 hspace=0 vspace=0></a></td>" & _
                        "<td width=24 align=center valign=center>" & ctlHTMLView.DefaultPadding & "</td>" & _
                        "<td width=16 align=center><a href='vb:IncDefaultPadding(1)' title='Increase the default padding'>" & _
                        "<img src='inc.gif' width=16 height=16 hspace=0 vspace=0></a></td>" & _
                        "</tr></table></td></tr>" & _
                        "<tr><td valign=center align=right>Default spacing:</td>" & _
                        "<td><table width=56 border=0 cellspacing=0 cellpadding=0><tr><td width=16 align=center><a href='vb:IncDefaultSpacing(-1)' title='Decrease the default spacing'>" & _
                        "<img src='dec.gif' width=16 height=16 hspace=0 vspace=0></a></td>" & _
                        "<td width=24 align=center valign=center>" & ctlHTMLView.DefaultSpacing & "</td>" & _
                        "<td width=16 align=center><a href='vb:IncDefaultSpacing(1)' title='Increase the default spacing'>" & _
                        "<img src='inc.gif' width=16 height=16 hspace=0 vspace=0></a></td>" & _
                        "</tr></table></td></tr>" & _
                        "<tr><td valign=top align=right>Default font:</td><td><a href='vb:SetFont()' title='Change the default font'><font color=blue>" & ctlHTMLView.DefaultFontName & ", " & ctlHTMLView.DefaultFontSize & "pt</font></a></td></tr>" & _
                        "<tr><td valign=center align=right>Background colour:</td><td><a href='vb:SetBackColor()' title='Change the default background colour'><font color=blue>" & ctlHTMLView.BackColor & "</font></a></td></tr>" & _
                        "<tr><td valign=center align=right>Border style:</td><td><a href='vb:SetOption(""BorderStyle"", " & IIf(ctlHTMLView.BorderStyle = 1, 0, 1) & ")' title=""Change the control's border style""><font color=blue>" & IIf(ctlHTMLView.BorderStyle = 1, "Single", "None") & "</font></a></td></tr>" & _
                        "<tr><td valign=center align=right>Appearance:</td><td><a href='vb:SetOption(""Appearance"", " & IIf(ctlHTMLView.Appearance = 1, 0, 1) & ")' title=""Change the control's appearance""><font color=blue>" & IIf(ctlHTMLView.Appearance = 1, "3D", "Flat") & "</font></a></td></tr>" & _
                        "</table>"
            strRetVal = strRetVal & "</td></tr>"
    End Select

    strRetVal = strRetVal & "<tr height=30><td>" & _
                "<a href=vb:ShowMyHTML() title='Show my HTML'>" & _
                "<img src=show.gif hspace=0 vspace=0 border=0></a>" & _
                "</td></tr>"

    strRetVal = strRetVal & "</table>" & _
                "</body></html>"

    mstrPanelHTML = strRetVal
End Function

'
' Control panel callbacks.
'

'
' SetView()
'
' Switch the Welcome/Options view to the specified view.
'
Public Sub SetView(strView As String)
    mstrCurrentView = strView

    ' Rebuild the options HTML document.
    ctlPanel.DocumentHTML = mstrPanelHTML()

    ' Refresh the display.
    ctlPanel.Refresh False
End Sub

'
' Actions HTML callbacks.
'

'
' FileOpen()
'
' Load the content of a file selected by the user inot the HTML source TextBox.
'
Public Sub FileOpen()
    Dim objFile As Object

    On Error GoTo ErrorHandler

    ' Prompt the user for a file to open.
    With ctlCommonDialog
        .CancelError = True
        .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist
        .Filter = "HTML files (*.html)|*.html|All files (*.*)|*.*"
        .DialogTitle = "Open a file..."
        .ShowOpen
    End With

    ' Load the file into the text box and render it.
    Set objFile = CreateObject("Scripting.FileSystemObject")
    txtHTMLSource.Text = objFile.OpenTextFile(ctlCommonDialog.FileName, 1).ReadAll()
    ShowMyHTML

ExitPoint:
    Set objFile = Nothing
    Exit Sub

ErrorHandler:
    If Err.Number <> cdlCancel Then
        ctlHTMLView.DocumentHTML = "<html><body>" & _
                                    "<p>Error:</p>" & _
                                    "<p>The file could not be opened.</p>" & _
                                    "</body></html>"
    End If
    Resume ExitPoint
End Sub
'
' FileSave()
'
' Save the user's current HTML document source to a file specified by the user.
'
Public Sub FileSave()
    Dim objFile     As Object

    On Error GoTo ErrorHandler

    ' Prompt the user for a file to open.
    With ctlCommonDialog
        .CancelError = True
        .Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
        .Filter = "HTML files (*.html)|*.html|All files (*.*)|*.*"
        .DialogTitle = "Save my HTML..."
        .ShowSave
    End With

    ' Load the file into the text box and render it.
    Set objFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(ctlCommonDialog.FileName, 2, True, 0)

    objFile.Write txtHTMLSource.Text
    objFile.Close

    ShowMyHTML

ExitPoint:
    Set objFile = Nothing
    Exit Sub

ErrorHandler:
    If Err.Number <> cdlCancel Then
        ctlHTMLView.DocumentHTML = "<html><body>" & _
                                    "<p>Error:</p>" & _
                                    "<p>The file could not be saved.</p>" & _
                                    "</body></html>"
    End If
    Resume ExitPoint
End Sub
'
' FileExit()
'
' Close the application.
'
Public Sub FileExit()
    Unload Me
End Sub
'
' FileCopyAsVBCode()
'
' Copy the user's current HTML document source onto the clipboard, formatted as Visual Basic code.
'
Public Sub FileCopyAsVBCode()
    Const intLineContinuations  As Integer = 20
    Const cintLineLength        As Integer = 60

    Dim intLine     As Integer
    Dim strVBCode   As String

    If Len(txtHTMLSource.Text) > 0 Then
        strVBCode = "strMyHTML = _ " & vbCrLf
        For intLine = 0 To Len(txtHTMLSource.Text) \ cintLineLength
            If intLine > 0 Then
                If intLine Mod intLineContinuations = 0 Then
                    strVBCode = strVBCode & vbCrLf & vbCrLf & "strMyHTML = strMyHTML "
                End If
                strVBCode = strVBCode & " & _ " & vbCrLf
            End If

            strVBCode = strVBCode & """" & _
                        Replace( _
                            Replace( _
                                Replace( _
                                    Replace( _
                                        Replace( _
                                            Mid(txtHTMLSource.Text, intLine * (cintLineLength) + 1, cintLineLength), _
                                        vbCrLf, " "), _
                                    vbCr, " "), _
                                vbLf, " "), _
                            vbTab, " "), _
                        """", """""") & """"
        Next intLine

        Clipboard.SetText strVBCode, vbCFText
    End If
End Sub
'
' HelpContents()
'
' Display the readme file.
'
Public Sub HelpContents()
    Dim objFile As Object

    On Error GoTo ErrorHandler

    Set objFile = CreateObject("Scripting.FileSystemObject")
    ctlHTMLView.DocumentHTML = objFile.OpenTextFile(App.Path & "\readme.html", 1).ReadAll()

ExitPoint:
    Set objFile = Nothing
    Exit Sub

ErrorHandler:
    ctlHTMLView.DocumentHTML = "<html><body>" & _
                                "<p>Error:</p>" & _
                                "<p>Either file README.HTML does not exist, or the " & _
                                "FileSystemObject is not correctly registered on your " & _
                                "system.</p>" & _
                                "</body></html>"
    Resume ExitPoint
End Sub
'
' HelpAbout()
'
' Display the About... box.
'
Public Sub HelpAbout()
    Dim frmAbout    As FAbout

    Set frmAbout = New FAbout
    frmAbout.Show vbModal
    Set frmAbout = Nothing
End Sub

'
' Options HTML callbacks.
'

'
' SetOption()
'
' Set the specified option to the specified setting.
'
' strOption     :   Identifier of the option to be changed.
' intSetting    :   The new setting for the option.
'
Public Sub SetOption(strOption As String, intSetting As Integer)
    ' Set the specified option's new value.
    Select Case strOption
        Case "UnderlineLinks"
            ctlHTMLView.UnderlineLinks = IIf(intSetting = 0, False, True)
        Case "EnableScroll"
            ctlHTMLView.EnableScroll = IIf(intSetting = 0, False, True)
        Case "EnableAnchors"
            ctlHTMLView.EnableAnchors = IIf(intSetting = 0, False, True)
        Case "EnableTooltips"
            ctlHTMLView.EnableTooltips = IIf(intSetting = 0, False, True)
        Case "BorderStyle"
            ctlHTMLView.BorderStyle = intSetting
        Case "Appearance"
            ctlHTMLView.Appearance = intSetting
        Case Else
    End Select

    ' Rebuild the options HTML document.
    ctlPanel.DocumentHTML = mstrPanelHTML()

    ' Refresh the display.
    ctlPanel.Refresh False
    ctlHTMLView.Refresh False
End Sub
'
' SetFont()
'
' Present the font selection common dialog and set the HTML view's font accordingly.
'
Public Sub SetFont()
    On Error GoTo ErrorHandler

    ' Set the control's font according to the user's selection.
    With ctlCommonDialog
        .CancelError = True
        .Flags = cdlCFBoth
        .ShowFont
        ctlHTMLView.DefaultFontName = .FontName
        ctlHTMLView.DefaultFontSize = .FontSize
    End With

    ' Rebuild the options HTML document.
    ctlPanel.DocumentHTML = mstrPanelHTML()

    ' Refresh the display.
    ctlPanel.Refresh False
    ctlHTMLView.Refresh False

ExitPoint:
    Exit Sub

ErrorHandler:
    Resume ExitPoint
End Sub
'
' SetBackColor()
'
' Present the colour selection common dialog and set the HTML view's background colour accordingly.
'
Public Sub SetBackColor()
    On Error GoTo ErrorHandler

    ' Set the control's background colour according to the user's selection.
    With ctlCommonDialog
        .Flags = cdlCCFullOpen
        .CancelError = True
        .ShowColor
        ctlHTMLView.BackColor = .Color
    End With

    ' Rebuild the options HTML document.
    ctlPanel.DocumentHTML = mstrPanelHTML()

    ' Refresh the display.
    ctlPanel.Refresh False
    ctlHTMLView.Refresh False

ExitPoint:
    Exit Sub

ErrorHandler:
    Resume ExitPoint
End Sub
'
' IncDefaultPadding()
'
' Increment the default padding by the specified amount.
'
Public Sub IncDefaultPadding(intIncrement As Integer)
    If ctlHTMLView.DefaultPadding + intIncrement > -1 Then
        ' Increment the default padding by the specified amount.
        ctlHTMLView.DefaultPadding = ctlHTMLView.DefaultPadding + intIncrement
    
        ' Rebuild the options HTML document.
        ctlPanel.DocumentHTML = mstrPanelHTML()
    
        ' Refresh the display.
        ctlPanel.Refresh False
        ctlHTMLView.Refresh False
    End If
End Sub
'
' IncDefaultSpacing()
'
' Increment the default spacing by the specified amount.
'
Public Sub IncDefaultSpacing(intIncrement As Integer)
    If ctlHTMLView.DefaultSpacing + intIncrement > -1 Then
        ' Increment the default padding by the specified amount.
        ctlHTMLView.DefaultSpacing = ctlHTMLView.DefaultSpacing + intIncrement
    
        ' Rebuild the options HTML document.
        ctlPanel.DocumentHTML = mstrPanelHTML()
    
        ' Refresh the display.
        ctlPanel.Refresh False
        ctlHTMLView.Refresh False
    End If
End Sub
'
' ShowMyHTML()
'
' "Show my HTML" command button handler - display the current HTML source in the HTMLLabel.
'
Public Sub ShowMyHTML()
    ctlHTMLView.DocumentHTML = txtHTMLSource.Text
End Sub
