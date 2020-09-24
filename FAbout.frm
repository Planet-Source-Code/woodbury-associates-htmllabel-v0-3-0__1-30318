VERSION 5.00
Begin VB.Form FAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About HTMLLabel"
   ClientHeight    =   3435
   ClientLeft      =   4755
   ClientTop       =   3705
   ClientWidth     =   5220
   Icon            =   "FAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   229
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   348
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin HTMLLabelEdit.HTMLLabel ctlHTML 
      Height          =   2505
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   4419
      Appearance      =   0
      BorderStyle     =   0
      BackColor       =   -2147483633
      EnableAnchors   =   -1  'True
      EnableScroll    =   0   'False
      EnableTooltips  =   -1  'True
      DefaultFontName =   "Tahoma"
      DefaultFontSize =   8
      UnderlineLinks  =   -1  'True
      DefaultPadding  =   4
      DefaultSpacing  =   4
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   2910
      Width           =   1215
   End
End
Attribute VB_Name = "FAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'
'   Form FAbout.
'
'   About... box.
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
' cmdOK_Click
'
' Dismiss the dialog.
'
Private Sub cmdOK_Click()
    Unload Me
End Sub
'
' ctlHTML_HyperlinkClick()
'
' Follow any clicked hyperlinks.
'
Private Sub ctlHTML_HyperlinkClick(Href As String)
    ShellExecute 0, "open", Href, "", "", 0
End Sub
'
' ctlHTML_LoadImage()
'
' HTMLLabel callback which is fired to obtain the specified image.
'
' Source    :   The SRC attribute from the HTML <IMG> tag.
' Image     :   A Picture object reference to be set to the loaded image.
'
Private Sub ctlHTML_LoadImage(Source As String, Image As stdole.Picture)
    On Error Resume Next
    Set Image = LoadPicture(App.Path & "\" & Source)
End Sub

'
' Form_Load
'
' Initialisation.
'
Private Sub Form_Load()
    ' Set screen display.
    ctlHTML.DocumentHTML = "<html><body>" & _
                           "<table border='0' cellspacing='0' cellpadding='0'><tr>" & _
                           "<td width='60'><img src='logo.gif' border='1' hspace='0'></td>" & _
                           "<td><font size='+4'><b>HTMLLabel</b></font><br><br>" & _
                           "<font face='tahoma' size='3'>Version: " & ctlHTML.Version & "<br><hr>" & _
                           "<p>For further information, version updates and other products, " & _
                           "<a href='http://www.damnet.freeserve.co.uk/products/'>" & _
                           "visit our products web site</a>.</p><br><hr>" & _
                           "<p>Copyright &copy; 2001-2002 <a href='http://www.woodbury.co.uk'>" & _
                           "Woodbury Associates</a></p>" & _
                           "</font></td>" & _
                           "</tr></table>" & _
                           "</body></html>"
End Sub
'
' Form_Unload
'
Private Sub Form_Unload(Cancel As Integer)
    Set mfrmParent = Nothing
End Sub
'
' Form_Resize()
'
' Refresh the HTMLLabel control.
'
Private Sub Form_Resize()
    ctlHTML.Refresh False
End Sub

