VERSION 5.00
Begin VB.UserControl ctlCURL 
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5880
   ScaleHeight     =   330
   ScaleWidth      =   5880
   ToolboxBitmap   =   "ctlCURL.ctx":0000
   Begin VB.Timer ProgressTimer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5895
      Top             =   0
   End
   Begin VB.CommandButton cmdPauseResume 
      Caption         =   "Pause"
      Height          =   240
      Left            =   5040
      TabIndex        =   3
      Top             =   45
      Width           =   780
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   240
      Left            =   4185
      TabIndex        =   2
      Top             =   45
      Width           =   780
   End
   Begin VB.PictureBox PBar 
      FillColor       =   &H00FFC0C0&
      Height          =   240
      Left            =   1395
      ScaleHeight     =   180
      ScaleWidth      =   1800
      TabIndex        =   0
      Top             =   45
      Width           =   1860
   End
   Begin VB.Label lblKbps 
      Caption         =   "kb/s"
      Height          =   240
      Left            =   3285
      TabIndex        =   4
      Top             =   45
      Width           =   825
   End
   Begin VB.Label lblDownloadItem 
      Height          =   240
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   1275
   End
End
Attribute VB_Name = "ctlCURL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Enum submitmethodtype
    GET_FORM = 0
    POST_FORM = 1
End Enum

Public Property Let item_label(ByVal itemlabel As String)
 m_itemlabel = itemlabel
 lblDownloadItem.Caption = m_itemlabel
End Property
Public Property Let file_url(ByVal fileurl As String)
 m_file_url = fileurl
End Property
Public Property Let savefileas(ByVal savefilename As String)
 m_savefileas = savefilename
End Property
Public Property Let proxy_url(ByVal proxyurl As String)
 m_proxy_url = proxyurl
End Property
Public Property Let post_fields(ByVal postfields As String)
 m_post_fields = postfields
End Property
Public Property Let submit_method(ByVal submitmethod As submitmethodtype)
 m_use_post = submitmethod
End Property
Public Property Get submit_method() As submitmethodtype
 submit_method = m_use_post
End Property
Private Sub cmdPauseResume_Click()
 ProgressTimer.Enabled = False
 stopcurl = True
End Sub
Private Sub cmdStart_Click()
 ProgressTimer.Enabled = True
 stopcurl = False
 EasyGet
End Sub
Private Sub ProgressTimer_Timer()
SetStatus PBar, m_progress
lblKbps.Caption = m_kbps & "kb/s"
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "SubmitMethod", m_use_post, "0"
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
m_use_post = PropBag.ReadProperty("SubmitMethod", "0")
End Sub

Private Sub SetStatus(Progressbar As Object, Percent As Integer, Optional Style As Integer, Optional Style2 As Integer)

    Progressbar.AutoRedraw = True
    Progressbar.Cls
    Progressbar.FontTransparent = True
    Progressbar.Tag = Percent
    Progressbar.ScaleWidth = 100
    Progressbar.ScaleHeight = 10
    Progressbar.DrawStyle = Style2
    Progressbar.DrawMode = 13
    Progressbar.FillStyle = Style
    Progressbar.Line (0, 0)-(Percent, Progressbar.ScaleHeight - 1), , BF
    Progressbar.Line (0, 0)-(Percent, Progressbar.ScaleHeight - 1), , B
    Progressbar.FontTransparent = False
    Progressbar.CurrentX = 58 - Progressbar.TextWidth(Percent & "%")
    Progressbar.CurrentY = (Progressbar.ScaleHeight / 2) - (Progressbar.TextHeight(Percent & "%") / 2)
    Progressbar.FontBold = True
    Progressbar.FontSize = 7
    Progressbar.FontName = "Tahoma"
    Progressbar.Print " " & Percent & "% "
    
End Sub

Private Sub UserControl_Resize()
UserControl.Width = 5880
UserControl.Height = 330
End Sub
Private Sub UserControl_Initialize()
'm_file_url = "http://download.microsoft.com/download/2/9/f/29fe7195-b4ce-4b81-94dd-559139a23409/vbcontrolsvideos.exe"
'm_savefileas = App.Path & "\temp.exe"
'lblDownloadItem.Caption = "Videos"
'm_use_post = 0
End Sub
