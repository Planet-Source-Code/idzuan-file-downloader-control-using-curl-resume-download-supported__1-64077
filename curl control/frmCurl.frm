VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "CURL File Downloader (Resume capability)"
   ClientHeight    =   945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   ScaleHeight     =   945
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
   Begin Project1.ctlCURL ctlCURL1 
      Height          =   330
      Left            =   135
      TabIndex        =   0
      Top             =   180
      Width           =   5880
      _extentx        =   10372
      _extenty        =   582
   End
   Begin VB.Label Label1 
      Caption         =   "Pleaze vote for me :), thank you..."
      Height          =   285
      Left            =   90
      TabIndex        =   1
      Top             =   1485
      Width           =   4155
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'ctlCURL1.file_url = "http://localhost:8080/easetup.exe"
'ctlCURL1.file_url = "http://download.microsoft.com/download/2/9/f/29fe7195-b4ce-4b81-94dd-559139a23409/vbcontrolsvideos.exe"
ctlCURL1.file_url = "http://download.mozilla.org/?product=thunderbird-1.5&os=win&lang=en-US"
'ctlCURL1.file_url = "http://www.planetsourcecode.com/vb/scripts/ShowZip.asp"
'ctlCURL1.post_fields = "lngWId=1&lngCodeId=43640&strZipAccessCode=tp/T436403813"
'ctlCURL1.savefileas = "easetup.exe"
'ctlCURL1.savefileas = "vbcontrolsvideos.exe"
ctlCURL1.savefileas = "mozilla.exe"
'ctlCURL1.savefileas = "vb.zip"
ctlCURL1.submit_method = GET_FORM
ctlCURL1.item_label = "vb.NET Videos"
End Sub
