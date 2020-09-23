VERSION 5.00
Object = "{6580F760-7819-11CF-B86C-444553540000}#1.0#0"; "EZFTP.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Please Register"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows Default
   Begin EZFTPLib.EZFTP ftp 
      Left            =   3000
      Top             =   840
      _Version        =   65536
      _ExtentX        =   800
      _ExtentY        =   800
      _StockProps     =   0
      LocalFile       =   ""
      RemoteFile      =   ""
      RemoteAddres    =   ""
      UserName        =   ""
      Password        =   ""
      Binary          =   0   'False
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   1080
      TabIndex        =   15
      Top             =   2280
      Width           =   4215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   735
      Left            =   5400
      TabIndex        =   13
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GO!,(You must be connected to the internet)"
      Height          =   735
      Left            =   5400
      TabIndex        =   12
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1080
      TabIndex        =   11
      Top             =   1920
      Width           =   4215
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1080
      TabIndex        =   10
      Top             =   1560
      Width           =   4215
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1080
      TabIndex        =   9
      Top             =   1200
      Width           =   4215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1080
      TabIndex        =   8
      Top             =   840
      Width           =   4215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Top             =   480
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label7 
      Caption         =   "E-mail:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Zip:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "State:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "City:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Address 2:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Ok guys, to put in your ftp information, you should see it
'down below, under command1_click
'ITs in big capitol letters!!
'pretty much all the important code is right below these comments
'you'll see the ini to write to, change it to wahtever
'fits your needs, then also under form load, it reads an ini
'i think you won;t have any problems, but if anything comes up
'e-mail me @ actorindp@juno.com
'
'PLEASE IF YOU FIND THIS USEFUL, VOTE AND LEAVE FEEDBACK FOR ME
'THANKS!!!
Option Explicit
Private Sub Command1_Click()
Dim ret
If Text1.Text = "" And Text2.Text = "" And Text4.Text = "" And Text5.Text = "" And Text6.Text = "" Then MsgBox "Please Fill In All Information" Else
   ret = writeini("register", "register", "yes", "c:\progconfig.ini")
   Open Text1.Text + ".txt" For Output As #1
    Print #1, Text2.Text + vbNewLine + Text3.Text + vbNewLine + Text4.Text + vbNewLine + Text5.Text + vbNewLine + Text6.Text + vbNewLine + Text7.Text
       Close #1
ftp.RemoteAddress = "FTP ADDRESS HERE"
ftp.Username = "USERNAME"
ftp.Password = "PASSWORD"
ftp.Connect
If Err <> 0 Then
        MsgBox "Unable to connect to the specified host", vbCritical
End If
ftp.Localfile = Text1.Text + ".txt"
ftp.Remotefile = Text1.Text + ".txt"
ftp.PutFile
If Err <> 0 Then
        MsgBox "Unable to register, please try again later", vbExclamation
        End If
If Err = 0 Then
        MsgBox "Thank You!", vbExclamation
        Form1.Hide
        Form2.Show
 End If
   End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
If "yes" = ReadINI("register", "register", "c:\progconfig.ini") Then
   Form1.Visible = False
   Form2.Visible = True
End If
End Sub
