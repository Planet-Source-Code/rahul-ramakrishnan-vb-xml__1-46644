VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C000&
   Caption         =   "Employee Records"
   ClientHeight    =   6750
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text7 
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   1440
      TabIndex        =   23
      Top             =   1200
      Width           =   3195
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   1440
      TabIndex        =   21
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton Edit 
      Caption         =   "Edit"
      Height          =   315
      Left            =   2340
      TabIndex        =   17
      Top             =   5820
      Width           =   675
   End
   Begin VB.CommandButton Delete 
      Caption         =   "Delete"
      Height          =   315
      Left            =   3120
      TabIndex        =   14
      Top             =   5340
      Width           =   675
   End
   Begin VB.CommandButton New 
      Caption         =   "New"
      Height          =   315
      Left            =   2340
      TabIndex        =   13
      Top             =   5340
      Width           =   675
   End
   Begin VB.CommandButton Save 
      Caption         =   "Save"
      Height          =   315
      Left            =   3900
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5340
      Width           =   675
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search"
      Height          =   735
      Left            =   60
      TabIndex        =   8
      Top             =   4380
      Width           =   4635
      Begin VB.CommandButton Search 
         Caption         =   "Search"
         Height          =   315
         Left            =   3420
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   1380
         TabIndex        =   10
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "First Name"
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   300
         Width           =   1395
      End
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   1440
      TabIndex        =   7
      Top             =   5400
      Width           =   495
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   1440
      TabIndex        =   6
      Top             =   3060
      Width           =   3195
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   1440
      TabIndex        =   4
      Top             =   2640
      Width           =   3195
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   2220
      Width           =   3195
   End
   Begin VB.CommandButton Next 
      Caption         =   "Next>>"
      Height          =   255
      Left            =   3840
      TabIndex        =   0
      Top             =   3900
      Width           =   675
   End
   Begin VB.Frame Frame2 
      Caption         =   "Action"
      Height          =   1095
      Left            =   60
      TabIndex        =   15
      Top             =   5160
      Width           =   4635
      Begin VB.Label Label6 
         Caption         =   "No. Records"
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   300
         Width           =   1275
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Navigation"
      Height          =   735
      Left            =   60
      TabIndex        =   18
      Top             =   3600
      Width           =   4635
      Begin VB.CommandButton Prev 
         Caption         =   "<<Prev"
         Height          =   255
         Left            =   3000
         TabIndex        =   19
         Top             =   300
         Width           =   675
      End
      Begin VB.Label Label4 
         Caption         =   "Current Record"
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Top             =   300
         Width           =   1155
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Source"
      Height          =   1155
      Left            =   60
      TabIndex        =   24
      Top             =   900
      Width           =   4635
      Begin VB.CommandButton nset 
         Caption         =   "Set"
         Height          =   255
         Left            =   3780
         TabIndex        =   26
         Top             =   780
         Width           =   675
      End
      Begin VB.Label Label7 
         Caption         =   "Path"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   195
      Left            =   420
      TabIndex        =   22
      Top             =   1080
      Width           =   15
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C000&
      Caption         =   "City"
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      Caption         =   "Middle Name"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2700
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "First Name"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   2220
      Width           =   855
   End
   Begin VB.Menu mnucon 
      Caption         =   "Configure"
      Begin VB.Menu mnunew 
         Caption         =   "New"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim inode As New MSXML.DOMDocument
Dim Nodes As MSXML.IXMLDOMNodeList
Dim dnode As MSXML.IXMLDOMNode
Dim nnode As MSXML.IXMLDOMNode 'Delete Node
Dim rnode As MSXML.IXMLDOMNode 'Edit Node
Dim fso As New Scripting.FileSystemObject
Dim i As Integer
Dim counter As Integer
Dim ncount As Integer
Dim nname As String
Private Sub Delete_Click()
If Text5.Text = "" Then
MsgBox "Please search for the record you wish to delete"
Exit Sub
End If
Set nnode = inode.getElementsByTagName("details").Item(i)
MsgBox nnode.Text
inode.documentElement.removeChild nnode
inode.Save nname
Form_Load
End Sub
Private Sub Edit_Click()
If Text5.Text = "" Then
MsgBox "Please search for the record you wish to Edit"
Exit Sub
End If
Set rnode = inode.getElementsByTagName("details").Item(i)
Set dnode = rnode.firstChild
dnode.firstChild.nodeValue = Text1.Text
Set dnode = dnode.nextSibling
dnode.firstChild.nodeValue = Text2.Text
Set dnode = dnode.nextSibling
dnode.firstChild.nodeValue = Text3.Text
inode.Save nname
Form_Load
End Sub
Private Sub Form_Load()
counter = 0
inode.Load nname
Set Nodes = inode.getElementsByTagName("FirstName")
ncount = Nodes.length
If Nodes.length <> 0 Then
Text1.Text = inode.getElementsByTagName("FirstName").Item(counter).Text
Text2.Text = inode.getElementsByTagName("LastName").Item(counter).Text
Text3.Text = inode.getElementsByTagName("City").Item(counter).Text
Text4.Text = ncount
Text5.Text = ""
Text6.Text = counter + 1
Else
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
''MsgBox "There are no records to be viewed,please select NEW to enter records"
End If
Save.Enabled = False
Save.Visible = False
nset.Enabled = False
Text7.Enabled = False
End Sub

Private Sub mnunew_Click()
nset.Enabled = True
Text7.Enabled = True
End Sub

Private Sub New_Click()
  Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Save.Enabled = True
    Save.Visible = True
End Sub
''Module to Edit
''save-completed
Private Sub Next_Click()
If counter <> ncount - 1 Then
counter = counter + 1
Else
counter = 0
End If
Text6.Text = counter + 1
Text1.Text = inode.getElementsByTagName("FirstName").Item(counter).Text
Text2.Text = inode.getElementsByTagName("LastName").Item(counter).Text
Text3.Text = inode.getElementsByTagName("City").Item(counter).Text
End Sub

Private Sub nset_Click()
Dim nresult As Variant
Dim nstr As String
Dim mfile As TextStream
nname = Text7.Text
If fso.FileExists(nname) = False Then
nresult = MsgBox("The file does not exist,would you like to create the file", vbQuestion + vbYesNo)
If nresult = vbYes Then
fso.CreateTextFile (nname)
Set mfile = fso.OpenTextFile(nname, ForWriting)
nstr = "<xml>" & vbCrLf
nstr = nstr & "</xml>"
mfile.WriteLine nstr
MsgBox "The file was succesfully created", vbInformation
Else
Exit Sub
End If
End If
Form_Load
End Sub

Private Sub Prev_Click()
If counter > 0 Then
counter = counter - 1
Text6.Text = counter + 1
End If
Text1.Text = inode.getElementsByTagName("FirstName").Item(counter).Text
Text2.Text = inode.getElementsByTagName("LastName").Item(counter).Text
Text3.Text = inode.getElementsByTagName("City").Item(counter).Text
End Sub
Private Sub Save_Click()
    Dim NewRecord As MSXML.IXMLDOMNode
    Set NewRecord = inode.createElement("details")
    inode.documentElement.appendChild NewRecord
    Dim Detail As MSXML.IXMLDOMNode
    Dim FirstName As MSXML.IXMLDOMNode
    Dim LastName As MSXML.IXMLDOMNode
    Dim City As MSXML.IXMLDOMNode
    Set Detail = inode.documentElement
    Set FirstName = inode.createElement("FirstName")
    Set LastName = inode.createElement("LastName")
    Set City = inode.createElement("City")
    FirstName.appendChild inode.createTextNode(Text1.Text)
    LastName.appendChild inode.createTextNode(Text2.Text)
    City.appendChild inode.createTextNode(Text3.Text)
    NewRecord.appendChild FirstName
    NewRecord.appendChild LastName
    NewRecord.appendChild City
    inode.Save nname
    Form_Load
End Sub
Private Sub Search_Click()
For i = 0 To ncount - 1
If Text5.Text = inode.getElementsByTagName("FirstName").Item(i).Text Then
Text6.Text = i + 1
Text1.Text = inode.getElementsByTagName("FirstName").Item(i).Text
Text2.Text = inode.getElementsByTagName("LastName").Item(i).Text
Text3.Text = inode.getElementsByTagName("City").Item(i).Text
Exit Sub
End If
Next
MsgBox "No record match was found"
Text5.Text = ""
End Sub

