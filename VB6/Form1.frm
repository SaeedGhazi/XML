VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12600
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   12600
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   7320
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   840
      Width           =   2055
   End
   Begin VB.ListBox List6 
      Height          =   2010
      Left            =   4920
      TabIndex        =   6
      Top             =   3720
      Width           =   2175
   End
   Begin VB.ListBox List5 
      Height          =   2790
      Left            =   4920
      TabIndex        =   5
      Top             =   840
      Width           =   2295
   End
   Begin VB.ListBox List4 
      Height          =   1035
      Left            =   3120
      TabIndex        =   4
      Top             =   1920
      Width           =   1695
   End
   Begin VB.ListBox List3 
      Height          =   1035
      Left            =   3120
      TabIndex        =   3
      Top             =   840
      Width           =   1695
   End
   Begin VB.ListBox List2 
      Height          =   1035
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   2895
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xmlDoc As MSXML.DOMDocument
Dim rootNode As MSXML.IXMLDOMNode
Dim attrNode As MSXML.IXMLDOMNode




Private Sub Command1_Click()
    Set xmlDoc = New MSXML.DOMDocument
    xmlDoc.Load App.Path + "/" + "eval.xml"
    Set rootNode = xmlDoc.documentElement

    ' Populate List1 with child nodes
    Dim childNode As MSXML.IXMLDOMNode
    For Each childNode In rootNode.childNodes
        List1.AddItem childNode.nodeName
    Next
End Sub

Private Sub List1_Click()
    ' Clear List2 and List3
    List2.Clear
    List3.Clear
    ' Get the selected child node
    Set childNode = rootNode.childNodes(List1.ListIndex)
    ' List child node attributes
    Dim attrNode As MSXML.IXMLDOMNode
    For Each attrNode In childNode.Attributes
        List2.AddItem attrNode.nodeName & ": " & attrNode.nodeValue
    Next
    ' List child nodes of the selected child node
    Dim grandChildNode As MSXML.IXMLDOMNode
    For Each grandChildNode In childNode.childNodes
        List3.AddItem grandChildNode.nodeName
    Next
End Sub
Private Sub List3_Click()
    ' Clear List4
    List4.Clear
    List5.Clear
    List6.Clear
    Text1.Text = ""
    ' Get the selected child node
    Set childNode = rootNode.childNodes(List1.ListIndex).childNodes(List3.ListIndex)

    ' List child node attributes
    Dim attrNode As MSXML.IXMLDOMNode
    For Each attrNode In childNode.Attributes
        List4.AddItem attrNode.nodeName & ": " & attrNode.nodeValue
    Next
    For Each childNode In rootNode.childNodes(List1.ListIndex).childNodes(List3.ListIndex).childNodes
        List5.AddItem childNode.nodeName
    Next
End Sub

Private Sub List5_Click()
    ' Clear List6
    List6.Clear

    ' Get the selected child node
    Set childNode = rootNode.childNodes(List1.ListIndex).childNodes(List3.ListIndex).childNodes(List5.ListIndex)
    Text1.Text = rootNode.childNodes(List1.ListIndex).childNodes(List3.ListIndex).childNodes(List5.ListIndex).nodeTypedValue
    ' List child node attributes
    Dim attrNode As MSXML.IXMLDOMNode
    For Each attrNode In childNode.Attributes
        List6.AddItem attrNode.nodeName & ": " & attrNode.nodeValue
    Next
End Sub
