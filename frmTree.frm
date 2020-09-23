VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTree 
   Caption         =   "Treeview example"
   ClientHeight    =   2850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   ScaleHeight     =   2850
   ScaleWidth      =   5970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   3720
      TabIndex        =   4
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtAdd 
      Height          =   285
      Left            =   3720
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   4683
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "frmTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'the basic code with treeviews is
'TreeView1.Nodes.Add((input another nodes key), twvChild, (a key for the new node), (a label for it))
'i had a small problem with it, but that might be because english isn`t my 1st language... =)

'i look at treeviews as an advanced listbox... very useful when i understood how to use them... =D




Private Sub cmdAdd_Click()
'checks if the textbox is empty.. =D
If txtAdd.Text = "" Then Exit Sub
'finds out if there are any "nodes"...
If TreeView1.Nodes.Count > 0 Then
'checks if any node has the same name, using the same as search
    For i = 1 To TreeView1.Nodes.Count
    If TreeView1.Nodes.Item(i).Key = txtAdd.Text Or TreeView1.Nodes.Item(i).Text = txtAdd.Text Then
'a node with the same name or key is found..
        MsgBox "The name you entered has allready been used..", vbOKOnly, "Error"
        txtAdd.Text = ""
        txtAdd.SetFocus
        Exit Sub
    End If
    Next
'ask if u want to add the new node under a exiting one (the one selected)
    answer = MsgBox("Do you want to add the new node under """ & (TreeView1.SelectedItem.Text) & """?", vbYesNoCancel, "Adding node")
'remember "treeview1.selecteditem", very useful.. you won`t get far without it.. lol
'If answer if yes:
    If answer = vbYes Then
        Set anode = TreeView1.Nodes.Add((TreeView1.SelectedItem.Key), tvwChild, txtAdd.Text, txtAdd.Text)
'twvchild tells u that it is "under" the selected item
'selects the new item, sets focus to treeview and returns focus to the textbox:
        TreeView1.Nodes.Item(TreeView1.Nodes.Count).Selected = True
        TreeView1.SetFocus
        txtAdd.SetFocus
'If answer if no:
    ElseIf answer = vbNo Then
' adds a new node without any parent
        Set anode = TreeView1.Nodes.Add(, , txtAdd.Text, txtAdd.Text)
'the "relative" & "relationship" field is left blank
'selects the new item, sets focus to treeview and returns focus to textbox
        TreeView1.Nodes.Item(TreeView1.Nodes.Count).Selected = True
        TreeView1.SetFocus
        txtAdd.SetFocus
'if answer = cancel:
    ElseIf answer = vbCancel Then
        Exit Sub
    End If
Else
'if there are no nodes, this will make a new one...
    Set anode = TreeView1.Nodes.Add(, , txtAdd.Text, txtAdd.Text)
    TreeView1.Nodes.Item(TreeView1.Nodes.Count).Selected = True
End If
End Sub

Private Sub cmdDelete_Click()
'removes the selected node if "yes" is pressed
If MsgBox("Do you want to delete node """ & (TreeView1.SelectedItem.Text) & """ ?", vbYesNo, "Remove") = vbYes Then
TreeView1.Nodes.Remove (TreeView1.SelectedItem.Index)
End If
End Sub

Private Sub cmdSearch_Click()
'To search trough a treeview i simply use "for"... very easy i think..
'looks at the search textbox:
If txtSearch.Text = "" Then Exit Sub
'starts search, by looking at all the nodes with "for"
For i = 1 To TreeView1.Nodes.Count
'checks the name and key of the node:
If TreeView1.Nodes.Item(i).Text = txtSearch.Text Or TreeView1.Nodes.Item(i).Key = txtSearch.Text Then
'If it`s the same, it will select it and set focus to treeview
TreeView1.Nodes.Item(i).Selected = True
TreeView1.SetFocus
'shows a msgbox
MsgBox "Found it!", vbOKOnly, "Search completed"
Exit Sub
End If
Next
'If it wasn`t found it will go trough the "for" and come here:
MsgBox "Nothing found", vbOKOnly, "Search completed"
End Sub

Private Sub Form_Load()
Dim anode As Node
End Sub

