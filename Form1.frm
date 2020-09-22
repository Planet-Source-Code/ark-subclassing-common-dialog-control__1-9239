VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1005
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Using
Private Sub Command1_Click()
  CommonDialog1.Flags = cdlOFNHelpButton
  SetControlOnDlg Me
  CommonDialog1.ShowOpen
  Caption = CommonDialog1.filename
End Sub

Public Sub Cdlg_Init()
   MoveDialog , , 600, 400
'  CenterDialog
  ModifyCtrl ID_OK, "&Insert"
  ModifyCtrl ID_CANCEL, "&Abort"
  ModifyCtrl ID_FOLDERLABEL, "Disable"
'  ModifyCtrl ID_FOLDER, , False
  ModifyCtrl ID_READONLY, , , False
  ModifyCtrl ID_FILETEXT, "Disabled. Use listbox to select a file", False
  ModifyCtrl ID_HELP, "Need Help?", False
  MoveCtrl ID_LIST, , , 310
  MoveCtrl ID_OK, 500
  MoveCtrl ID_CANCEL, 500
  MoveCtrl ID_HELP, 500
End Sub

Public Sub Cdlg_UserAction(id As CtrlID, bCancel As Boolean, nAction As ActionType)
   Dim ret As Long
   Select Case id
      Case ID_CANCEL
         ret = MsgBox("Do you really want to exit dialog?" & vbCrLf & "Press Yes to exit dilalog or No to continue", vbYesNo)
         If ret = vbNo Then bCancel = True
      Case ID_OK
         ret = MsgBox("Are you sure you select proper file?" & vbCrLf & "Press Yes to exit dilalog or No to continue", vbYesNo)
         If ret = vbNo Then bCancel = True
      Case ID_NEWFOLDER
         ret = MsgBox("Are you sure you want to create new folder?" & vbCrLf & "Press Yes to create folder or No to Cancel", vbYesNo)
         If ret = vbNo Then bCancel = True
      Case ID_PARENTFOLDER
         ret = MsgBox("Are you sure you want to choose parent folder?" & vbCrLf & "Press Yes to go up or No to Cancel", vbYesNo)
         If ret = vbNo Then bCancel = True
      Case ID_FILETEXT
         Select Case nAction
             Case EN_CHANGE
                  MsgBox "Text changed. New text " & vbCrLf & GetCtrlText(ID_FILETEXT)
             Case EN_UPDATE
                  MsgBox "Text about to be change. New text will be" & vbCrLf & GetCtrlText(ID_FILETEXT)
         End Select
   End Select
End Sub

Private Sub Form_Load()
  Command1.Caption = "Start Dialog"
End Sub
