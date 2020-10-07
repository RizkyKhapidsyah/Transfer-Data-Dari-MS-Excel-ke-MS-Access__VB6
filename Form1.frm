VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Transfer Data dari MS Excel ke MS Access"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6540
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   2160
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Created by Rizky Khapidsyah
'Source Code program di mulai dari sini

Private Sub Command1_Click()
On Error GoTo errHandler
Dim mExcelFile As String
Dim mAccessFile As String
Dim mWorkSheet As String
Dim mTableName As String
Dim mdatabase As Database
mExcelFile = App.Path & "\Book1.xls"
mAccessFile = App.Path & "\Db1.mdb"
mWorkSheet = "Sheet1"
mTableName = "Table1"
  'Anda bisa memakai "Excel 7.0" or 8.0 tergantung dari
  'ISAM yang terinstall di PC Anda.
  Set mdatabase = OpenDatabase(mExcelFile, True, _
                  False, "Excel 5.0")
  mdatabase.Execute "SELECT * into [;database=" & _
           mAccessFile & "]." & _
           mTableName & " FROM [" & mWorkSheet & "$]"
  MsgBox "Sukses. Buka dengan MS Access untuk melihat tabel " & mTableName
  Exit Sub
errHandler:
  If Err.Number = 3010 Then
     MsgBox mTableName & " sudah ada!" & vbCrLf & _
     "Hapus " & mTableName & _
      " dulu atau pilih nama yang lain."
  Else
     MsgBox Err.Number & " " & Err.Description
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   End
End Sub


