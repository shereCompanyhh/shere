VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F3_ProgressForm 
   Caption         =   "スペルと文章のチェック"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "F3_ProgressForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "F3_ProgressForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




'キャンセル処理用フラグ
Public IsCancel As Boolean

Private Sub UserForm_Initialize()
    IsCancel = False
End Sub

'キャンセルボタンクリックイベント
Private Sub ButtonCancel_Click()
    'キャンセルフラグにTrueを設定
    IsCancel = True
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        IsCancel = True
    End If
End Sub

