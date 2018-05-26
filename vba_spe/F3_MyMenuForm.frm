VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F3_MyMenuForm 
   Caption         =   "スペルと文章のチェック"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5415
   OleObjectBlob   =   "F3_MyMenuForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "F3_MyMenuForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub ComboBox1_Change()

End Sub

Private Sub CommandButton1_Click()
    allDisp
End Sub

Private Sub CommandButton2_Click()
    errorDisp
End Sub

Private Sub CommandButton3_Click()
    Unload F3_MyMenuForm 'Formを閉じる
End Sub

Private Sub Label4_Click()

End Sub

Private Sub OptionButton1_Click()

End Sub

Private Sub UserForm_Initialize()
    With ComboBox1
        .AddItem "対象ファイルと同じ階層"
        .AddItem "デスクトップ"
    End With
    
    With ComboBox2
        .AddItem "通常の文(校正用)"
        .AddItem "公用文(校正用)"
        .AddItem "くだけた文"
        .AddItem "ユーザー設定1"
        .AddItem "ユーザー設定2"
        .AddItem "ユーザー設定3"
        .AddItem "通常の文"
    End With
    
    OptionButton0.Value = True
    
End Sub

