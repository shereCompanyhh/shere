VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    Dim oCommandBar As CommandBar
    
    'Excel2007の場合はリボンではなくコマンドバーズで設定
    If (Application.Version <= 12) Then
        For Each oCommandBar In Application.CommandBars
            If oCommandBar.Name = "Cell" Then
                'コマンドバーの取得
                Set cmdBr = Application.CommandBars("cell")
            
                'コマンドバー項目の作成
                ' (1)改行数指定
                Set cmdBtn1 = oCommandBar.Controls.Add(msoControlButton, , , , True)
                
                'コマンドバー項目に属性を追加
                With cmdBtn1
                    '呼び出すサブプロシージャ名
                    .OnAction = "m_SpellAndGrammarChk.selSpellAndGramaCheck"
                    '項目の表示名
                    .Caption = "スペルと文章のチェック"
                    .TooltipText = "" 'ヒントは表示しない
                End With
            End If
        Next
    End If
End Sub
