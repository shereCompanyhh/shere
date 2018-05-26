Attribute VB_Name = "m_SpellAndGrammarChk"
Option Explicit
'===========================================================================================
'
' スペルと文章のチェック
'
'===========================================================================================
Type cellInfoType 'セルデータと結果の格納用
    type As String 'タイプ
    rowclm As String '行_列
    text As String '文字
    textOrg As String 'オリジナルの文字
    pointout As String '指摘
End Type

'グローバル定数定義
Const divNum = 50 '分割処理の単位
Const progressMax = 70 'プログレスバーのカウントの最大値
Const proggressInitWidth = 225 'プログレスバーの上位置
Const proggressInitLeft = 5 'プログレスバーの左位置

'グローバル変数定義
Dim startTime As Single '時間計測用(開始)
Dim finishTime As Single '時間計測用(終了)
Dim oWord As Object 'Wordアプリケーション
Dim oDoc As Object 'Wordファイル
Dim objOptions As Object 'Wordオプション
Dim initSpellChkVal As Boolean '実行前のスペルチェックオプションの値
Dim progressCnt As Integer  'プログレスバーのカウント

'=============================================================================
' ユーザフォーム用
'=============================================================================
'「スペルと文章のチェック」のショートカットを選択した際に表示
Public Sub selSpellAndGramaCheck()
    Dim RowH As Long, RowL As Long '選択範囲(行上、行下)
    F3_MyMenuForm.Label1.Caption = vbNewLine & "　　　　　　スペルと文章のチェックを実行します"
    F3_MyMenuForm.CommandButton2.SetFocus '「指摘のみ表示」にフォーカス
    F3_MyMenuForm.Show vbModal
End Sub

'「全て表示」を実行時
Public Sub allDisp()
    Dim savePathBoxText As String
    Dim docStyleBoxText As String
    Dim searchAllSheet As Boolean
    savePathBoxText = F3_MyMenuForm.ComboBox1.text 'Formを閉じる前に値を保存
    docStyleBoxText = F3_MyMenuForm.ComboBox2.text 'Formを閉じる前に値を保存
    searchAllSheet = F3_MyMenuForm.OptionButton0 'Formを閉じる前に値を保存
    Unload F3_MyMenuForm 'Formを閉じる
    Call spellAndGramaCheck(True, savePathBoxText, docStyleBoxText, searchAllSheet)
End Sub

'「指摘のみ表示」を実行時
Public Sub errorDisp()
    Dim savePathBoxText As String
    Dim docStyleBoxText As String
    Dim searchAllSheet As Boolean
    savePathBoxText = F3_MyMenuForm.ComboBox1.text 'Formを閉じる前に値を保存
    docStyleBoxText = F3_MyMenuForm.ComboBox2.text 'Formを閉じる前に値を保存
    searchAllSheet = F3_MyMenuForm.OptionButton0 'Formを閉じる前に値を保存
    Unload F3_MyMenuForm 'Formを閉じる
    Call spellAndGramaCheck(False, savePathBoxText, docStyleBoxText, searchAllSheet)
End Sub


'=============================================================================
' メイン
'=============================================================================
Public Sub spellAndGramaCheck(allDispFlag As Boolean, savePathBoxText As String, docStyleBoxText As String, searchAllSheet As Boolean)
    Dim MaxRow As Long, MaxClm As Long 'アクティブな範囲(行、列)
    Dim iRow As Long, iClm As Long '行、列の変数
    Dim i As Integer, n As Integer, m As Integer, k As Integer '　ループ文用
    Dim cellInfo() As cellInfoType 'セル情報の配列
    Dim cellInfoCnt As Long 'セル情報の配列の数
    Dim tmpStr As String '一時的な文字列保持用
    Dim iniProgressCnt As Integer '分割処理に入る前の進捗値
    Dim progressApi As String '処理中アピール用
    Dim myWorkbookName As String 'マクロ実行対象のワークブック名
    Dim mySheet As Worksheet '選択シート
    
    Dim mySheetCnt As Long
    Dim mySheetName() As String
    
    On Error GoTo ErrorProcess 'エラーが発生した時は、ErrorProcess処理へ飛ぶ
    
    ' 開始時間を記録
    startTime = Timer
    '描画停止
    Application.ScreenUpdating = False
    'カーソルを砂時計にする
    Application.Cursor = xlWait
    'カウント数初期化
    cellInfoCnt = 0
    '配列の初期化
    Erase cellInfo
    'ワークブック名の取得
    myWorkbookName = ActiveWorkbook.Name
    
    '------------------
    ' 非表示シートについてもチェックするかを確認
    '------------------
    Dim checkHiddenSheetFlag As Boolean
    Dim ansToMsgbox As Integer
    checkHiddenSheetFlag = False
    
    If (searchAllSheet = True) Then
        For Each mySheet In Worksheets
            If (mySheet.Visible = xlSheetHidden) Then
                ansToMsgbox = MsgBox("非表示シートがあるようです。" & vbNewLine & _
                "非表示シートに対してもチェックを行いますか？", vbYesNo)
                If (ansToMsgbox = vbYes) Then '
                    checkHiddenSheetFlag = True
                End If
                Exit For
            End If
        Next
    End If
    
    '------------------
    'プログレスバーの設定
    '------------------
    progressCnt = 0
    F3_ProgressForm.Show vbModeless                   'プログレスバーFormを表示
          
    '------------------
    ' Wordファイルを作成
    '------------------
    progressCnt = progressCnt + 10
    Call setProgress(progressCnt, "起動中")
     
     'デスクトップのアドレスを取得
    Dim saveDirectory As String, WSH As Variant
    Set WSH = CreateObject("Wscript.Shell")
    
    '保存先の取得
    If (savePathBoxText = "デスクトップ") Then
        saveDirectory = WSH.SpecialFolders("Desktop")
    ElseIf (savePathBoxText = "対象ファイルと同じ階層") Then
        saveDirectory = ActiveWorkbook.Path
    Else
        MsgBox ("コンボボックスの値が正しくありません。(" & savePathBoxText & ")")
        Application.ScreenUpdating = True
        Application.Cursor = xlDefault 'カーソルを元に戻す
        End '誤った値であるため終了
    End If
     
    '文書のスタイルの取得
    If Not ((docStyleBoxText = "通常の文") Or _
        (docStyleBoxText = "通常の文(校正用)") Or _
        (docStyleBoxText = "公用文(校正用)") Or _
        (docStyleBoxText = "くだけた文") Or _
        (docStyleBoxText = "ユーザー設定1") Or _
        (docStyleBoxText = "ユーザー設定2") Or _
        (docStyleBoxText = "ユーザー設定3") _
        ) Then
        
        MsgBox ("コンボボックスの値が正しくありません。(" & docStyleBoxText & ")")
        Application.ScreenUpdating = True
        Application.Cursor = xlDefault 'カーソルを元に戻す
        End '誤った値であるため終了
    End If
     
    
    Dim wordFile As String 'ファイルの場所＋Wordファイル名
    Dim fileNameOnly As String 'ファイル名のみ
    fileNameOnly = "スペルと文章のチェック結果（対象ファイル：" & myWorkbookName & "）.doc"
    wordFile = saveDirectory & "\" & fileNameOnly 'OpenするWord文書ﾌｧｲﾙ名をﾊﾟｽ名付きで入れる
        
    '------------------
    'Microsoft Wordを起動
    '------------------
    Dim Task
    Set oWord = CreateObject("Word.Application")
    oWord.Visible = False '非表示
'   oWord.Visible = True '非表示になっていたWordを表示
    
    '------------------
    'アドイン実行前の文章校正の「入力時にスペルチェックを行う」オプションの設定
    '------------------
'   initSpellChkVal = objOptions.CheckSpellingAsYouType
    initSpellChkVal = True 'アドインを実行すると、オプションをONする。(OFFしていたらスペルチェックのNG箇所が見えないため)
    Set objOptions = oWord.Options 'オプションのオブジェクトを取得
    objOptions.CheckSpellingAsYouType = False '入力時のスペルチェックをオフする。文字列が多すぎる場合に警告が表示されるのを抑えるため。
    
    '------------------
    '既に同じ名前のファイルが開いていないか確認
    '------------------
    For Each Task In oWord.Tasks
        If Task.Visible = True And (InStr(Task.Name, fileNameOnly) <> 0) Then
            MsgBox "チェック結果ファイルが開いています。" & vbNewLine & _
             "チェック結果ファイルを閉じてから再度実行してください。", vbInformation
            objOptions.CheckSpellingAsYouType = initSpellChkVal '入力時のスペルチェックを元に戻す
            oWord.Quit 'Wordを閉じる
            Application.ScreenUpdating = True '描画再開
            Application.Cursor = xlDefault 'カーソルを元に戻す
            Unload F3_ProgressForm 'プログレスバーFormを閉じる
            Exit Sub
        End If
    Next
    
    '------------------
    'Wordファイル(新規)を開く
    '------------------
    On Error GoTo canNotSaveDeskpotProcess 'デスクトップに保存できない場合は、強制的に対象ファイルと同じ階層に保存する
    Set oDoc = oWord.Documents.Add
    oDoc.ActiveWritingStyle(1041) = docStyleBoxText '文書のスタイルを指定.japanese(=1041)
    oDoc.SaveAs wordFile 'wordFileのファイル名でWordファイルを保存
    If (False) Then '通常は通らず、エラーの時のみ通る
canNotSaveDeskpotProcess:
        If (savePathBoxText = "デスクトップ") Then
            MsgBox ("デスクトップに保存できないようです。" & vbNewLine & "結果の保存場所を「対象ファイルと同じ階層」に指定して再度実行してください")
        Else
            MsgBox ("対象ファイルが保存先を指定されたファイルではないようです。" & vbNewLine & "対象ファイルを保存してから再度実行してください")
        End If
        
        oWord.Quit 'Wordを閉じる
        Application.ScreenUpdating = True '描画再開
        Application.Cursor = xlDefault 'カーソルを元に戻す
        Unload F3_ProgressForm 'プログレスバーFormを閉じる
        End 'デスクトップに保存できないため終了
    End If
    On Error GoTo ErrorProcess 'エラーが発生した時は、ErrorProcess処理へ飛ぶ
    
    oWord.WindowState = 2 ' wdWindowStateMinimize(=2).ウィンドウを最小化
    Workbooks(myWorkbookName).Activate '選択シートをアクティブ。（最小化から戻れるように）
        
    '------------------
    ' 出力表のページ設定
    '------------------
    oDoc.PageSetup.LeftMargin = 30  '左余白
    oDoc.PageSetup.RightMargin = 30 '右余白
    oDoc.PageSetup.TopMargin = 50 '上余白
    oDoc.PageSetup.BottomMargin = 50  '下余白
    oDoc.Sections(1).Headers(1).Range.text = fileNameOnly & "  [文書スタイル:" & docStyleBoxText & "]" 'ヘッダーを付ける
        
    '------------------
    ' 各シートに対して処理
    '------------------
    Call setProgress(progressCnt, "対象ファイルの文字情報を取得中")
        
    For Each mySheet In Worksheets
        Dim tmpSheetName As String
        tmpSheetName = mySheet.Name
        
        If (searchAllSheet = False) Then
            tmpSheetName = ActiveSheet.Name
        End If
        
        '非表示シートの確認
        If ((Worksheets(tmpSheetName).Visible = xlSheetHidden) And (checkHiddenSheetFlag = False)) Then
            GoTo NextSheetSheet
        End If
        
        'シート名の設定
        ' ダブルコーテーションやカッコを含むシート名のために「'」で囲む
        Dim nameOfmySheet As String
        nameOfmySheet = "'" & Worksheets(tmpSheetName).Name & "'"
        
        '------------------
        ' セルの情報を取得
        '------------------
        MaxRow = Worksheets(tmpSheetName).UsedRange.Row + Worksheets(tmpSheetName).UsedRange.Rows.Count - 1 '有効範囲の最大行を取得
        For iRow = 1 To MaxRow
            MaxClm = Worksheets(tmpSheetName).Cells(iRow, Columns.Count).End(xlToLeft).Column 'iRow行の最終データ列を取得
            For iClm = 1 To MaxClm
                DoEvents
                If F3_ProgressForm.IsCancel = True Then '動作停止(Cancel)
                    Call vbaSuspend 'VBA実行中断
                End If
                
                If (Worksheets(tmpSheetName).Cells(iRow, iClm).text <> "") Then '何も記載がない場合は無視
                    ReDim Preserve cellInfo(cellInfoCnt + 1) '配列を再定義(1加算)
                    cellInfo(cellInfoCnt).rowclm = nameOfmySheet & "!" & Replace(Cells(iRow, iClm).Address, "$", "")
                    cellInfo(cellInfoCnt).type = "セル"
                    cellInfo(cellInfoCnt).text = Worksheets(tmpSheetName).Cells(iRow, iClm).text
                    cellInfoCnt = cellInfoCnt + 1
                End If
            Next
        Next
        
        '------------------
        ' 図形の情報を取得
        '------------------
        '選択しているシート上のShapes数をカウント
        Dim shapeCnt As Integer
        shapeCnt = Worksheets(tmpSheetName).Shapes.Count
        '配列strObjNameにオブジェクト名を代入
        For i = 1 To shapeCnt
            With Worksheets(tmpSheetName).Shapes(i)
                If (.type = msoComment) Then
                    'コメントの時は何もしない
                ElseIf (.type = msoGroup) Then 'グループ化された図形の場合
                    For k = 1 To .GroupItems.Count
                        DoEvents
                        If F3_ProgressForm.IsCancel = True Then '動作停止(Cancel)
                            Call vbaSuspend 'VBA実行中断
                        End If
                        
                        On Error Resume Next 'テキストを持たないシェイプの場合の対策
                        If (.GroupItems(k).DrawingObject.Characters.text = "") Then '何も記載がない場合
                            '何もしない
                        Else
                            ReDim Preserve cellInfo(cellInfoCnt + 1) '配列を再定義(1加算)
                            'シート名とセルのアドレス(絶対参照の$を削除)
                            
                            cellInfo(cellInfoCnt).rowclm = nameOfmySheet & "!" & _
                             Replace(Cells(.GroupItems(k).TopLeftCell.Row, .GroupItems(k).TopLeftCell.Column).Address, "$", "")
                             
                            cellInfo(cellInfoCnt).type = "図形"  '
                            cellInfo(cellInfoCnt).text = .GroupItems(k).DrawingObject.Characters.text  '
                            cellInfoCnt = cellInfoCnt + 1
                        End If
                    Next
                Else
                    DoEvents
                    If F3_ProgressForm.IsCancel = True Then '動作停止(Cancel)
                        Call vbaSuspend 'VBA実行中断
                    End If
                                            
                    On Error Resume Next 'テキストを持たないシェイプの場合の対策
                    If (.DrawingObject.Characters.text = "") Then '何も記載がない場合
                        '何もしない
                    Else
                        ReDim Preserve cellInfo(cellInfoCnt + 1) '配列を再定義(1加算)
                        cellInfo(cellInfoCnt).rowclm = nameOfmySheet & "!" & _
                         Replace(Cells(.TopLeftCell.Row, .TopLeftCell.Column).Address, "$", "")
                        cellInfo(cellInfoCnt).type = "図形"  '
                        cellInfo(cellInfoCnt).text = .DrawingObject.Characters.text  '
                        cellInfoCnt = cellInfoCnt + 1
                    End If
                End If
            End With
        Next i

        '------------------
        ' コメントの情報を取得
        '------------------
        Dim tempCom As comment
        For Each tempCom In Worksheets(tmpSheetName).Comments
            DoEvents
            If F3_ProgressForm.IsCancel = True Then '動作停止(Cancel)
                Call vbaSuspend 'VBA実行中断
            End If
            
            If (tempCom.text <> "") Then '何も記載がない場合は無視
                Dim ret As Long
                tmpStr = tempCom.text
                ret = InStr(1, tempCom.text, ":" & Chr(10), vbBinaryCompare) ' ":"+"改行"の左側はユーザ名の可能性あり
                If (InStr(1, Left(tempCom.text, ret), Chr(10)) = 0) Then 'ユーザ名対象の文字列に改行がある場合は0となり、ユーザ名として扱わない
                    tmpStr = "<User Name>:" & Right(tempCom.text, Len(tempCom.text) - ret)
                End If
                
                ReDim Preserve cellInfo(cellInfoCnt + 1) '配列を再定義(1加算)
                cellInfo(cellInfoCnt).rowclm = nameOfmySheet & "!" & tempCom.Parent.Address(0, 0) '
                cellInfo(cellInfoCnt).type = "コメント"  '
                cellInfo(cellInfoCnt).text = tmpStr  '
                cellInfo(cellInfoCnt).textOrg = tempCom.text '最後にユーザ名を元に戻すため
                cellInfoCnt = cellInfoCnt + 1
            End If
        Next tempCom
NextSheetSheet:  'シートが非表示だった場合はここへ飛ぶ
        
        If (searchAllSheet = False) Then
            Exit For
        End If
    
    Next
    
    '------------------
    '表に文字情報を書き込み
    '------------------
    Dim tbl
    Dim tblCnt As Integer
    Dim oRange
    
    '文字が一つもない場合は終了
    If (cellInfoCnt = 0) Then '一つも文字がない場合
        Set oRange = oDoc.Range()
        oDoc.Tables.Add oRange, 2, 4 '行、列の数を指定し、表作成
        Set tbl = oDoc.Tables(1)
        tbl.Style = "表 (格子)"
        tbl.Rows(1).Shading.BackgroundPatternColorIndex = 15 '
        For i = 1 To 4
            tbl.Cell(1, i).Range.Font.Color = RGB(255, 255, 255)
        Next
        tbl.Cell(1, 1).Range.text = "位置"
        tbl.Cell(1, 2).Range.text = "タイプ"
        tbl.Cell(1, 3).Range.text = "原文"
        tbl.Cell(1, 4).Range.text = "指摘"
        tbl.Cell(2, 1).Range.text = "ファイル内に文字がありません"
                
        Call vbaEnd(wordFile, fileNameOnly, savePathBoxText) '終了処理
    End If
    
    '------------------
    '分割処理(divNum単位で処理する)
    '------------------
    Dim tblCntMax As Integer '表をいくつに分割するかの数
    Dim tableLineNum As Long '作成する表の行数
    Dim progressWord As String '進捗表示文字列
        
    tblCntMax = Int(cellInfoCnt / divNum)
    If (cellInfoCnt Mod divNum = 0) Then
        tblCntMax = tblCntMax - 1
    End If
    
    iniProgressCnt = progressCnt
    For tblCnt = 0 To tblCntMax
        Set oRange = oDoc.Range()
        oDoc.Tables.Add oRange, divNum, 1 '行、列(=2)の数を指定し、表作成
        Set tbl = oDoc.Tables(1)
                                    
        '分割した際の処理行数
        If (((tblCnt + 1) * divNum) > cellInfoCnt) Then '
            tableLineNum = cellInfoCnt - (tblCnt * divNum)
        Else
            tableLineNum = divNum
        End If
        
        '表に文字情報を書きこむ
        For iRow = 0 To tableLineNum - 1
            DoEvents
            If F3_ProgressForm.IsCancel = True Then '動作停止(Cancel)
                Call vbaSuspend 'VBA実行中断
            End If
            tbl.Cell(iRow + 1, 1).Range.text = cellInfo(iRow + (tblCnt * divNum)).text
        Next
        
        '進捗表示
        progressCnt = iniProgressCnt + 40 * ((tblCnt + 1) / (tblCntMax + 1))
        progressWord = "スペルと文章のチェック中(" & (tblCnt + 1) & "/" & (tblCntMax + 1) & ")"
        progressApi = ""
        progressApi = progressApiGen(progressApi)
        Call setProgress(progressCnt, progressWord & " " & progressApi)
        
        '------------------
        ' 文法チェック
        '------------------
        '日本語の文法上の誤りを列挙して指摘をコメントとして追加する
        Dim rngGrammaticalError
        Dim ctl
        Dim cnt As Long
        Dim s As String
        '文法上の誤りを列挙
        For Each rngGrammaticalError In oDoc.GrammaticalErrors
            DoEvents
            If F3_ProgressForm.IsCancel = True Then '動作停止(Cancel)
                Call vbaSuspend 'VBA実行中断
            End If
                            
            '進捗表示
            progressApi = progressApiGen(progressApi)
            Call setProgress(progressCnt, progressWord & " " & progressApi)
            
'           If (rngGrammaticalError.LanguageID = 1041) Then '1041(=wdJapanese).日本語のみ
'            End If
            'Excel2016は右クリックの値を取得できないため誤字を含んでいることだけ伝える
            '設定された言語に対して文章チェックを行う
            s = "日本語誤字あり[左記参照]"
            For i = 1 To Len(rngGrammaticalError.text)
                cnt = 0 '初期化
                rngGrammaticalError.Characters(i).Select
                '指摘をCommandBarControlから取得
                For Each ctl In oWord.CommandBars("Grammar").Controls
                    DoEvents
                    If F3_ProgressForm.IsCancel = True Then '動作停止(Cancel)
                        Call vbaSuspend 'VBA実行中断
                    End If

                    '[IDが「0」のもの = 指摘]として取得
                    If ctl.ID = 0 Then
                         If cnt < 1 Then
                           s = s & vbNewLine & ctl.Caption
                         Else
                          s = s & "," & ctl.Caption
                         End If
                         cnt = cnt + 1
                    End If
                Next
                If cnt > 0 Then Exit For
            Next
            '指摘を追加
            s = Replace(s, Chr(2), " ") 'コメントに改行コードがあるため削除
            tmpStr = cellInfo(rngGrammaticalError.Rows(1).Index + (divNum * tblCnt) - 1).pointout
            If (tmpStr <> "") Then '既にコメントがあった場合は、改行
                tmpStr = tmpStr & Chr(13)
            End If
            cellInfo(rngGrammaticalError.Rows(1).Index + (divNum * tblCnt) - 1).pointout = tmpStr & s
        Next
        
        '------------------
        ' スペルチェック
        '------------------
        Dim iCount As Integer 'スペル間違いの指摘数
        Dim sErrors
        Dim suggestWord As String 'スペルの候補
        
        ' スペルチェック
        Set sErrors = oDoc.SpellingErrors
        iCount = sErrors.Count
        If iCount <> 0 Then
            For i = 1 To iCount
                DoEvents
                If F3_ProgressForm.IsCancel = True Then '動作停止(Cancel)
                    Call vbaSuspend 'VBA実行中断
                End If
                
                '進捗表示
                progressApi = progressApiGen(progressApi)
                Call setProgress(progressCnt, progressWord & " " & progressApi)
                
                suggestWord = "スペルミス？ "
                On Error Resume Next
                suggestWord = suggestWord & sErrors.Item(i).GetSpellingSuggestions.Item(1)
                tmpStr = cellInfo(sErrors.Item(i).Rows(1).Index + (divNum * tblCnt) - 1).pointout
                If (tmpStr <> "") Then '既にコメントがあった場合は、改行
                    tmpStr = tmpStr & Chr(13)
                End If
                If (Asc(tmpStr) <> 13) Then '一文字も記載されない場合、ASCIIコード:13が入る
                    cellInfo(sErrors.Item(i).Rows(1).Index + (divNum * tblCnt) - 1).pointout = tmpStr & suggestWord
                Else
                    cellInfo(sErrors.Item(i).Rows(1).Index + (divNum * tblCnt) - 1).pointout = suggestWord
                End If
            Next
        End If
        tbl.Delete
    Next tblCnt
                        
    '進捗経過記録
    iniProgressCnt = progressCnt
    
    '------------------
    ' ハイパーリンクのためのファイルパスを作成
    '------------------
    ' 絶対アドレスを相対アドレスに変換
    Dim absFileAdr As String ' 絶対アドレスのパス
    Dim refFileAdr As String ' 相対アドレスのパス
    absFileAdr = Workbooks(myWorkbookName).Path
    refFileAdr = absToRelativePath(absFileAdr, saveDirectory)
        
    '------------------
    '　モードによって表の編集法を変える
    '------------------
    If (allDispFlag = True) Then
        '-----------------------
        '全て表示の場合
        '-----------------------
        Dim roopNum As Integer '表を追加する回数
        tblCnt = 1
        roopNum = Int(cellInfoCnt / divNum)
        If (cellInfoCnt Mod divNum = 0) Then '割り切れる場合
            roopNum = roopNum - 1
        End If
        iniProgressCnt = progressCnt
        For k = 0 To roopNum
            DoEvents
            If F3_ProgressForm.IsCancel = True Then '動作停止(Cancel)
                Call vbaSuspend 'VBA実行中断
            End If
            
            If (((k + 1) * divNum) > cellInfoCnt) Then '
                tableLineNum = cellInfoCnt - (k * divNum)
            Else
                tableLineNum = divNum
            End If
            
            oWord.Selection.EndKey 6, 0   ' Unit:=wdLine, Extend:=wdExtend:ドキュメントの最終行を選択
            oDoc.Range.InsertAfter "__LINE_FOR_TABLE_UNIT__"   'divNum単位で区切る（入れないとテーブルを分けて認識しないため）
            oWord.Selection.EndKey 6, 0   ' Unit:=wdLine, Extend:=wdExtend:ドキュメントの最終行を選択
            
            '表を挿入
            Set oRange = oWord.Selection.Range
            If (k = 0) Then
                oDoc.Tables.Add oRange, tableLineNum + 1, 4  '行、列の数を指定し、表を追加. 目次分１行多め
                Set tbl = oDoc.Tables(tblCnt) '追加した表を選択
                tbl.Style = "表 (格子)"
                tbl.Shading.BackgroundPatternColorIndex = 16 '薄いグレイ
                tbl.Rows(1).Shading.BackgroundPatternColorIndex = 15 '
                tbl.Cell(1, 1).Range.text = "位置"
                tbl.Cell(1, 2).Range.text = "タイプ"
                tbl.Cell(1, 3).Range.text = "原文"
                tbl.Cell(1, 4).Range.text = "指摘"
                For i = 1 To 4
                    tbl.Cell(1, i).Range.Font.Color = RGB(255, 255, 255)
                Next
            Else
                oDoc.Tables.Add oRange, tableLineNum, 4   '行、列の数を指定し、表を追加
                Set tbl = oDoc.Tables(tblCnt) '追加した表を選択
                tbl.Style = "表 (格子)"
                tbl.Shading.BackgroundPatternColorIndex = 16 '薄いグレイ
            End If
            
            '表の幅を調整
            tbl.Columns(1).Width = 140
            tbl.Columns(2).Width = 50
            tbl.Columns(3).Width = 230
            tbl.Columns(4).Width = 130
                        
            For iRow = 0 To tableLineNum - 1
                DoEvents
                If F3_ProgressForm.IsCancel = True Then '動作停止(Cancel)
                    Call vbaSuspend 'VBA実行中断
                End If
                
                '配列の位置を指定
                Dim arrayPointer As Long
                arrayPointer = iRow + (divNum * k)
                If (iRow Mod 10 = 0) Then
                    progressCnt = iniProgressCnt + 20 * (arrayPointer / cellInfoCnt)
                    Call setProgress(progressCnt, "結果表の行の色を編集中(" & arrayPointer & "/" & cellInfoCnt & ")")
                End If
                
                '表の行を指定
                Dim tblPointer As Long
                If (k = 0) Then
                    tblPointer = iRow + 2
                Else
                    tblPointer = iRow + 1
                End If
                
                'ハイパーリンク
                If (cellInfo(iRow).rowclm <> "") Then
                    oDoc.Hyperlinks.Add Anchor:=tbl.Cell(tblPointer, 1).Range, _
                    Address:=(refFileAdr & "\" & myWorkbookName & "#" & cellInfo(arrayPointer).rowclm), _
                    SubAddress:="", _
                    ScreenTip:="", _
                    TextToDisplay:=cellInfo(arrayPointer).rowclm
                End If
        
                tbl.Cell(tblPointer, 2).Range.text = cellInfo(arrayPointer).type
                
                'コメントを元に戻す
                If (cellInfo(arrayPointer).textOrg <> "") Then ' textOrgにはコメントのオリジナル文字列しか入っていない
                    tbl.Cell(tblPointer, 3).Range.text = cellInfo(arrayPointer).textOrg
                Else
                    tbl.Cell(tblPointer, 3).Range.text = cellInfo(arrayPointer).text
                End If
                
                tbl.Cell(tblPointer, 4).Range.text = cellInfo(arrayPointer).pointout
                If (cellInfo(arrayPointer).pointout <> "") Then '指摘がある場合は背景色を白色
                    tbl.Rows(tblPointer).Shading.BackgroundPatternColorIndex = 0 ' White
                End If
            Next
            tblCnt = tblCnt + 1 'テーブル数を加算
        Next
        
        '-----------------------------
        '"LINE_FOR_TABLE_UNIT"の行を削除
        '-----------------------------
        oWord.Selection.HomeKey 6, 0   ' Unit:=wdStory, Extend:=wdMove:ドキュメントの文頭を選択
        '検索する方向を指定
        With oWord.Selection
            .Find.text = "__LINE_FOR_TABLE_UNIT__"
            .Find.Replacement.text = ""
            .Find.Forward = True
            Do While .Find.Execute
                .HomeKey 5, 0  ' Unit:=wdLine, Extend:=wdMove:ドキュメントの分頭を選択
                .EndKey 5, 1   ' Unit:=wdLine, Extend:=wdMove:ドキュメントの分頭を選択
                .Delete
            Loop
        End With
        
    Else
        '-----------------------
        ' 指摘のみ表示の場合
        '-----------------------
         Dim CommentCnt As Integer
        CommentCnt = 1
        Call setProgress(progressCnt, "結果表の不要な行を削除中")
        For iRow = 0 To cellInfoCnt - 1
            DoEvents
            If F3_ProgressForm.IsCancel = True Then '動作停止(Cancel)
                Call vbaSuspend 'VBA実行中断
            End If
            '指摘があった数を求める
            If (cellInfo(iRow).pointout <> "") Then
                CommentCnt = CommentCnt + 1
            End If
        Next
                
        If (CommentCnt = 1) Then '指摘がない場合は2行
            Set tbl = oDoc.Tables.Add(Range:=oDoc.Range, NumRows:=2, NumColumns:=4)
        Else '指摘がある場合は、CommentCnt行作成
            Set tbl = oDoc.Tables.Add(Range:=oDoc.Range, NumRows:=CommentCnt, NumColumns:=4)
        End If
        
        '表の幅を調整
        tbl.Columns(1).Width = 140
        tbl.Columns(2).Width = 50
        tbl.Columns(3).Width = 230
        tbl.Columns(4).Width = 130
        
        tbl.Style = "表 (格子)"
        tbl.Rows(1).Shading.BackgroundPatternColorIndex = 15 '濃いグレイ
        For i = 1 To 4
            tbl.Cell(1, i).Range.Font.Color = RGB(255, 255, 255)
        Next
        tbl.Cell(1, 1).Range.text = "位置"
        tbl.Cell(1, 2).Range.text = "タイプ"
        tbl.Cell(1, 3).Range.text = "原文"
        tbl.Cell(1, 4).Range.text = "指摘"
                                
        '指摘がない場合
        If (CommentCnt = 1) Then
            tbl.Cell(2, 1).Range.text = "指摘はありません"
            Call vbaEnd(wordFile, fileNameOnly, savePathBoxText) '終了処理
        End If
        
        iniProgressCnt = progressCnt
        i = 1 '書きこむ行の指定
        For iRow = 0 To cellInfoCnt
            DoEvents
            If F3_ProgressForm.IsCancel = True Then '動作停止(Cancel)
                Call vbaSuspend 'VBA実行中断
            End If
            If (iRow Mod 10 = 0) Then
                progressCnt = iniProgressCnt + 20 * ((iRow + 1) / (cellInfoCnt + 1))
                Call setProgress(progressCnt, "表に書き込み中(" & (iRow + 1) & "/" & (cellInfoCnt + 1) & ")")
            End If
            
            '指摘箇所があった箇所のみ記載
            If (cellInfo(iRow).pointout <> "") Then '
                'ハイパーリンク
                oDoc.Hyperlinks.Add Anchor:=tbl.Cell(i + 1, 1).Range, _
                Address:=(refFileAdr & "\" & myWorkbookName & "#" & cellInfo(iRow).rowclm), _
                SubAddress:="", _
                ScreenTip:="", _
                TextToDisplay:=cellInfo(iRow).rowclm
                
                tbl.Cell(i + 1, 2).Range.text = cellInfo(iRow).type
                                
                If (cellInfo(iRow).textOrg <> "") Then ' textOrgにはコメントのオリジナル文字列しか入っていない
                    tbl.Cell(i + 1, 3).Range.text = cellInfo(iRow).textOrg
                Else
                    tbl.Cell(i + 1, 3).Range.text = cellInfo(iRow).text
                End If
                
                tbl.Cell(i + 1, 4).Range.text = cellInfo(iRow).pointout
                i = i + 1
            End If
        Next
    End If
        
    Call setProgress(progressMax, "完了")
    Call vbaEnd(wordFile, fileNameOnly, savePathBoxText) '終了処理
    
    Exit Sub
ErrorProcess:
    'プログレスバーFormを閉じる
    Unload F3_ProgressForm
    '描画再開
    Application.ScreenUpdating = True
    'カーソルを元に戻す
    Application.Cursor = xlDefault
'    objOptions.CheckSpellingAsYouType = initSpellChkVal '入力時のスペルチェックを元に戻す
    MsgBox "エラー内容：" & Err.Description
End Sub

'---------------------------------------------------------------------
' 中断処理
'---------------------------------------------------------------------
Private Sub vbaSuspend()
    'ステータスバーを閉じる
    objOptions.CheckSpellingAsYouType = initSpellChkVal '入力時のスペルチェックを元に戻す
    Unload F3_ProgressForm
    oDoc.Close SaveChanges:=False
    oWord.Quit

    MsgBox "処理を中断しました。"
    Application.Cursor = xlDefault  '砂時計表示を戻す
    End '処理を停止
End Sub

'---------------------------------------------------------------------
' 終了処理
'---------------------------------------------------------------------
Private Sub vbaEnd(wordFile As String, fileNameOnly As String, saveDir As String)
    '------------------
    '　AutoFit
    '------------------
    oDoc.SelectAllEditableRanges
    oDoc.Sections(1).Range.Font.Size = 9
    oDoc.Paragraphs.DisableLineHeightGrid = True
    oWord.Selection.HomeKey 6, 0   ' Unit:=wdStory, Extend:=wdMove:ドキュメントの文頭を選択
        
    '保存
    oDoc.Save
    '描画再開
    Application.ScreenUpdating = True
    'カーソルを元に戻す
    Application.Cursor = xlDefault
    'プログレスバーFormを閉じる
    Unload F3_ProgressForm
    finishTime = Timer
    
    'MsgBox ("実行時間：" & finishTime - startTime & "秒")
    
    MsgBox ("処理が完了しました。(実行時間：" & Round((finishTime - startTime), 2) & "秒)" & vbNewLine & vbNewLine & _
     "※チェック結果を" & saveDir & "に保存しました。" & vbNewLine & _
     "※選択された文書スタイルをMS Wordに設定しました。" & vbNewLine & vbNewLine & _
     "チェック結果を開きます。")
    
    
    objOptions.CheckSpellingAsYouType = initSpellChkVal '入力時のスペルチェックを元に戻す
    oWord.Visible = True '非表示になっていたWordを表示
    oWord.WindowState = 0 'ウィンドウを標準サイズに戻す
    
    '既に同じ名前のファイルが開いていないか確認
    
    Dim Task2
    For Each Task2 In oWord.Tasks
        If Task2.Visible = True And (InStr(Task2.Name, fileNameOnly) <> 0) Then
            AppActivate Task2.Name
        End If
    Next
    
    End 'VBAを終了
End Sub

'---------------------------------------------------------------------
' プログレスバーを更新
'---------------------------------------------------------------------
Private Sub setProgress(i As Integer, msg As String)
    Dim percent As Double
    Dim percentInt As Integer
    Dim progressVal As Double
        
    percent = i / progressMax
    percentInt = Int(percent * 100)
    progressVal = proggressInitWidth * percentInt * 0.01
    
    F3_ProgressForm.Label1.Caption = percentInt & "%完了" & "  " & msg & ""
    F3_ProgressForm.Progress_FG.Width = 0.1 + proggressInitWidth - progressVal
    F3_ProgressForm.Progress_FG.Left = proggressInitLeft + progressVal
    
    '最小化していた場合は元に戻す
    If (Application.WindowState = xlMinimized) Then
        Application.ScreenUpdating = True
        Application.WindowState = xlNormal
        Application.ScreenUpdating = False
    End If
    
End Sub


'---------------------------------------------------------------------
' 絶対パスを相対パスに変換
'---------------------------------------------------------------------
Function absToRelativePath(filePath As String, currentPath As String) As String
    Dim i As Integer
    Dim arrayF() As String
    Dim arrayC() As String
    
    arrayF = Split(filePath, "\")
    arrayC = Split(currentPath, "\")
    
    'ドライブが異なる場合は絶対パスを返す
    If (arrayF(0) <> arrayC(0)) Then
        absToRelativePath = filePath
        Exit Function
    End If
    
    '最上位の階層から同じものを確認し文字列削除
    Dim sameArrayNum As Integer
    If (UBound(arrayF) - UBound(arrayC)) > 0 Then
        sameArrayNum = UBound(arrayC)
    Else
        sameArrayNum = UBound(arrayF)
    End If
    For i = 0 To sameArrayNum
        DoEvents
        If (arrayF(i) = arrayC(i)) Then
            arrayF(i) = ""
            arrayC(i) = ""
        End If
    Next
    
    '検索対象のフォルダのパスを編集
    For i = 0 To UBound(arrayF) - 1
        DoEvents
        If arrayF(i) <> "" Then
            arrayF(i) = arrayF(i) & "\"
        End If
    Next
    
    'VBA実行中のフォルダのパスを編集
    For i = 0 To UBound(arrayC)
        DoEvents
        If arrayC(i) <> "" Then
            arrayC(i) = "..\"
        End If
    Next
    absToRelativePath = Join(arrayC, "") & Join(arrayF, "")
    If (absToRelativePath = "") Then
        absToRelativePath = "." '実行VBAと同じ場合は.に変換
    End If
End Function


'---------------------------------------------------------------------
' 進捗アピール用
'---------------------------------------------------------------------
Function progressApiGen(progressApi As String) As String
    If (progressApi = "") Then
        progressApiGen = "."
    ElseIf (progressApi = ".") Then
        progressApiGen = ".."
    ElseIf (progressApi = "..") Then
        progressApiGen = "..."
    Else
        progressApiGen = ""
    End If
End Function



