' ==============================================================================
' RubyManager - Word VBA ルビ一括管理システム (Reverse Processing Edition)
' ==============================================================================
' [動作環境]   Windows版 Microsoft Word デスクトップ版専用 (Mac/Web版非対応)
' [セキュリティ] 外部ネットワーク通信なし (完全ローカル処理)
' [機能概要]   文書末尾から先頭に向かう逆順移動（Step -1）処理により、Rangeオフセットズレを
'             防ぎつつルビの追加・削除・単漢字対話学習・レイアウト補正・INI設定保存を実行します。
' 
' [アクセス制限] Alt+F8 メニューには RubyManagerSystem_RubyFirst のみが表示されます。
' ==============================================================================
' ==============================================================================
' 逆順ルビ削除処理（Step -1）
' ============================================================================

Option Explicit

' -----------------------------------------------------------------
' 【メインマクロ】※Alt+F8の実行リストにはこれのみ表示されます
' -----------------------------------------------------------------
Public Sub RubyManagerSystem_RubyFirst()
    Dim doc As Document
    Set doc = ActiveDocument
    
    Dim customDict As Object
    Set customDict = CreateObject("Scripting.Dictionary")
    
    ' プログラム内初期定義（初期値）
    Call AddYomiToDict(customDict, "副交感神経", "ふくこうかんしんけい")
    Call AddYomiToDict(customDict, "人工知能", "じんこうちのう")
    Call AddYomiToDict(customDict, "部分", "ぶぶん")
    Call AddYomiToDict(customDict, "表", "おもて")
    Call AddYomiToDict(customDict, "角", "つの")
    Call AddYomiToDict(customDict, "角", "かく")
    Call AddYomiToDict(customDict, "角", "かど")
    
    ' 1. 統合設定ファイル（ruby_config.ini）の読み込み
    Dim configPath As String
    configPath = ""
    
    Dim defMode As String: defMode = "1"
    Dim defGrade As String: defGrade = "2"
    Dim defSingle As String: defSingle = "1"
    Dim defScope As String: defScope = "1"
    Dim defLayout As String: defLayout = "0"
    Dim defRatio As String: defRatio = "0.5"
    
    If doc.Path <> "" Then
        configPath = doc.Path & "\ruby_config.ini"
        Call LoadConfigFile(configPath, customDict, defMode, defGrade, defSingle, defScope, defLayout, defRatio)
    End If
    
    ' 2. モード選択（状況依存ヘルプ対応）
    Dim modeInput As String
    Dim mode As Integer
    Do
        modeInput = InputBoxWithHelp( _
            "実行する処理の番号を入力してください：" & vbCrLf & vbCrLf & _
            "【ルビ追加】" & vbCrLf & _
            "1 ： すべてにルビを振る（全部振る）" & vbCrLf & _
            "2 ： 【指定範囲（ページ単位）】ページの最初に出た単語だけにルビを振る" & vbCrLf & _
            "3 ： 【段落単位】段落の最初に出た単語だけにルビを振る" & vbCrLf & vbCrLf & _
            "【ルビ消去】" & vbCrLf & _
            "4 ： 範囲内のルビをすべて消去する（全部取る）" & vbCrLf & _
            "5 ： 【指定範囲（ページ単位）】ページの頭を残して重複ルビを取る" & vbCrLf & _
            "6 ： 【段落単位】段落の頭を残して重複ルビを取る", _
            "1. モード選択", defMode, _
            "【モード選択のヘルプ】" & vbCrLf & vbCrLf & _
            "■ 1 ： すべて振る（ルビ追加）" & vbCrLf & _
            "出現回数に関わらず、すべての対象文字にルビを振ります。" & vbCrLf & vbCrLf & _
            "■ 2 ： ページの最初（ルビ追加）" & vbCrLf & _
            "各ページ（指定範囲）の中で、最初に出てきた単語だけにルビを振ります。" & vbCrLf & vbCrLf & _
            "■ 3 ： 段落の最初（ルビ追加）" & vbCrLf & _
            "段落が変わるごとにカウントをリセットし、各段落で最初に出た単語だけにルビを振ります。" & vbCrLf & vbCrLf & _
            "■ 4 ： 全部取る（ルビ消去）" & vbCrLf & _
            "指定範囲内の既存のルビ（フリガナ）をすべて削除します。" & vbCrLf & vbCrLf & _
            "■ 5 ： ページの頭を残して取る（ルビ消去）" & vbCrLf & _
            "各ページ（指定範囲）の中で、最初のルビだけを残し、2回目以降の重複ルビを取り除きます。" & vbCrLf & vbCrLf & _
            "■ 6 ： 段落の頭を残して取る（ルビ消去）" & vbCrLf & _
            "各段落の中で、最初のルビだけを残し、2回目以降の重複ルビを取り除きます。")
            
        If modeInput = "" Then Exit Sub
        
        If IsNumeric(modeInput) Then
            mode = CInt(modeInput)
            If mode >= 1 And mode <= 6 Then Exit Do
            MsgBox "1 ～ 6 の番号を入力してください。", vbExclamation, "入力エラー"
        Else
            MsgBox "正しい番号を入力してください。", vbExclamation, "入力エラー"
        End If
    Loop
    
    ' 各種設定の入力（ルビ追加モード 1?3 の場合のみ実行）
    Dim learningGrade As Integer
    learningGrade = 0
    Dim kanjiSingleMode As Integer
    kanjiSingleMode = 1
    
    If mode >= 1 And mode <= 3 Then
        Dim gradeInput As String
        Do
            gradeInput = InputBoxWithHelp( _
                "対象とする学習者のレベルを選択してください：" & vbCrLf & vbCrLf & _
                "1 ： 【入門】 すべての漢字 ＋ カタカナ" & vbCrLf & _
                "2 ： 【初級】 基礎漢字をスキップ （カタカナOFF）" & vbCrLf & _
                "3 ： 【中級】 日常漢字もスキップ （専門用語メイン）" & vbCrLf & _
                "4 ： 【カスタム】 登録した特定用語・熟語のみ", _
                "2. ルビのレベル選択", defGrade, _
                "【ルビレベル選択のヘルプ】" & vbCrLf & vbCrLf & _
                "■ 1 ： 入門レベル" & vbCrLf & _
                "すべての漢字に加え、カタカナにもひらがなルビを振ります。" & vbCrLf & vbCrLf & _
                "■ 2 ： 初級レベル" & vbCrLf & _
                "JLPT N5レベル（日、月、人、本、小など約100字）の基本漢字を自動スキップします。" & vbCrLf & vbCrLf & _
                "■ 3 ： 中級レベル" & vbCrLf & _
                "N5およびN4レベル（合計約300字）の常用漢字を自動スキップし、専門用語メインに絞ります。" & vbCrLf & vbCrLf & _
                "■ 4 ： カスタム" & vbCrLf & _
                "自動判定を行わず、辞書（ruby_config.ini）に登録した特定単語のみにルビを振ります。")
            If gradeInput = "" Then Exit Sub
            
            If IsNumeric(gradeInput) Then
                learningGrade = CInt(gradeInput)
                If learningGrade >= 1 And learningGrade <= 4 Then Exit Do
                MsgBox "1 ～ 4 の番号を入力してください。", vbExclamation, "入力エラー"
            Else
                MsgBox "正しい番号を入力してください。", vbExclamation, "入力エラー"
            End If
        Loop
        
        If learningGrade <= 3 Then
            Dim singleModeInput As String
            Do
                singleModeInput = InputBoxWithHelp( _
                    "漢字【1文字】に対するフリガナの判定方法を選んでください：" & vbCrLf & vbCrLf & _
                    "1 ： 自動（標準：音読み優先）" & vbCrLf & _
                    "2 ： 必ず【訓読み】にする" & vbCrLf & _
                    "3 ： 読み選択・直接入力（履歴機能付き）", _
                    "漢字1文字の読み設定", defSingle, _
                    "【1文字漢字の読み設定のヘルプ】" & vbCrLf & vbCrLf & _
                    "熟語ではない「1文字の漢字」（例：「角」「表」「重」など）の読み指定です。" & vbCrLf & vbCrLf & _
                    "■ 1 ： 自動（標準）" & vbCrLf & _
                    "Word/Excelの標準辞書にしたがって自動設定します（音読み優先）。" & vbCrLf & vbCrLf & _
                    "■ 2 ： 必ず訓読み" & vbCrLf & _
                    "1文字の漢字について、可能な限り訓読み（例：「角」→「つの」）を適用します。" & vbCrLf & vbCrLf & _
                    "■ 3 ： 読み選択・直接入力" & vbCrLf & _
                    "1文字漢字が登場するたびにダイアログが表示され、読みの選択や直接入力ができます（選択結果は自動記憶されます）。")
                If singleModeInput = "" Then Exit Sub
                
                If IsNumeric(singleModeInput) Then
                    kanjiSingleMode = CInt(singleModeInput)
                    If kanjiSingleMode >= 1 And kanjiSingleMode <= 3 Then Exit Do
                    MsgBox "1 ～ 3 の番号を入力してください。", vbExclamation, "入力エラー"
                Else
                    MsgBox "正しい番号を入力してください。", vbExclamation, "入力エラー"
                End If
            Loop
        End If
    End If
    
    Dim totalPages As Long
    Selection.EndKey Unit:=wdStory
    totalPages = Selection.Information(wdActiveEndPageNumber)
    Selection.HomeKey Unit:=wdStory
    
    Dim scopeInput As String
    Do
        scopeInput = InputBoxWithHelp( _
            "処理するページ範囲の方式を選んでください：" & vbCrLf & vbCrLf & _
            "1 ： 文書全体 （1 ～ " & totalPages & " ページ）" & vbCrLf & _
            "2 ： ページを範囲指定する", _
            "3. ページ範囲方式", defScope, _
            "【ページ範囲方式のヘルプ】" & vbCrLf & vbCrLf & _
            "■ 1 ： 文書全体" & vbCrLf & _
            "文書の最初のページから最終ページまですべて処理します。" & vbCrLf & vbCrLf & _
            "■ 2 ： ページを範囲指定する" & vbCrLf & _
            "開始ページと終了ページを指定して、特定の部分だけを処理します。")
        If scopeInput = "" Then Exit Sub
        
        If scopeInput = "1" Or scopeInput = "2" Then Exit Do
        MsgBox "1 または 2 を入力してください。", vbExclamation, "入力エラー"
    Loop
    
    Dim startPage As Long
    Dim endPage As Long
    If scopeInput = "2" Then
        Dim startInput As String
        Do
            startInput = InputBoxWithHelp( _
                "開始ページ番号を入力してください (1 ～ " & totalPages & ")：", _
                "開始ページ指定", "1", _
                "【開始ページのヘルプ】" & vbCrLf & vbCrLf & _
                "処理を開始したいページ番号を半角数字で入力してください。" & vbCrLf & _
                "（例：「3」と入力すると3ページ目から処理を開始します）")
            If startInput = "" Then Exit Sub
            
            If IsNumeric(startInput) Then
                startPage = CLng(startInput)
                If startPage >= 1 And startPage <= totalPages Then Exit Do
                MsgBox "1 ～ " & totalPages & " の範囲で入力してください。", vbExclamation, "入力エラー"
            Else
                MsgBox "半角数字で入力してください。", vbExclamation, "入力エラー"
            End If
        Loop
        
        Dim endInput As String
        Do
            endInput = InputBoxWithHelp( _
                "終了ページ番号を入力してください (" & startPage & " ～ " & totalPages & ")：", _
                "終了ページ指定", CStr(totalPages), _
                "【終了ページのヘルプ】" & vbCrLf & vbCrLf & _
                "処理を終了したいページ番号を半角数字で入力してください。" & vbCrLf & _
                "（例：「5」と入力すると指定した開始ページから5ページ目までを処理します）")
            If endInput = "" Then Exit Sub
            
            If IsNumeric(endInput) Then
                endPage = CLng(endInput)
                If endPage >= startPage And endPage <= totalPages Then Exit Do
                MsgBox startPage & " ～ " & totalPages & " の範囲で入力してください。", vbExclamation, "入力エラー"
            Else
                MsgBox "半角数字で入力してください。", vbExclamation, "入力エラー"
            End If
        Loop
    Else
        startPage = 1
        endPage = totalPages
    End If
    
    Dim layoutInput As String
    Do
        layoutInput = InputBoxWithHelp( _
            "【試験的機能】適用するレイアウト補正の番号を入力（例：13）" & vbCrLf & vbCrLf & _
            "0 ： 補正を行わない（デフォルト）" & vbCrLf & _
            "1 ： 本文の行間固定" & vbCrLf & _
            "2 ： テキストボックスの文字あふれ調整" & vbCrLf & _
            "3 ： ページ境界ズレ調整" & vbCrLf & _
            "4 ： すべての補正を適用する", _
            "4. レイアウト補正オプション", defLayout, _
            "【レイアウト補正オプションのヘルプ】" & vbCrLf & vbCrLf & _
            "ルビ追加・変更による崩れを防ぐ自動補正機能です。数字を組み合わせて指定可能です（例：12）。" & vbCrLf & vbCrLf & _
            "■ 0 ： 補正なし" & vbCrLf & _
            "■ 1 ： 本文の行間固定（ルビによる行間の広がりを抑えます）" & vbCrLf & _
            "■ 2 ： テキストボックス文字あふれ調整（枠内余白・行間を自動最適化）" & vbCrLf & _
            "■ 3 ： ページ境界ズレ調整（処理前の総ページ数を極力保持）" & vbCrLf & _
            "■ 4 ： すべて適用（1～3の全補正を適用）")
        If layoutInput = "" Then Exit Sub
        Exit Do
    Loop
    
    Dim optLineSpace As Boolean
    Dim optShape As Boolean
    Dim optPageKeep As Boolean
    If InStr(layoutInput, "4") > 0 Then
        optLineSpace = True
        optShape = True
        optPageKeep = True
    ElseIf InStr(layoutInput, "0") > 0 Then
        optLineSpace = False
        optShape = False
        optPageKeep = False
    Else
        optLineSpace = (InStr(layoutInput, "1") > 0)
        optShape = (InStr(layoutInput, "2") > 0)
        optPageKeep = (InStr(layoutInput, "3") > 0)
    End If

    Dim rubyRatio As Double
    rubyRatio = 0.5
    If mode >= 1 And mode <= 3 Then
        Dim sizeRatioInput As String
        Do
            sizeRatioInput = InputBoxWithHelp( _
                "ルビのフォントサイズ比率を入力してください（例：0.5）", _
                "5. ルビサイズ比率", defRatio, _
                "【ルビサイズ比率のヘルプ】" & vbCrLf & vbCrLf & _
                "本文に対するルビ文字の大きさの比率を設定します。" & vbCrLf & vbCrLf & _
                "・ 0.5 （デフォルト）: 本文の半分の大きさ" & vbCrLf & _
                "・ 0.4 : 小さめのルビ（すっきり表示）" & vbCrLf & _
                "・ 0.6 : 大きめのルビ（読みやすさ重視）")
            If sizeRatioInput = "" Then Exit Sub
            
            If IsNumeric(sizeRatioInput) Then
                rubyRatio = CDbl(sizeRatioInput)
                If rubyRatio > 0 And rubyRatio <= 2 Then Exit Do
                MsgBox "0.1 ～ 2.0 の範囲で入力してください。", vbExclamation, "入力エラー"
            Else
                MsgBox "数値で入力してください（例：0.5）。", vbExclamation, "入力エラー"
            End If
        Loop
    End If

    ' 設定・辞書の自動保存（ruby_config.ini へ一括保存）
    If configPath <> "" Then
        Call SaveConfigFile(configPath, customDict, CStr(mode), CStr(learningGrade), CStr(kanjiSingleMode), scopeInput, layoutInput, CStr(rubyRatio))
    End If

    ' 3. 正規表現パターンと除外リストの生成
    Dim targetKanji As Boolean
    Dim targetKatakana As Boolean
    targetKanji = False
    targetKatakana = False
    
    If mode >= 1 And mode <= 3 Then
        Select Case learningGrade
            Case 1
                targetKanji = True
                targetKatakana = True
            Case 2, 3
                targetKanji = True
                targetKatakana = False
            Case 4
                targetKanji = False
                targetKatakana = False
        End Select
    End If
    
    Dim patternStr As String
    patternStr = ""
    If targetKanji Then patternStr = "[\u4E00-\u9FFF々〇]+"
    If targetKatakana Then
        Dim kataPattern As String
        kataPattern = "[\u30A0-\u30FFーヽヾ?]+"
        If patternStr <> "" Then
            patternStr = patternStr & "|" & kataPattern
        Else
            patternStr = kataPattern
        End If
    End If
    
    Dim skipKanjiDict As Object
    Set skipKanjiDict = CreateObject("Scripting.Dictionary")
    If (mode >= 1 And mode <= 3) And (learningGrade = 2 Or learningGrade = 3) Then
        Call BuildSkipKanjiDictionary(skipKanjiDict, learningGrade)
    End If
    
    ' 4. 環境準備とExcelオブジェクト起動
    On Error Resume Next
    Application.ScreenUpdating = False
    If ActiveWindow.View.Type <> wdPrintView Then ActiveWindow.View.Type = wdPrintView
    doc.Repaginate
    
    Dim baseFontSize As Single
    baseFontSize = doc.Content.Font.Size
    If baseFontSize <= 0 Then baseFontSize = 10.5
    
    If optLineSpace Then
        With doc.Content.ParagraphFormat
            .LineSpacingRule = wdLineSpaceExactly
            .LineSpacing = baseFontSize + 8
        End With
    End If
    
    Dim oldGrammar As Boolean
    Dim oldSpelling As Boolean
    oldGrammar = Options.CheckGrammarAsYouType
    oldSpelling = Options.CheckSpellingAsYouType
    Options.CheckGrammarAsYouType = False
    Options.CheckSpellingAsYouType = False
    
    Dim ai As Object
    For Each ai In Application.COMAddIns
        If InStr(LCase(ai.ProgID), "scwordnativeaddin") > 0 Then ai.Connect = False
    Next ai
    On Error GoTo 0
    
    Dim xlApp As Object
    If mode >= 1 And mode <= 3 Then Set xlApp = CreateObject("Excel.Application")
    
    ' 5. メインループ（ページ単位処理）
    Dim pageRng As Range
    Dim currentPageNum As Long
    
    For currentPageNum = endPage To startPage Step -1
        Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=currentPageNum
        Set pageRng = Selection.Bookmarks("\Page").Range
        
        If Not pageRng Is Nothing Then
            If pageRng.Start < pageRng.End Then
                Call ProcessSinglePage(doc, pageRng, currentPageNum, mode, learningGrade, kanjiSingleMode, optShape, optPageKeep, baseFontSize, rubyRatio, patternStr, customDict, skipKanjiDict, xlApp)
            End If
        End If
    Next currentPageNum
    
    ' 6. 後処理と設定ファイルの最終保存
    If configPath <> "" Then
        Call SaveConfigFile(configPath, customDict, CStr(mode), CStr(learningGrade), CStr(kanjiSingleMode), scopeInput, layoutInput, CStr(rubyRatio))
    End If
    
    Selection.HomeKey Unit:=wdStory
    If Not xlApp Is Nothing Then
        xlApp.Quit
        Set xlApp = Nothing
    End If
    For Each ai In Application.COMAddIns
        If InStr(LCase(ai.ProgID), "scwordnativeaddin") > 0 Then ai.Connect = True
    Next ai
    Options.CheckGrammarAsYouType = oldGrammar
    Options.CheckSpellingAsYouType = oldSpelling
    Application.ScreenUpdating = True
    
    MsgBox "処理が完了しました。", vbInformation, "完了"
End Sub

' -----------------------------------------------------------------
' 【汎用コンテキスト（状況依存）ヘルプ入力ダイアログ】
' -----------------------------------------------------------------
Private Function InputBoxWithHelp(ByVal promptStr As String, ByVal titleStr As String, ByVal defaultVal As String, ByVal helpMsg As String) As String
    Dim userInput As String
    Do
        userInput = InputBox(promptStr & vbCrLf & vbCrLf & "※「?」または「H」で【この項目の詳細ヘルプ】を表示", titleStr, defaultVal)
        If userInput = "" Then
            InputBoxWithHelp = ""
            Exit Function
        End If
        
        userInput = Trim(userInput)
        If UCase(userInput) = "H" Or userInput = "？" Or userInput = "?" Then
            MsgBox helpMsg, vbInformation, "ヘルプ ： " & titleStr
        Else
            InputBoxWithHelp = userInput
            Exit Function
        End If
    Loop
End Function

' -----------------------------------------------------------------
' 【以下、Privateサブプロシージャ群】
' -----------------------------------------------------------------

Private Sub ProcessSinglePage(doc As Document, pageRng As Range, currentPageNum As Long, mode As Integer, learningGrade As Integer, kanjiSingleMode As Integer, optShape As Boolean, optPageKeep As Boolean, baseFontSize As Single, rubyRatio As Double, patternStr As String, customDict As Object, skipKanjiDict As Object, xlApp As Object)
    On Error Resume Next
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim delFields As Collection
    Set delFields = New Collection
    
    Dim shp As Object
    
    If mode >= 1 And mode <= 4 Then
        Call ClearAllRubyUnconditionallyInPage(pageRng)
        For Each shp In doc.Shapes
            If shp.Anchor.Start >= pageRng.Start And shp.Anchor.Start <= pageRng.End Then
                Call ClearAllRubyUnconditionallyInShape(shp)
                If optShape Then Call AdjustTextBoxSizeInternal(shp)
            End If
        Next shp
    End If
    
    If mode = 4 Then
        If optPageKeep Then Call KeepPageBoundaryBiDirectional(pageRng, currentPageNum, baseFontSize, True)
        Exit Sub
    End If
    
    Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=currentPageNum
    Set pageRng = Selection.Bookmarks("\Page").Range
    
    If mode = 1 Or mode = 2 Then
        Call ProcessCustomDictionaryPriority(pageRng, customDict, dict, mode, learningGrade)
    End If
    
    Call ProcessPageItemsInOrder(doc, pageRng, mode, dict, delFields, xlApp, patternStr, customDict, kanjiSingleMode, skipKanjiDict, learningGrade)
    
    ' 重複ルビ削除（モード5：指定範囲(ページ)単位 / モード6：段落単位）
    If mode = 5 Or mode = 6 Then
        Dim idx As Long
        Dim targetFld As Field
        For idx = delFields.Count To 1 Step -1
            Set targetFld = delFields(idx)
            targetFld.Select
            Selection.Range.PhoneticGuide ""
        Next idx
    End If
    
    If mode >= 1 And mode <= 3 Then Call BulkAdjustRubySizeInPage(pageRng, rubyRatio)
    
    If optShape Then
        For Each shp In doc.Shapes
            If shp.Anchor.Start >= pageRng.Start And shp.Anchor.Start <= pageRng.End Then
                Call AdjustTextBoxSizeInternal(shp)
            End If
        Next shp
    End If
    
    If optPageKeep Then Call KeepPageBoundaryBiDirectional(pageRng, currentPageNum, baseFontSize, False)
End Sub

Private Sub ProcessPageItemsInOrder(doc As Document, pageRng As Range, mode As Integer, dict As Object, delFields As Collection, xlApp As Object, patternStr As String, customDict As Object, kanjiSingleMode As Integer, skipKanjiDict As Object, learningGrade As Integer)
    On Error Resume Next
    
    Dim items As Collection
    Set items = New Collection
    
    Dim para As Paragraph
    Dim item As Object
    For Each para In pageRng.Paragraphs
        Set item = CreateObject("Scripting.Dictionary")
        Set item("Range") = para.Range
        item("Start") = para.Range.Start
        items.Add item
    Next para
    
    Dim shp As Object
    For Each shp In doc.Shapes
        If shp.Anchor.Start >= pageRng.Start And shp.Anchor.Start <= pageRng.End Then
            Call CollectShapesInternal(shp, items)
        End If
    Next shp
    
    Dim itemCount As Long
    itemCount = items.Count
    If itemCount = 0 Then Exit Sub
    
    Dim mainSequence() As Object
    ReDim mainSequence(1 To itemCount)
    Dim i As Long
    For i = 1 To itemCount
        Set mainSequence(i) = items(i)
    Next i
    
    Dim j As Long
    Dim temp As Object
    For i = 1 To itemCount - 1
        For j = i + 1 To itemCount
            If mainSequence(i)("Start") > mainSequence(j)("Start") Then
                Set temp = mainSequence(i)
                Set mainSequence(i) = mainSequence(j)
                Set mainSequence(j) = temp
            End If
        Next j
    Next i
    
    For i = 1 To itemCount
        Dim r As Range
        Set r = mainSequence(i)("Range")
        
        Dim currentDict As Object
        If mode = 3 Or mode = 6 Then
            ' 段落単位（モード3, 6）：段落ごとに初出辞書をリセット
            Set currentDict = CreateObject("Scripting.Dictionary")
            If mode = 3 Then Call ProcessCustomDictionaryPriority(r, customDict, currentDict, mode, learningGrade)
        Else
            Set currentDict = dict
        End If
        
        Call ProcessRangeByModeSafe(r, mode, currentDict, delFields, xlApp, patternStr, customDict, kanjiSingleMode, skipKanjiDict)
    Next i
End Sub

Private Sub ProcessRangeByModeSafe(rng As Range, mode As Integer, dict As Object, delFields As Collection, xlApp As Object, patternStr As String, customDict As Object, kanjiSingleMode As Integer, skipKanjiDict As Object)
    On Error Resume Next
    
    ' モード 5・6: 重複ルビフィールドの検出
    If mode = 5 Or mode = 6 Then
        Call ProcessMode5Fields(rng, dict, delFields)
    End If
    
    ' モード 1・2・3: ルビ付与処理
    If mode >= 1 And mode <= 3 Then
        If patternStr = "" Then Exit Sub
        Call ApplyRubyToRange(rng, mode, dict, xlApp, patternStr, customDict, kanjiSingleMode, skipKanjiDict)
    End If
End Sub

Private Sub ProcessMode5Fields(rng As Range, dict As Object, delFields As Collection)
    On Error Resume Next
    Dim i As Long
    Dim fld As Field
    For i = 1 To rng.Fields.Count
        Set fld = rng.Fields(i)
        If IsRubyField(fld) Then
            Dim parentText As String
            parentText = ExtractParentText(fld)
            If parentText <> "" Then
                If dict.Exists(parentText) Then
                    delFields.Add fld
                Else
                    dict.Add parentText, True
                End If
            End If
        End If
    Next i
End Sub

Private Sub ApplyRubyToRange(rng As Range, mode As Integer, dict As Object, xlApp As Object, patternStr As String, customDict As Object, kanjiSingleMode As Integer, skipKanjiDict As Object)
    On Error Resume Next
    
    Dim txt As String
    txt = rng.Text
    If txt = "" Then Exit Sub
    
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    reg.Pattern = patternStr
    reg.Global = True
    
    Dim matches As Object
    Set matches = reg.Execute(txt)
    If matches.Count = 0 Then Exit Sub
    
    Dim kataReg As Object
    Set kataReg = CreateObject("VBScript.RegExp")
    kataReg.Pattern = "^[\u30A0-\u30FFーヽヾ?]+$"
    
    Dim kanjiReg As Object
    Set kanjiReg = CreateObject("VBScript.RegExp")
    kanjiReg.Pattern = "^[\u4E00-\u9FFF〇]$"
    
    Dim m As Object
    For Each m In matches
        Dim wordText As String
        wordText = m.Value
        
        Dim isFirst As Boolean
        isFirst = Not dict.Exists(wordText)
        If (mode = 2 Or mode = 3) And Not isFirst Then GoTo NextMatch
        
        If skipKanjiDict.Count > 0 And Not kataReg.Test(wordText) Then
            If IsAllSkipKanji(wordText, skipKanjiDict) Then GoTo NextMatch
        End If
        
        Call FindAndApplyRubyWord(rng, wordText, mode, dict, xlApp, customDict, kanjiSingleMode, kataReg, kanjiReg)
NextMatch:
    Next m
End Sub

Private Function IsAllSkipKanji(wordText As String, skipKanjiDict As Object) As Boolean
    On Error Resume Next
    IsAllSkipKanji = True
    
    Dim cIdx As Long
    For cIdx = 1 To Len(wordText)
        Dim char As String
        char = Mid(wordText, cIdx, 1)
        If Not skipKanjiDict.Exists(char) And char <> "々" And char <> "〇" Then
            IsAllSkipKanji = False
            Exit Function
        End If
    Next cIdx
End Function

Private Sub FindAndApplyRubyWord(rng As Range, wordText As String, mode As Integer, dict As Object, xlApp As Object, customDict As Object, kanjiSingleMode As Integer, kataReg As Object, kanjiReg As Object)
    On Error Resume Next
    
    Dim searchRng As Range
    Set searchRng = rng.Duplicate
    
    With searchRng.Find
        .ClearFormatting
        .Text = wordText
        .MatchWildcards = False
        .Forward = True
        .Wrap = wdFindStop
        
        Do While .Execute
            If searchRng.Start < rng.Start Or searchRng.End > rng.End Then Exit Do
            
            Dim applied As Boolean
            applied = False
            
            If Not IsInsideAnyField(searchRng, rng) Then
                Dim yomi As String
                yomi = DetermineYomi(wordText, xlApp, customDict, kanjiSingleMode, searchRng, kataReg, kanjiReg)
                
                If yomi <> "" Then
                    searchRng.PhoneticGuide Text:=yomi, Alignment:=wdPhoneticGuideAlignmentOneTwoOne
                    applied = True
                    If Not dict.Exists(wordText) Then dict.Add wordText, True
                End If
            End If
            
            If (mode = 2 Or mode = 3) And applied Then Exit Do
            searchRng.Collapse wdCollapseEnd
        Loop
    End With
End Sub

Private Function DetermineYomi(wordText As String, xlApp As Object, customDict As Object, kanjiSingleMode As Integer, targetRng As Range, kataReg As Object, kanjiReg As Object) As String
    On Error Resume Next
    DetermineYomi = ""
    
    Dim customYomi As String
    customYomi = ""
    If customDict.Exists(wordText) Then
        If customDict(wordText).Count > 0 Then customYomi = customDict(wordText)(1)
    End If
    
    If kanjiReg.Test(wordText) Then
        DetermineYomi = ResolveSingleKanjiPhonetic(wordText, xlApp, kanjiSingleMode, targetRng, customDict)
    ElseIf customYomi <> "" Then
        DetermineYomi = customYomi
    ElseIf kataReg.Test(wordText) Then
        DetermineYomi = StrConv(wordText, vbHiragana)
    Else
        Dim xlYomi As String
        If Not xlApp Is Nothing Then
            xlYomi = xlApp.GetPhonetic(wordText)
            If xlYomi <> "" And xlYomi <> wordText Then
                DetermineYomi = StrConv(xlYomi, vbHiragana)
            Else
                DetermineYomi = wordText
            End If
        Else
            DetermineYomi = wordText
        End If
    End If
End Function

Private Function ResolveSingleKanjiPhonetic(wordText As String, xlApp As Object, kanjiSingleMode As Integer, targetRng As Range, ByRef customDict As Object) As String
    On Error Resume Next
    Dim onYomi As String
    Dim kunYomi As String
    Dim customYomiTop As String
    customYomiTop = ""
    
    If Not customDict Is Nothing Then
        If customDict.Exists(wordText) Then
            If customDict(wordText).Count > 0 Then customYomiTop = customDict(wordText)(1)
        End If
    End If
    
    If Not xlApp Is Nothing Then
        onYomi = xlApp.GetPhonetic(wordText)
        If onYomi <> "" And onYomi <> wordText Then onYomi = StrConv(onYomi, vbHiragana)
    End If
    If onYomi = "" Or onYomi = wordText Then onYomi = wordText
    
    If kanjiSingleMode = 1 Then
        If customYomiTop <> "" Then
            ResolveSingleKanjiPhonetic = customYomiTop
        Else
            ResolveSingleKanjiPhonetic = onYomi
        End If
        Exit Function
    End If
    
    If Not xlApp Is Nothing Then
        Dim testStr As String
        Dim rawYomi As String
        testStr = wordText & "る"
        rawYomi = xlApp.GetPhonetic(testStr)
        If rawYomi <> "" And rawYomi <> testStr Then
            rawYomi = StrConv(rawYomi, vbHiragana)
            If Right(rawYomi, 1) = "る" Then
                kunYomi = Left(rawYomi, Len(rawYomi) - 1)
            Else
                kunYomi = rawYomi
            End If
        End If
    End If
    If kunYomi = "" Or kunYomi = wordText Then kunYomi = onYomi
    
    If kanjiSingleMode = 2 Then
        If customYomiTop <> "" Then
            ResolveSingleKanjiPhonetic = customYomiTop
        Else
            ResolveSingleKanjiPhonetic = kunYomi
        End If
        Exit Function
    End If
    
    If kanjiSingleMode = 3 Then
        ResolveSingleKanjiPhonetic = PromptUserForSingleKanjiYomi(wordText, onYomi, kunYomi, targetRng, customDict)
    End If
End Function

Private Function PromptUserForSingleKanjiYomi(wordText As String, onYomi As String, kunYomi As String, targetRng As Range, ByRef customDict As Object) As String
    On Error Resume Next
    Application.ScreenUpdating = True
    If Not targetRng Is Nothing Then
        targetRng.Select
        ActiveWindow.ScrollIntoView Selection.Range, True
    End If
    
    Dim candidates As Collection
    Set candidates = New Collection
    Dim addedSet As Object
    Set addedSet = CreateObject("Scripting.Dictionary")
    
    If Not customDict Is Nothing Then
        If customDict.Exists(wordText) Then
            Dim yomiObj As Variant
            For Each yomiObj In customDict(wordText)
                Dim sItm As String
                sItm = Trim(CStr(yomiObj))
                If sItm <> "" And Not addedSet.Exists(sItm) Then
                    candidates.Add sItm
                    addedSet.Add sItm, "履歴"
                End If
            Next yomiObj
        End If
    End If
    
    If onYomi <> "" And onYomi <> wordText And Not addedSet.Exists(onYomi) Then
        candidates.Add onYomi
        addedSet.Add onYomi, "推測"
    End If
    If kunYomi <> "" And kunYomi <> wordText And Not addedSet.Exists(kunYomi) Then
        candidates.Add kunYomi
        addedSet.Add kunYomi, "推測"
    End If
    If candidates.Count = 0 Then
        candidates.Add wordText
        addedSet.Add wordText, "推測"
    End If
    
    Dim promptMsg As String
    promptMsg = "漢字【 " & wordText & " 】のフリガナ番号を選択するか、直接ルビを入力してください：" & vbCrLf & vbCrLf
    
    Dim idx As Long
    Dim tag As String
    For idx = 1 To candidates.Count
        tag = addedSet(candidates(idx))
        If tag = "履歴" And idx = 1 Then
            promptMsg = promptMsg & idx & " ： " & candidates(idx) & " （前回選択・最優先）" & vbCrLf
        ElseIf tag = "履歴" Then
            promptMsg = promptMsg & idx & " ： " & candidates(idx) & " （過去の履歴）" & vbCrLf
        Else
            promptMsg = promptMsg & idx & " ： " & candidates(idx) & " （推測）" & vbCrLf
        End If
    Next idx
    promptMsg = promptMsg & "0 ： ルビを振らない（付与スキップ）" & vbCrLf & vbCrLf & "※Enterキーのみで「1」を適用 / 「?」でヘルプを表示"
    
    Dim selectedYomi As String
    selectedYomi = ""
    Dim isDecided As Boolean
    isDecided = False
    
    Do While Not isDecided
        Dim userChoice As String
        userChoice = InputBox(promptMsg, "1文字漢字のフリガナ選択（" & wordText & "）", "1")
        userChoice = Trim(userChoice)
        
        Do While Left(userChoice, 1) = " "
            userChoice = Mid(userChoice, 2)
        Loop
        Do While Right(userChoice, 1) = " "
            userChoice = Left(userChoice, Len(userChoice) - 1)
        Loop
        
        If userChoice = "?" Or userChoice = "？" Or UCase(userChoice) = "H" Then
            MsgBox "【1文字漢字のルビ選択のヘルプ】" & vbCrLf & vbCrLf & _
                   "・提示された番号（1, 2...）を選択して「OK」を押します。" & vbCrLf & _
                   "・リストにない読みを設定したい場合は、直接ひらがなでルビを入力してください。" & vbCrLf & _
                   "・「0」を入力すると、この漢字にはルビを振りません。" & vbCrLf & _
                   "・選択・入力した読みは次回から最優先として自動記憶されます。", vbInformation, "ヘルプ ： 1文字漢字ルビ選択"
        ElseIf userChoice = "" Then
            selectedYomi = ""
            isDecided = True
        ElseIf IsNumeric(userChoice) Then
            Dim choiceNum As Long
            choiceNum = CLng(userChoice)
            If choiceNum = 0 Then
                selectedYomi = ""
                isDecided = True
            ElseIf choiceNum >= 1 And choiceNum <= candidates.Count Then
                selectedYomi = candidates(choiceNum)
                isDecided = True
            Else
                MsgBox "番号一覧から選択してください。", vbExclamation, "確認"
            End If
        Else
            If MsgBox("「 " & userChoice & " 」 をルビとして設定しますか？", vbYesNo + vbQuestion, "確認") = vbYes Then
                selectedYomi = userChoice
                isDecided = True
            End If
        End If
    Loop
    
    Application.ScreenUpdating = False
    
    If selectedYomi <> "" And Not customDict Is Nothing Then
        Call AddYomiToDict(customDict, wordText, selectedYomi)
    End If
    PromptUserForSingleKanjiYomi = selectedYomi
End Function

' -----------------------------------------------------------------
' 【統合設定・辞書ファイル（ruby_config.ini）の入出力処理】
' -----------------------------------------------------------------

Private Sub LoadConfigFile(filePath As String, customDict As Object, ByRef defMode As String, ByRef defGrade As String, ByRef defSingle As String, ByRef defScope As String, ByRef defLayout As String, ByRef defRatio As String)
    On Error Resume Next
    If filePath = "" Then Exit Sub
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then Exit Sub
    
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "UTF-8"
    stm.Open
    stm.LoadFromFile filePath
    
    Dim currentSection As String
    currentSection = ""
    
    Dim lineStr As String
    Dim parts() As String
    Do Until stm.EOS
        lineStr = Trim(stm.ReadText(-2))
        If lineStr <> "" And Left(lineStr, 1) <> "#" And Left(lineStr, 1) <> "'" Then
            If Left(lineStr, 1) = "[" And Right(lineStr, 1) = "]" Then
                currentSection = UCase(Mid(lineStr, 2, Len(lineStr) - 2))
            Else
                If currentSection = "SETTINGS" Then
                    parts = Split(lineStr, "=")
                    If UBound(parts) = 1 Then
                        Dim key As String, val As String
                        key = UCase(Trim(parts(0)))
                        val = Trim(parts(1))
                        Select Case key
                            Case "MODE": If val <> "" Then defMode = val
                            Case "GRADE": If val <> "" Then defGrade = val
                            Case "SINGLEMODE": If val <> "" Then defSingle = val
                            Case "SCOPE": If val <> "" Then defScope = val
                            Case "LAYOUT": If val <> "" Then defLayout = val
                            Case "RATIO": If val <> "" Then defRatio = val
                        End Select
                    End If
                ElseIf currentSection = "DICTIONARY" Then
                    parts = Split(lineStr, ",")
                    If UBound(parts) >= 1 Then
                        Dim wordKey As String
                        wordKey = Trim(parts(0))
                        If wordKey <> "" Then
                            Dim i As Long
                            For i = UBound(parts) To 1 Step -1
                                Dim yomiValue As String
                                yomiValue = Trim(parts(i))
                                If yomiValue <> "" Then Call AddYomiToDict(customDict, wordKey, yomiValue)
                            Next i
                        End If
                    End If
                End If
            End If
        End If
    Loop
    stm.Close
    Set stm = Nothing
    Set fso = Nothing
End Sub

Private Sub SaveConfigFile(filePath As String, customDict As Object, mode As String, grade As String, singleMode As String, scope As String, layout As String, ratio As String)
    On Error Resume Next
    If filePath = "" Then Exit Sub
    
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "UTF-8"
    stm.Open
    
    stm.WriteText "# Ruby Manager System Integrated Config & Dictionary File" & vbCrLf, 1
    stm.WriteText "[SETTINGS]" & vbCrLf, 1
    stm.WriteText "MODE=" & mode & vbCrLf, 1
    stm.WriteText "GRADE=" & grade & vbCrLf, 1
    stm.WriteText "SINGLEMODE=" & singleMode & vbCrLf, 1
    stm.WriteText "SCOPE=" & scope & vbCrLf, 1
    stm.WriteText "LAYOUT=" & layout & vbCrLf, 1
    stm.WriteText "RATIO=" & ratio & vbCrLf, 1
    stm.WriteText vbCrLf, 1
    
    stm.WriteText "[DICTIONARY]" & vbCrLf, 1
    If Not customDict Is Nothing Then
        Dim key As Variant
        For Each key In customDict.Keys
            If CStr(key) <> "" And customDict(key).Count > 0 Then
                Dim lineStr As String
                lineStr = CStr(key)
                Dim itm As Variant
                For Each itm In customDict(key)
                    lineStr = lineStr & "," & CStr(itm)
                Next itm
                stm.WriteText lineStr & vbCrLf, 1
            End If
        Next key
    End If
    
    stm.SaveToFile filePath, 2
    stm.Close
    Set stm = Nothing
End Sub

' -----------------------------------------------------------------
' 【辞書・その他レイアウト補助プロシージャ群】
' -----------------------------------------------------------------

Private Sub AddYomiToDict(dict As Object, key As String, yomi As String)
    If yomi = "" Then Exit Sub
    If Not dict.Exists(key) Then dict.Add key, New Collection
    
    Dim col As Collection
    Set col = dict(key)
    Dim i As Long
    For i = col.Count To 1 Step -1
        If col(i) = yomi Then col.Remove i
    Next i
    
    If col.Count > 0 Then
        col.Add yomi, Before:=1
    Else
        col.Add yomi
    End If
End Sub

Private Sub BuildSkipKanjiDictionary(skipDict As Object, learningGrade As Integer)
    On Error Resume Next
    Dim n5Str As String
    Dim n5Arr() As String
    Dim k As Long
    
    n5Str = "日,月,火,水,木,金,土,人,男,女,子,私,何,一,二,三,四,五,六,七,八,九,十,百,千,万,円,年,上,下,左,右,中,大,小,本,半,分,力,明,休,体,好,林,森,間,畑,岩,目,耳,手,足,雨,竹,米,貝,石,糸,花,茶,肉,文,字,物,牛,馬,鳥,魚,新,古,長,短,高,安,低,暗,多,少,行,来,帰,食,飲,見,聞,読,書,話,買,教,朝,昼,夕,夜,晩,春,夏,秋,冬,今,先,毎,時,週,算,語,生,学,校,友,父,母,兄,弟,姉,妹,国,天,気,空,白,黒,赤,青,名,外,前,後,午,早,車,駅,道,歩,店,社"
    n5Arr = Split(n5Str, ",")
    For k = 0 To UBound(n5Arr)
        skipDict(n5Arr(k)) = True
    Next k
    
    If learningGrade = 3 Then
        Dim n4Str As String
        Dim n4Arr() As String
        n4Str = "心,配,思,急,立,起,寝,働,終,始,言,知,覚,忘,考,決,住,売,貸,借,返,使,作,建,洗,走,止,動,習,勉,強,引,開,閉,去,死,集,泣,笑,泳,飛,乗,降,渡,通,過,違,遊,呼,頼,登,落,出,入,町,村,都,県,市,区,院,医,病,薬,局,写,真,館,屋,室,堂,用,意,味,注,洋,和,風,家,族,親,音,楽,歌,画,声,理,科,野,菜,果,飯,色,黄,料,紙,切,服,着,送,待,持,打,押,捨,拾,受,取,払,直,治,痛,熱,冷,温,寒,暑,遠,近,遅,速,重,軽,広,狭,弱,太,細,若,老,悪,良,同,別"
        n4Arr = Split(n4Str, ",")
        For k = 0 To UBound(n4Arr)
            skipDict(n4Arr(k)) = True
        Next k
    End If
End Sub

Private Sub ProcessCustomDictionaryPriority(targetRng As Range, customDict As Object, dict As Object, mode As Integer, learningGrade As Integer)
    On Error Resume Next
    Dim key As Variant
    Dim searchRng As Range
    
    For Each key In customDict.Keys
        Dim cWord As String
        cWord = CStr(key)
        
        Dim cYomi As String
        cYomi = ""
        If customDict(key).Count > 0 Then cYomi = customDict(key)(1)
        
        If cYomi <> "" Then
            If Len(cWord) >= 2 Or learningGrade = 4 Then
                Set searchRng = targetRng.Duplicate
                With searchRng.Find
                    .ClearFormatting
                    .Text = cWord
                    .MatchWildcards = False
                    .Forward = True
                    .Wrap = wdFindStop
                    
                    Do While .Execute
                        If searchRng.Start >= targetRng.Start And searchRng.End <= targetRng.End Then
                            Dim applied As Boolean
                            applied = False
                            If Not IsInsideAnyField(searchRng, searchRng.Paragraphs(1).Range) Then
                                searchRng.PhoneticGuide Text:=cYomi, Alignment:=wdPhoneticGuideAlignmentOneTwoOne
                                If Not dict.Exists(cWord) Then dict.Add cWord, True
                                applied = True
                            End If
                            If (mode = 2 Or mode = 3) And applied Then Exit Do
                            searchRng.Collapse wdCollapseEnd
                        Else
                            Exit Do
                        End If
                    Loop
                End With
            End If
        End If
    Next key
End Sub

Private Sub ClearAllRubyUnconditionallyInPage(pageRng As Range)
    On Error Resume Next
    Dim i As Long
    For i = pageRng.Fields.Count To 1 Step -1
        If IsRubyField(pageRng.Fields(i)) Then
            pageRng.Fields(i).Select
            Selection.Range.PhoneticGuide ""
        End If
    Next i
End Sub

Private Sub ClearAllRubyUnconditionallyInShape(shp As Object)
    On Error Resume Next
    Dim subShp As Object
    Dim i As Long
    If shp.Type = 6 Then
        For Each subShp In shp.GroupItems
            Call ClearAllRubyUnconditionallyInShape(subShp)
        Next subShp
    Else
        If shp.TextFrame.HasText <> 0 Then
            For i = shp.TextFrame.TextRange.Fields.Count To 1 Step -1
                If IsRubyField(shp.TextFrame.TextRange.Fields(i)) Then
                    shp.TextFrame.TextRange.Fields(i).Select
                    Selection.Range.PhoneticGuide ""
                End If
            Next i
        End If
    End If
End Sub

Private Sub KeepPageBoundaryBiDirectional(targetPageRng As Range, targetPageNum As Long, baseSize As Single, isDeleteMode As Boolean)
    On Error Resume Next
    Dim currentLastPage As Long
    Dim currentLineSpacing As Single
    Dim loopCount As Integer
    
    loopCount = 0
    ActiveDocument.Repaginate
    targetPageRng.Paragraphs(targetPageRng.Paragraphs.Count).Range.Select
    currentLastPage = Selection.Information(wdActiveEndPageNumber)
    
    If currentLastPage > targetPageNum And Not isDeleteMode Then
        currentLineSpacing = baseSize + 8
        Do
            If loopCount >= 5 Then Exit Do
            currentLineSpacing = currentLineSpacing - 0.5
            If currentLineSpacing < (baseSize + 3) Then Exit Do
            With targetPageRng.ParagraphFormat
                .LineSpacingRule = wdLineSpaceExactly
                .LineSpacing = currentLineSpacing
            End With
            ActiveDocument.Repaginate
            targetPageRng.Paragraphs(targetPageRng.Paragraphs.Count).Range.Select
            If Selection.Information(wdActiveEndPageNumber) <= targetPageNum Then Exit Do
            loopCount = loopCount + 1
        Loop
    ElseIf currentLastPage < targetPageNum Or isDeleteMode Then
        currentLineSpacing = baseSize + 8
        Do
            If loopCount >= 5 Then Exit Do
            ActiveDocument.Repaginate
            Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=targetPageNum + 1
            If currentLastPage < targetPageNum Then
                currentLineSpacing = currentLineSpacing + 0.5
                If currentLineSpacing > (baseSize + 14) Then Exit Do
                With targetPageRng.ParagraphFormat
                    .LineSpacingRule = wdLineSpaceExactly
                    .LineSpacing = currentLineSpacing
                End With
            Else
                Exit Do
            End If
            targetPageRng.Paragraphs(targetPageRng.Paragraphs.Count).Range.Select
            currentLastPage = Selection.Information(wdActiveEndPageNumber)
            loopCount = loopCount + 1
        Loop
    End If
End Sub

Private Sub CollectShapesInternal(shp As Object, items As Collection)
    On Error Resume Next
    Dim subShp As Object
    Dim item As Object
    If shp.Type = 6 Then
        For Each subShp In shp.GroupItems
            Call CollectShapesInternal(subShp, items)
        Next subShp
    Else
        If shp.Anchor.StoryType = 1 And shp.TextFrame.HasText <> 0 Then
            Set item = CreateObject("Scripting.Dictionary")
            Set item("Range") = shp.TextFrame.TextRange
            item("Start") = shp.Anchor.Start
            items.Add item
        End If
    End If
End Sub

Private Sub BulkAdjustRubySizeInPage(targetRng As Range, ratio As Double)
    On Error Resume Next
    Dim f As Field
    Dim codeTxt As String
    Dim pFontSize As Single
    Dim targetHps As Long
    Dim sizeReg As Object
    Set sizeReg = CreateObject("VBScript.RegExp")
    sizeReg.Pattern = "\\hps\d+"
    sizeReg.Global = True
    
    For Each f In targetRng.Fields
        If IsRubyField(f) Then
            codeTxt = f.Code.Text
            If sizeReg.Test(codeTxt) Then
                pFontSize = f.Result.Font.Size
                If pFontSize <= 0 Then pFontSize = 10.5
                targetHps = Int(pFontSize * ratio * 4)
                If targetHps < 6 Then targetHps = 8
                f.Code.Text = sizeReg.Replace(codeTxt, "\hps" & targetHps)
                f.Update
            End If
        End If
    Next f
End Sub

Private Sub AdjustTextBoxSizeInternal(shp As Object)
    On Error Resume Next
    Dim subShp As Object
    If shp.Type = 6 Then
        For Each subShp In shp.GroupItems
            Call AdjustTextBoxSizeInternal(subShp)
        Next subShp
    Else
        If shp.TextFrame.HasText <> 0 Then
            With shp.TextFrame
                .MarginTop = 0
                .MarginBottom = 0
                .MarginLeft = 2
                .MarginRight = 2
                .AutoSize = True
            End With
            Dim tFontSize As Single
            tFontSize = shp.TextFrame.TextRange.Font.Size
            If tFontSize <= 0 Then tFontSize = 10.5
            With shp.TextFrame.TextRange.ParagraphFormat
                .LineSpacingRule = wdLineSpaceExactly
                .LineSpacing = tFontSize + 7
            End With
        End If
    End If
End Sub

Private Function IsInsideAnyField(targetRng As Range, parentRng As Range) As Boolean
    On Error Resume Next
    IsInsideAnyField = False
    If parentRng.Fields.Count = 0 Then Exit Function
    
    Dim f As Field
    For Each f In parentRng.Fields
        Dim minStart As Long
        minStart = f.Code.Start
        If f.Result.Start > 0 And f.Result.Start < minStart Then minStart = f.Result.Start
        
        Dim maxEnd As Long
        maxEnd = f.Code.End
        If f.Result.End > maxEnd Then maxEnd = f.Result.End
        
        If targetRng.Start < maxEnd And targetRng.End > minStart Then
            IsInsideAnyField = True
            Exit Function
        End If
    Next f
End Function

Private Function IsRubyField(fld As Field) As Boolean
    On Error Resume Next
    IsRubyField = False
    If fld.Type = wdFieldFormula Or fld.Type = 92 Then
        If InStr(1, fld.Code.Text, "\s\up", vbTextCompare) > 0 Then IsRubyField = True
    End If
End Function

Private Function ExtractParentText(fld As Field) As String
    On Error Resume Next
    ExtractParentText = ""
    ExtractParentText = Trim(fld.Result.Text)
    If ExtractParentText <> "" Then Exit Function
    
    Dim codeTxt As String
    codeTxt = fld.Code.Text
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    reg.Pattern = ",\s*([^)]+)\)\s*$"
    Dim matches As Object
    Set matches = reg.Execute(codeTxt)
    If matches.Count > 0 Then ExtractParentText = Trim(matches(0).SubMatches(0))
End Function
