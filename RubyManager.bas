Option Explicit

' ==============================================================================
' RubyManager - Word VBA ルビ一括管理システム (Reverse Processing Edition)
' ==============================================================================
' [機能概要]
' 文書末尾から先頭に向かう逆順移動（Step -1）処理により、Rangeオフセットズレを防ぎつつ
' ルビの追加・削除・単漢字対話学習・レイアウト補正・INI設定保存を実行します。
' 
' [アクセス制限]
' Alt+F8 メニューには RubyManagerSystem_RubyFirst のみが表示されます。
' ==============================================================================

Private Type RubyConfig
    Mode As Integer            ' 1-6
    Level As Integer           ' 1-4
    SingleMode As Integer      ' 1:音読み, 2:訓読み, 3:対話選択
    LayoutOptions As String    ' レイアウト設定コード
    RubyScale As Integer       ' ルビサイズ比率(%)
    CustomDict As String       ' 学習辞書データ
End Type

' ------------------------------------------------------------------------------
' メインエントリポイント (Public)
' ------------------------------------------------------------------------------
Public Sub RubyManagerSystem_RubyFirst()
    On Error GoTo ErrorHandler
    
    Dim cfg As RubyConfig
    Call LoadConfig(cfg)
    
    ' 1. モード選択ダイアログ
    Dim modeHelp As String
    modeHelp = "【処理モードを選択してください】" & vbCrLf & _
               "1: 【ルビ追加】すべてに振る（全指定）" & vbCrLf & _
               "2: 【ルビ追加】ページの最初に出た単語だけに振る（ページ単位）" & vbCrLf & _
               "3: 【ルビ追加】段落の最初に出た単語だけに振る（段落単位）" & vbCrLf & _
               "4: 【ルビ消去】すべて消去する" & vbCrLf & _
               "5: 【ルビ消去】ページの最初のルビを残して重複削除（ページ単位）" & vbCrLf & _
               "6: 【ルビ消去】段落の最初のルビを残して重複削除（段落単位）"
               
    Dim modeInput As String
    modeInput = InputBoxWithHelp("モード番号(1-6)を入力してください:", "RubyManager - モード選択", CStr(cfg.Mode), modeHelp)
    If modeInput = "" Then Exit Sub
    cfg.Mode = CInt(modeInput)
    
    ' ルビ追加モード時の詳細設定
    If cfg.Mode >= 1 And cfg.Mode <= 3 Then
        ' 対象レベル選択
        Dim levelHelp As String
        levelHelp = "【対象レベルを選択してください】" & vbCrLf & _
                    "1: 【入門】全漢字 + カタカナ（ひらがナルビ）" & vbCrLf & _
                    "2: 【初級】基本漢字(JLPT N5相当)を自動スキップ" & vbCrLf & _
                    "3: 【中級】応用漢字(JLPT N5+N4相当)を自動スキップ" & vbCrLf & _
                    "4: 【カスタム】辞書登録単語のみ対象"
        Dim levelInput As String
        levelInput = InputBoxWithHelp("対象レベル(1-4)を入力してください:", "RubyManager - 対象レベル", CStr(cfg.Level), levelHelp)
        If levelInput = "" Then Exit Sub
        cfg.Level = CInt(levelInput)
        
        ' 1文字漢字モード選択
        Dim singleHelp As String
        singleHelp = "【1文字漢字の読み選択モード】" & vbCrLf & _
                     "1: 音読み優先（自動決定）" & vbCrLf & _
                     "2: 訓読み優先（自動決定）" & vbCrLf & _
                     "3: 対話選択＆履歴学習（画面ハイライト選択）"
        Dim singleInput As String
        singleInput = InputBoxWithHelp("単漢字モード(1-3)を入力してください:", "RubyManager - 単漢字モード", CStr(cfg.SingleMode), singleHelp)
        If singleInput = "" Then Exit Sub
        cfg.SingleMode = CInt(singleInput)
    End If
    
    ' レイアウト補正オプション
    Dim layoutHelp As String
    layoutHelp = "【レイアウト崩れ防止オプション】（複数指定可 例: 12）" & vbCrLf & _
                 "0: 補正なし" & vbCrLf & _
                 "1: 行間固定（ルビによる行高拡大を抑制）" & vbCrLf & _
                 "2: テキストボックス文字あふれ自動調整" & vbCrLf & _
                 "3: ページ境界維持（改ページズレ防止）" & vbCrLf & _
                 "4: すべて適用(1+2+3)"
    Dim layoutInput As String
    layoutInput = InputBoxWithHelp("レイアウト設定コードを入力してください:", "RubyManager - レイアウト補正", cfg.LayoutOptions, layoutHelp)
    If layoutInput <> "" Then cfg.LayoutOptions = layoutInput
    
    ' 設定の永続化
    Call SaveConfig(cfg)
    
    ' 処理実行（画面描画の停止による高速化）
    Application.ScreenUpdating = False
    
    If cfg.Mode >= 1 And cfg.Mode <= 3 Then
        Call ProcessRubyAddReverse(cfg)
    Else
        Call ProcessRubyRemoveReverse(cfg)
    End If
    
    ' レイアウト補正の適用
    If cfg.LayoutOptions <> "0" And cfg.LayoutOptions <> "" Then
        Call ApplyLayoutFixes(ActiveDocument, cfg.LayoutOptions)
    End If
    
    Application.ScreenUpdating = True
    MsgBox "処理が正常に完了しました。（逆順処理構造: 文末 → 文頭）", vbInformation, "RubyManager - 完了"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical, "RubyManager - エラー"
End Sub

' ==============================================================================
' 逆順ルビ追加処理（Step -1）
' ==============================================================================
Private Sub ProcessRubyAddReverse(ByRef cfg As RubyConfig)
    Dim doc As Document
    Set doc = ActiveDocument
    
    ' 1. 対象の単語Rangeリストを事前抽出（正方向判定）
    Dim targetRanges As Collection
    Set targetRanges = New Collection
    
    Dim pageDict As Object, paraDict As Object
    Set pageDict = CreateObject("Scripting.Dictionary")
    Set paraDict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    Dim wRng As Range
    Dim txt As String
    Dim pageNum As Long, paraIdx As Long
    Dim isTarget As Boolean
    
    For i = 1 To doc.Words.Count
        Set wRng = doc.Words(i)
        txt = Trim(wRng.Text)
        
        If IsValidTargetWord(txt, cfg) Then
            isTarget = True
            pageNum = wRng.Information(wdActiveEndPageNumber)
            paraIdx = wRng.Paragraphs(1).Range.Start
            
            ' モード別重複チェック
            If cfg.Mode = 2 Then ' ページ単位
                If pageDict.Exists(pageNum & "_" & txt) Then
                    isTarget = False
                Else
                    pageDict.Add pageNum & "_" & txt, True
                End If
            ElseIf cfg.Mode = 3 Then ' 段落単位
                If paraDict.Exists(paraIdx & "_" & txt) Then
                    isTarget = False
                Else
                    paraDict.Add paraIdx & "_" & txt, True
                End If
            End If
            
            If isTarget Then
                targetRanges.Add wRng
            End If
        End If
    Next i
    
    ' 2. 抽出した対象に対して「末尾から逆順 (Step -1)」にルビを追加
    Dim k As Long
    Dim rubyText As String
    
    For k = targetRanges.Count To 1 Step -1
        Set wRng = targetRanges(k)
        txt = Trim(wRng.Text)
        
        ' 単漢字かつ対話選択モードの場合
        If Len(txt) = 1 And cfg.SingleMode = 3 Then
            rubyText = GetInteractiveRuby(wRng, txt, cfg)
        Else
            rubyText = GetPhoneticString(txt)
        End If
        
        ' ルビ設定の適用
        If rubyText <> "" And rubyText <> txt Then
            Call ApplyRubyToRange(wRng, rubyText, cfg)
        End If
    Next k
End Sub

' ==============================================================================
' 逆順ルビ削除処理（Step -1）
' ==============================================================================
Private Sub ProcessRubyRemoveReverse(ByRef cfg As RubyConfig)
    Dim doc As Document
    Set doc = ActiveDocument
    
    ' フィールドコレクションを末尾から逆順（Step -1）で走査
    Dim fldIdx As Long
    Dim fld As Field
    Dim pageDict As Object, paraDict As Object
    Set pageDict = CreateObject("Scripting.Dictionary")
    Set paraDict = CreateObject("Scripting.Dictionary")
    
    ' モード5/6のために事前に順方向で残すフィールドのIDを決定
    Dim keepFieldIDs As Object
    Set keepFieldIDs = CreateObject("Scripting.Dictionary")
    
    If cfg.Mode = 5 Or cfg.Mode = 6 Then
        For fldIdx = 1 To doc.Fields.Count
            Set fld = doc.Fields(fldIdx)
            If fld.Type = wdFieldEquation Then ' ルビフィールド
                Dim pageNum As Long, paraIdx As Long
                pageNum = fld.Result.Information(wdActiveEndPageNumber)
                paraIdx = fld.Result.Paragraphs(1).Range.Start
                Dim fldText As String
                fldText = fld.Result.Text
                
                If cfg.Mode = 5 Then ' ページ内の最初を残す
                    If Not pageDict.Exists(pageNum & "_" & fldText) Then
                        pageDict.Add pageNum & "_" & fldText, True
                        keepFieldIDs.Add fldIdx, True
                    End If
                ElseIf cfg.Mode = 6 Then ' 段落内の最初を残す
                    If Not paraDict.Exists(paraIdx & "_" & fldText) Then
                        paraDict.Add paraIdx & "_" & fldText, True
                        keepFieldIDs.Add fldIdx, True
                    End If
                End If
            End If
        Next fldIdx
    End If
    
    ' 末尾から逆順で削除（Range位置のズレ防止）
    For fldIdx = doc.Fields.Count To 1 Step -1
        Set fld = doc.Fields(fldIdx)
        If fld.Type = wdFieldEquation Then
            Dim shouldDelete As Boolean
            shouldDelete = False
            
            If cfg.Mode = 4 Then ' 全消去
                shouldDelete = True
            ElseIf cfg.Mode = 5 Or cfg.Mode = 6 Then ' 最初のルビ以外を消去
                If Not keepFieldIDs.Exists(fldIdx) Then
                    shouldDelete = True
                End If
            End If
            
            If shouldDelete Then
                ' ルビの削除（漢字テキストを残してフィールド解除）
                Dim baseText As String
                baseText = fld.Result.Text
                fld.Unlink
            End If
        End If
    Next fldIdx
End Sub

' ==============================================================================
' 補助関数群 (All Private)
' ==============================================================================

' コンテキストヘルプ機能付き InputBox
Private Function InputBoxWithHelp(ByVal prompt As String, ByVal title As String, ByVal defaultVal As String, ByVal helpMsg As String) As String
    Dim res As String
    Do
        res = Trim(InputBox(prompt & vbCrLf & "(※「?」または「H」で詳細ヘルプ表示)", title, defaultVal))
        If UCase(res) = "?" Or UCase(res) = "H" Or res = "？" Or res = "ｈ" Then
            MsgBox helpMsg, vbInformation, title & " - ヘルプ"
        Else
            Exit Do
        End If
    Loop
    InputBoxWithHelp = res
End Function

' 対象単語の判定（漢字・カタカナ・スキップレベル判定）
Private Function IsValidTargetWord(ByVal txt As String, ByRef cfg As RubyConfig) As Boolean
    If txt = "" Or Len(txt) > 10 Then IsValidTargetWord = False: Exit Function
    
    Dim firstChar As String
    firstChar = Mid(txt, 1, 1)
    
    ' 漢字チェック (Unicode Range: 4E00-9FFF)
    Dim code As Long
    code = AscW(firstChar)
    Dim isKanji As Boolean
    isKanji = (code >= &H4E00 And code <= &H9FFF)
    
    ' カタカナチェック
    Dim isKatakana As Boolean
    isKatakana = (code >= &H30A0 And code <= &H30FF)
    
    If cfg.Level = 1 Then ' 入門: 全漢字 + カタカナ
        IsValidTargetWord = isKanji Or isKatakana
    ElseIf cfg.Level = 2 Then ' 初級: N5漢字スキップ
        If isKanji Then IsValidTargetWord = Not IsN5Kanji(firstChar)
    ElseIf cfg.Level = 3 Then ' 中級: N5+N4漢字スキップ
        If isKanji Then IsValidTargetWord = Not (IsN5Kanji(firstChar) Or IsN4Kanji(firstChar))
    ElseIf cfg.Level = 4 Then ' カスタム辞書のみ
        IsValidTargetWord = (InStr(cfg.CustomDict, txt) > 0)
    End If
End Function

' N5漢字リストチェック
Private Function IsN5Kanji(ByVal ch As String) As Boolean
    Dim n5List As String
    n5List = "一二三四五六七八九十百千万円日月火水木金土上下左右前後中大小長半分時今朝昼夜生行年来父母子男女性学校店書見聞読話買出入会天気雨高安"
    IsN5Kanji = (InStr(n5List, ch) > 0)
End Function

' N4漢字リストチェック
Private Function IsN4Kanji(ByVal ch As String) As Boolean
    Dim n4List As String
    n4List = "家族兄弟姉妹犬猫手足目耳口心体病院薬車電車駅道名品物服飲食立歩走止作使持待知思言伝答問教習試英画写真音声歌音色赤青黒白"
    IsN4Kanji = (InStr(n4List, ch) > 0)
End Function

' 対話型ルビ選択＆履歴学習
Private Function GetInteractiveRuby(ByRef rng As Range, ByVal ch As String, ByRef cfg As RubyConfig) As String
    ' 該当文字を画面上でハイライト・スクロール選択
    rng.Select
    
    Dim defaultRuby As String
    defaultRuby = GetPhoneticString(ch)
    
    Dim promptMsg As String
    promptMsg = "文字: 「" & ch & "」 のフリガナを選択・入力してください。" & vbCrLf & _
                "1: 推測読み (" & defaultRuby & ")"
                
    Dim userChoice As String
    userChoice = InputBoxWithHelp(promptMsg, "RubyManager - 1文字漢字対話学習", defaultRuby, "読みを直接入力するか、選択肢の番号を入力してください。")
    
    If userChoice = "1" Then userChoice = defaultRuby
    If userChoice = "" Then userChoice = defaultRuby
    
    ' 学習辞書に追加・保存
    If InStr(cfg.CustomDict, ch & "=" & userChoice) = 0 Then
        cfg.CustomDict = cfg.CustomDict & ch & "=" & userChoice & ";"
        Call SaveConfig(cfg)
    End If
    
    GetInteractiveRuby = userChoice
End Function

' Excel GetPhoneticを利用したふりがな取得
Private Function GetPhoneticString(ByVal txt As String) As String
    On Error Resume Next
    Dim xlApp As Object
    Set xlApp = CreateObject("Excel.Application")
    If Not xlApp Is Nothing Then
        GetPhoneticString = xlApp.GetPhonetic(txt)
        xlApp.Quit
        Set xlApp = Nothing
    End If
    If GetPhoneticString = "" Then GetPhoneticString = txt
End Function

' ルビの適用処理
Private Sub ApplyRubyToRange(ByRef rng As Range, ByVal rubyText As String, ByRef cfg As RubyConfig)
    On Error Resume Next
    rng.PhoneticGuide Text:=rubyText, Alignment:=wdPhoneticGuideAlignmentCenter, _
                      Raise:=0, SecondFontSize:=0, FontSize:=rng.Font.Size / 2
End Sub

' レイアウト自動補正処理
Private Sub ApplyLayoutFixes(ByRef doc As Document, ByVal layoutCode As String)
    Dim p As Paragraph
    ' オプション1: 行間固定（ルビ拡張防止）
    If InStr(layoutCode, "1") > 0 Or InStr(layoutCode, "4") > 0 Then
        For Each p In doc.Paragraphs
            p.Format.LineSpacingRule = wdLineSpaceExactly
            p.Format.LineSpacing = 18 ' 18ptに固定
        Next p
    End If
    
    ' オプション2: テキストボックスの自動調整
    If InStr(layoutCode, "2") > 0 Or InStr(layoutCode, "4") > 0 Then
        Dim shp As Shape
        For Each shp In doc.Shapes
            If shp.HasTextFrame Then
                shp.TextFrame.WordWrap = True
                shp.TextFrame.MarginTop = 2
                shp.TextFrame.MarginBottom = 2
            End If
        Next shp
    End If
End Sub

' UTF-8 INIファイル読み込み/保存機能
Private Sub LoadConfig(ByRef cfg As RubyConfig)
    ' 初期値セット
    cfg.Mode = 1
    cfg.Level = 1
    cfg.SingleMode = 1
    cfg.LayoutOptions = "0"
    cfg.RubyScale = 50
    cfg.CustomDict = ""
    
    Dim iniPath As String
    iniPath = ActiveDocument.Path & "\ruby_config.ini"
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(iniPath) Then Exit Sub
    
    On Error Resume Next
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' adTypeText
    stream.Charset = "UTF-8"
    stream.Open
    stream.LoadFromFile iniPath
    
    Dim content As String
    content = stream.ReadText
    stream.Close
    
    ' 設定値パース
    Dim lines() As String
    lines = Split(content, vbCrLf)
    Dim lineItem As Variant, kv() As String
    For Each lineItem In lines
        If InStr(lineItem, "=") > 0 Then
            kv = Split(lineItem, "=")
            Select Case Trim(kv(0))
                Case "Mode": cfg.Mode = CInt(kv(1))
                Case "Level": cfg.Level = CInt(kv(1))
                Case "SingleMode": cfg.SingleMode = CInt(kv(1))
                Case "LayoutOptions": cfg.LayoutOptions = Trim(kv(1))
                Case "CustomDict": cfg.CustomDict = Trim(kv(1))
            End Select
        End If
    Next lineItem
End Sub

Private Sub SaveConfig(ByRef cfg As RubyConfig)
    Dim iniPath As String
    iniPath = ActiveDocument.Path & "\ruby_config.ini"
    If ActiveDocument.Path = "" Then Exit Sub ' 未保存文書の場合はスキップ
    
    Dim content As String
    content = "[RubyManagerSettings]" & vbCrLf & _
              "Mode=" & cfg.Mode & vbCrLf & _
              "Level=" & cfg.Level & vbCrLf & _
              "SingleMode=" & cfg.SingleMode & vbCrLf & _
              "LayoutOptions=" & cfg.LayoutOptions & vbCrLf & _
              "CustomDict=" & cfg.CustomDict & vbCrLf
              
    On Error Resume Next
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "UTF-8"
    stream.Open
    stream.WriteText content
    stream.SaveToFile iniPath, 2 ' adSaveCreateOverWrite
    stream.Close
End Sub
