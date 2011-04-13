Option Explicit
'On Error Resume Next
Dim arg ,fso ,regEx ,WshShell ,oExec ,msg ,i ,j ,tempFile ,tempLine ,matchLog ,TimeAdj1 ,TimeAdj1s ,TimeAdj2 ,TimeAdj2s ,DiskFree ,strChHashTag ,strRecState ,strTweet
Dim twtcnslPath ,intMaxRetry ,intRetryTime ,intSleepTime ,MatchStrict ,strRepFile1 ,strRepFile2 ,HashTag ,intDiskFree ,addChTag ,strChList ,strTagList
Set arg = WScript.Arguments
Set fso = CreateObject("Scripting.FileSystemObject")
Set regEx = New RegExp
Set WshShell = WScript.CreateObject("WScript.Shell")



'----------設定項目------------------------------------------------------------
'TweetConsoleのフルパス（※必須）
twtcnslPath = "C:\TvRock\TweetConsole\twtcnsl.exe"

'Twitter投稿エラー時のリトライ上限回数（投稿は2分弱でタイムアウトするようです）
intMaxRetry = 3 '初期設定：3

'Twitter投稿エラー時のリトライ間隔（ミリ秒）
intRetryTime = 30000 '初期設定：30000（30秒）

'録画開始・終了時にログファイルを開くまでの待機時間（ミリ秒）
intSleepTime = 3500 '初期設定：3500（3.5秒）

'録画開始・終了時のログ検索で予約タイトル名が見つからない場合（0で続行、1で終了）
MatchStrict = 0 '初期設定：0（「連続した予約は録画を終了しない」場合の不具合対策）

'録画開始・終了でファイル名利用する場合の置換条件1（当てはまらなければ条件2へ）
strRepFile1 = "(.+?)_\d{6}.*" 'ファイル名の末尾の「_年月日…」を取り除く
'strRepFile1 = "(.+?)_" & arg(4) & ".*" 'ファイル名の末尾の「_放送局名…」を取り除く
'strRepFile1 = "\d{12}_(.+?) _.+?$" 'SCRename「@YY@MM@DD@SH@SM_@TT _@CH」からタイトル以外を取り除く

'録画開始・終了でファイル名利用する場合の置換条件2（当てはまらなければファイル名をそのまま使用）
'strRepFile1 = "(.+?)_\d{6}.*" 'ファイル名の末尾の「_年月日…」を取り除く
'strRepFile2 = "(.+?)_" & arg(4) & ".*" 'ファイル名の末尾の「_放送局名…」を取り除く
strRepFile2 = "\d{12}_(.+?) _.+?$" 'SCRename「@YY@MM@DD@SH@SM_@TT _@CH」からタイトル以外を取り除く

'TvRockのハッシュタグの表記（0でタグなし、1で「通常タグ」のみ、2で「短いタグ 通常タグ」）
HashTag = 2 '初期設定：2

'空き容量が指定%未満で末尾に「※空き容量○%」の追加（0で追加しない）2TBだと5%で約100GB
intDiskFree = 5 '初期設定：5

'視聴中のチャンネルハッシュタグを追加（0で付加しない、1で付加する）
addChTag = 0 '初期設定：0

'チャンネルハッシュタグ置換用の放送局名（TvRockで設定してある放送局名にしてください）
strChList = Array("ＮＨＫ総合","ＮＨＫ教育","日本テレビ","テレビ朝日","ＴＢＳテレビ","テレビ東京","フジテレビ","ＭＸテレビ", _
                  "ＮＨＫ衛星第一","ＮＨＫ　ＢＳ１","ＮＨＫハイビジョン","ＮＨＫ　ＢＳプレミアム","ＢＳ日テレ","ＢＳ朝日","ＢＳ－ＴＢＳ","ＢＳジャパン","ＢＳフジ","ＢＳ１１")
'チャンネルハッシュタグ（放送局名と順番を合わせてください）
strTagList = Array("nhk","etv","ntv","tvasahi","tbs","tvtokyo","fujitv","mxtv", _
                   "bs1","bs1","nhkbsp","nhkbsp","bsntv","bsasahi","bstbs","bsj","bsfuji","bs11")
'------------------------------------------------------------------------------



'処理振り分け
Select Case Len(arg(0))
	Case 2
		RecLog '録画開始・終了
	Case 3
		Watching '視聴中
	Case 4
		RecSet '録画予約と時間調整
End Select
Set arg = Nothing
Set tempFile = Nothing



'録画予約と時間調整　引数の例："録画予約" "11/25" "21:00" "22:00" "twrb12345678" "放送局名" "予約タイトル名"
Sub RecSet()
strTweet = arg(0) & " ["
Select Case Len(arg(1))
	Case 4,5 '標準
		strTweet = strTweet & Replace(arg(1),"/0","/") & " " & arg(2) & "～" & arg(3)
	Case 7,8 '28時間表示
		If DateDiff("h","0:00",arg(2)) < 4 Then
			If DateDiff("h",arg(2),arg(3)) < 0 Then
				strTweet = strTweet & Mid(Replace(DateAdd("d",-1,DateValue(Mid(arg(1),4))),"/0","/"),6) & " " & Left(arg(2),1)+24 & Right(arg(2),3) & "～" & Left(arg(3),1)+48 & Right(arg(3),3)
			Else
				strTweet = strTweet & Mid(Replace(DateAdd("d",-1,DateValue(Mid(arg(1),4))),"/0","/"),6) & " " & Left(arg(2),1)+24 & Right(arg(2),3) & "～" & Left(arg(3),1)+24 & Right(arg(3),3)
			End If
		ElseIf DateDiff("h",arg(2),arg(3)) < 0 Then
			strTweet = strTweet & Replace(Mid(arg(1),4),"/0","/") & " " & arg(2) & "～" & Left(arg(3),1)+24 & Right(arg(3),3)
		Else
			strTweet = strTweet & Replace(Mid(arg(1),4),"/0","/") & " " & arg(2) & "～" & arg(3)
		End If
	Case 9,10 '曜日表示
		strTweet = strTweet & Replace(Mid(arg(1),6),"/0","/") & "(" & WeekdayName(Weekday(DateValue(arg(1))),True) & ") " & arg(2) & "～" & arg(3)
	Case 12,13 '曜日・28時間表示
		If DateDiff("h","0:00",arg(2)) < 4 Then
			If DateDiff("h",arg(2),arg(3)) < 0 Then
				strTweet = strTweet & Mid(Replace(DateAdd("d",-1,DateValue(Mid(arg(1),4))),"/0","/"),6) & "(" & WeekdayName(Weekday(DateAdd("d",-1,DateValue(Mid(arg(1),4)))),True) & ") " & Left(arg(2),1)+24 & Right(arg(2),3) & "～" & Left(arg(3),1)+48 & Right(arg(3),3)
			Else
				strTweet = strTweet & Mid(Replace(DateAdd("d",-1,DateValue(Mid(arg(1),4))),"/0","/"),6) & "(" & WeekdayName(Weekday(DateAdd("d",-1,DateValue(Mid(arg(1),4)))),True) & ") " & Left(arg(2),1)+24 & Right(arg(2),3) & "～" & Left(arg(3),1)+24 & Right(arg(3),3)
			End If
		ElseIf DateDiff("h",arg(2),arg(3)) < 0 Then
			strTweet = strTweet & Replace(Mid(arg(1),9),"/0","/") & "(" & WeekdayName(Weekday(DateValue(Mid(arg(1),4))),True) & ") " & arg(2) & "～" & Left(arg(3),1)+24 & Right(arg(3),3)
		Else
			strTweet = strTweet & Replace(Mid(arg(1),9),"/0","/") & "(" & WeekdayName(Weekday(DateValue(Mid(arg(1),4))),True) & ") " & arg(2) & "～" & arg(3)
		End If
End Select

If arg.Count > 7 Then '予約タイトル名にダブルクオーテーションがある場合の処理
	For i=6 To arg.Count-1
		matchLog = matchLog & " " & arg(i)
	Next
Else
	matchLog = " " & arg(6)
End If

strTweet = strTweet & "]" & matchLog & " [" & arg(5) & ":"

If StrComp(arg(0),"時間調整") = 0 Then
	WScript.Sleep 10000 '念のためtvrock.logの更新を待つ(10秒)
	regEx.Pattern = "\[.+?\]:番組「(.+?)」の(.+?)時間を(.+?)分(\d+?)秒調整しました"
	Set tempFile = fso.OpenTextFile("tvrock.log") 'このVBSファイルをTvRock設定の作業フォルダ以外に置いた場合は各自の環境に合わせて下さい
	Do Until tempFile.AtEndOfStream
		tempLine = tempFile.ReadLine
		If regEx.Test(tempLine) Then
			If DateDiff("s",Mid(tempLine,2,17),Now) < 40 Then 'ログの時間調整情報が実行時から40秒未満の場合
				If StrComp(regEx.Replace(tempLine,"$1"),arg(6)) = 0 Then
					If StrComp(regEx.Replace(tempLine,"$2"),"開始") = 0 Then
						TimeAdj1 = regEx.Replace(tempLine,"$3")
						TimeAdj1s = regEx.Replace(tempLine,"$4")
					End If
					If StrComp(regEx.Replace(tempLine,"$2"),"終了") = 0 Then
						TimeAdj2 = regEx.Replace(tempLine,"$3")
						TimeAdj2s = regEx.Replace(tempLine,"$4")
					End If
				End If
			End If
		End If
	Loop
	tempFile.Close
	
	If TimeAdj1 >= 0 Then i = "開始+"
	If TimeAdj1 <  0 Then i = "開始-"
	If TimeAdj2 >= 0 Then j = "終了+"
	If TimeAdj2 <  0 Then j = "終了-"
	Select Case StrComp(TimeAdj1s,"00")
		Case -1
			TimeAdj1s = ""
		Case 0
			TimeAdj1s = ","
		Case Else
			TimeAdj1s = TimeAdj1s & "秒,"
	End Select
	Select Case StrComp(TimeAdj2s,"00")
		Case -1
			TimeAdj2s = ""
		Case 0
			TimeAdj2s = ","
		Case Else
			TimeAdj2s = TimeAdj2s & "秒,"
	End Select
	If IsEmpty(TimeAdj1) = False Then
		TimeAdj1 = Abs(TimeAdj1)
		If TimeAdj1 > 1440 Then
			TimeAdj1 = i & Fix(TimeAdj1/1440) & "日" & Fix((TimeAdj1-Fix(TimeAdj1/1440)*1440)/60) & "時間" & (TimeAdj1-Fix(TimeAdj1/1440)*1440)-Fix((TimeAdj1-Fix(TimeAdj1/1440)*1440)/60)*60 & "分"
'			TimeAdj1 = i & Fix(TimeAdj1/1440) & "日 " & Left(TimeSerial(0,TimeAdj1-Fix(TimeAdj1/1440)*1440,0),5)
		ElseIf TimeAdj1 > 60 Then
			TimeAdj1 = i & Fix(TimeAdj1/60) & "時間" & TimeAdj1-Fix(TimeAdj1/60)*60 & "分"
		ElseIf TimeAdj1 > 0 Then
			TimeAdj1 = i & TimeAdj1 & "分"
'			TimeAdj1 = i & Left(TimeSerial(0,TimeAdj1,0),5)
		ElseIf TimeAdj1 = 0 Then
			TimeAdj1 = i
		End If
	End If
	If IsEmpty(TimeAdj2) = False Then
		TimeAdj2 = Abs(TimeAdj2)
		If TimeAdj2 > 1440 Then
			TimeAdj2 = j & Fix(TimeAdj2/1440) & "日" & Fix((TimeAdj2-Fix(TimeAdj2/1440)*1440)/60) & "時間" & (TimeAdj2-Fix(TimeAdj2/1440)*1440)-Fix((TimeAdj2-Fix(TimeAdj2/1440)*1440)/60)*60 & "分"
'			TimeAdj2 = j & Fix(TimeAdj2/1440) & "日 " & Left(TimeSerial(0,TimeAdj2-Fix(TimeAdj2/1440)*1440,0),5)
		ElseIf TimeAdj2 > 60 Then
			TimeAdj2 = j & Fix(TimeAdj2/60) & "時間" & TimeAdj2-Fix(TimeAdj2/60)*60 & "分"
		ElseIf TimeAdj2 > 0 Then
			TimeAdj2 = j & TimeAdj2 & "分"
'			TimeAdj2 = j & Left(TimeSerial(0,TimeAdj2,0),5)
		ElseIf TimeAdj2 = 0 Then
			TimeAdj2 = j
		End If
	End If
	strTweet = strTweet & TimeAdj1 & TimeAdj1s & TimeAdj2 & TimeAdj2s
Else
	Set tempFile = fso.OpenTextFile("tvrock.log2") 'このVBSファイルをTvRock設定の作業フォルダ以外に置いた場合は各自の環境に合わせて下さい
	tempLine = tempFile.ReadLine
	tempFile.Close
End If

If Len(arg(4)) > 0 Then 'TvRockのハッシュタグ
	Select Case HashTag
		Case 0
			HashTag = ""
		Case 1
			HashTag = " " & arg(4)
		Case 2
			HashTag = " " & Left(arg(4),Len(arg(4))-4) & " " & arg(4)
	End Select
Else
	HashTag = ""
End If

regEx.Pattern = "^\[\d\d/\d\d/\d\d \d\d:\d\d:\d\d (.+?)\]:.+"
strTweet = strTweet & regEx.Replace(tempLine,"TvRock V$1") & HashTag & "]"
TweetPost
End Sub



'視聴中　引数の例："視聴中" "twrb12345678" "放送局名" "ジャンル名" "タイトル名"
Sub Watching()
'If arg(3) = "ニュース／報道" Then Exit Sub '視聴中のツイートしないジャンルの場合、行頭の「'」を削除
'If arg(3) = "スポーツ" Then Exit Sub '同上
'If arg(3) = "ドラマ" Then Exit Sub '同上
'If arg(3) = "音楽" Then Exit Sub '同上
'If arg(3) = "バラエティー" Then Exit Sub '同上
'If arg(3) = "映画" Then Exit Sub '同上
'If arg(3) = "アニメ／特撮" Then Exit Sub '同上
'If arg(3) = "情報／ワイドショー" Then Exit Sub '同上
'If arg(3) = "ドキュメンタリー／教養 " Then Exit Sub '同上
'If arg(3) = "劇場／公演" Then Exit Sub '同上
'If arg(3) = "趣味／教育" Then Exit Sub '同上

If arg.Count > 5 Then 'タイトル名にダブルクオーテーションがある場合の処理
	For i=4 To arg.Count-1
		matchLog = matchLog & " " & arg(i)
	Next
Else
	matchLog = " " & arg(4)
End If

If Len(arg(1)) > 0 Then 'TvRockのハッシュタグ
	Select Case HashTag
		Case 0
			HashTag = ""
		Case 1
			HashTag = " #" & arg(1)
		Case 2
			HashTag = " #" & Left(arg(1),Len(arg(1))-4) & " #" & arg(1)
	End Select
Else
	HashTag = ""
End If

If addChTag = 1 Then ChTag

strTweet = arg(0) & matchLog & HashTag & strChHashTag
TweetPost
End Sub

Sub ChTag() '放送局名ハッシュタグ
For i = LBound(strChList) To UBound(strChList)
	If InStr(arg(2),strChList(i)) <> 0 Then Exit For
Next
strChHashTag = " #" & strTagList(i)
End Sub



'録画開始・終了　引数の例(予約タイトル名までは必須)： "終了" "T2" "twrb12345678" "予約タイトル名" "放送局名" "ファイル名" "ジャンル名"
Sub RecLog()
If arg.Count > 6 Then '指定ジャンルではツイートしない場合の処理
'	If arg(arg.Count-1) = "ニュース／報道" Then Exit Sub '録画開始・終了のツイートしないジャンルの場合、行頭の「'」を削除
'	If arg(arg.Count-1) = "スポーツ" Then Exit Sub '同上
'	If arg(arg.Count-1) = "ドラマ" Then Exit Sub '同上
'	If arg(arg.Count-1) = "音楽" Then Exit Sub '同上
'	If arg(arg.Count-1) = "バラエティー" Then Exit Sub '同上
'	If arg(arg.Count-1) = "映画" Then Exit Sub '同上
'	If arg(arg.Count-1) = "アニメ／特撮" Then Exit Sub '同上
'	If arg(arg.Count-1) = "情報／ワイドショー" Then Exit Sub '同上
'	If arg(arg.Count-1) = "ドキュメンタリー／教養 " Then Exit Sub '同上
'	If arg(arg.Count-1) = "劇場／公演" Then Exit Sub '同上
'	If arg(arg.Count-1) = "趣味／教育" Then Exit Sub '同上
End If

For i = 0 To 1
	WScript.Sleep intSleepTime '念のためtvrock.logの更新を待つ
	regEx.Pattern = "\[" & arg(1) & "\]番組「.+?」 録画" & arg(0)
	Set tempFile = fso.OpenTextFile("tvrock.log") 'このVBSファイルをTvRock設定の作業フォルダ以外に置いた場合は各自の環境に合わせて下さい
	Do Until tempFile.AtEndOfStream
		tempLine = tempFile.ReadLine
		If regEx.Test(tempLine) Then matchLog = tempLine
	Loop
	tempFile.Close
	regEx.Pattern = "^.*:\d\d (.+?)\]:\[.+?\]番組「(.+?)」 (.+?) Card=.+?Sig=(.+?), Bitrate=(.+?)Mbps, Drop=(.+?), Scrambling.+? BcTimeDiff=(.+?)sec, TimeAdj=(.+?)sec, CPU.+?DiskFree=(.+?)%\."
	If StrComp(regEx.Replace(matchLog,"$2"),arg(3)) = 0 Then Exit For
Next

If DateDiff("s",Mid(matchLog,2,17),Now) > 20 Then Exit Sub 'ログの録画情報が実行時から20秒を超えている場合に処理を終了

If Len(arg(2)) > 0 Then 'TvRockのハッシュタグ
	Select Case HashTag
		Case 0
			HashTag = ""
		Case 1
			HashTag = " " & arg(2)
		Case 2
			HashTag = " " & Left(arg(2),Len(arg(2))-4) & " " & arg(2)
	End Select
Else
	HashTag = ""
End If

If CSng(regEx.Replace(matchLog,"$9")) < intDiskFree Then DiskFree = " ※空き容量" & regEx.Replace(matchLog,"$9") & "%" '空き容量が指定%未満の場合にメッセージを末尾へ追加

If StrComp(regEx.Replace(matchLog,"$2"),arg(3)) <> 0 Then
	If InStr(regEx.Replace(matchLog,"$2"),"""") <> 0 Then '予約タイトル名にダブルクオーテーションがある場合の処理
		strTweet = regEx.Replace(matchLog,"$3 $2 [Sg$4,Br$5,Dr$6,Td$7,Ta$8,TvRock V$1" & HashTag & "]" & DiskFree)
		TweetPost
		Exit Sub
	End If
	If MatchStrict = 1 Then WScript.Quit(224) 'ログ検索で予約タイトル名が見つからなければ終了する場合、異常終了コード(0xe0)を返して終了
	strTweet = regEx.Replace(matchLog,"$3 $2 [Sg$4,Br$5,Dr$6,Td$7,Ta$8,TvRock V$1" & HashTag & "]" & DiskFree)
	TweetPost
	Exit Sub
End If

strRecState = regEx.Replace(matchLog,"Sg$4,Br$5,Dr$6,Td$7,Ta$8,TvRock V$1" & HashTag & "]" & DiskFree)
Select Case arg.Count
	Case 4
		strTweet = regEx.Replace(matchLog,"$3 $2 [") & strRecState 'Ver 0.9t8準拠の表記
	Case 5
		strTweet = regEx.Replace(matchLog,"$3 $2 [") & arg(4) & ":" & strRecState '放送局名追加
	Case Else
		strTweet = regEx.Replace(matchLog,"$3 ") '予約タイトル名の代わりにファイル名を使う
		regEx.Pattern = strRepFile1 'ファイル名置換条件1
		If regEx.Test(arg(5)) Then
			strTweet = strTweet & regEx.Replace(arg(5),"$1") & " [" & arg(4) & ":" & strRecState
		Else
			regEx.Pattern = strRepFile2 'ファイル名置換条件2
			If regEx.Test(arg(5)) Then
				strTweet = strTweet & regEx.Replace(arg(5),"$1") & " [" & arg(4) & ":" & strRecState
			Else
				strTweet = strTweet & arg(5) & " [" & arg(4) & ":" & strRecState '上記の置換条件に当てはまらない場合ファイル名をそのまま使う
			End If
		End If
End Select
TweetPost
End Sub



'ツイートを投稿
Sub TweetPost()
strTweet = Replace(strTweet,"草・剛","草彅剛") 'TvRockで文字化けする文字列を置換
strTweet = Replace(strTweet,"・田延彦","髙田延彦")
strTweet = Replace(strTweet,"・橋大輔","髙橋大輔")
strTweet = Replace(strTweet,"・橋真梨子","髙橋真梨子")
strTweet = Replace(strTweet,"・永英明","德永英明")
strTweet = Replace(strTweet,"宮・あおい","宮﨑あおい")
regEx.Pattern = "録画(予約|開始|終了)(.+?)イカ娘(.+?)"
strTweet = regEx.Replace(strTweet,"録画$1でゲソ！$2イカ娘$3")
regEx.Pattern = "(.+?)\[(.+?)(:|,)TvRock.*?\]$"
If Len(regEx.Replace(strTweet,"$1\[$2")) <= 140 Then '「,TvRock」以下を除いたら140文字以内なら末尾の「,TvRock」以下を削除して投稿
	For j=1 To intMaxRetry '投稿のリトライ回数
		Set oExec = WshShell.Exec("""" & twtcnslPath & """ /t """ & Left(strTweet,140))
		msg = oExec.StdOut.ReadAll
		If InStr(msg,"つぶやきました") <> 0 Then Exit For
		If j=intMaxRetry Then WScript.Quit(225) 'リトライ回数を超えたら異常終了コード(0xe1)を返して終了
		WScript.Sleep intRetryTime '投稿できなかった場合リトライまで待機
	Next
Else
	For i=0 To Int(Len(strTweet)/134+0.9)-1 '上記で140文字を超えるなら先頭に分割数を追加して分割投稿
		For j=1 To intMaxRetry '投稿のリトライ回数
			Set oExec = WshShell.Exec("""" & twtcnslPath & """ /t (""" & i+1 & "/" & Int(Len(strTweet)/134+0.9) & ") " & Mid(strTweet,1+134*i,134))
			msg = i & oExec.StdOut.ReadAll
			If InStr(msg,i & "つぶやきました") <> 0 Then Exit For
			If j=intMaxRetry Then WScript.Quit(225+i) 'リトライ回数を超えたら異常終了コード(0xe1)以降を返して終了
			WScript.Sleep intRetryTime '投稿できなかった場合リトライまで待機
		Next
	Next
End If
Set WshShell = Nothing
End Sub