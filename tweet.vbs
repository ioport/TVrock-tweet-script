Option Explicit
'On Error Resume Next
Dim arg ,fso ,regEx ,WshShell ,oExec ,msg ,i ,j ,tempFile ,tempLine ,matchLog ,TimeAdj1 ,TimeAdj1s ,TimeAdj2 ,TimeAdj2s ,DiskFree ,strChHashTag ,strRecState ,strTweet
Dim twtcnslPath ,intMaxRetry ,intRetryTime ,intSleepTime ,MatchStrict ,strRepFile1 ,strRepFile2 ,HashTag ,intDiskFree ,addChTag ,strChList ,strTagList
Set arg = WScript.Arguments
Set fso = CreateObject("Scripting.FileSystemObject")
Set regEx = New RegExp
Set WshShell = WScript.CreateObject("WScript.Shell")



'----------�ݒ荀��------------------------------------------------------------
'TweetConsole�̃t���p�X�i���K�{�j
twtcnslPath = "C:\TvRock\TweetConsole\twtcnsl.exe"

'Twitter���e�G���[���̃��g���C����񐔁i���e��2����Ń^�C���A�E�g����悤�ł��j
intMaxRetry = 3 '�����ݒ�F3

'Twitter���e�G���[���̃��g���C�Ԋu�i�~���b�j
intRetryTime = 30000 '�����ݒ�F30000�i30�b�j

'�^��J�n�E�I�����Ƀ��O�t�@�C�����J���܂ł̑ҋ@���ԁi�~���b�j
intSleepTime = 3500 '�����ݒ�F3500�i3.5�b�j

'�^��J�n�E�I�����̃��O�����ŗ\��^�C�g������������Ȃ��ꍇ�i0�ő��s�A1�ŏI���j
MatchStrict = 0 '�����ݒ�F0�i�u�A�������\��͘^����I�����Ȃ��v�ꍇ�̕s��΍�j

'�^��J�n�E�I���Ńt�@�C�������p����ꍇ�̒u������1�i���Ă͂܂�Ȃ���Ώ���2�ցj
strRepFile1 = "(.+?)_\d{6}.*" '�t�@�C�����̖����́u_�N�����c�v����菜��
'strRepFile1 = "(.+?)_" & arg(4) & ".*" '�t�@�C�����̖����́u_�����ǖ��c�v����菜��
'strRepFile1 = "\d{12}_(.+?) _.+?$" 'SCRename�u@YY@MM@DD@SH@SM_@TT _@CH�v����^�C�g���ȊO����菜��

'�^��J�n�E�I���Ńt�@�C�������p����ꍇ�̒u������2�i���Ă͂܂�Ȃ���΃t�@�C���������̂܂܎g�p�j
'strRepFile1 = "(.+?)_\d{6}.*" '�t�@�C�����̖����́u_�N�����c�v����菜��
'strRepFile2 = "(.+?)_" & arg(4) & ".*" '�t�@�C�����̖����́u_�����ǖ��c�v����菜��
strRepFile2 = "\d{12}_(.+?) _.+?$" 'SCRename�u@YY@MM@DD@SH@SM_@TT _@CH�v����^�C�g���ȊO����菜��

'TvRock�̃n�b�V���^�O�̕\�L�i0�Ń^�O�Ȃ��A1�Łu�ʏ�^�O�v�̂݁A2�Łu�Z���^�O �ʏ�^�O�v�j
HashTag = 2 '�����ݒ�F2

'�󂫗e�ʂ��w��%�����Ŗ����Ɂu���󂫗e�ʁ�%�v�̒ǉ��i0�Œǉ����Ȃ��j2TB����5%�Ŗ�100GB
intDiskFree = 5 '�����ݒ�F5

'�������̃`�����l���n�b�V���^�O��ǉ��i0�ŕt�����Ȃ��A1�ŕt������j
addChTag = 0 '�����ݒ�F0

'�`�����l���n�b�V���^�O�u���p�̕����ǖ��iTvRock�Őݒ肵�Ă�������ǖ��ɂ��Ă��������j
strChList = Array("�m�g�j����","�m�g�j����","���{�e���r","�e���r����","�s�a�r�e���r","�e���r����","�t�W�e���r","�l�w�e���r", _
                  "�m�g�j�q�����","�m�g�j�@�a�r�P","�m�g�j�n�C�r�W����","�m�g�j�@�a�r�v���~�A��","�a�r���e��","�a�r����","�a�r�|�s�a�r","�a�r�W���p��","�a�r�t�W","�a�r�P�P")
'�`�����l���n�b�V���^�O�i�����ǖ��Ə��Ԃ����킹�Ă��������j
strTagList = Array("nhk","etv","ntv","tvasahi","tbs","tvtokyo","fujitv","mxtv", _
                   "bs1","bs1","nhkbsp","nhkbsp","bsntv","bsasahi","bstbs","bsj","bsfuji","bs11")
'------------------------------------------------------------------------------



'�����U�蕪��
Select Case Len(arg(0))
	Case 2
		RecLog '�^��J�n�E�I��
	Case 3
		Watching '������
	Case 4
		RecSet '�^��\��Ǝ��Ԓ���
End Select
Set arg = Nothing
Set tempFile = Nothing



'�^��\��Ǝ��Ԓ����@�����̗�F"�^��\��" "11/25" "21:00" "22:00" "twrb12345678" "�����ǖ�" "�\��^�C�g����"
Sub RecSet()
strTweet = arg(0) & " ["
Select Case Len(arg(1))
	Case 4,5 '�W��
		strTweet = strTweet & Replace(arg(1),"/0","/") & " " & arg(2) & "�`" & arg(3)
	Case 7,8 '28���ԕ\��
		If DateDiff("h","0:00",arg(2)) < 4 Then
			If DateDiff("h",arg(2),arg(3)) < 0 Then
				strTweet = strTweet & Mid(Replace(DateAdd("d",-1,DateValue(Mid(arg(1),4))),"/0","/"),6) & " " & Left(arg(2),1)+24 & Right(arg(2),3) & "�`" & Left(arg(3),1)+48 & Right(arg(3),3)
			Else
				strTweet = strTweet & Mid(Replace(DateAdd("d",-1,DateValue(Mid(arg(1),4))),"/0","/"),6) & " " & Left(arg(2),1)+24 & Right(arg(2),3) & "�`" & Left(arg(3),1)+24 & Right(arg(3),3)
			End If
		ElseIf DateDiff("h",arg(2),arg(3)) < 0 Then
			strTweet = strTweet & Replace(Mid(arg(1),4),"/0","/") & " " & arg(2) & "�`" & Left(arg(3),1)+24 & Right(arg(3),3)
		Else
			strTweet = strTweet & Replace(Mid(arg(1),4),"/0","/") & " " & arg(2) & "�`" & arg(3)
		End If
	Case 9,10 '�j���\��
		strTweet = strTweet & Replace(Mid(arg(1),6),"/0","/") & "(" & WeekdayName(Weekday(DateValue(arg(1))),True) & ") " & arg(2) & "�`" & arg(3)
	Case 12,13 '�j���E28���ԕ\��
		If DateDiff("h","0:00",arg(2)) < 4 Then
			If DateDiff("h",arg(2),arg(3)) < 0 Then
				strTweet = strTweet & Mid(Replace(DateAdd("d",-1,DateValue(Mid(arg(1),4))),"/0","/"),6) & "(" & WeekdayName(Weekday(DateAdd("d",-1,DateValue(Mid(arg(1),4)))),True) & ") " & Left(arg(2),1)+24 & Right(arg(2),3) & "�`" & Left(arg(3),1)+48 & Right(arg(3),3)
			Else
				strTweet = strTweet & Mid(Replace(DateAdd("d",-1,DateValue(Mid(arg(1),4))),"/0","/"),6) & "(" & WeekdayName(Weekday(DateAdd("d",-1,DateValue(Mid(arg(1),4)))),True) & ") " & Left(arg(2),1)+24 & Right(arg(2),3) & "�`" & Left(arg(3),1)+24 & Right(arg(3),3)
			End If
		ElseIf DateDiff("h",arg(2),arg(3)) < 0 Then
			strTweet = strTweet & Replace(Mid(arg(1),9),"/0","/") & "(" & WeekdayName(Weekday(DateValue(Mid(arg(1),4))),True) & ") " & arg(2) & "�`" & Left(arg(3),1)+24 & Right(arg(3),3)
		Else
			strTweet = strTweet & Replace(Mid(arg(1),9),"/0","/") & "(" & WeekdayName(Weekday(DateValue(Mid(arg(1),4))),True) & ") " & arg(2) & "�`" & arg(3)
		End If
End Select

If arg.Count > 7 Then '�\��^�C�g�����Ƀ_�u���N�I�[�e�[�V����������ꍇ�̏���
	For i=6 To arg.Count-1
		matchLog = matchLog & " " & arg(i)
	Next
Else
	matchLog = " " & arg(6)
End If

strTweet = strTweet & "]" & matchLog & " [" & arg(5) & ":"

If StrComp(arg(0),"���Ԓ���") = 0 Then
	WScript.Sleep 10000 '�O�̂���tvrock.log�̍X�V��҂�(10�b)
	regEx.Pattern = "\[.+?\]:�ԑg�u(.+?)�v��(.+?)���Ԃ�(.+?)��(\d+?)�b�������܂���"
	Set tempFile = fso.OpenTextFile("tvrock.log") '����VBS�t�@�C����TvRock�ݒ�̍�ƃt�H���_�ȊO�ɒu�����ꍇ�͊e���̊��ɍ��킹�ĉ�����
	Do Until tempFile.AtEndOfStream
		tempLine = tempFile.ReadLine
		If regEx.Test(tempLine) Then
			If DateDiff("s",Mid(tempLine,2,17),Now) < 40 Then '���O�̎��Ԓ�����񂪎��s������40�b�����̏ꍇ
				If StrComp(regEx.Replace(tempLine,"$1"),arg(6)) = 0 Then
					If StrComp(regEx.Replace(tempLine,"$2"),"�J�n") = 0 Then
						TimeAdj1 = regEx.Replace(tempLine,"$3")
						TimeAdj1s = regEx.Replace(tempLine,"$4")
					End If
					If StrComp(regEx.Replace(tempLine,"$2"),"�I��") = 0 Then
						TimeAdj2 = regEx.Replace(tempLine,"$3")
						TimeAdj2s = regEx.Replace(tempLine,"$4")
					End If
				End If
			End If
		End If
	Loop
	tempFile.Close
	
	If TimeAdj1 >= 0 Then i = "�J�n+"
	If TimeAdj1 <  0 Then i = "�J�n-"
	If TimeAdj2 >= 0 Then j = "�I��+"
	If TimeAdj2 <  0 Then j = "�I��-"
	Select Case StrComp(TimeAdj1s,"00")
		Case -1
			TimeAdj1s = ""
		Case 0
			TimeAdj1s = ","
		Case Else
			TimeAdj1s = TimeAdj1s & "�b,"
	End Select
	Select Case StrComp(TimeAdj2s,"00")
		Case -1
			TimeAdj2s = ""
		Case 0
			TimeAdj2s = ","
		Case Else
			TimeAdj2s = TimeAdj2s & "�b,"
	End Select
	If IsEmpty(TimeAdj1) = False Then
		TimeAdj1 = Abs(TimeAdj1)
		If TimeAdj1 > 1440 Then
			TimeAdj1 = i & Fix(TimeAdj1/1440) & "��" & Fix((TimeAdj1-Fix(TimeAdj1/1440)*1440)/60) & "����" & (TimeAdj1-Fix(TimeAdj1/1440)*1440)-Fix((TimeAdj1-Fix(TimeAdj1/1440)*1440)/60)*60 & "��"
'			TimeAdj1 = i & Fix(TimeAdj1/1440) & "�� " & Left(TimeSerial(0,TimeAdj1-Fix(TimeAdj1/1440)*1440,0),5)
		ElseIf TimeAdj1 > 60 Then
			TimeAdj1 = i & Fix(TimeAdj1/60) & "����" & TimeAdj1-Fix(TimeAdj1/60)*60 & "��"
		ElseIf TimeAdj1 > 0 Then
			TimeAdj1 = i & TimeAdj1 & "��"
'			TimeAdj1 = i & Left(TimeSerial(0,TimeAdj1,0),5)
		ElseIf TimeAdj1 = 0 Then
			TimeAdj1 = i
		End If
	End If
	If IsEmpty(TimeAdj2) = False Then
		TimeAdj2 = Abs(TimeAdj2)
		If TimeAdj2 > 1440 Then
			TimeAdj2 = j & Fix(TimeAdj2/1440) & "��" & Fix((TimeAdj2-Fix(TimeAdj2/1440)*1440)/60) & "����" & (TimeAdj2-Fix(TimeAdj2/1440)*1440)-Fix((TimeAdj2-Fix(TimeAdj2/1440)*1440)/60)*60 & "��"
'			TimeAdj2 = j & Fix(TimeAdj2/1440) & "�� " & Left(TimeSerial(0,TimeAdj2-Fix(TimeAdj2/1440)*1440,0),5)
		ElseIf TimeAdj2 > 60 Then
			TimeAdj2 = j & Fix(TimeAdj2/60) & "����" & TimeAdj2-Fix(TimeAdj2/60)*60 & "��"
		ElseIf TimeAdj2 > 0 Then
			TimeAdj2 = j & TimeAdj2 & "��"
'			TimeAdj2 = j & Left(TimeSerial(0,TimeAdj2,0),5)
		ElseIf TimeAdj2 = 0 Then
			TimeAdj2 = j
		End If
	End If
	strTweet = strTweet & TimeAdj1 & TimeAdj1s & TimeAdj2 & TimeAdj2s
Else
	Set tempFile = fso.OpenTextFile("tvrock.log2") '����VBS�t�@�C����TvRock�ݒ�̍�ƃt�H���_�ȊO�ɒu�����ꍇ�͊e���̊��ɍ��킹�ĉ�����
	tempLine = tempFile.ReadLine
	tempFile.Close
End If

If Len(arg(4)) > 0 Then 'TvRock�̃n�b�V���^�O
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



'�������@�����̗�F"������" "twrb12345678" "�����ǖ�" "�W��������" "�^�C�g����"
Sub Watching()
'If arg(3) = "�j���[�X�^��" Then Exit Sub '�������̃c�C�[�g���Ȃ��W�������̏ꍇ�A�s���́u'�v���폜
'If arg(3) = "�X�|�[�c" Then Exit Sub '����
'If arg(3) = "�h���}" Then Exit Sub '����
'If arg(3) = "���y" Then Exit Sub '����
'If arg(3) = "�o���G�e�B�[" Then Exit Sub '����
'If arg(3) = "�f��" Then Exit Sub '����
'If arg(3) = "�A�j���^���B" Then Exit Sub '����
'If arg(3) = "���^���C�h�V���[" Then Exit Sub '����
'If arg(3) = "�h�L�������^���[�^���{ " Then Exit Sub '����
'If arg(3) = "����^����" Then Exit Sub '����
'If arg(3) = "��^����" Then Exit Sub '����

If arg.Count > 5 Then '�^�C�g�����Ƀ_�u���N�I�[�e�[�V����������ꍇ�̏���
	For i=4 To arg.Count-1
		matchLog = matchLog & " " & arg(i)
	Next
Else
	matchLog = " " & arg(4)
End If

If Len(arg(1)) > 0 Then 'TvRock�̃n�b�V���^�O
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

Sub ChTag() '�����ǖ��n�b�V���^�O
For i = LBound(strChList) To UBound(strChList)
	If InStr(arg(2),strChList(i)) <> 0 Then Exit For
Next
strChHashTag = " #" & strTagList(i)
End Sub



'�^��J�n�E�I���@�����̗�(�\��^�C�g�����܂ł͕K�{)�F "�I��" "T2" "twrb12345678" "�\��^�C�g����" "�����ǖ�" "�t�@�C����" "�W��������"
Sub RecLog()
If arg.Count > 6 Then '�w��W�������ł̓c�C�[�g���Ȃ��ꍇ�̏���
'	If arg(arg.Count-1) = "�j���[�X�^��" Then Exit Sub '�^��J�n�E�I���̃c�C�[�g���Ȃ��W�������̏ꍇ�A�s���́u'�v���폜
'	If arg(arg.Count-1) = "�X�|�[�c" Then Exit Sub '����
'	If arg(arg.Count-1) = "�h���}" Then Exit Sub '����
'	If arg(arg.Count-1) = "���y" Then Exit Sub '����
'	If arg(arg.Count-1) = "�o���G�e�B�[" Then Exit Sub '����
'	If arg(arg.Count-1) = "�f��" Then Exit Sub '����
'	If arg(arg.Count-1) = "�A�j���^���B" Then Exit Sub '����
'	If arg(arg.Count-1) = "���^���C�h�V���[" Then Exit Sub '����
'	If arg(arg.Count-1) = "�h�L�������^���[�^���{ " Then Exit Sub '����
'	If arg(arg.Count-1) = "����^����" Then Exit Sub '����
'	If arg(arg.Count-1) = "��^����" Then Exit Sub '����
End If

For i = 0 To 1
	WScript.Sleep intSleepTime '�O�̂���tvrock.log�̍X�V��҂�
	regEx.Pattern = "\[" & arg(1) & "\]�ԑg�u.+?�v �^��" & arg(0)
	Set tempFile = fso.OpenTextFile("tvrock.log") '����VBS�t�@�C����TvRock�ݒ�̍�ƃt�H���_�ȊO�ɒu�����ꍇ�͊e���̊��ɍ��킹�ĉ�����
	Do Until tempFile.AtEndOfStream
		tempLine = tempFile.ReadLine
		If regEx.Test(tempLine) Then matchLog = tempLine
	Loop
	tempFile.Close
	regEx.Pattern = "^.*:\d\d (.+?)\]:\[.+?\]�ԑg�u(.+?)�v (.+?) Card=.+?Sig=(.+?), Bitrate=(.+?)Mbps, Drop=(.+?), Scrambling.+? BcTimeDiff=(.+?)sec, TimeAdj=(.+?)sec, CPU.+?DiskFree=(.+?)%\."
	If StrComp(regEx.Replace(matchLog,"$2"),arg(3)) = 0 Then Exit For
Next

If DateDiff("s",Mid(matchLog,2,17),Now) > 20 Then Exit Sub '���O�̘^���񂪎��s������20�b�𒴂��Ă���ꍇ�ɏ������I��

If Len(arg(2)) > 0 Then 'TvRock�̃n�b�V���^�O
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

If CSng(regEx.Replace(matchLog,"$9")) < intDiskFree Then DiskFree = " ���󂫗e��" & regEx.Replace(matchLog,"$9") & "%" '�󂫗e�ʂ��w��%�����̏ꍇ�Ƀ��b�Z�[�W�𖖔��֒ǉ�

If StrComp(regEx.Replace(matchLog,"$2"),arg(3)) <> 0 Then
	If InStr(regEx.Replace(matchLog,"$2"),"""") <> 0 Then '�\��^�C�g�����Ƀ_�u���N�I�[�e�[�V����������ꍇ�̏���
		strTweet = regEx.Replace(matchLog,"$3 $2 [Sg$4,Br$5,Dr$6,Td$7,Ta$8,TvRock V$1" & HashTag & "]" & DiskFree)
		TweetPost
		Exit Sub
	End If
	If MatchStrict = 1 Then WScript.Quit(224) '���O�����ŗ\��^�C�g������������Ȃ���ΏI������ꍇ�A�ُ�I���R�[�h(0xe0)��Ԃ��ďI��
	strTweet = regEx.Replace(matchLog,"$3 $2 [Sg$4,Br$5,Dr$6,Td$7,Ta$8,TvRock V$1" & HashTag & "]" & DiskFree)
	TweetPost
	Exit Sub
End If

strRecState = regEx.Replace(matchLog,"Sg$4,Br$5,Dr$6,Td$7,Ta$8,TvRock V$1" & HashTag & "]" & DiskFree)
Select Case arg.Count
	Case 4
		strTweet = regEx.Replace(matchLog,"$3 $2 [") & strRecState 'Ver 0.9t8�����̕\�L
	Case 5
		strTweet = regEx.Replace(matchLog,"$3 $2 [") & arg(4) & ":" & strRecState '�����ǖ��ǉ�
	Case Else
		strTweet = regEx.Replace(matchLog,"$3 ") '�\��^�C�g�����̑���Ƀt�@�C�������g��
		regEx.Pattern = strRepFile1 '�t�@�C�����u������1
		If regEx.Test(arg(5)) Then
			strTweet = strTweet & regEx.Replace(arg(5),"$1") & " [" & arg(4) & ":" & strRecState
		Else
			regEx.Pattern = strRepFile2 '�t�@�C�����u������2
			If regEx.Test(arg(5)) Then
				strTweet = strTweet & regEx.Replace(arg(5),"$1") & " [" & arg(4) & ":" & strRecState
			Else
				strTweet = strTweet & arg(5) & " [" & arg(4) & ":" & strRecState '��L�̒u�������ɓ��Ă͂܂�Ȃ��ꍇ�t�@�C���������̂܂܎g��
			End If
		End If
End Select
TweetPost
End Sub



'�c�C�[�g�𓊍e
Sub TweetPost()
strTweet = Replace(strTweet,"���E��","���g��") 'TvRock�ŕ����������镶�����u��
strTweet = Replace(strTweet,"�E�c���F","���c���F")
strTweet = Replace(strTweet,"�E�����","�������")
strTweet = Replace(strTweet,"�E���^���q","�����^���q")
strTweet = Replace(strTweet,"�E�i�p��","���i�p��")
strTweet = Replace(strTweet,"�{�E������","�{��������")
regEx.Pattern = "�^��(�\��|�J�n|�I��)(.+?)�C�J��(.+?)"
strTweet = regEx.Replace(strTweet,"�^��$1�ŃQ�\�I$2�C�J��$3")
regEx.Pattern = "(.+?)\[(.+?)(:|,)TvRock.*?\]$"
If Len(regEx.Replace(strTweet,"$1\[$2")) <= 140 Then '�u,TvRock�v�ȉ�����������140�����ȓ��Ȃ疖���́u,TvRock�v�ȉ����폜���ē��e
	For j=1 To intMaxRetry '���e�̃��g���C��
		Set oExec = WshShell.Exec("""" & twtcnslPath & """ /t """ & Left(strTweet,140))
		msg = oExec.StdOut.ReadAll
		If InStr(msg,"�Ԃ₫�܂���") <> 0 Then Exit For
		If j=intMaxRetry Then WScript.Quit(225) '���g���C�񐔂𒴂�����ُ�I���R�[�h(0xe1)��Ԃ��ďI��
		WScript.Sleep intRetryTime '���e�ł��Ȃ������ꍇ���g���C�܂őҋ@
	Next
Else
	For i=0 To Int(Len(strTweet)/134+0.9)-1 '��L��140�����𒴂���Ȃ�擪�ɕ�������ǉ����ĕ������e
		For j=1 To intMaxRetry '���e�̃��g���C��
			Set oExec = WshShell.Exec("""" & twtcnslPath & """ /t (""" & i+1 & "/" & Int(Len(strTweet)/134+0.9) & ") " & Mid(strTweet,1+134*i,134))
			msg = i & oExec.StdOut.ReadAll
			If InStr(msg,i & "�Ԃ₫�܂���") <> 0 Then Exit For
			If j=intMaxRetry Then WScript.Quit(225+i) '���g���C�񐔂𒴂�����ُ�I���R�[�h(0xe1)�ȍ~��Ԃ��ďI��
			WScript.Sleep intRetryTime '���e�ł��Ȃ������ꍇ���g���C�܂őҋ@
		Next
	Next
End If
Set WshShell = Nothing
End Sub