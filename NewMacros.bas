Attribute VB_Name = "NewMacros"
Sub SetBibleVersesLinks()
  ' disable screen update
  Application.ScreenUpdating = False
  bibleBooksMatch_string = "Бытие:ge:Быт:ge:Исход:ex:Исх:ex:Левит:le:Лев:le:Числа:nu:Чис:nu:Второзаконие:de:Втор:de:Иисус Навин:jos:Навин:jos:Нав:jos:Судей:jud:Суд:jud:Руфь:ru:Руф:ru:1 Царств:1ki:1 Цар:1ki:2 Царств:2ki:2 Цар:2ki:3 Царств:3ki:3 Цар:3ki:4 Царств:4ki:4 Цар:4ki:" & _
    "1Царств:1ki:1Цар:1ki:2Царств:2ki:2Цар:2ki:3Царств:3ki:3Цар:3ki:4Царств:4ki:4Цар:4ki:1 Паралипоменон:1ch:1 Пар:1ch:2 Паралипоменон:2ch:2 Пар:2ch:1Паралипоменон:1ch:1Пар:1ch:2Паралипоменон:2ch:2Пар:2ch:Ездра:ezr:Езд:ezr:Неемия:ne:Неем:ne:Есфирь:es:" & _
    "Есф:es:Иов:job:Псалтирь:ps:Псалтырь:ps:Псалом:ps:Пс:ps:Притчи:pr:Прит:pr:Екклесиаст:ec:Еккл:ec:Песни Песней:so:Песн:so:Исаия:isa:Ис:isa:Иеремия:jer:Иер:jer:Плач Иеремии:la:Плач:la:Иезекииль:eze:Иез:eze:Даниил:da:Дан:da:Осия:ho:Ос:ho:Иоиль:joe:" & _
    "Иоил:joe:Амос:am:Ам:am:Авдий:ob:Авд:ob:Иона:jon:Ион:jon:Михей:mic:Мих:mic:Наум:na:Аввакум:hab:Авв:hab:Софония:sof:Соф:sof:Аггей:hag:Агг:hag:Захария:zec:Зах:zec:Малахия:mal:Мал:mal:Матфея:mt:Мф:mt:Марка:mr:Мк:mr:Луки:lu:Лк:lu:" & _
    "Деяния:ac:Деян:ac:Иакова:jas:Иак:jas:1 Петра:1pe:1 Пет:1pe:2 Петра:2pe:2 Пет:2pe:1Петра:1pe:1Пет:1pe:2Петра:2pe:2Пет:2pe:1 Иоанна:1jo:1 Ин:1jo:1Ин:1jo:2 Иоанна:2jo:2 Ин:2jo:2Ин:2jo:3 Иоанна:3jo:3 Ин:3jo:3Ин:3jo:1Иоанна:1jo:1Ин:1jo:2Иоанна:2jo:2Ин:2jo:3Иоанна:3jo:3Ин:3jo:" & _
    "Иоанна:joh:Ин:joh:Иуды:jude:Иуд:jude:Иуда:jude:Римлянам:ro:Рим:ro:1 Коринфянам:1co:1 Кор:1co:2 Коринфянам:2co:2 Кор:2co:1Коринфянам:1co:1Кор:1co:2Коринфянам:2co:2Кор:2co:Галатам:ga:Гал:ga:Ефесянам:eph:Еф:eph:Филиппийцам:php:Флп:php:Колоссянам:col:Кол:col:" & _
    "1 Фессалоникийцам:1th:1 Фес:1th:2 Фессалоникийцам:2th:2 Фес:2th:1Фессалоникийцам:1th:1Фес:1th:2Фессалоникийцам:2th:2Фес:2th:1 Тимофею:1ti:1 Тим:1ti:2 Тимофею:2ti:2 Тим:2ti:1 Тимофея:1ti:1 Тим:1ti:2 Тимофея:2ti:2 Тим:2ti:1Тимофею:1ti:1Тим:1ti:" & _
    "2Тимофею:2ti:2Тим:2ti:1Тимофея:1ti:1Тим:1ti:2Тимофея:2ti:2Тим:2ti:Титу:tit:Тит:tit:Филимону:phm:Флм:phm:Филимон:phm:Евреям:heb:Евр:heb:Откровение:re:Откр:re:Апок:re:"

  ' variables
  Dim versePosStart As Long
  Dim versePosEnd As Long
  Dim docRng As Word.Range
  Dim i As Integer
  Dim j As Long
  Dim endOfDocumentPos As Long
  Dim linkCount As Long
  
  ' init vars
  'Set docRng = ActiveDocument.Content
  'endOfDocumentPos = docRng.End
  linkCount = 0

  ' get bible book names and short link name accordances
  bibleBooksMatchArray = Split(bibleBooksMatch_string, ":")
  
  ' lookup all occurrence of book names
  For i = LBound(bibleBooksMatchArray) To UBound(bibleBooksMatchArray) - 2 Step 2
    ' find current book name position in document
    bookFindStr = bibleBooksMatchArray(i)
    bookShortName = bibleBooksMatchArray(i + 1)
    
    Set docRng = ActiveDocument.Content
    Dim versesPos As Collection
    
    With docRng.Find
      .Text = bookFindStr
      ' to find with end
      .Forward = False

      ' store all positions of book verses occuruences
      Set versesPos = New Collection
      While .Execute
        'Debug.Print docRng.Start & " " & docRng.End
        versesPos.Add docRng.Start
      Wend
    End With

    
    ' process book verses
    For j = 1 To versesPos.Count
      versePosStart = versesPos(j)
      linkCount = linkCount + 1
      
      ' collect test info
        versePosStart & " - " & strTheText

      Dim verseNum As String
      Dim chapterNum As String
      Dim curChar As String
      Dim isVerseNum As Boolean
      Dim isNewVerse As Boolean
      Dim setLink As Boolean
      Dim forbiddenCharsCount As Integer

      ' init vars
      curPos = versePosStart + Len(bookFindStr) + 1
      versePosEnd = curPos
      forbiddenCharsCount = 0
      isVerseNum = False
      isNewVerse = False
      setLink = False
      isVersesInNewChapter = True
      chapterNum = ""
      verseNum = ""
      bibleSiteAddress = "http://allbible.info/bible/sinodal/"  'example - "http://allbible.info/bible/sinodal/ge/1#1-7"
      ' read char by char to procees bible verses
      Do While True
        curChar = ActiveDocument.Range(curPos, curPos + 1).Text

        If IsNumeric(curChar) Then
          If isNewVerse Then
            versePosStart = curPos
            verseNum = ""
            bibleSiteAddress = "http://allbible.info/bible/sinodal/"  'example - "http://allbible.info/bible/sinodal/ge/1#1-7"
            isNewVerse = False
            versePosStart = curPos
            If isVersesInNewChapter Then
              chapterNum = ""
            End If
          End If

          If isVerseNum Then
            verseNum = verseNum & curChar
          Else
            chapterNum = chapterNum & curChar
          End If
          forbiddenCharsCount = 0
          versePosEnd = curPos
        Else
          Select Case curChar
          Case ":"
            isVerseNum = True
          Case ";"
            isNewVerse = True
            isVerseNum = False
            setLink = True
            isVersesInNewChapter = True
          Case ","
            isNewVerse = True
            isVerseNum = True
            setLink = True
            isVersesInNewChapter = False
          Case "-"
            verseNum = verseNum & "-"
          Case ")"
            setLink = True
            forbiddenCharsCount = 6
          Case " "
            ' Continue Do
          Case Else
            forbiddenCharsCount = forbiddenCharsCount + 1
            ' Continue Do
          End Select

          If setLink Then
            If Len(chapterNum) > 0 Then
              ' forms link
              bibleSiteAddress = bibleSiteAddress & bookShortName & "/" & chapterNum
              If verseNum <> "" Then
                bibleSiteAddress = bibleSiteAddress & "#" & verseNum
              End If
              
              ' select text and insert link
              ActiveDocument.Range(versePosStart, versePosEnd + 1).Select
              ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:=bibleSiteAddress
              'Set docRng = ActiveDocument.Content
              setLink = False
              curPos = curPos + Len(bibleSiteAddress) + 18
            End If
          End If
          
          If forbiddenCharsCount > 4 Then
            Exit Do
          End If
        End If
        
        curPos = curPos + 1
        
        ' get end of document position
        Set docRng = ActiveDocument.Content
        endOfDocumentPos = docRng.End
        
        ' check on out of bound of document
        If curPos + 1 > docRng.End Then
          Exit Do
        End If
        
      Loop
    Next j
  Next i
  
  Call ResetFormatToGoogleDocs
  
  ' enable screen update
  Application.ScreenUpdating = True
End Sub

' Insert figure caption
Sub InsertCaption()
  Selection.InsertCaption Label:="Рисунок", TitleAutoText:="InsertCaption1", _
    Title:=" " & Chr(150) & " Пример", Position:=wdCaptionPositionBelow, ExcludeLabel:=0
End Sub

Sub Set_crossreference_as_num()
  Application.ScreenUpdating = False
  Selection.Fields.ToggleShowCodes
  Application.Keyboard (1033)
  Application.Keyboard (1049)
  Application.Keyboard (1033)
  Selection.TypeText Text:="\# \0 "
  Selection.Fields.Update
  Application.ScreenUpdating = True
End Sub

Sub Silent_save_to_PDF()
Attribute Silent_save_to_PDF.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Silent_save_to_PDF"
'
' Silent Save_to_PDF Macro
'
  ActiveDocument.ExportAsFixedFormat OutputFileName:= _
    Replace(ActiveDocument.FullName, ".docx", ".pdf"), _
    ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:= _
    wdExportOptimizeForPrint, Range:=wdExportAllDocument, Item:= _
    wdExportDocumentContent, IncludeDocProps:=False, KeepIRM:=True, _
    CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
    BitmapMissingFonts:=True, UseISO19005_1:=False
    
  On Error GoTo errorHandler
  ActiveDocument.Close _
    SaveChanges:=wdPromptToSaveChanges, _
    OriginalFormat:=wdPromptUser
errorHandler:
  If Err = 4198 Then MsgBox "Document was not closed"
End Sub

' save to pdf
Sub Save_to_PDF()
Attribute Save_to_PDF.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Save_to_PDF"
'
' Save_to_PDF Macro
'
  With Dialogs(wdDialogFileSaveAs)
      .Format = wdFormatPDF
      .Show
  End With
End Sub

'
Sub SetDefaultTextStyle()
Attribute SetDefaultTextStyle.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.SetDefaultTextStyle"
  Application.ScreenUpdating = False
  '''''''''''''''''
  ' Clear BG Color
  Options.DefaultHighlightColorIndex = wdNoHighlight
  Selection.Range.HighlightColorIndex = wdNoHighlight
  Selection.Font.Color = wdColorAutomatic
  With Selection.ParagraphFormat
      With .Shading
        .Texture = wdTextureNone
        .ForegroundPatternColor = wdColorAutomatic
        .BackgroundPatternColor = wdColorAutomatic
      End With
      .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
      .Borders(wdBorderRight).LineStyle = wdLineStyleNone
      .Borders(wdBorderTop).LineStyle = wdLineStyleNone
      .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
      .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
      With .Borders
        .DistanceFromTop = 1
        .DistanceFromLeft = 4
        .DistanceFromBottom = 1
        .DistanceFromRight = 4
        .Shadow = False
      End With
  End With
  With Options
    .DefaultBorderLineStyle = wdLineStyleSingle
    .DefaultBorderLineWidth = wdLineWidth050pt
    .DefaultBorderColor = wdColorAutomatic
  End With

  ''''''''''''''''''''''''
  ' Set text style to mine
    With Selection.ParagraphFormat
      .LeftIndent = InchesToPoints(0)
      .RightIndent = InchesToPoints(0)
      .SpaceBefore = 0
      .SpaceBeforeAuto = False
      .SpaceAfter = 0
      .SpaceAfterAuto = False
      .LineSpacingRule = wdLineSpace1pt5
      .Alignment = wdAlignParagraphLeft
      .WidowControl = True
      .KeepWithNext = False
      .KeepTogether = False
      .PageBreakBefore = False
      .NoLineNumber = False
      .Hyphenation = True
      .FirstLineIndent = InchesToPoints(0.49)
      .OutlineLevel = wdOutlineLevelBodyText
      .CharacterUnitLeftIndent = 0
      .CharacterUnitRightIndent = 0
      .CharacterUnitFirstLineIndent = 0
      .LineUnitBefore = 0
      .LineUnitAfter = 0
      .MirrorIndents = False
      .TextboxTightWrap = wdTightNone
      .CollapsedByDefault = False
    End With
    
    With Selection.Font
      .Name = "Times New Roman"
      .Size = 14
      .Bold = False
      .Italic = False
      .Underline = wdUnderlineNone
      .UnderlineColor = wdColorAutomatic
      .StrikeThrough = False
      .DoubleStrikeThrough = False
      .Outline = False
      .Emboss = False
      .Shadow = False
      .Hidden = False
      .SmallCaps = False
      .AllCaps = False
      .Color = wdColorAutomatic
      .Engrave = False
      .Superscript = False
      .Subscript = False
      .Spacing = 0
      .Scaling = 100
      .Position = 0
      .Kerning = 0
      .Animation = wdAnimationNone
      .Ligatures = wdLigaturesNone
      .NumberSpacing = wdNumberSpacingDefault
      .NumberForm = wdNumberFormDefault
      .StylisticSet = wdStylisticSetDefault
      .ContextualAlternates = 0
    End With
    
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    Application.ScreenUpdating = True
End Sub

'
Sub remove_paragraphs()
Attribute remove_paragraphs.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.remove_paragraphs"
  Dim sText As Range
  Dim vFindText As Variant
  Dim vReplText As Variant
  Dim i As Long
  Set sText = Selection.Range
  
  Application.ScreenUpdating = False
  
  vFindText = Array("^p")
  vReplText = Array("")
  With sText.Find
    .ClearFormatting
    .Replacement.ClearFormatting
    .Forward = True
    .Wrap = wdFindStop
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
    .Format = False
    .MatchCase = True
    For i = LBound(vFindText) To UBound(vFindText)
      .Text = vFindText(i)
      .Replacement.Text = vReplText(i)
      .Execute Replace:=wdReplaceAll
    Next i
  End With
  Application.ScreenUpdating = True
End Sub

' Convert youtube links with time in seconds at end of link to Time text with it's link
Sub convertYouTubeLinksTextToTime()
Attribute convertYouTubeLinksTextToTime.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.test1"
  Dim doc As Document
  Dim link, i, linkCount
  Set doc = Application.ActiveDocument
  
  ' init values
  linkCount = 0
  
  ' Loop through all hyperlinks.
  For i = 1 To doc.Hyperlinks.Count
    link = doc.Hyperlinks(i).Address
    
    ' check if link has timestamp
    If InStrRev(link, "t=") > 0 And (InStrRev(link, "youtu.be") > 0 Or InStrRev(link, "youtube") > 0) Then
      ' get seconds count from link (in youtube specifies as parametr "t=secondsCount")
      secCount = CInt(Right(link, Len(link) - InStrRev(link, "t=") - 1))

      ' Convert seconds to time format hh:mm:ss
      If secCount < 3600 Then
        minutes = secCount \ 60
        seconds = secCount Mod 60
        myTime = minutes & ":" & Format(CStr(seconds), "00")
      Else
        hours = secCount \ 3600
        minutes = (secCount - 3600 * (secCount \ 3600)) \ 60
        seconds = (secCount - 3600 * (secCount \ 3600)) Mod 60
        
        myTime = hours & ":" & Format(CStr(minutes), "00") & ":" & Format(CStr(seconds), "00")
      End If
      
      ' set link text
      doc.Hyperlinks(i).TextToDisplay = myTime
      ' counts converted links
      linkCount = linkCount + 1
    Else
      ' highlight non youtube links or without timestamp
      With doc.Hyperlinks(i).Range
        .Bold = 0
        .Italic = 0
        .Shading.BackgroundPatternColor = wdColorGray375
        .HighlightColorIndex = wdYellow
      End With
    End If
  Next
  
  Call ResetFormatToGoogleDocs
  
  If doc.Hyperlinks.Count <> linkCount Then
      MsgBox "Total links count: " & doc.Hyperlinks.Count & vbNewLine & "Link(s) converted: " & linkCount
  Else
      
  End If
End Sub

' Reset Format To Google Docs Style
Sub ResetFormatToGoogleDocs()
Attribute ResetFormatToGoogleDocs.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro1"
  Selection.WholeStory
  Selection.ClearFormatting
  Selection.ParagraphFormat.LineSpacing = LinesToPoints(1)
  Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
  With Selection.ParagraphFormat
    .LeftIndent = CentimetersToPoints(0)
    .SpaceBeforeAuto = False
    .SpaceAfterAuto = False
  End With
  Selection.Font.Name = "Arial"
  Selection.Font.Size = 12
End Sub



Sub getAllLinks()
  Dim doc As Document
  Dim link, i, linkCount, prevLink, linkText, linkAddress
  Set doc = Application.ActiveDocument

  
  ' init values
  linkCount = 0
  
  ' clear formatting for whole document
  Selection.WholeStory
  Selection.ClearFormatting
  
  ' move cursor to end of document
  Selection.EndKey Unit:=wdStory
  Selection.TypeParagraph
  
  ' pick up current cursor position
  Set currentPosition = Selection.Range
  
  ' Loop through all hyperlinks.
  For i = 1 To doc.Hyperlinks.Count
    linkAddress = doc.Hyperlinks(i).Address
    linkText = doc.Hyperlinks(i).TextToDisplay
    
    If InStr(LCase(linkText), "недельная глава") > 0 Then
      ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:=linkAddress, _
      SubAddress:="", ScreenTip:="", TextToDisplay:=linkText
      Selection.TypeParagraph
    End If
  Next
      
  currentPosition.Select 'return cursor to original position
  
  ' delete all before
  Selection.HomeKey Unit:=wdStory, Extend:=wdExtend
  Selection.Delete Unit:=wdCharacter, Count:=1
  
  MsgBox "Total links count: " & doc.Hyperlinks.Count
End Sub