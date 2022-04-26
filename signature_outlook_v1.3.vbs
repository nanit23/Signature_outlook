On Error Resume Next

Set objSysInfo = CreateObject("ADSystemInfo")

strUser = objSysInfo.UserName
Set objUser = GetObject("LDAP://" & strUser)

strZpov = "С уважением,"
strPostIndex = ObjUser.postalCode
strName = objUser.FullName 
strTitle = objUser.Title
strDepartment = objUser.Department
strCompany = objUser.Company
strPhone = objUser.telephoneNumber
strweb = objuser.wWWHomePage
strCity = objuser.l
strStreet = objuser.streetAddress
strOblast = objuser.st
strCountry = objuser.c
strfax = objuser.facsimileTelephoneNumber
strIntPhone = objuser.ipPhone
strEmail = objuser.mail

Set objWord = CreateObject("Word.Application")
Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection
Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature
Set objSignatureEntries = objSignatureObject.EmailSignatureEntries
Set objRange = objDoc.Range()
strLogo = "https://aldocoppola.ru/images/aldo-mail.png"

'формируем табличку в которую будут подставлены нужные записи в соответствующие блоки.
'большая подпись представляет из себя табличку из 3 строчных блоков, 2 строка разделена на 2 ячейки

objDoc.Tables.Add objRange,1,1
Set objTable = objDoc.Tables(1)

objTable.Rows(1).select  ' строка 1, выделяем 
objSelection.Cells.Merge ' обьеденяем в единую строку во всю ширину таблички

objTable.Cell(1, 1).select ' выделяем 1 строку и задаем ей ширину
objTable.Cell(1, 1).Width = 200

' начинаем наполнять ячейку текстом о сотруднике ( ФИО, Должность, обьект, моб, почта)
' адрес почты делаем кликабельным для быстрой отправки письма mailto:

objSelection.ParagraphFormat.Space1
objSelection.TypeText CHR(11)
objSelection.TypeText CHR(11)
objSelection.TypeText CHR(11)
objselection.font.name = "Arial"
objSelection.Font.Size = "9"
objSelection.Font.Color = RGB(101,114,118)
objSelection.TypeText strZpov
objSelection.TypeText CHR(11)
objselection.font.name = "Arial"
objSelection.Font.Size = "9"
objSelection.Font.Color = RGB(175,45,55)
objSelection.Font.Bold = true
objSelection.TypeText strName
objSelection.Font.Bold = false
objselection.font.name = "Arial"
objSelection.Font.Size = "9"
objSelection.Font.Color = RGB(101,114,118)
if len(trim(strTitle))<>0 then
objSelection.TypeText CHR(11)
objSelection.TypeText strTitle & " " & strDepartment
end if
objSelection.TypeText CHR(11)
objSelection.TypeText CHR(11)
objselection.font.name = "Arial"
objSelection.Font.Size = "9"
objSelection.Font.Color = RGB(175,45,55)
objSelection.Font.Bold = true
objSelection.TypeText strCompany
objSelection.Font.Bold = false
objselection.font.name = "Arial"
objSelection.Font.Size = "9"
objSelection.Font.Color = RGB(101,114,118)
if len(trim(strStreet))<>0 then
objSelection.TypeText CHR(11)
objSelection.TypeText strPostIndex & ", "
if len(trim(strOblast))<>0 then
objSelection.TypeText strOblast & ", "
end if
if len(trim(strCity))<>0 then
objSelection.TypeText strCity & ", "
end if
objSelection.TypeText CHR(11)
objSelection.TypeText strStreet
end if
objSelection.TypeText CHR(11)
objSelection.TypeText CHR(11)
objSelection.Font.underline = true
objselection.font.name = "Arial"
objSelection.Font.Size = "9"
objSelection.Font.Color = RGB(101,114,118)
objSelection.TypeText strWeb
objSelection.Font.Bold = false
objSelection.Font.underline = false
objSelection.TypeText CHR(11)
if len(trim(strintPhone))<>0 then
objSelection.TypeText "Тел. раб.:  " & strfax & ", доб.: " & strintPhone
end if
if len(trim(strPhone))<>0 then
objSelection.TypeText CHR(11)
objSelection.TypeText "Тел. моб.: " & strPhone
end if
objSelection.TypeText CHR(11)
objSelection.TypeText CHR(11)

objSelection.InlineShapes.AddPicture(strLogo)

objSelection.TypeText CHR(11)
objSelection.TypeText CHR(11)

objSelection.Font.underline = true
objselection.font.name = "Arial"
objSelection.Font.Size = "9"
objSelection.Font.Color = RGB(101,114,118)
objLink = objSelection.Hyperlinks.Add(objSelection.Range,"https://t.me/aldocoppola",,"Telegram Aldo Coppola", "Telegram")


objSelection.Font.underline = false
objselection.font.name = "Calibri"
objSelection.Font.Size = "11"
objSelection.Font.Color = RGB(101,114,118)
objSelection.TypeText "  |  "


objSelection.Font.underline = true
objselection.font.name = "Arial"
objSelection.Font.Size = "9"
objSelection.Font.Color = RGB(101,114,118)
objLink = objSelection.Hyperlinks.Add(objSelection.Range,"https://vk.com/aldocoppolarussia",,"VK Aldo Coppola", "VK")

objSelection.TypeText CHR(11)

objselection.font.name = "Calibri"
objSelection.Font.Size = "9"
objSelection.Font.Color = RGB(101,114,118)
objSelection.TypeText "_____________________________________________________________________________"
objSelection.TypeText CHR(11)
objSelection.Font.underline = false

objSelection.TypeText "Читая данное сообщение, Вы также соглашаетесь с " 
Set objLink = objSelection.Hyperlinks.Add(objSelection.Range,"https://aldocoppola.ru/disclaimer/",,"", "ограничением об ответственности")
objLink.Range.Font.Name = "Calibri" 
objLink.Range.Font.Size = 9
objLink.Range.Font.Bold = false
'objLink.Range.Font.underline = false
'objLink.Range.Font.Color = RGB (101,114,118)

Set objSelection = objDoc.Range()

objSignatureEntries.Add "Личная подпись", objSelection
objSignatureObject.NewMessageSignature = "Личная подпись"

objDoc.Saved = True

' Create Short Standard Signature

strZpov = "С уважением,"
strPostIndex = ObjUser.postalCode
strName = objUser.FullName 
strTitle = objUser.Title
strDepartment = objUser.Department
strCompany = objUser.Company
strPhone = objUser.telephoneNumber
strweb = objuser.wWWHomePage
strgorod = objuser.l
strstreet = objuser.streetAddress
strfax = objuser.facsimileTelephoneNumber
strIntPhone = objuser.ipPhone
strEmail = objuser.mail

Set objWord = CreateObject("Word.Application")
Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection
Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature
Set objSignatureEntries = objSignatureObject.EmailSignatureEntries
Set objRange = objDoc.Range()
strLogo = "https://aldocoppola.ru/images/aldo-mail.png"

'формируем табличку в которую будут подставлены нужные записи в соответствующие блоки.
'большая подпись представляет из себя табличку из 3 строчных блоков, 2 строка разделена на 2 ячейки

objDoc.Tables.Add objRange,1,1
Set objTable = objDoc.Tables(1)

objTable.Rows(1).select  ' строка 1, выделяем 
objSelection.Cells.Merge ' обьеденяем в единую строку во всю ширину таблички

objTable.Cell(1, 1).select ' выделяем 1 строку и задаем ей ширину
objTable.Cell(1, 1).Width = 200

' начинаем наполнять ячейку текстом о сотруднике ( ФИО, Должность, обьект, моб, почта)
' адрес почты делаем кликабельным для быстрой отправки письма mailto:

objSelection.ParagraphFormat.Space1
objSelection.TypeText CHR(11)
objSelection.TypeText CHR(11)
objSelection.TypeText CHR(11)
objselection.font.name = "Arial"
objSelection.Font.Size = "9"
objSelection.Font.Color = RGB(101,114,118)
objSelection.TypeText strZpov
objSelection.TypeText CHR(11)
objselection.font.name = "Arial"
objSelection.Font.Size = "9"
objSelection.Font.Color = RGB(175,45,55)
objSelection.Font.Bold = true
objSelection.TypeText strName
objSelection.Font.Bold = false
objselection.font.name = "Arial"
objSelection.Font.Size = "9"
objSelection.Font.Color = RGB(101,114,118)
if len(trim(strTitle))<>0 then
objSelection.TypeText CHR(11)
objSelection.TypeText strTitle & " " & strDepartment
end if
objSelection.TypeText CHR(11)
objSelection.TypeText CHR(11)
objselection.font.name = "Arial"
objSelection.Font.Size = "9"
objSelection.Font.Color = RGB(175,45,55)
objSelection.Font.Bold = true
objSelection.TypeText strCompany
objSelection.Font.Bold = false
objselection.font.name = "Arial"
objSelection.Font.Size = "9"
objSelection.Font.Color = RGB(101,114,118)
if len(trim(strStreet))<>0 then
objSelection.TypeText CHR(11)
objSelection.TypeText strPostIndex & ", "
if len(trim(strOblast))<>0 then
objSelection.TypeText strOblast & ", "
end if
if len(trim(strCity))<>0 then
objSelection.TypeText strCity & ", "
end if
objSelection.TypeText CHR(11)
objSelection.TypeText strStreet
end if
objSelection.TypeText CHR(11)
objSelection.TypeText CHR(11)
objSelection.Font.underline = true
objselection.font.name = "Arial"
objSelection.Font.Size = "9"
objSelection.Font.Color = RGB(101,114,118)
objSelection.TypeText strWeb
objSelection.Font.Bold = false
objSelection.Font.underline = false
objSelection.TypeText CHR(11)
if len(trim(strintPhone))<>0 then
objSelection.TypeText "Тел. раб.:  " & strfax & ", доб.: " & strintPhone
end if
if len(trim(strPhone))<>0 then
objSelection.TypeText CHR(11)
objSelection.TypeText "Тел. моб.: " & strPhone
end if

objSelection.TypeText CHR(11)

objselection.font.name = "Calibri"
objSelection.Font.Size = "9"
objSelection.Font.Color = RGB(101,114,118)
objSelection.TypeText "_____________________________________________________________________________"
objSelection.TypeText CHR(11)
objSelection.Font.underline = false

objSelection.TypeText "Читая данное сообщение, Вы также соглашаетесь с " 
Set objLink = objSelection.Hyperlinks.Add(objSelection.Range,"https://aldocoppola.ru/disclaimer/",,"", "ограничением об ответственности")
objLink.Range.Font.Name = "Calibri" 
objLink.Range.Font.Size = 9
objLink.Range.Font.Bold = false
'objLink.Range.Font.underline = false
'objLink.Range.Font.Color = RGB (101,114,118)

Set objSelection = objDoc.Range()
 
objSignatureEntries.Add "Личная подпись (Короткая)", objSelection
objSignatureObject.ReplyMessageSignature = "Личная подпись (Короткая)"
 
objDoc.Saved = True
objDoc.Close
objWord.Quit
objOutlook.Quit

Dim WshShell 
 
set WshShell = WScript.CreateObject("WScript.Shell") 

'WshShell.Run "taskkill /f /IM WINWORD.EXE",0
'WshShell.Run "taskkill /f /IM OUTLOOK.EXE",0

WScript.Quit