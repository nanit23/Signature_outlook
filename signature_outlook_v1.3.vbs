On Error Resume Next

Set objSysInfo = CreateObject("ADSystemInfo")

strUser = objSysInfo.UserName
Set objUser = GetObject("LDAP://" & strUser)

strZpov = "Ñ óâàæåíèåì,"
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

objTable.Rows(1).select  ' ñòðîêà 1, âûäåëÿåì 
objSelection.Cells.Merge ' îáüåäåíÿåì â åäèíóþ ñòðîêó âî âñþ øèðèíó òàáëè÷êè

objTable.Cell(1, 1).select ' âûäåëÿåì 1 ñòðîêó è çàäàåì åé øèðèíó
objTable.Cell(1, 1).Width = 200

' íà÷èíàåì íàïîëíÿòü ÿ÷åéêó òåêñòîì î ñîòðóäíèêå ( ÔÈÎ, Äîëæíîñòü, îáüåêò, ìîá, ïî÷òà)
' àäðåñ ïî÷òû äåëàåì êëèêàáåëüíûì äëÿ áûñòðîé îòïðàâêè ïèñüìà mailto:

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
objSelection.TypeText "Òåë. ðàá.:  " & strfax & ", äîá.: " & strintPhone
end if
if len(trim(strPhone))<>0 then
objSelection.TypeText CHR(11)
objSelection.TypeText "Òåë. ìîá.: " & strPhone
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

objSelection.TypeText "×èòàÿ äàííîå ñîîáùåíèå, Âû òàêæå ñîãëàøàåòåñü ñ " 
Set objLink = objSelection.Hyperlinks.Add(objSelection.Range,"https://aldocoppola.ru/disclaimer/",,"", "îãðàíè÷åíèåì îá îòâåòñòâåííîñòè")
objLink.Range.Font.Name = "Calibri" 
objLink.Range.Font.Size = 9
objLink.Range.Font.Bold = false
'objLink.Range.Font.underline = false
'objLink.Range.Font.Color = RGB (101,114,118)

Set objSelection = objDoc.Range()

objSignatureEntries.Add "Ëè÷íàÿ ïîäïèñü", objSelection
objSignatureObject.NewMessageSignature = "Ëè÷íàÿ ïîäïèñü"

objDoc.Saved = True

' Create Short Standard Signature

strZpov = "Ñ óâàæåíèåì,"
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

'ôîðìèðóåì òàáëè÷êó â êîòîðóþ áóäóò ïîäñòàâëåíû íóæíûå çàïèñè â ñîîòâåòñòâóþùèå áëîêè.
'áîëüøàÿ ïîäïèñü ïðåäñòàâëÿåò èç ñåáÿ òàáëè÷êó èç 3 ñòðî÷íûõ áëîêîâ, 2 ñòðîêà ðàçäåëåíà íà 2 ÿ÷åéêè

objDoc.Tables.Add objRange,1,1
Set objTable = objDoc.Tables(1)

objTable.Rows(1).select  ' ñòðîêà 1, âûäåëÿåì 
objSelection.Cells.Merge ' îáüåäåíÿåì â åäèíóþ ñòðîêó âî âñþ øèðèíó òàáëè÷êè

objTable.Cell(1, 1).select ' âûäåëÿåì 1 ñòðîêó è çàäàåì åé øèðèíó
objTable.Cell(1, 1).Width = 200

' íà÷èíàåì íàïîëíÿòü ÿ÷åéêó òåêñòîì î ñîòðóäíèêå ( ÔÈÎ, Äîëæíîñòü, îáüåêò, ìîá, ïî÷òà)
' àäðåñ ïî÷òû äåëàåì êëèêàáåëüíûì äëÿ áûñòðîé îòïðàâêè ïèñüìà mailto:

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
objSelection.TypeText "Òåë. ðàá.:  " & strfax & ", äîá.: " & strintPhone
end if
if len(trim(strPhone))<>0 then
objSelection.TypeText CHR(11)
objSelection.TypeText "Òåë. ìîá.: " & strPhone
end if

objSelection.TypeText CHR(11)

objselection.font.name = "Calibri"
objSelection.Font.Size = "9"
objSelection.Font.Color = RGB(101,114,118)
objSelection.TypeText "_____________________________________________________________________________"
objSelection.TypeText CHR(11)
objSelection.Font.underline = false

objSelection.TypeText "×èòàÿ äàííîå ñîîáùåíèå, Âû òàêæå ñîãëàøàåòåñü ñ " 
Set objLink = objSelection.Hyperlinks.Add(objSelection.Range,"https://aldocoppola.ru/disclaimer/",,"", "îãðàíè÷åíèåì îá îòâåòñòâåííîñòè")
objLink.Range.Font.Name = "Calibri" 
objLink.Range.Font.Size = 9
objLink.Range.Font.Bold = false
'objLink.Range.Font.underline = false
'objLink.Range.Font.Color = RGB (101,114,118)

Set objSelection = objDoc.Range()
 
objSignatureEntries.Add "Ëè÷íàÿ ïîäïèñü (Êîðîòêàÿ)", objSelection
objSignatureObject.ReplyMessageSignature = "Ëè÷íàÿ ïîäïèñü (Êîðîòêàÿ)"
 
objDoc.Saved = True
objDoc.Close
objWord.Quit
objOutlook.Quit

Dim WshShell 
 
set WshShell = WScript.CreateObject("WScript.Shell") 

'WshShell.Run "taskkill /f /IM WINWORD.EXE",0
'WshShell.Run "taskkill /f /IM OUTLOOK.EXE",0

WScript.Quit
