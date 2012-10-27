<%
  ' NOTE: There are a lot of configuration parameters in this fist section. Please read the
  '       comments located before each configuration parameters. Be sure to keep a copy of
  '       the original file before changing anything.

  '*************************************************************************************************
  '** Possible values are "CDOSYS", "CDONTS", "ASPEmail", "ASPMail", "Jmail"  ' line #5            *
  '*************************************************************************************************
  Const C_MAIL_USE = "CDOSYS"
  
  
  '*************************************************************************************************
  '** If your email server is not your web server, then enter the address in the following line    *
  '** Note for Godaddy users: The mail server is "relay-hosting.secureserver.net"                  *
  '** Note for 1and1.com users: The mail server is "mrelay.perfora.net"                            *
  '** Note for 1and1.co.uk and 1and1.de users: The mail server is "mrvnet.kundenserver.de"         *
  '*************************************************************************************************
  Const SMTP_SERVER = "localhost"     
  

  '*************************************************************************************************
  '** Authentication required? If yes, specify username and password                               *
  '*************************************************************************************************
  Const HDW_USE_AUTHENTICATION = False
  Const HDW_AUTH_USERNAME = ""
  Const HDW_AUTH_PASSWORD = ""
  
  
  '*************************************************************************************************
  '** send CC to another emails, separated by semicolon,                                           *
  '** example: Const HDW_SEND_CC_TO = "email-1@domain.com;email-2@domain.com"                      *
  '*************************************************************************************************
  Const HDW_SEND_CC_TO = ""
  

  '*************************************************************************************************
  '** send CC to the user with a "Thank you message"                                               *
  '** If enabled you MUST specify the name of the user's email field in the form                   *
  '*************************************************************************************************
  Const HDW_SEND_THANKYOU = False
  Const HDW_USER_EMAIL_FIELD_NAME = "email"
  Const HDW_INCLUDE_SUBMITTED_DATA = False
  Const HDW_THANKYOU_SUBJECT = "Thank you for your message."
  Const HDW_THANKYOU_MSG = "Thank you for your message. We will reply you as soon as possible."
  
  
  '*************************************************************************************************
  '** Email subject, you can change it here                                                        *
  '*************************************************************************************************
  Dim emailsubject
  emailsubject= "Form sent from " & Request.ServerVariables("SERVER_NAME")
  
  
  '*************************************************************************************************
  '** Use the email and/or subject entered by the user as FROM/SUBJECT.                            *
  '** IMPORTANT!!! You need to specify the name of the email/subject fields in your form.          *
  '*************************************************************************************************
  Const HDW_FROM_EMAIL_FIELD_NAME = ""
  Const HDW_SUBJECT_FIELD_NAME = ""

  
  '*************************************************************************************************  
  '** Exclude fields, add one AddExcludedField call for each excluded field                        *
  '*************************************************************************************************
  Dim excluded_fields(100), excluded_fields_count : excluded_fields_count = 0
  AddExcludedField "submit"
  AddExcludedField "Submit"
  AddExcludedField "hdcaptcha"
  AddExcludedField "hdwfail"
  AddExcludedField "sample_excluded_field" 
  
    
  '*************************************************************************************************
  '** Fix destination email. This helps to increate the form security                              *
  '** Example: Const HDW_FIX_DESTINATION_EMAIL = "email@sample.com"                                *
  '*************************************************************************************************
  Const HDW_FIX_DESTINATION_EMAIL = ""    


  '*************************************************************************************************
  '** Set this param to true to resend the info as GET parameters to the "Thank You" page.         *
  '*************************************************************************************************
  Const HDW_ENABLE_DEBUG_MESSAGES = False  

  
  '*************************************************************************************************
  '** Set this param to true to resend the info as GET parameters to the "Thank You" page.         *
  '*************************************************************************************************
  Const HOTDW_RESEND_PARAMS = False

  '*************************************************************************************************
  '** Use plain text instead HTML emails. This helps to prevent spam filters.                      *
  '*************************************************************************************************
  Const USE_PLAIN_TEXT_EMAILS = False

  '*************************************************************************************************
  '** Some antivirus prevent the filesystem object. In that case you can disable it.               *
  '*************************************************************************************************
  Const USE_FILESYSTEM_OBJECT = True

  
%> 
<!--METADATA TYPE="typelib"
      UUID="00000205-0000-0010-8000-00AA006D2EA4"
     NAME="ADODB Type Library"
-->
<% 
  
  Server.ScriptTimeout = 1000000000
  Response.Expires = 0
  Response.Buffer = True    
  Const SMTP_PORT = 25 
  Const HDW_F2M_EMAIL = "hdwemail"
  Const HDW_F2M_OK = "hdwok"
  Const HDW_F2M_NO_OK = "hdwnook"  
  Dim localpath 
  localpath = Server.MapPath("Form2Mail.asp.mdb")
  localpath = Left(localpath, Len(localpath)-Len("Form2Mail.asp.mdb") )  
  Dim fso, MyFile  
  Dim Attachments, globalbuffer
  Attachments = False  ' Do not modify this  
  Function getCountryID(ip)
      Dim cip
      cip = IPAddress2IPNumber(ip)
      On Error Resume Next
      Dim conn, index, rs
      Set conn = Server.CreateObject("ADODB.Connection")
      conn.open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("Form2Mail.asp.mdb") & ";Persist Security Info=False")    
      Set rs = conn.Execute("SELECT country FROM iptocountry WHERE ip1<="&cip&" AND ip2>="&cip)
      If rs.EOF Then getCountryID = "LOCAL, INTRANET OR UNKNOWN" Else getCountryID = countryname(rs.Fields.Item("country"))
  End Function    
  Function IPAddress2IPNumber(IPaddress)
    Dim i, pos, PrevPos, num
    If IPaddress = "" Then
      IPAddress2IPNumber = 0
    Else
      For i = 1 To 4
          pos = InStr(PrevPos + 1, IPaddress, ".", 1)
          If i = 4 Then pos = Len(IPaddress) + 1
          num = Int(Mid(IPaddress, PrevPos + 1, pos - PrevPos - 1))
          PrevPos = pos
          IPAddress2IPNumber = ((num Mod 256) * (256 ^ (4 - i))) + IPAddress2IPNumber
      Next
    End If
  End Function  
  Function ereg(find,str)
      if Instr(1, str, find, 1) <> 0 Then
          ereg = True
      Else
          ereg = False
      End If    
  End Function  
  Function ckbrowser(user_agent)
      Dim browser
 	  if((ereg("Netscape", user_agent))) Then
 	    browser = "Netscape"
 	  elseif(ereg("Firefox", user_agent)) Then
 	    browser = "Firefox"
      elseif(ereg("Safari", user_agent)) Then
        browser = "Safari"
      elseif(ereg("SAFARI", user_agent)) Then
        browser = "SAFARI"
      elseif(ereg("MSIE", user_agent)) Then
        browser = "MSIE"
      elseif(ereg("Lynx", user_agent)) Then
        browser = "Lynx"
      elseif(ereg("Opera", user_agent)) Then
        browser = "Opera"
      elseif(ereg("Gecko", user_agent)) Then 
        browser = "Mozilla"
      elseif(ereg("WebTV", user_agent)) Then 
        browser = "WebTV"
      elseif(ereg("Konqueror", user_agent)) Then 
        browser = "Konqueror"
      else
        browser = "bot"
      End If      
      ckbrowser = browser
  End Function
  
  ' country list  
  Dim countryname(212)
  countryname(0)  = "UNKNOWN" : countryname(1)  = "AFGHANISTAN"
  countryname(2)  = "ALBANIA" : countryname(3)  = "ALGERIA"
  countryname(4)  = "AMERICAN SAMOA" : countryname(5)  = "ANDORRA"
  countryname(6)  = "ANGOLA" : countryname(7)  = "ANTIGUA AND BARBUDA"
  countryname(8)  = "ARGENTINA" : countryname(9)  = "ARMENIA"
  countryname(10) = "AUSTRALIA": countryname(11) = "AUSTRIA"
  countryname(12) = "AZERBAIJAN" : countryname(13) = "BAHAMAS"
  countryname(14) = "BAHRAIN"  :  countryname(15) = "BANGLADESH"
  countryname(16) = "BARBADOS" :  countryname(17) = "BELARUS"
  countryname(18) = "BELGIUM" :  countryname(19) = "BELIZE"
  countryname(20) = "BENIN" :  countryname(21) = "BERMUDA"
  countryname(22) = "BHUTAN" 
  countryname(23) = "BOLIVIA"
  countryname(24) = "BOSNIA AND HERZEGOVINA"
  countryname(25) = "BOTSWANA"
  countryname(26) = "BRAZIL"
  countryname(27) = "BRITISH INDIAN OCEAN TERRITORY"
  countryname(28) = "BRUNEI DARUSSALAM"
  countryname(29) = "BULGARIA"
  countryname(30) = "BURKINA FASO"
  countryname(31) = "BURUNDI"
  countryname(32) = "CAMBODIA"
  countryname(33) = "CAMEROON"
  countryname(34) = "CANADA"
  countryname(35) = "CAPE VERDE"
  countryname(36) = "CAYMAN ISLANDS"
  countryname(37) = "CENTRAL AFRICAN REPUBLIC"
  countryname(38) = "CHAD"
  countryname(39) = "CHILE"
  countryname(40) = "CHINA"
  countryname(41) = "COLOMBIA"
  countryname(42) = "COMOROS"
  countryname(43) = "CONGO"
  countryname(44) = "COOK ISLANDS"
  countryname(45) = "COSTA RICA"
  countryname(46) = "COTE D""IVOIRE"
  countryname(47) = "CROATIA"
  countryname(48) = "CUBA"
  countryname(49) = "CYPRUS"
  countryname(50) = "CZECH REPUBLIC"
  countryname(51) = "DENMARK"
  countryname(52) = "DJIBOUTI"
  countryname(53) = "DOMINICAN REPUBLIC"
  countryname(54) = "EAST TIMOR"
  countryname(55) = "ECUADOR"
  countryname(56) = "EGYPT"
  countryname(57) = "EL SALVADOR"
  countryname(58) = "EQUATORIAL GUINEA"
  countryname(59) = "ERITREA"
  countryname(60) = "ESTONIA"
  countryname(61) = "ETHIOPIA"
  countryname(62) = "FALKLAND ISLANDS (MALVINAS)"
  countryname(63) = "FAROE ISLANDS"
  countryname(64) = "FIJI"
  countryname(65) = "FINLAND"
  countryname(66) = "FRANCE"
  countryname(67) = "FRENCH POLYNESIA"
  countryname(68) = "GABON"
  countryname(69) = "GAMBIA"
  countryname(70) = "GEORGIA"
  countryname(71) = "GERMANY"
  countryname(72) = "GHANA"
  countryname(73) = "GIBRALTAR"
  countryname(74) = "GREECE"
  countryname(75) = "GREENLAND"
  countryname(76) = "GRENADA"
  countryname(77) = "GUADELOUPE"
  countryname(78) = "GUAM"
  countryname(79) = "GUATEMALA"
  countryname(80) = "GUINEA"
  countryname(81) = "GUINEA-BISSAU"
  countryname(82) = "HAITI"
  countryname(83) = "HOLY SEE(VATICAN CITY STATE)"
  countryname(84) = "HONDURAS"
  countryname(85) = "HONG KONG"
  countryname(86) = "HUNGARY"
  countryname(87) = "ICELAND"
  countryname(88) = "INDIA"
  countryname(89) = "INDONESIA"
  countryname(90) = "IRAQ"
  countryname(91) = "IRELAND"
  countryname(92) = "ISLAMIC REPUBLIC OF IRAN"
  countryname(93) = "ISRAEL"
  countryname(94) = "ITALY"
  countryname(95) = "JAMAICA"
  countryname(96) = "JAPAN"
  countryname(97) = "JORDAN"
  countryname(98) = "KAZAKHSTAN"
  countryname(99) = "KENYA"
  countryname(100) = "KIRIBATI"
  countryname(101) = "KUWAIT"
  countryname(102) = "KYRGYZSTAN"
  countryname(103) = "LAO PEOPLE""S DEMOCRATIC REPUBLIC"
  countryname(104) = "LATVIA"
  countryname(105) = "LEBANON"
  countryname(106) = "LESOTHO"
  countryname(107) = "LIBERIA"
  countryname(108) = "LIBYAN ARAB JAMAHIRIYA"
  countryname(109) = "LIECHTENSTEIN"
  countryname(110) = "LITHUANIA"
  countryname(111) = "LUXEMBOURG"
  countryname(112) = "MACAO"
  countryname(113) = "MADAGASCAR"
  countryname(114) = "MALAWI"
  countryname(115) = "MALAYSIA"
  countryname(116) = "MALDIVES"
  countryname(117) = "MALI"
  countryname(118) = "MALTA"
  countryname(119) = "MARTINIQUE"
  countryname(120) = "MAURITANIA"
  countryname(121) = "MAURITIUS"
  countryname(122) = "MEXICO"
  countryname(123) = "MONACO"
  countryname(124) = "MONGOLIA"
  countryname(125) = "MOROCCO"
  countryname(126) = "MOZAMBIQUE"
  countryname(127) = "MYANMAR"
  countryname(128) = "NAMIBIA"
  countryname(129) = "NAURU"
  countryname(130) = "NEPAL"
  countryname(131) = "NETHERLANDS"
  countryname(132) = "NETHERLANDS ANTILLES"
  countryname(133) = "NEW CALEDONIA"
  countryname(134) = "NEW ZEALAND"
  countryname(135) = "NICARAGUA"
  countryname(136) = "NIGER"
  countryname(137) = "NIGERIA"
  countryname(138) = "NORTHERN MARIANA ISLANDS"
  countryname(139) = "NORWAY"
  countryname(140) = "OMAN"
  countryname(141) = "PAKISTAN"
  countryname(142) = "PALAU"
  countryname(143) = "PALESTINIAN TERRITORY"
  countryname(144) = "PANAMA"
  countryname(145) = "PAPUA NEW GUINEA"
  countryname(146) = "PARAGUAY"
  countryname(147) = "PERU"
  countryname(148) = "PHILIPPINES"
  countryname(149) = "POLAND"
  countryname(150) = "PORTUGAL"
  countryname(151) = "PUERTO RICO"
  countryname(152) = "QATAR"
  countryname(153) = "REPUBLIC OF KOREA"
  countryname(154) = "REPUBLIC OF MOLDOVA"
  countryname(155) = "REUNION"
  countryname(156) = "ROMANIA"
  countryname(157) = "RUSSIAN FEDERATION"
  countryname(158) = "RWANDA"
  countryname(159) = "SAMOA"
  countryname(160) = "SAN MARINO"
  countryname(161) = "SAO TOME AND PRINCIPE"
  countryname(162) = "SAUDI ARABIA"
  countryname(163) = "SENEGAL"
  countryname(165) = "SERBIA AND MONTENEGRO"
  countryname(166) = "SEYCHELLES"
  countryname(167) = "SIERRA LEONE"
  countryname(168) = "SINGAPORE"
  countryname(169) = "SLOVAKIA"
  countryname(170) = "SLOVENIA"
  countryname(171) = "SOLOMON ISLANDS"
  countryname(172) = "SOMALIA"
  countryname(173) = "SOUTH AFRICA"
  countryname(174) = "SPAIN"
  countryname(175) = "SRI LANKA"
  countryname(176) = "SUDAN"
  countryname(177) = "SURINAME"
  countryname(178) = "SWAZILAND"
  countryname(179) = "SWEDEN"
  countryname(180) = "SWITZERLAND"
  countryname(181) = "SYRIAN ARAB REPUBLIC"
  countryname(182) = "TAIWAN"
  countryname(183) = "TAJIKISTAN"
  countryname(184) = "THAILAND"
  countryname(185) = "THE DEMOCRATIC REPUBLIC OF THE CONGO"
  countryname(186) = "THE FORMER YUGOSLAV REPUBLIC OF MACEDONIA"
  countryname(187) = "TOGO"
  countryname(188) = "TOKELAU"
  countryname(189) = "TONGA"
  countryname(190) = "TRINIDAD AND TOBAGO"
  countryname(191) = "TUNISIA"
  countryname(192) = "TURKEY"
  countryname(193) = "TURKMENISTAN"
  countryname(194) = "TUVALU"
  countryname(195) = "UGANDA"
  countryname(196) = "UKRAINE"
  countryname(197) = "UNITED ARAB EMIRATES"
  countryname(198) = "UNITED KINGDOM"
  countryname(199) = "UNITED REPUBLIC OF TANZANIA"
  countryname(200) = "UNITED STATES"
  countryname(201) = "URUGUAY"
  countryname(202) = "UZBEKISTAN"
  countryname(203) = "VANUATU"
  countryname(204) = "VENEZUELA"
  countryname(205) = "VIET NAM"
  countryname(206) = "VIRGIN ISLANDS"
  countryname(207) = "WESTERN SAHARA"
  countryname(208) = "YEMEN"
  countryname(209) = "ZAMBIA"
  countryname(210) = "ZIMBABWE"

  Sub AddExcludedField(value)
     excluded_fields_count = excluded_fields_count + 1 
     excluded_fields(excluded_fields_count) = value
  End Sub
  Function notInThisArray(value)
    Dim i, found
    found = False
    For i = 1 To excluded_fields_count
       If (excluded_fields(i)=value) Then found = True
    Next         
    notInThisArray = Not found
  End Function

  Dim emailaddress, fromaddress, body, item, getStr      
    
    
  body ="SUBMITTED INFORMATION<br />" &_
        "***************************<br />"
  getStr = ""
  If (InStr(1,Request.ServerVariables("CONTENT_TYPE"), "multipart/form-data", 1) <= 0) Then
     Dim name
     for i = 1 to Request.Form.Count
       for each name in Request.Form
         if Request.Form(name) is Request.Form(i) AND (name <> HDW_F2M_OK) And (name <> HDW_F2M_NO_OK) And (name <> HDW_F2M_EMAIL) And notInThisArray(name) then
           body = body & ""&name&": "&Request.Form(name)&"<br /><br />"
           getStr = getStr & "&"&name&"="&Server.URLEncode(Request.Form(name))
         end if
       next
     next     
     for i = 1 to Request.QueryString.Count
       for each name in Request.QueryString
         if Request.QueryString(name) is Request.QueryString(i) AND (name <> HDW_F2M_OK) And (name <> HDW_F2M_NO_OK) And (name <> HDW_F2M_EMAIL) And notInThisArray(name) then
           body = body & ""&name&": "&Request.QueryString(name)&"<br /><br />"
           getStr = getStr & "&"&name&"="&Server.URLEncode(Request.QueryString(name))
         end if
       next
     next       
     
     emailaddress = Replace(Request(HDW_F2M_EMAIL),"+","@")
     If HDW_FIX_DESTINATION_EMAIL <> "" Then emailaddress = HDW_FIX_DESTINATION_EMAIL
     fromaddress = emailaddress
     If (HDW_FROM_EMAIL_FIELD_NAME <> "") Then fromaddress = Request(HDW_FROM_EMAIL_FIELD_NAME)
     If (HDW_SUBJECT_FIELD_NAME <> "") Then emailsubject = Request(HDW_SUBJECT_FIELD_NAME)
  Else  
     Dim UploadRequest, byteCount, RequestBin, keys, i
     Set UploadRequest = CreateObject("Scripting.Dictionary")
     byteCount = Request.TotalBytes
     RequestBin = Request.BinaryRead(byteCount)
     BuildUploadRequest  RequestBin   
     Attachments = True    
     keys = UploadRequest.Keys            
     For i = 0 To UploadRequest.Count -1 
       If Not (UploadRequest.Item(keys(i)).Exists("FileName")) Then
           If (keys(i) <> HDW_F2M_OK) And (keys(i) <> HDW_F2M_NO_OK) And (keys(i) <> HDW_F2M_EMAIL) And notInThisArray(keys(i)) Then
              body = body & ""&keys(i)&": "&UploadRequest.Item(keys(i)).Item("Value")&"<br /><br />"
              getStr = getStr & "&"&keys(i)&"="&Server.URLEncode(UploadRequest.Item(keys(i)).Item("Value"))
           End If   
       Else  
           If notInThisArray(keys(i)) Then                   
              body = body & ""&keys(i)&": "&UploadRequest.Item(keys(i)).Item("FileName")&"<br /><br />"
              getStr = getStr & "&"&keys(i)&"="&Server.URLEncode(UploadRequest.Item(keys(i)).Item("FileName"))
           End If
       End If         
     Next
     emailaddress = UploadRequest.Item(HDW_F2M_EMAIL).Item("Value")
     emailaddress = Replace(emailaddress,"+","@")
     If HDW_FIX_DESTINATION_EMAIL <> "" Then emailaddress = HDW_FIX_DESTINATION_EMAIL
     fromaddress = emailaddress
     If (HDW_FROM_EMAIL_FIELD_NAME <> "") Then fromaddress = UploadRequest.Item(HDW_FROM_EMAIL_FIELD_NAME).Item("Value")
     If (HDW_SUBJECT_FIELD_NAME <> "") Then emailsubject = UploadRequest.Item(HDW_SUBJECT_FIELD_NAME).Item("Value")
  End If
  
  
  getStr = "hdw=1" & getStr
  body  = body & "SUPPORT INFORMATION<br />" &_
        "***************************<br />" &_  
        "Country: " &getCountryID(Request.ServerVariables("REMOTE_HOST"))&"<br />" &_
        "User IP: "&Request.ServerVariables("REMOTE_ADDR")&"<br />" &_
        "User Host: "&Request.ServerVariables("REMOTE_HOST")&"<br />" &_
        "Referer: "&Request.ServerVariables("HTTP_REFERER")&"<br />" &_
        "Server Time: "&Date & " "& Time&"<br />" &_
        "Browser: "&ckbrowser(Request.ServerVariables("HTTP_USER_AGENT"))&"<br />" &_
        "User Agent: "&Request.ServerVariables("HTTP_USER_AGENT")&"<br /><br />" &_                  
        "<hr />Delivered by HotDreamweaver Form2Mail Script"         

  If (USE_PLAIN_TEXT_EMAILS) Or (C_MAIL_USE = "ASPMail") Then      
      body = Replace(body,"","")
      body = Replace(body,"","")
      body = Replace(body,"<br />", vbNewLine)
      body = Replace(body,"<hr />","___________________________" & vbNewLine)        
  End If  
  
  If HDW_SEND_THANKYOU Then
      Dim thk_msg, thk_email
      thk_msg = HDW_THANKYOU_MSG
      If HDW_INCLUDE_SUBMITTED_DATA Then thk_msg = thk_msg & "<br /><br />" & body   
      If Not Attachments Then thk_email = Request.Form(HDW_USER_EMAIL_FIELD_NAME) Else thk_email = UploadRequest.Item(HDW_USER_EMAIL_FIELD_NAME).Item("Value")  
      If (thk_email = fromaddress) Then
          SendMail thk_email,emailaddress,HDW_THANKYOU_SUBJECT,thk_msg,HDW_INCLUDE_SUBMITTED_DATA    
      Else 
          SendMail thk_email,fromaddress,HDW_THANKYOU_SUBJECT,thk_msg,HDW_INCLUDE_SUBMITTED_DATA  
      End If
  End If
  
  If Attachments Then
    If (InStr(UploadRequest.Item(HDW_F2M_OK).Item("Value"),"?") > 0) And (getStr <> "") Then getStr = "&" & getStr Else getStr = "?" & getStr 
    If Not HOTDW_RESEND_PARAMS Then getStr = ""
    If (SendMail(emailaddress, fromaddress, emailsubject, body, True)) Then
        Response.Redirect UploadRequest.Item(HDW_F2M_OK).Item("Value") & getStr
    Else  
        Response.Redirect UploadRequest.Item(HDW_F2M_NO_OK).Item("Value")
    End If      
  Else
  	If (InStr(Request(HDW_F2M_OK),"?") > 0) And (getStr <> "") Then getStr = "&" & getStr Else getStr = "?" & getStr 
  	If Not HOTDW_RESEND_PARAMS Then getStr = ""
    If (SendMail(emailaddress, fromaddress, emailsubject, body, True)) Then
        Response.Redirect Request(HDW_F2M_OK) & getStr
    Else  
        Response.Redirect Request(HDW_F2M_NO_OK)
    End If      
  End If


%><%



  Function SendMail (var_ToAddress, var_FromAddress, var_Subject, var_Message, var_FlagAttachments)   
      Dim  smtp, objMail, i, j, value, iBp, Flds, Binary
      Dim objStream 
      Dim Stm 
      Dim buffer
      
     
      ' Send the email
      If C_MAIL_USE = "CDOSYS" Then   
          Set objMail = Server.CreateObject("CDO.Message")
          'objMail.MailFormat = 1  
          objMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
          
          objMail.Configuration.Fields.Item _
              ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTP_SERVER
          objMail.Configuration.Fields.Item _
              ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = SMTP_PORT
    
          If HDW_USE_AUTHENTICATION Then
              objMail.Configuration.Fields.Item _
                  ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1          
              objMail.Configuration.Fields.Item _
                  ("http://schemas.microsoft.com/cdo/configuration/sendusername") = HDW_AUTH_USERNAME
              objMail.Configuration.Fields.Item _ 
                  ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = HDW_AUTH_PASSWORD
          End If
            
          objMail.Configuration.Fields.Update
          
          objMail.To = var_ToAddress
          objMail.From = var_FromAddress
          objMail.Subject = var_Subject
          If Not USE_PLAIN_TEXT_EMAILS Then
              objMail.HTMLBody = var_Message
          Else
              objMail.TextBody = var_Message
          End If    
          If (HDW_SEND_CC_TO <> "") Then objMail.CC = HDW_SEND_CC_TO
          If Attachments And (var_FlagAttachments) Then
            For i = 0 To UploadRequest.Count -1 
              If (UploadRequest.Item(keys(i)).Exists("FileName")) Then
                       
                  Set iBp = objMail.Attachments.Add
                  Set Flds = iBp.Fields
                  With Flds
                     .Item("urn:schemas:mailheader:content-type") = "binary; name="&UploadRequest.Item(keys(i)).Item("FileName")
                     .Item("urn:schemas:mailheader:content-transfer-encoding") = "base64"
                     .Update
                  End With   
                                    
                  Set Stm = iBp.GetDecodedContentStream
     
                  Set value = UploadRequest.Item(keys(i)).Item("Value")                  
                  On Error Resume Next
                  Stm.Write (value.Read)
                  value.Position = 0 
                   
                  Stm.Flush
                  Set Stm = Nothing
              End If                                            
            Next                                                
          End If 
          
          On Error Resume Next
          objMail.Send

          Set objMail = Nothing   
          If Err.Number = 0 Then SendMail = True      
          If (Err.Number <> 0) And (HDW_ENABLE_DEBUG_MESSAGES) Then
              Response.Write Err.Description
              Response.End
          End If
      End If
      
      If (C_MAIL_USE = "CDONTS") Or ((C_MAIL_USE = "CDOSYS") AND (Err.Number <> 0)) Then     
          Set objMail = Server.CreateObject("CDONTS.NewMail")
          objMail.MailFormat = 0
          If Not USE_PLAIN_TEXT_EMAILS Then
              objMail.BodyFormat = 0
          Else                      
              objMail.BodyFormat = 1
          End If    
          objMail.To = var_ToAddress
          If (HDW_SEND_CC_TO <> "") Then objMail.CC = HDW_SEND_CC_TO
          objMail.From = var_FromAddress
          objMail.Subject = var_Subject
          objMail.Body = var_Message
          Err.Clear
          If Attachments And (var_FlagAttachments) Then
            For i = 0 To UploadRequest.Count -1 
              If (UploadRequest.Item(keys(i)).Exists("FileName")) Then              
                  UploadRequest.Item(keys(i)).Item("Value").saveToFile (localpath&"_uploadedfile-"&UploadRequest.Item(keys(i)).Item("FileName"))
                  objMail.AttachFile localpath&"_uploadedfile-"&UploadRequest.Item(keys(i)).Item("FileName"), UploadRequest.Item(keys(i)).Item("FileName")
              End If                                            
            Next                                                
          End If            
          objMail.Send
          Set objMail = Nothing   
          If Err.Number = 0 Then SendMail = True Else SendMail = False
          
          If (Err.Number <> 0) And (HDW_ENABLE_DEBUG_MESSAGES) Then
              Response.Write Err.Description
              Response.End
          End If
          
          
          If (Attachments) And (USE_FILESYSTEM_OBJECT) And (var_FlagAttachments) Then
            For i = 0 To UploadRequest.Count -1 
              If (UploadRequest.Item(keys(i)).Exists("FileName")) Then              
                  Set fso = CreateObject("Scripting.FileSystemObject")
                  Set MyFile = fso.GetFile( localpath&"_uploadedfile-"&UploadRequest.Item(keys(i)).Item("FileName") )
                  MyFile.Delete 
              End If                                            
            Next                                                
          End If 
          
          
      ElseIf (C_MAIL_USE = "ASPEmail") Then   
          Set objMail = Server.CreateObject("Persits.MailSender")
          objMail.Host = SMTP_SERVER
          objMail.From = var_FromAddress 
          objMail.FromName = var_FromAddress
          objMail.AddAddress var_ToAddress, var_ToAddress   
          If HDW_USE_AUTHENTICATION Then
              objMail.Username = HDW_AUTH_USERNAME
              objMail.Password = HDW_AUTH_PASSWORD
          End If    
           If Attachments And (var_FlagAttachments) Then
             For i = 0 To UploadRequest.Count -1 
               If (UploadRequest.Item(keys(i)).Exists("FileName")) Then              
                   UploadRequest.Item(keys(i)).Item("Value").saveToFile (localpath&"_uploadedfile-"&UploadRequest.Item(keys(i)).Item("FileName"))
                   objMail.AddAttachment localpath&"_uploadedfile-"&UploadRequest.Item(keys(i)).Item("FileName")
               End If                                            
             Next                                                
           End If   
          If (HDW_SEND_CC_TO <> "") Then objMail.AddCC HDW_SEND_CC_TO 
          objMail.Subject = var_Subject
          objMail.Body = var_Message
          objMail.IsHTML  = Not USE_PLAIN_TEXT_EMAILS
          objMail.Send      
          
          If (Attachments) And (USE_FILESYSTEM_OBJECT) And (var_FlagAttachments) Then
            For i = 0 To UploadRequest.Count -1 
              If (UploadRequest.Item(keys(i)).Exists("FileName")) Then              
                  Set fso = CreateObject("Scripting.FileSystemObject")
                  Set MyFile = fso.GetFile( localpath&"_uploadedfile-"&UploadRequest.Item(keys(i)).Item("FileName") )
                  MyFile.Delete 
              End If                                            
            Next                                                
          End If 
          SendMail = True  
                    
      ElseIf (C_MAIL_USE = "ASPMail") Then 
          set objMail = Server.CreateObject("SMTPsvg.Mailer")
          objMail.RemoteHost = SMTP_SERVER
          objMail.FromAddress = var_FromAddress 
          objMail.FromName = var_FromAddress 
          objMail.AddRecipient  var_ToAddress, var_ToAddress 
          If (HDW_SEND_CC_TO <> "") Then objMail.AddCC HDW_SEND_CC_TO, HDW_SEND_CC_TO 
          If Attachments And (var_FlagAttachments) Then
            For i = 0 To UploadRequest.Count -1 
              If (UploadRequest.Item(keys(i)).Exists("FileName")) Then     
                 UploadRequest.Item(keys(i)).Item("Value").saveToFile (localpath&"_uploadedfile-"&UploadRequest.Item(keys(i)).Item("FileName"))
                 objMail.AddAttachment localpath&"_uploadedfile-"&UploadRequest.Item(keys(i)).Item("FileName")
              End If                                            
            Next                                                
          End If   

          objMail.Subject = var_Subject
          objMail.BodyText = var_Message
          objMail.ContentType   = "text/html"
          objMail.SendMail      
          
          If (Attachments) And (USE_FILESYSTEM_OBJECT) And (var_FlagAttachments) Then
            For i = 0 To UploadRequest.Count -1 
              If (UploadRequest.Item(keys(i)).Exists("FileName")) Then              
                  Set fso = CreateObject("Scripting.FileSystemObject")
                  Set MyFile = fso.GetFile( localpath&"_uploadedfile-"&UploadRequest.Item(keys(i)).Item("FileName") )
                  MyFile.Delete 
              End If                                            
            Next                                                
          End If 
          SendMail = True  
                    
      ElseIf (C_MAIL_USE = "Jmail") Then 
          set objMail = Server.CreateObject("JMail.SMTPMail")
          objMail.ServerAddress  = SMTP_SERVER
          objMail.Sender  = var_FromAddress 
          objMail.AddRecipient var_ToAddress
          If (HDW_SEND_CC_TO <> "") Then objMail.CC = HDW_SEND_CC_TO
           If Attachments And (var_FlagAttachments) Then
             For i = 0 To UploadRequest.Count -1 
               If (UploadRequest.Item(keys(i)).Exists("FileName")) Then              
                   UploadRequest.Item(keys(i)).Item("Value").saveToFile (localpath&"_uploadedfile-"&UploadRequest.Item(keys(i)).Item("FileName"))
                   objMail.AddAttachment localpath&"_uploadedfile-"&UploadRequest.Item(keys(i)).Item("FileName")
               End If                                            
             Next                                                
           End If           
          
          objMail.Subject = var_Subject
          objMail.HTMLBody = var_Message
          objMail.Execute      
          
          If (Attachments) And (USE_FILESYSTEM_OBJECT) And (var_FlagAttachments) Then
            For i = 0 To UploadRequest.Count -1 
              If (UploadRequest.Item(keys(i)).Exists("FileName")) Then              
                  Set fso = CreateObject("Scripting.FileSystemObject")
                  Set MyFile = fso.GetFile( localpath&"_uploadedfile-"&UploadRequest.Item(keys(i)).Item("FileName") )
                  MyFile.Delete 
              End If                                            
            Next                                                
          End If 
          SendMail = True            
      End If  
      
  End Function
  
%>  
<%
Sub BuildUploadRequest(RequestBin)
	'Get the boundary
	Dim PosBeg, PosEnd, boundary, boundaryPos, Pos, Name, PosFile, PosBound, FileName, i, Value, gd, tmp
	Dim ContentType
	PosBeg = 1
	PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))
	boundary = MidB(RequestBin,PosBeg,PosEnd-PosBeg)
	boundaryPos = InstrB(1,RequestBin,boundary)
	'Get all data inside the boundaries
	Do until (boundaryPos=InstrB(RequestBin,boundary & getByteString("--")))
		'Members variable of objects are put in a dictionary object
		Dim UploadControl
		Set UploadControl = CreateObject("Scripting.Dictionary")
		'Get an object name
		Pos = InstrB(BoundaryPos,RequestBin,getByteString("Content-Disposition"))
		Pos = InstrB(Pos,RequestBin,getByteString("name="))
		PosBeg = Pos+6
		PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(34)))
		Name = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
		PosFile = InstrB(BoundaryPos,RequestBin,getByteString("filename="))
		PosBound = InstrB(PosEnd,RequestBin,boundary)
		'Test if object is of file type
		If  PosFile<>0 AND (PosFile<PosBound) Then
			'Get Filename, content-type and content of file
			PosBeg = PosFile + 10
			PosEnd =  InstrB(PosBeg,RequestBin,getByteString(chr(34)))
			FileName = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
			'Add filename to dictionary object
			UploadControl.Add "FileName", CleanFileName(FileName)
			Pos = InstrB(PosEnd,RequestBin,getByteString("Content-Type:"))
			If Pos = 0 Then
			  PosEnd = PosEnd+1
			  ' Esto es por el problema de las machintosh con los
			  ' PDF que se esta tragando el context type
			Else
			  PosBeg = Pos+14
			  PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))
			  'Add content-type to dictionary object
			  ContentType = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
			  UploadControl.Add "ContentType",ContentType
			End If
			
			'Get content of object
			PosBeg = PosEnd+4
			tmp = PosBeg-1
			PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
			Value = MidB(RequestBin,PosBeg,PosEnd-PosBeg)
			
			Set gd = CreateObject("ADODB.Stream")
            gd.Type = 1 ' adTypeBinary
            gd.Open
            gd.Write RequestBin
            gd.Flush
            gd.Position=tmp
            
            Set globalbuffer = CreateObject("ADODB.Stream")
            globalbuffer.Type = 1 ' adTypeBinary
            globalbuffer.Open
            On Error Resume Next
            globalbuffer.Write gd.Read(PosEnd-PosBeg)
            globalbuffer.Flush
            globalbuffer.Position=0
   
			 UploadControl.Add "Value" , globalbuffer
		Else
			'Get content of object
			Pos = InstrB(Pos,RequestBin,getByteString(chr(13)))
			PosBeg = Pos+4
			PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
			Value = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))	
			
			UploadControl.Add "Value" , Value		
		End If
		'Add content to dictionary object
	
		'Add dictionary object to main dictionary
	If Not UploadRequest.Exists(name) Then
     	    UploadRequest.Add name, UploadControl
     	Else
     	    ' este nuevo cambio es por los select tipo multiple,
     	    ' que la version original no se los tragaba
     	    UploadRequest.Item(name).Item("Value") = UploadRequest.Item(name).Item("Value") + ";"+Value
     	End If
		'Loop to next object
		BoundaryPos=InstrB(BoundaryPos+LenB(boundary),RequestBin,boundary)
	Loop

End Sub

'String to byte string conversion
Function getByteString(StringStr)
 Dim i, char
 For i = 1 to Len(StringStr)
 	char = Mid(StringStr,i,1)
	getByteString = getByteString & chrB(AscB(char))
 Next
End Function

'Byte string to string conversion
Function getString(StringBin)
 Dim intCount
 getString =""
 For intCount = 1 to LenB(StringBin)
	getString = getString & chr(AscB(MidB(StringBin,intCount,1)))
 Next
End Function

Function CleanFileName (fname)
  CleanFileName = Right(fname, Len(fname) - InStrRev(fname,"\"))
End Function

%>