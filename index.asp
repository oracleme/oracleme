<!--#include file="Inc/SysProduct.asp" -->
<%
function cutstr(tempstr,tempwid)
if len(tempstr)>tempwid then
cutstr=left(tempstr,tempwid)&"..."
else
cutstr=tempstr
end if
end function%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>网站首页-<%=SiteTitle%></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<META content=<%=Sitekey%> name=keywords>
<META content=<%=Sitedes%> name=description>
<link rel="stylesheet" type="text/css" href="Style.css">
</head>
<!--#include file="head.asp" -->

<DIV id=main>
<DIV id=main1>
<FORM action="search.asp" method=get name=myform id="myform">
<DIV id=search>
<UL>
  <LI>
  <P class=stext>关键字：</P>
  <P class=sform><INPUT id=keyword size=11 name=keyword>
  </P></LI>
  <LI>
  <P class=stext>类　别：</P>
  <P class=sform>

<SELECT id=id style="WIDTH: 100px" name=id>
 <OPTION value="" selected>全部产品</OPTION>
<%
set rsbig = server.CreateObject ("adodb.recordset")
		sql="select * from BigClass"
		rsbig.open sql,conn,1,1
if not rsbig.eof then
	do while not rsbig.eof
%>
 <OPTION value="<%=rsbig("BigClassName")%>" ><%=rsbig("BigClassName")%></OPTION>
<%
rsbig.movenext
Loop
end if
rsbig.close
rsbig=nothing
%>

  </SELECT> 
</P>
  <DIV id=searchbtn><INPUT id=button style="MARGIN-TOP: 5px" type=image height=21 width=166 src="images/111111.gif" value=Submit name=button></DIV></LI></UL></DIV></FORM>
<SCRIPT language=javascript>
       function myclick(){
		   document.getElementsByTagName("from1").submit();
		   document.getElementById("from1").submit();
		   alert('dddd');
		   }
    </SCRIPT>

<DIV id=about>
<DIV id=aboutnav><A href="about.asp"><IMG height=19 src="images/more.jpg" width=52 border=0></A></DIV>
<DIV id=aboutcont>

<%
  Set rs = Server.CreateObject("ADODB.Recordset")
  sql="select Content from Aboutus where Title='企业简介'"
  rs.open sql,conn,1,3
%>
<%=cutstr(rs("content"),138)%>
<%
rs.close
rs=nothing
%>
<a href="Aboutus.asp?Title=企业简介">[详细]</a>


</DIV></DIV>
<DIV id=case>
<DIV id=casel>
	
<script language="JavaScript" type="text/javascript">
var swf_width=277
var swf_height=134
var files='UploadFiles/1.jpg|UploadFiles/2.jpg|UploadFiles/3.jpg|UploadFiles/4.jpg|UploadFiles/5.jpg'
var links ='CompVisualize.asp|CompVisualize.asp|CompVisualize.asp|CompVisualize.asp|CompVisualize.asp'
var texts //='#|#'
document.write('<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="'+ swf_width +'" height="'+ swf_height +'">');
document.write('<param name="movie" value="images/ShowPic.swf"><param name="quality" value="high">');
document.write('<param name="menu" value="false"><param name=wmode value="opaque">');
document.write('<param name="FlashVars" value="bcastr_file='+files+'&bcastr_link='+links+'&bcastr_title='+texts+'">');
document.write('<embed src="images/ShowPic.swf" wmode="opaque" FlashVars="bcastr_file='+files+'&bcastr_link='+links+'&bcastr_title='+texts+'& menu="false" quality="high" width="'+ swf_width +'" height="'+ swf_height +'" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />'); document.write('</object>'); 
  </script>














</DIV>
<DIV id=caser><IMG height=136 src="images/case.jpg"width=24></DIV></DIV></DIV>

<DIV id=main2>
<DIV id=main2l>
<A href="Aboutus.asp?Title=联系我们"><IMG src="images/cnotact.jpg" width=210 height=76 border="0"></A>
<A href="Feedback.asp"><IMG src="images/feedback.jpg" width=210 height=76 border="0"></A></DIV>
<DIV id=main2r>
<DIV id=prol><IMG height=172 src="images/pro.jpg" width=60></DIV>
<DIV id=maisydemo style="FLOAT: left; OVERFLOW: hidden; WIDTH: 620px" 
align=center>
<TABLE cellSpacing=0 cellPadding=0 border=0>
  <TBODY>
  <TR>
    <TD id=maisydemo1>
      <DIV id=prom>
      <UL>
              <%
set rs_Product=server.createobject("adodb.recordset")
sqltext="select top 6 * from Product where Passed=True and Elite=true  order by UpdateTime desc"
rs_Product.open sqltext,conn,1,1

           If rs_Product.eof and rs_Product.bof then
           response.write "<li><p align='center'><font color='#ff0000'>还没任何产品</font></p></li>"
           else

			  Do While Not rs_Product.EOF%>	   
        <LI>
        <P class=propicture>
		<a href="ProductShow.asp?ID=<%=rs_Product("id")%>"><img src=<%=rs_Product("DefaultPicUrl")%> width="142" height="107" border=0 ></a></P>
        <P class=protext><a class=prolink href="Product_Show.asp?ID=<%=rs_Product("id")%>"><%=rs_Product("Title")%></a></P></LI>
 
                 <%
rs_Product.MoveNext
row_count=row_count+1
Loop
end if
rs_Product.close
%>      
	    
    

       
	    
                                
	   
	   </UL></DIV></TD>
    <TD id=maisydemo2 width=1></TD>
    <TD id=maisydemo3 width=1></TD></TR></TBODY></TABLE>
<SCRIPT>
var speed=30;
maisydemo2.innerHTML=maisydemo1.innerHTML;
maisydemo3.innerHTML=maisydemo1.innerHTML;
function Marqueemaisy(){
if(maisydemo2.offsetWidth-maisydemo.scrollLeft<=0)
     maisydemo.scrollLeft-=maisydemo1.offsetWidth;
else{
     maisydemo.scrollLeft++;
    }
}
var MyMarmaisy=setInterval(Marqueemaisy,speed);
maisydemo.onmouseover=function() {clearInterval(MyMarmaisy);}
maisydemo.onmouseout=function() {MyMarmaisy=setInterval(Marqueemaisy,speed);}
</SCRIPT>
</DIV>
<DIV id=pror><IMG height=172 src="images/pror.jpg" width=16></DIV></DIV></DIV></DIV>

<!--#include file="Foot.asp" -->

