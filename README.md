<div align="center">

## Windows Messages and Subclassing


</div>

### Description

Subclassing ofers great advantages to VB programmres. This article should teach you all about the message system Windows Uses, and how to implement it into your Visual Basic Programs.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[IRBMe](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/irbme.md)
**Level**          |Intermediate
**User Rating**    |4.8 (163 globes from 34 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/irbme-windows-messages-and-subclassing__1-34112/archive/master.zip)





### Source Code

```
<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:w="urn:schemas-microsoft-com:office:word"
xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<meta name=ProgId content=Word.Document>
<meta name=Generator content="Microsoft Word 9">
<meta name=Originator content="Microsoft Word 9">
<link rel=File-List href="./Windows%20Messages%20tutorial_files/filelist.xml">
<title>Windows Programming</title>
<!--[if gte mso 9]><xml>
 <o:DocumentProperties>
 <o:Author>Christopher Waddell</o:Author>
 <o:LastAuthor>Christopher Waddell</o:LastAuthor>
 <o:Revision>3</o:Revision>
 <o:TotalTime>94</o:TotalTime>
 <o:Created>2002-04-25T17:06:00Z</o:Created>
 <o:LastSaved>2002-04-26T15:50:00Z</o:LastSaved>
 <o:Pages>5</o:Pages>
 <o:Words>1596</o:Words>
 <o:Characters>9098</o:Characters>
 <o:Company>Developement</o:Company>
 <o:Lines>75</o:Lines>
 <o:Paragraphs>18</o:Paragraphs>
 <o:CharactersWithSpaces>11172</o:CharactersWithSpaces>
 <o:Version>9.4402</o:Version>
 </o:DocumentProperties>
</xml><![endif]-->
<style>
<!--
 /* Style Definitions */
p.MsoNormal, li.MsoNormal, div.MsoNormal
	{mso-style-parent:"";
	margin:0cm;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:12.0pt;
	font-family:"Times New Roman";
	mso-fareast-font-family:"Times New Roman";}
h2
	{mso-style-next:Normal;
	margin-top:12.0pt;
	margin-right:0cm;
	margin-bottom:3.0pt;
	margin-left:0cm;
	mso-pagination:widow-orphan;
	page-break-after:avoid;
	mso-outline-level:2;
	font-size:14.0pt;
	font-family:Arial;
	font-style:italic;}
p.MsoBodyText, li.MsoBodyText, div.MsoBodyText
	{margin:0cm;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:10.0pt;
	font-family:"Courier New";
	mso-fareast-font-family:"Times New Roman";}
p.MsoBodyText2, li.MsoBodyText2, div.MsoBodyText2
	{margin:0cm;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	mso-layout-grid-align:none;
	text-autospace:none;
	font-size:10.0pt;
	font-family:"Courier New";
	mso-fareast-font-family:"Times New Roman";
	color:black;}
a:link, span.MsoHyperlink
	{color:blue;
	text-decoration:underline;
	text-underline:single;}
a:visited, span.MsoHyperlinkFollowed
	{color:purple;
	text-decoration:underline;
	text-underline:single;}
@page Section1
	{size:595.3pt 841.9pt;
	margin:72.0pt 90.0pt 72.0pt 90.0pt;
	mso-header-margin:35.4pt;
	mso-footer-margin:35.4pt;
	mso-paper-source:0;}
div.Section1
	{page:Section1;}
 /* List Definitions */
@list l0
	{mso-list-id:1167982792;
	mso-list-type:hybrid;
	mso-list-template-ids:1525298616 67698711 67698713 67698715 67698703 67698713 67698715 67698703 67698713 67698715;}
@list l0:level1
	{mso-level-number-format:alpha-lower;
	mso-level-text:"%1\)";
	mso-level-tab-stop:36.0pt;
	mso-level-number-position:left;
	text-indent:-18.0pt;}
ol
	{margin-bottom:0cm;}
ul
	{margin-bottom:0cm;}
-->
</style>
</head>
<body lang=EN-GB link=blue vlink=purple style='tab-interval:36.0pt'>
<div class=Section1>
<h2 align=center style='text-align:center'>Windows Programming</h2>
<h2 align=center style='text-align:center'>Part 1 – Messages</h2>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>How does the Window Operating System know what you are
doing? How does it know when you click, where you click and with what button
you click? How does it know when you press a key, what key you pressed and what
window you are typing in?<span style="mso-spacerun: yes">  </span>There are
many questions with only one simple answer. The answer being a message system.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>There are many hundreds of common Windows messages, which
include the left mouse click, the right mouse click and also the key down, and
key up messages. There are other messages other than those used to indicate
user input. There is also a message for instance that tells a window to repaint
(or redraw) itself and also a timer message.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>So how do applications receive these messages? The answer is
a “window procedure”, although not official, it is generally agreed that it
should be called “WindowProc”. The window procedure is a function that will be
called every time a message is sent to that window. It must be declared as a
public function in a module! It looks like this:</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>Public</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'> </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>Function</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'> WindowProc(</span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>ByVal</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'> hwnd </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>As</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'> </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>Long</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'>, </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>ByVal</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'> uMsg </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>As Long</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'>, _ <o:p></o:p></span></p>
<p class=MsoNormal style='margin-left:72.0pt;text-indent:36.0pt;mso-layout-grid-align:
none;text-autospace:none'><span style='font-size:10.0pt;font-family:"Courier New";
color:navy'>ByVal</span><span style='font-size:10.0pt;font-family:"Courier New";
color:black'> wParam </span><span style='font-size:10.0pt;font-family:"Courier New";
color:navy'>As Long</span><span style='font-size:10.0pt;font-family:"Courier New";
color:black'>, </span><span style='font-size:10.0pt;font-family:"Courier New";
color:navy'>ByVal</span><span style='font-size:10.0pt;font-family:"Courier New";
color:black'> lParam </span><span style='font-size:10.0pt;font-family:"Courier New";
color:navy'>As Long</span><span style='font-size:10.0pt;font-family:"Courier New";
color:black'>) </span><span style='font-size:10.0pt;font-family:"Courier New";
color:navy'>As Long</span><span style='font-size:10.0pt;font-family:"Courier New";
color:black'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New";
color:navy'>End Function</span><span style='color:navy'><o:p></o:p></span></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Parameters: -</p>
<p class=MsoNormal><span style='mso-tab-count:1'>            </span>hwnd – The window
handle of your window. A window handle is a unique number, which is assigned to
your window. Whenever you call an API function that wants to do something with
your window, you must pass the hwnd property</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>uMsg – This is the number of the message that was sent your
window. For example:</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New";
color:navy'>Public</span><span style='font-size:10.0pt;font-family:"Courier New";
color:black'> </span><span style='font-size:10.0pt;font-family:"Courier New";
color:navy'>Const</span><span style='font-size:10.0pt;font-family:"Courier New";
color:black'> WM_DRAWCLIPBOARD = &amp;H308<span style="mso-spacerun: yes"> 
</span></span><span style='font-size:10.0pt;font-family:"Courier New";
color:green'>‘Declare this message as a<span style="mso-spacerun:
yes">           </span>const, making it easier to deal with.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New";
color:green'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New"'>You
would then use it like this:<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New";
color:navy'>Declare Function</span><span style='font-size:10.0pt;font-family:
"Courier New";color:black'> CallWindowProc </span><span style='font-size:10.0pt;
font-family:"Courier New";color:navy'>Lib</span><span style='font-size:10.0pt;
font-family:"Courier New";color:black'> &quot;user32&quot; </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>Alias</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'>
&quot;CallWindowProcA&quot; (</span><span style='font-size:10.0pt;font-family:
"Courier New";color:navy'>ByVal</span><span style='font-size:10.0pt;font-family:
"Courier New";color:black'> lpPrevWndFunc </span><span style='font-size:10.0pt;
font-family:"Courier New";color:navy'>As</span><span style='font-size:10.0pt;
font-family:"Courier New";color:black'> </span><span style='font-size:10.0pt;
font-family:"Courier New";color:navy'>Long</span><span style='font-size:10.0pt;
font-family:"Courier New";color:black'>, </span><span style='font-size:10.0pt;
font-family:"Courier New";color:navy'>ByVal</span><span style='font-size:10.0pt;
font-family:"Courier New";color:black'> hwnd </span><span style='font-size:
10.0pt;font-family:"Courier New";color:navy'>As Long</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'>, </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>ByVal</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'> Msg </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>As Long</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'>, </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>ByVal</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'> wParam </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>As Long</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'>, </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>ByVal</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'> lParam </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>As Long</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'>) </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>As Long</span><span
style='font-size:10.0pt;font-family:"Courier New"'><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New";
color:green'>‘...In windowproc<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New";
color:navy'>Select case</span><span style='font-size:10.0pt;font-family:"Courier New"'>
uMsg<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New"'><span
style="mso-spacerun: yes">  </span><span style='color:navy'>Case</span> <span
style='color:black'>DRAWCLIPBOARD<o:p></o:p></span></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New";
color:black'><span style="mso-spacerun: yes">    </span></span><span
style='font-size:10.0pt;font-family:"Courier New";color:green'>‘The data in the
clipboard has changed, so do something<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New";
color:black'><span style="mso-spacerun: yes">  </span></span><span
style='font-size:10.0pt;font-family:"Courier New";color:green'>‘Case ... Other
messages go here<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New";
color:green'><span style="mso-spacerun: yes">  </span></span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>Case Else<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><span
style="mso-spacerun: yes">      </span>WindowProc = CallWindowProc(PrevProc,
hwnd, uMsg, wParam,<span style="mso-spacerun: yes">   </span>lParam) </span><span
style='font-size:10.0pt;font-family:"Courier New";color:green'>‘Process all
those other messages that we don’t care about<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New";
color:navy'>End</span><span style='font-size:10.0pt;font-family:"Courier New";
color:black'> </span><span style='font-size:10.0pt;font-family:"Courier New";
color:navy'>select<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New";
color:navy'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal>wParam/lParam – These are general parameters and can store pretty
much any values including other sub-messages. If memory serves me correctly
then the mouse move message comes with the X and Y coordinates of the mouse
stored in the wParam and lParam parameters.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Now some of you may be thinking, “I hope I don’t have to
process all of the hundreds of messages, my code could be thousands of lines
long”. For those of you who weren’t, well you are now. The answer is thankfully
no. There is a default window procedure that will carry out the basic commands
like painting your window, resizing it, moving it, giving it focus, and all of
the hundreds of other things.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>We have a lot of control when it comes to messages. We can
create our own messages, send messages to the system and look at all the
messages in the message queue. Consider the following API functions:</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>Declare Function </span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'>GetMessage </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>Lib</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'>
&quot;user32&quot; </span><span style='font-size:10.0pt;font-family:"Courier New";
color:navy'>Alias</span><span style='font-size:10.0pt;font-family:"Courier New";
color:black'> &quot;GetMessageA&quot; (lpMsg </span><span style='font-size:
10.0pt;font-family:"Courier New";color:navy'>As</span><span style='font-size:
10.0pt;font-family:"Courier New";color:black'> Msg, </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>ByVal</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'> hWnd </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>As</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'> </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>Long</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'>, </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>ByVal</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'> wMsgFilterMin </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>As Long</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'>, </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>ByVal</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'> wMsgFilterMax </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>As Long</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'>) </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>As Long</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>Declare Function</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'>
TranslateMessage </span><span style='font-size:10.0pt;font-family:"Courier New";
color:navy'>Lib</span><span style='font-size:10.0pt;font-family:"Courier New";
color:black'> &quot;user32&quot; (lpMsg </span><span style='font-size:10.0pt;
font-family:"Courier New";color:navy'>As</span><span style='font-size:10.0pt;
font-family:"Courier New";color:black'> Msg) </span><span style='font-size:
10.0pt;font-family:"Courier New";color:navy'>As Long</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>Declare Function</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'> DispatchMessage
</span><span style='font-size:10.0pt;font-family:"Courier New";color:navy'>Lib</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'>
&quot;user32&quot; </span><span style='font-size:10.0pt;font-family:"Courier New";
color:navy'>Alias</span><span style='font-size:10.0pt;font-family:"Courier New";
color:black'> &quot;DispatchMessageA&quot; (lpMsg </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>As</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'> Msg) </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>As Long</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><o:p></o:p></span></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>Type</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'> POINTAPI<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><span
style="mso-spacerun: yes">    </span>x </span><span style='font-size:10.0pt;
font-family:"Courier New";color:navy'>As Long</span><span style='font-size:
10.0pt;font-family:"Courier New";color:black'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><span
style="mso-spacerun: yes">    </span>y </span><span style='font-size:10.0pt;
font-family:"Courier New";color:navy'>As Long</span><span style='font-size:
10.0pt;font-family:"Courier New";color:black'><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New";
color:navy'>End Type<o:p></o:p></span></p>
<p class=MsoNormal><span style='color:navy'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>Type Msg<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><span
style="mso-spacerun: yes">    </span>hWnd </span><span style='font-size:10.0pt;
font-family:"Courier New";color:navy'>As Long</span><span style='font-size:
10.0pt;font-family:"Courier New";color:black'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><span
style="mso-spacerun: yes">    </span>message </span><span style='font-size:
10.0pt;font-family:"Courier New";color:navy'>As Long</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><span
style="mso-spacerun: yes">    </span>wParam </span><span style='font-size:10.0pt;
font-family:"Courier New";color:navy'>As Long<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><span
style="mso-spacerun: yes">    </span>lParam </span><span style='font-size:10.0pt;
font-family:"Courier New";color:navy'>As Long</span><span style='font-size:
10.0pt;font-family:"Courier New";color:black'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><span
style="mso-spacerun: yes">    </span>time </span><span style='font-size:10.0pt;
font-family:"Courier New";color:navy'>As Long</span><span style='font-size:
10.0pt;font-family:"Courier New";color:black'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><span
style="mso-spacerun: yes">    </span>pt </span><span style='font-size:10.0pt;
font-family:"Courier New";color:navy'>As</span><span style='font-size:10.0pt;
font-family:"Courier New";color:black'> POINTAPI<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New";
color:navy'>End Type<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New";
color:navy'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoBodyText>Complicated looking isn’t it? We can use these API
functions as follows:</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style='color:navy'>Dim</span> aMsg <span
style='color:navy'>as</span> Msg</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>Call</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'> GetMessage
(aMsg, 0, 0, 0)<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>Call</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'>
TranslateMessage (aMsg)<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New";
color:navy'>Call</span><span style='font-size:10.0pt;font-family:"Courier New";
color:black'> DispatchMessage (aMsg)<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New";
color:black'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal>I think that is pretty self-explanatory. </p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>VB has a built in message handler in its form object. This
is where the events come from on your forms, and also the controls as well.
These events are just generated whenever the corresponding messages are detected
in the window Procedure. And the X and Y values in the MouseDown event for
example are just extracted from the lParam and wParam arguments in the
WindowProc function.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Now, why would you want to write our own message handler if
VB already provides a perfectly good one?</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='margin-left:36.0pt;text-indent:-18.0pt;mso-list:l0 level1 lfo2;
tab-stops:list 36.0pt'><![if !supportLists]>a)<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span><![endif]>VB hides a lot of the Messages from us</p>
<p class=MsoNormal style='margin-left:36.0pt;text-indent:-18.0pt;mso-list:l0 level1 lfo2;
tab-stops:list 36.0pt'><![if !supportLists]>b)<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span><![endif]>VB deals with some messages in a way that might not suit what
we want</p>
<p class=MsoNormal style='margin-left:36.0pt;text-indent:-18.0pt;mso-list:l0 level1 lfo2;
tab-stops:list 36.0pt'><![if !supportLists]>c)<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span><![endif]>VB processes its messages before sending us the event. What if
we don’t want it to do anything?</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Let us consider the rather complicated topic of Winsock API.
The way Winsock lets us know what is going on is through messages sent to our
window’s message handler. However VB hides these ones from us. In order to see
them, we will have to create a window procedure of our own.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Now, how do we tell windows to send messages to our new
window procedure? Like so:</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>Private Declare</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'> </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>Function</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'> GetWindowLong </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>Lib</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'> _<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:black'>&quot;user32&quot;
</span><span style='font-size:10.0pt;font-family:"Courier New";color:navy'>Alias</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'>
&quot;GetWindowLongA&quot; (</span><span style='font-size:10.0pt;font-family:
"Courier New";color:navy'>ByVal</span><span style='font-size:10.0pt;font-family:
"Courier New";color:black'> hWnd _<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>As Long</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'>, </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>ByVal</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'> nIndex </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>As</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'> </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>Long</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'>) </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>As Long</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>Private Declare
Function</span><span style='font-size:10.0pt;font-family:"Courier New";
color:black'> SetWindowLong </span><span style='font-size:10.0pt;font-family:
"Courier New";color:navy'>Lib</span><span style='font-size:10.0pt;font-family:
"Courier New";color:black'> _<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:black'>&quot;user32&quot;
</span><span style='font-size:10.0pt;font-family:"Courier New";color:navy'>Alias</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'>
&quot;SetWindowLongA&quot; (</span><span style='font-size:10.0pt;font-family:
"Courier New";color:navy'>ByVal</span><span style='font-size:10.0pt;font-family:
"Courier New";color:black'> hWnd _<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>As Long</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'>, </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>ByVal</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'> nIndex </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>As Long</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'>, </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>ByVal</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'> dwNewLong _<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New";
color:navy'>As Long</span><span style='font-size:10.0pt;font-family:"Courier New";
color:black'>) </span><span style='font-size:10.0pt;font-family:"Courier New";
color:navy'>As Long<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New";
color:navy'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoBodyText>Those are 2 new API calls, one creates a window procedure,
and the other returns the address of a window procedure given the hwnd (window
handle remember)</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>So, to set up a window procedure, we do this:</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New";
color:navy'>Public</span><span style='font-size:10.0pt;font-family:"Courier New";
color:black'> </span><span style='font-size:10.0pt;font-family:"Courier New";
color:navy'>Const</span><span style='font-size:10.0pt;font-family:"Courier New";
color:black'> GWL_WNDPROC = -4</span></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style='color:navy'>Private Sub</span> Form_Load() <span
style='color:green'>‘Of course it doesn’t have to go in form load<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><span
style="mso-spacerun: yes">   </span>PrevProc = SetWindowLong(hwnd, GWL_WNDPROC,
</span><span style='font-size:10.0pt;font-family:"Courier New";color:navy'>AddressOf</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'> WindowProc)<o:p></o:p></span></p>
<p class=MsoNormal><span style='color:navy'>End sub<o:p></o:p></span></p>
<p class=MsoNormal><span style='color:navy'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal>You can replace the “AddressOf WindowProc” with the name you
have given to your window procedure, but I suggest you keep the name to
WindowProc. Also remember WindowProc must be a public Function, written with
the correct parameters and everything, in a public Module.</p>
<p class=MsoNormal><span style='color:navy'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal>This API call returns the handle to the previous window
procedure if one exists</p>
<p class=MsoNormal>We must store a value into PrevProc so that we can return
the default Window Procedure when we are finished. So, how do we return the
previous window procedure? Like this:</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style='color:navy'>Private Sub</span>
Form_Unload(Cancel <span style='color:navy'>as Integer</span>) <span
style='color:green'>‘Again, doesn’t have to be in Form_Unload</span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'><span
style="mso-spacerun: yes">    </span>If</span><span style='font-size:10.0pt;
font-family:"Courier New";color:black'> PrevProc &lt;&gt; 0 </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>Then</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><span
style="mso-spacerun: yes">       </span>SetWindowLong hwnd, GWL_WNDPROC,
PrevProc<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><span
style="mso-spacerun: yes">       </span>PrevProc = 0<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New";
color:black'><span style="mso-spacerun: yes">    </span></span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>End If<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New";
color:navy'>End Sub</span><span style='font-size:10.0pt;font-family:"Courier New"'><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New"'>So
now we know how to:<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New"'>Create
the WindowProc Function.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New"'>Set
the WindowProc function as a window procedure.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New"'>Look
for messages that we want.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New"'>Extract
values from the lParam and wParam arguments.<o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New"'>Process
all the other messages with the default handler.<o:p></o:p></span></p>
<p class=MsoBodyText>Remove our window procedure.</p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New"'>Here
is a small example taken from <a href="http://www.allapi.net/">AllApi.Net</a><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:green'>'Create a new
project, add a module to it<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:green'>'Add a command
button to Form1<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:green'>'In the form</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>Private Sub</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'> Form_Load()<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><span
style="mso-spacerun: yes">    </span></span><span style='font-size:10.0pt;
font-family:"Courier New";color:green'>'KPD-Team 1999<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:green'><span
style="mso-spacerun: yes">    </span>'URL: http://www.allapi.net/<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:green'><span
style="mso-spacerun: yes">    </span>'E-Mail: KPDTeam@Allapi.net<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:green'><span
style="mso-spacerun: yes">    </span>'Subclass this form<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><span
style="mso-spacerun: yes">    </span>HookForm Me<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><span
style="mso-spacerun: yes">    </span></span><span style='font-size:10.0pt;
font-family:"Courier New";color:green'>'Register this form as a Clipboardviewer<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><span
style="mso-spacerun: yes">    </span>SetClipboardViewer Me.hwnd<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>End Sub<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>Private Sub</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'>
Form_Unload(Cancel </span><span style='font-size:10.0pt;font-family:"Courier New";
color:navy'>As Integer</span><span style='font-size:10.0pt;font-family:"Courier New";
color:black'>)<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><span
style="mso-spacerun: yes">    </span></span><span style='font-size:10.0pt;
font-family:"Courier New";color:green'>'Unhook the form<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><span
style="mso-spacerun: yes">    </span>UnHookForm Me<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>End Sub<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>Private Sub</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'>
Command1_Click()<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><span
style="mso-spacerun: yes">    </span></span><span style='font-size:10.0pt;
font-family:"Courier New";color:green'>'Change the clipboard<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><span
style="mso-spacerun: yes">    </span>Clipboard.Clear<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><span
style="mso-spacerun: yes">    </span>Clipboard.SetText &quot;Hello !&quot;<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>End Sub<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:green'>'In a module<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:green'>'These routines
are explained in our subclassing tutorial.<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:green'>'http://www.allapi.net/vbtutor/subclass.php<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>Declare Function</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'> SetWindowLong </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>Lib</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'>
&quot;user32&quot; </span><span style='font-size:10.0pt;font-family:"Courier New";
color:navy'>Alias</span><span style='font-size:10.0pt;font-family:"Courier New";
color:black'> &quot;SetWindowLongA&quot; (</span><span style='font-size:10.0pt;
font-family:"Courier New";color:navy'>ByVal hwnd As Long, ByVal nIndex As Long,
ByVal dwNewLong As Long) As Long</span><span style='font-size:10.0pt;
font-family:"Courier New";color:black'><o:p></o:p></span></p>
<p class=MsoBodyText2><span style='color:navy'>Declare Function</span>
CallWindowProc <span style='color:navy'>Lib</span> &quot;user32&quot; <span
style='color:navy'>Alias</span> &quot;CallWindowProcA&quot; (<span
style='color:navy'>ByVal</span> lpPrevWndFunc <span style='color:navy'>As Long</span>,
<span style='color:navy'>ByVal</span> hwnd <span style='color:navy'>As Long</span>,
<span style='color:navy'>ByVal</span> Msg <span style='color:navy'>As Long</span>,
<span style='color:navy'>ByVal</span> wParam <span style='color:navy'>As Long</span>,
<span style='color:navy'>ByVal</span> lParam <span style='color:navy'>As Long</span>)
<span style='color:navy'>As Long</span></p>
<p class=MsoBodyText2><span style='color:navy'>Declare Function</span>
SetClipboardViewer <span style='color:navy'>Lib</span> &quot;user32&quot; (<span
style='color:navy'>ByVal</span> hwnd <span style='color:navy'>As Long</span>) <span
style='color:navy'>As Long<o:p></o:p></span></p>
<p class=MsoBodyText2><span style='color:navy'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>Public Const</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'>
WM_DRAWCLIPBOARD = &amp;H308<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>Public Const</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'> GWL_WNDPROC =
(-4)<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>Dim</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'> PrevProc </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>As Long<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>Public Sub</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'> HookForm(F As
Form)<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><span
style="mso-spacerun: yes">    </span>PrevProc = SetWindowLong(F.hwnd,
GWL_WNDPROC, </span><span style='font-size:10.0pt;font-family:"Courier New";
color:navy'>AddressOf</span><span style='font-size:10.0pt;font-family:"Courier New";
color:black'> WindowProc)<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>End Sub<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>Public Sub</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'> UnHookForm(F </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>As</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'> Form)<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><span
style="mso-spacerun: yes">    </span>SetWindowLong F.hwnd, GWL_WNDPROC,
PrevProc<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>End Sub<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>Public Function</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'> WindowProc(</span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>ByVal</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'> hwnd </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>As Long</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'>, </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>ByVal</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'> uMsg </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>As Lon</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'>g, </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>ByVal</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'> wParam </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>As Long</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'>, </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>ByVal</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'> lParam </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>As Long</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'>) </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>As Long</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><span
style="mso-spacerun: yes">    </span>WindowProc = CallWindowProc(PrevProc,
hwnd, uMsg, wParam, lParam)<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><span
style="mso-spacerun: yes">    </span></span><span style='font-size:10.0pt;
font-family:"Courier New";color:navy'>If</span><span style='font-size:10.0pt;
font-family:"Courier New";color:black'> uMsg = WM_DRAWCLIPBOARD </span><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>Then</span><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:black'><span
style="mso-spacerun: yes">        </span>MsgBox &quot;Clipboard changed
...&quot;<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'><span
style="mso-spacerun: yes">    </span>End If<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'>End Function<o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoBodyText style='mso-layout-grid-align:none;text-autospace:none'>If
you want, you can create your own windows messages. However, problems can
arise. Imagine you use a message in a DLL as follows:</p>
<p class=MsoBodyText style='mso-layout-grid-align:none;text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoBodyText style='mso-layout-grid-align:none;text-autospace:none'><span
style='color:navy'>Const</span> MYMSG = WM_USER + 7</p>
<p class=MsoBodyText style='mso-layout-grid-align:none;text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoBodyText style='mso-layout-grid-align:none;text-autospace:none'>However,
lets then imagine that another DLL uses the exact same message for something
completely different. Now to make matters worse, some poor person tries to use
the two DLL’s in the same project. Let the errors and bugs and problems
commence. Well, there is a way around this:</p>
<p class=MsoBodyText style='mso-layout-grid-align:none;text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoBodyText style='mso-layout-grid-align:none;text-autospace:none'><span
style='color:navy'>Declare Function</span> RegisterWindowMessage <span
style='color:navy'>Lib</span> &quot;user32&quot; <span style='color:navy'>Alias</span>
&quot;RegisterWindowMessageA&quot; (<span style='color:navy'>ByVal</span>
lpString <span style='color:navy'>As String</span>) <span style='color:navy'>As
Long</span></p>
<p class=MsoBodyText style='mso-layout-grid-align:none;text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoBodyText style='mso-layout-grid-align:none;text-autospace:none'>What
this will do is allow you to create unique message numbers. Lets say you wanted
to create your own message, you would do something like this:</p>
<p class=MsoBodyText style='mso-layout-grid-align:none;text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoBodyText style='mso-layout-grid-align:none;text-autospace:none'>MY_MESSAGE
= RegisterWindowMessage (“MyUniqueString”)</p>
<p class=MsoBodyText style='mso-layout-grid-align:none;text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoBodyText style='mso-layout-grid-align:none;text-autospace:none'>This
will assign MY_MESSAGE a new unique message number every time it is run.
However, if you put this in a DLL then how will the applications using the DLL
know what the number of your message is? They do EXACTLY the same thing as
above. When they enter “MyUniqueString” into the lpString Parameter, because it
already exists (it was originally made by your DLL remember), it will now
return the number that it assigned to MY_MESSAGE. Consider the following
example:</p>
<p class=MsoBodyText style='mso-layout-grid-align:none;text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoBodyText style='mso-layout-grid-align:none;text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoBodyText style='mso-layout-grid-align:none;text-autospace:none'>MESSAGE_ONE
= RegisterWindowMessage (“MyFirstString”)</p>
<p class=MsoBodyText style='mso-layout-grid-align:none;text-autospace:none'>Msgbox
“Your first new message is “ &amp; MESSAGE_ONE</p>
<p class=MsoBodyText style='mso-layout-grid-align:none;text-autospace:none'>MEASSAGE_TWO
= RegisterWindowMessage (“MySecondString”)</p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'>Msgbox
“Your second new message is “ &amp; MESSAGE_TWO</p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'>Msgbox
“How do we retrieve message one? Like this: “ &amp; RegisterWindowMessage (“MyFirstString”)</p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'>Msgbox
“How do we retrieve message two? Like this: “ &amp; RegisterWindowMessage (“MySecondString”)<span
style='font-size:10.0pt;font-family:"Courier New"'><o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal style='mso-layout-grid-align:none;text-autospace:none'><span
style='font-size:10.0pt;font-family:"Courier New";color:navy'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoBodyText style='mso-layout-grid-align:none;text-autospace:none'>Well,
that’s the end of this tutorial. Let me just tell you that the technical name
for this is called Sub classing, in case you ever hear it referred to as that.</p>
<p class=MsoBodyText style='mso-layout-grid-align:none;text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoBodyText style='mso-layout-grid-align:none;text-autospace:none'>I
hope that after reading this you understand everything, however if there is
anything you still don’t understand then visit <a href="http://www.allapi.net/">http://www.AllAPI.net</a>
and search for one of the API declarations mentioned in the tutorials.
Alternately, search for WindowProc, or Subclass. They should get you something.</p>
<p class=MsoBodyText style='mso-layout-grid-align:none;text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoBodyText style='mso-layout-grid-align:none;text-autospace:none'>I’d
just like to say how long it took me to highlight all that code in its correct
colouring, so if anybody has a good program to do that automatically, I’d be
grateful!</p>
<p class=MsoBodyText style='mso-layout-grid-align:none;text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoBodyText style='mso-layout-grid-align:none;text-autospace:none'>Also,
I know there are loads of people out there who know the ins and outs of Windows
messaging, and have read this for whatever reason. I know I read tutorials on
things I know inside out anyway. So, for any of you experts who have read this,
any concerns with the tutorial (Misinformation, bugs in code, even typo’s),
then I’d like to know, so leave a comment if you want.</p>
<p class=MsoBodyText style='mso-layout-grid-align:none;text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoBodyText style='mso-layout-grid-align:none;text-autospace:none'>I
also like to know if I have helped people, and if so, how much. So some
comments there wouldn’t go amiss.</p>
<p class=MsoBodyText style='mso-layout-grid-align:none;text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoBodyText style='mso-layout-grid-align:none;text-autospace:none'>Enjoy!</p>
<p class=MsoBodyText style='mso-layout-grid-align:none;text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoBodyText style='mso-layout-grid-align:none;text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoBodyText style='mso-layout-grid-align:none;text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoBodyText style='mso-layout-grid-align:none;text-autospace:none'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:"Courier New"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
</div>
</body>
</html>
```

