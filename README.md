<div align="center">

## DoEvents evolution; the API approach\. \(Method for 100% optimized loops\)


</div>

### Description

Do you want to make your loops 100% faster? Here's how :

Many of us have used several times DoEvents, to supply a bit of air to our App, on Heavy-Duty times such as loops for updates or inserts on recordsets etc. As we most know, DoEvents processes Windows messages currently in the message queue. But what if we wanted to execute DoEvents only in times, when we want to allow user (Keyboard and Mouse) input? ( A special "thank you" to all of you who rated this article)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2001-12-20 11:22:16
**By**             |[John Galanopoulos](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/john-galanopoulos.md)
**Level**          |Intermediate
**User Rating**    |5.0 (283 globes from 57 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[DoEvents\_e4315712202001\.zip](https://github.com/Planet-Source-Code/john-galanopoulos-doevents-evolution-the-api-approach-method-for-100-optimized-loops__1-29735/archive/master.zip)





### Source Code

<p class="MsoTitle" style="margin-top:0cm;margin-right:-3.4pt;margin-bottom:0cm;
margin-left:0cm;margin-bottom:.0001pt" align="left"><span lang="EN-GB"><b><font color="#0000FF" size="4">DoEvents
evolution; the API approach</font></b></span></p>
<p class="MsoTitle" style="margin-top:0cm;margin-right:-3.4pt;margin-bottom:0cm;
margin-left:0cm;margin-bottom:.0001pt" align="left">                                                                        </p>
<p class="MsoBodyText3"><span style="font-size: 10.0pt; font-family: Verdana" lang="EN-GB">If there was such a function to
inspect the message queue for user input, we would have a main benefit:</span></p>
<p class="MsoBodyText" style="margin-right:1.3pt"><i><span lang="EN-GB" style="font-size:10.0pt;font-family:Verdana;color:windowtext"><span style="mso-spacerun: yes">      
</span>We would speed up our loops ‘cause we would process all the messages in
the queue (with DoEvents) only on user input. It’s <u>faster</u> to check for
a message than to process all messages every time. <o:p>
</o:p>
</span></i></p>
<p class="MsoBodyText3"><span lang="EN-GB"> <span style="font-size: 10.0pt; font-family: Verdana">API provides us with such a
function:</span></span></p>
<p class="MsoNormal" style="margin-right:1.3pt;mso-margin-top-alt:auto;
mso-margin-bottom-alt:auto"><span lang="EN-GB" style="font-size:10.0pt;
font-family:Verdana"><span style="mso-spacerun: yes"> </span>It’s called
GetInputState and you can locate it in user32 library.<o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-right:1.3pt;mso-margin-top-alt:auto;
mso-margin-bottom-alt:auto"><span lang="EN-GB" style="font-size:10.0pt;
font-family:Verdana"> Here is the declaration: <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-right:1.3pt;mso-margin-top-alt:auto;
mso-margin-bottom-alt:auto"><span lang="EN-GB" style="font-size:10.0pt;
font-family:Verdana"><span style="mso-spacerun: yes">    </span> 
<span style="color:#333399">Public Declare Function GetInputState Lib
"user32" () As Long <o:p>
</o:p>
</span></span></p>
<p class="MsoBodyText3"><span lang="EN-GB" style="font-size:10.0pt;
font-family:Verdana"> </span><span style="font-size: 10.0pt; font-family: Verdana" lang="EN-GB">The GetInputState function
determines whether there are mouse-button or keyboard messages in the calling
thread's message queue.</span></p>
<p class="MsoNormal" style="margin-right:1.3pt;mso-margin-top-alt:auto;
mso-margin-bottom-alt:auto"><span lang="EN-GB" style="font-size:10.0pt;
font-family:Verdana">If the queue contains one or more new mouse-button or
keyboard messages, the return value is nonzero else if there are no new
mouse-button or keyboard messages in the queue, the return value is zero. <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-right:1.4pt;mso-margin-top-alt:auto;
mso-margin-bottom-alt:auto"><span lang="EN-GB" style="font-size:10.0pt;
font-family:Verdana">So we can create an improved DoEvents with a Subroutine
like this : <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-right:1.4pt;mso-margin-top-alt:auto;
mso-margin-bottom-alt:auto"><span lang="EN-GB" style="font-size:10.0pt;
font-family:Verdana;color:#333399">Public Sub newDoEvents() <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-right:1.4pt;mso-margin-top-alt:auto;
mso-margin-bottom-alt:auto"><span lang="EN-GB" style="font-size:10.0pt;
font-family:Verdana;color:#333399"><span style="mso-spacerun: yes">       
</span>If GetInputState() <> 0 then DoEvents <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-right:1.4pt;mso-margin-top-alt:auto;
mso-margin-bottom-alt:auto"><span lang="EN-GB" style="font-size:10.0pt;
font-family:Verdana;color:#333399">End Sub <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="mso-margin-top-alt:auto;mso-margin-bottom-alt:auto"><span lang="EN-GB" style="font-size:10.0pt;font-family:Verdana"> You
can use GetInputState() with many variations for example :<span style="mso-spacerun:
yes">            </span><o:p>
</o:p>
</span></p>
<p class="MsoNormal"><span lang="EN-GB" style="font-size:10.0pt;font-family:Verdana;
color:#333399"><span style="mso-spacerun: yes">        
</span>uCancelMode = False <o:p>
</o:p>
</span></p>
<p class="MsoNormal"><span lang="EN-GB" style="font-size:10.0pt;font-family:Verdana;
color:#333399"><span style="mso-spacerun: yes">         
       </span>Do until rs.Eof <o:p>
</o:p>
</span></p>
<p class="MsoNormal"><span lang="EN-GB" style="font-size:10.0pt;font-family:Verdana;
color:#333399"><span style="mso-spacerun: yes">                        
</span>Rs.AddNew  <o:p>
</o:p>
</span></p>
<p class="MsoNormal"><span lang="EN-GB" style="font-size:10.0pt;font-family:Verdana;
color:#333399"><span style="mso-spacerun:
yes">                                 
</span>(..your source here)<o:p>
</o:p>
</span></p>
<p class="MsoNormal"><span lang="EN-GB" style="font-size:10.0pt;font-family:Verdana;
color:#333399"><span style="mso-spacerun: yes">                          
</span>Rs.Update <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-right:1.4pt"><span lang="EN-GB" style="font-size:10.0pt;font-family:Verdana;color:#333399"><span style="mso-spacerun: yes">                          
</span>Rs.MoveNext <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-right:1.4pt"><span lang="EN-GB" style="font-size:10.0pt;font-family:Verdana;color:#333399"> <span style="mso-spacerun: yes">                            
</span>If GetInputState() <> 0 then <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-right:1.4pt"><span lang="EN-GB" style="font-size:10.0pt;font-family:Verdana;color:#333399"><span style="mso-spacerun: yes">                                
</span>DoEvents <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-right:1.4pt"><span lang="EN-GB" style="font-size:10.0pt;font-family:Verdana;color:#333399"><span style="mso-spacerun: yes">                                 
</span>If uCancelMode Then Exit Do <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-right:1.4pt;mso-margin-top-alt:auto;
mso-margin-bottom-alt:auto"><span lang="EN-GB" style="font-size:10.0pt;
font-family:Verdana;color:#333399"><span style="mso-spacerun:
yes">                             
</span>End If <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-right:1.4pt"><span lang="EN-GB" style="font-size:10.0pt;font-family:Verdana;color:#333399"><span style="mso-spacerun: yes">      
</span><span style="mso-tab-count:1">   </span><span style="mso-spacerun: yes">          </span><o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-right:1.4pt"><span lang="EN-GB" style="font-size:10.0pt;font-family:Verdana;color:#333399"><span style="mso-spacerun: yes">                        
</span>Loop <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-right:1.4pt"><span lang="EN-GB" style="font-size:10.0pt;font-family:Verdana;color:#333399"><span style="mso-spacerun: yes">                        
</span><o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-right:1.4pt"><span lang="EN-GB" style="font-size:10.0pt;font-family:Verdana;color:#333399"><span style="mso-spacerun: yes">                   
</span><span style="mso-spacerun:
yes">     </span>Msgbox “Finished.” <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-right:1.4pt"><span lang="EN-GB" style="font-size:10.0pt;font-family:Verdana"> <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-right:1.4pt"><span lang="EN-GB" style="font-size:10.0pt;font-family:Verdana">…or
we could use it in a ScreenSaver e.t.c. <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-right:1.4pt"><span lang="EN-GB" style="font-size:10.0pt;font-family:Verdana">Let’s
go a little further now and see what exactly is behind GetInputState(). <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-right:1.4pt"><span lang="EN-GB" style="font-size:10.0pt;font-family:Verdana">It
is another API function located in User32 as well; GetQueueStatus() <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-right:1.4pt;mso-margin-top-alt:auto;
mso-margin-bottom-alt:auto"><span lang="EN-GB" style="font-size:10.0pt;
font-family:Verdana">The GetQueueStatus function indicates the type of messages
found in the calling thread's message queue. Here are the flags that
GetQueueStatus uses : <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana">QS_ALLEVENTS<span style="mso-spacerun:
yes">             </span>An
input, WM_TIMER, WM_PAINT,<span style="mso-spacerun:
yes"> </span>WM_HOTKEY, or posted message is in the queue. <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana">QS_ALLINPUT <span style="mso-tab-count:1">         
</span><span style="mso-spacerun: yes">    </span>Any
message is in the queue. <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana">QS_ALLPOSTMESSAGE<span style="mso-spacerun: yes">   
</span>A posted message (other than those listed here) is in the queue. <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana">QS_HOTKEY<span style="mso-spacerun: yes">   
</span><span style="mso-tab-count:1">         
</span><span style="mso-spacerun: yes">    </span>A
WM_HOTKEY message is in the queue. <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana">QS_INPUT <span style="mso-tab-count:1">                 
</span><span style="mso-spacerun:
yes"> </span>An input message is in the queue. <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana">QS_KEY <span style="mso-tab-count:1">                 
</span><span style="mso-spacerun:
yes">    </span>A WM_KEYUP, WM_KEYDOWN,<span style="mso-spacerun:
yes"> </span>WM_SYSKEYUP, or WM_SYSKEYDOWN<span style="mso-tab-count:1">                                 
</span><span style="mso-spacerun: yes">                                
 </span>message is in the queue. <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana">QS_MOUSE <span style="mso-tab-count:1">                
</span>A WM_MOUSEMOVE message or mouse-button<span style="mso-spacerun:
yes"> </span>message (WM_LBUTTONUP, WM_RBUTTONDOWN, <span style="mso-tab-count:
1">          </span><span style="mso-tab-count:1">                   
</span><span style="mso-spacerun:
yes">    </span>and so on). <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana">QS_MOUSEBUTTON <span style="mso-tab-count:1">     
</span>A mouse-button message (WM_LBUTTONUP, WM_RBUTTONDOWN, and so on). <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana">QS_MOUSEMOVE<span style="mso-tab-count:1">         
</span>A WM_MOUSEMOVE message is in the queue. <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana">QS_PAINT <span style="mso-tab-count:1">                 
</span>A WM_PAINT message is in the queue. <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana">QS_POSTMESSAGE <span style="mso-tab-count:1">     
</span>A posted message (other than those listed<span style="mso-tab-count:1"> </span>here)
is in the queue. <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana">QS_SENDMESSAGE <span style="mso-tab-count:1">     
</span>A message sent by another thread or<span style="mso-tab-count:1"> </span>application
is in the queue. <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana">QS_TIMER <span style="mso-tab-count:1">                 
</span>A WM_TIMER message is in the queue. <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana"> <o:p>
</o:p>
</span></p>
<p class="MsoBlockText"><span lang="EN-GB" style="color:#333399">         
</span><span lang="EN-GB" style="font-size: 10.0pt; font-family: Verdana">(I believe that GetInputState() is a GetQueueStatus(QS_HOTKEY Or QS_KEY Or
QS_MOUSEBUTTON)) <o:p>
</span><span lang="EN-GB" style="color:#333399">
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana">With these constants you can create your own
GetInputState function that fits your needs. For example you can create a custom
function that issues DoEvents when it’ll detects not only a Keyboard or Mouse<br>
Key input, but also a WM_PAINT signal. <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana">Why’s that? ‘cause in your loop you might need
to update the screen so you must let your custom function process the specific
signal. <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana">Look at this : <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana"> <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana;color:#333399">Public Const QS_HOTKEY = &H80 <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana;color:#333399">Public Const QS_KEY = &H1 <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana;color:#333399">Public Const QS_MOUSEBUTTON = &H4 <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana;color:#333399">Public Const QS_MOUSEMOVE = &H2 <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana;color:#333399">Public Const QS_PAINT = &H20 <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana;color:#333399">Public Const QS_POSTMESSAGE = &H8 <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana;color:#333399">Public Const QS_SENDMESSAGE = &H40
<o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana;color:#333399">Public Const QS_TIMER = &H10 <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana;color:#333399">Public Const QS_ALLINPUT = (QS_SENDMESSAGE
Or QS_PAINT Or<span style="mso-spacerun: yes"> </span>QS_TIMER Or
QS_POSTMESSAGE Or<span style="mso-tab-count:1">                            
</span><span style="mso-tab-count: 1; font-size: 10.0pt; font-family: Verdana; color: #333399" lang="EN-GB">             
</span><span style="mso-spacerun:
yes">   </span>QS_MOUSEBUTTON Or QS_MOUSEMOVE Or QS_HOTKEY Or
QS_KEY) <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana;color:#333399">Public Const QS_MOUSE = (QS_MOUSEMOVE
Or QS_MOUSEBUTTON) <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana;color:#333399">Public Const QS_INPUT = (QS_MOUSE Or
QS_KEY) <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana;color:#333399">Public Const QS_ALLEVENTS = (QS_INPUT
Or QS_POSTMESSAGE Or<span style="mso-spacerun: yes"> </span>QS_TIMER Or
QS_PAINT Or QS_HOTKEY) <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana;color:#333399"> <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana;color:#333399">Public Declare Function GetQueueStatus
Lib "user32" (ByVal qsFlags As Long) As Long <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana;color:#333399"> <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana;color:#333399"> <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana;color:#333399">Public Function cGetInputState() <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt;text-indent:36.0pt"><span lang="EN-GB" style="font-size:10.0pt;font-family:Verdana;color:#333399">Dim
qsRet As Long <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:108.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana;color:#333399">qsRet = GetQueueStatus(QS_HOTKEY Or
QS_KEY Or QS_MOUSEBUTTON Or QS_PAINT) <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana;color:#333399"><span style="mso-spacerun:
yes">                  
</span>cGetInputState = qsRet <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana;color:#333399">End Function </span><span lang="EN-GB" style="font-size:10.0pt;font-family:Verdana"><o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana"> <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana">With this function you can trigger the DoEvents to
be executed only when the message queue contains Key input, Mouse button or a
WM_PAINT signal. <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana"> <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana">Call it like this…. <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana;color:#333399">. . if cGetInputState() <> 0
then DoEvents <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana"> <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-right:1.3pt"><span lang="EN-GB" style="font-size:10.0pt;font-family:Verdana"><span style="mso-spacerun:
yes">          </span>This was
tested and proved to optimise a loop<span style="mso-spacerun: yes">  </span>by
100% !!!!!!!!!<o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana"> <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana">I wrote this article believing that the API is a
powerfull part on Windows programming and deserves your attention. I was stuck
several times and API prooved to be a problem solver. API is a large world but
with little effort, you can take advantage of it. You will create more
sophisticated and user aware programs. <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana"> <o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana">I hope I helped.<o:p>
</o:p>
</span></p>
<p class="MsoNormal" style="margin-top:0cm;margin-right:1.3pt;margin-bottom:0cm;
margin-left:36.0pt;margin-bottom:.0001pt"><span lang="EN-GB" style="font-size:
10.0pt;font-family:Verdana">Any comments or suggestions are always welcomed. </span><span lang="EN-GB" style="font-size:10.0pt;font-family:Verdana"><o:p>
</o:p>
</span></p>
<h1><span lang="EN-GB"><font size="4">        
</font><font size="3">John
Galanopoulos</font></span></h1>
<p class="MsoNormal"><span lang="EN-GB">          
(Below there is a link to the .doc version of this article, for you to download. <br>If you want to implement this source in your projects, download the Class Module posted by http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=33401 John Baughman in this address <br>Also, you can check out Olav Jordan's article : http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=37888&lngWId=1 Optimized loop (no more doevents) </span></p>
<br><b>Need Oracle tips? try here : http://aboutoracle.blogspot.com<b><br>

