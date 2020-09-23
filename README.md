<div align="center">

## Beginners Tutorial


</div>

### Description

This tutorial is intended as a guide for people new to Visual Basic.

It shows common coding conventions and basic use of VB.

----

1. Option Explicit

2. Code Formatting

3. Commenting

4. Variable Types

5. Static Variables

6. Global/Local Variables

7. Public/Private Functions

8. Arrays

9. Constants

10. Control Names

11. Variable/Constant/Procedure Names

----

NB: The zip file has the HTML file of the tutorial.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2002-02-12 22:27:22
**By**             |[JoshD](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/joshd.md)
**Level**          |Beginner
**User Rating**    |4.4 (79 globes from 18 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Beginners\_547972122002\.zip](https://github.com/Planet-Source-Code/joshd-beginners-tutorial__1-31732/archive/master.zip)





### Source Code

<style type="text/css">
<!--
.heading { font-family: Arial, Helvetica, sans-serif; font-size: 20px; font-weight: bold}
.text { font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px}
.code { font-family: "Courier New", Courier, mono; font-size: 12px; color: #003399; clip:  rect(  )}
.comment { font-family: "Courier New", Courier, mono; font-size: 12px; color: #339900}
.note { font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold; color: #FF0000}
.a:hover { text-decoration:underline}
.link { font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px ; text-decoration: none }
.linktop { font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px ; text-decoration: none ; font-weight: normal}
-->
</style>
<p><span class="heading">Beginners Tutorial</span></p>
<p><br>
 This is my first tutorial to PSC, so please be kind with any comments... </p>
<p>The tutorial is targeted at users who are new to Visual Basic and provides
 a few simple basic tips that will help you in your coding. These are all useful
 tips that will speed up your programming and, as your programs get larger, your
 debugging.</p>
<p>Like many people I am sick of downloading a program to see that the code is
 along the lines of:<br>
</p>
<p><span class="code">Private Sub Command4_Click()<br>
 Dim a, b, c, d<br>
 a = Text12.Text<br>
 b = Text3.Text<br>
 If a = b Then<br>
 ...</span></p>
<p>This tutorial forms a guide on how to write your code to look professional.</p>
<p>Syntax:<br>
 <span class="code">Code</span> represents code examples. <br>
 <span class="comment">Text</span> represents a comment. <br>
 <span class="note">Note:</span> Explains some of the code in the examples.<br>
 Procedure: a Sub or function.</p>
<p class="heading">Contents<a name="top"></a></p>
<p><a href="#option" class="link">1. Option Explicit</a><br>
 <a href="#formatting" class="link">2. Code Formatting</a><br>
 <a href="#commenting" class="link">3. Commenting<br>
 </a><a href="#types" class="link">4. Variable Types<br>
 </a><a href="#static" class="link">5. Static Variables<br>
 </a><a href="#global" class="link">6. Global/Local Variables<br>
 </a><a href="#public" class="link">7. Public/Private Functions<br>
 </a><a href="#arrays" class="link">8. Arrays<br>
 </a><a href="#constants" class="link">9. Constants<br>
 </a><a href="#control" class="link">10. Control Names<br>
 </a><a href="#variable" class="link">11. Variable/Constant/Procedure Names</a></p>
<p class="heading">Option Explicit<a name="option"></a> <a href="#top" class="linktop">(top)</a></p>
<p>Option Explicit is an extremely useful function of VB, if you refer to a variable
 that does not exist VB will stop executing the code and inform you of this.
 If you do not include Option Explicit and misname a variable then VB will assume
 create a new variable with a value of 0 or null. For example:</p>
<p class="code">Dim myName as String<br>
 myName = &quot;BOB&quot;<br>
 If myyName = &quot;&quot; Then<br>
 &nbsp;&nbsp;&nbsp;MsgBox
 &quot;Hello &quot; &amp; myName<br>
 End If</p>
<p class="comment">
<p>If <span class="code">Option Explicit</span> is used VB will stop when the
 IF statement is reached. If <span class="comment">Option Explicit</span> is
 not used the code will not display a message, as VB will assume <span class="code">myyName</span>
 is different from <span class="code">myName</span>. It will create the variable
 <span class="code">myyName</span> (of type Variant) with an initial value of
 <span class="code"> &quot;&quot;</span>, because of this the message box will
 not be displayed.</p>
<p>While this mistake appears obvious, and would be easy to debug, when thousands
 of lines are used a small mistake can stop the entire code from working.</p>
<p>The <span class="code">Option Explicit</span> keywords should be the topmost
 line of code for a form or module, before any declarations or procedures.</p>
<p class="heading">Code Formatting<a name="formatting"></a> <a href="#top" class="linktop">(top)</a></p>
<p>The most important aspect when writing code for your programs is the format
 of the code. This is the first thing another programmer will see when they open
 your program, and they will be more likely to peruse and attempt to understand
 your code if it appears friendlier. Formatting is easy to remember:</p>
<p>Inside a procedure code will be should indented one tab.</p>
<p class="code"> Public Sub cmdAbout_Click()<br>
 &nbsp;&nbsp;&nbsp;MsgBox
 &quot;(c) me, 2002&quot;<br>
 End Sub</p>
<p>For each <span class="code">IF</span> statement, <span class="code">FOR</span>
 or <span class="code">DO</span> loop the code should be indented an additional
 tab.</p>
<p class="code"> Public Sub cmdDisplay_Click()<br>
 &nbsp;&nbsp;&nbsp;Dim
 i as Integer<br>
 &nbsp;&nbsp;&nbsp;For
 i = 1 to 10<br>
 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;txtResults.Text
 = txtResults.Text &amp; i &amp; vbTab &amp; i^2 &amp; vbCrLf<br>
 &nbsp;&nbsp;&nbsp;Next
 i <br>
 &nbsp;&nbsp;&nbsp;MsgBox
 &quot;(c) me, 2002&quot;<br>
 End Sub</p>
<p>This is far easier to read than</p>
<p class="code"> Public Sub cmdDisplay_Click()<br>
 Dim i as integer<br>
 For i = 1 to 10<br>
 txtResults.text = txtResults.text &amp; i &amp; vbTab &amp; i^2 &amp; vbCrLf<br>
 Next i <br>
 MsgBox &quot;(c) me, 2002&quot;<br>
 End Sub</p>
<p><span class="note">Note:</span> <span class="code">vbTab</span> adds a tab
 to the text, <span class="code">vbCrLf</span> adds a Carriage Return and Line
 Feed characters, which together make the text go to a new line. These are both
 VB constants, explained below. <span class="code">i^2</span> is simply a mathematical
 function, representing<span class="code"> i x i</span><span class="text">,</span><span class="code">
 ^</span> represents the mathematical function power. Therefore<span class="code">
 i^3</span> means <span class="code">i x i x i</span>.</p>
<p>VB does not insert a Tab character, instead it will insert spaces that act
 in a similar way. The number of spaces per tab can be set in the program options.</p>
<p class="heading">Commenting<a name="commenting"></a> <a href="#top" class="linktop">(top)</a></p>
<p>A comment is a piece of text in your program that is not intended to be executed
 as a command. A comment is indicated by a single quotation mark before the text.
 The comment can appear on the same line as code but anything after the quotation
 will be considered a comment</p>
<p>Commenting is very useful to remind yourself what code does and is very important
 in indicating to other programmers unfamiliar with your code what each piece
 of code does.</p>
<p>If submitting a program to PSC, a good idea is to add your name, e-mail, the
 project name and the date at the top of the first from/module this makes it
 easier for other users to contact you if they wish to use your code. You also
 may wish to add a short description of your program. e.g.</p>
<p class="comment"> '****************************************<br>
 'Program: Date Finder<br>
 'Author: I. Rule &lt;rulei@fictitional.com&gt;<br>
 'Date: 1/1/2002<br>
 '<br>
 'This program will accept a date and check<br>
 'whether it is a leap year.<br>
 '****************************************</p>
<p>Comments should be used within the code to explain its function. </p>
<p> <span class="code">Dim leap as Boolean<br>
 leap = (year Mod 4 = 0) </span><span class="comment">&nbsp;&nbsp;'Is this a
 leap year? </span></p>
<p>If a line of code is long it may be easier to write the comment before the
 line of code, to save yourself continually scrolling to see the comments.</p>
<p class="code"> Dim leap as Boolean<br>
 <span class="comment">'Is this a leap year? </span><br>
 leap = (year Mod 4 = 0)</p>
<p>It is recommended that you do not over comment this can make code appear cluttered
 and if unnecessary, will slow you down. For example:</p>
<p> <span class="code">Dim leap as Boolean, year as Integer</span><br>
 <span class="code">year = Val(cmdYear.Text)</span><span class="comment">&nbsp;'Get
 the year from the text box</span><br>
 <span class="code">leap = (year Mod 4 = 0) </span><span class="comment">&nbsp;'Is
 this a leap year? </span><br>
 <span class="code">MsgBox leap </span><span class="comment">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;'Display
 a message box indicating whether this is a leap year.</span></p>
<p><span class="note">Note:</span> <span class="code">Val(text)</span> will give
 you the numerical value text, this is useful incase the user types a letter
 instead of a number, which would cause the program to crash, instead <span class="code">Val(text)</span>
 will return 0. The <span class="code">Mod</span> operator gives the remainder,
 so <span class="code">5 Mod 2</span> would give 1(the remainder of 5 divided
 by 2). The formatting of <span class="code">leap = (year Mod 4 = 0)</span> is
 an easy way to get a <span class="code">True/False</span> value. <span class="code">year
 Mod 4</span> gives a number, if this number is 0 then <span class="code">(year
 Mod 4 = 0)</span> gives <span class="code">True</span>, if not it gives <span class="code">False</span>.
 Therefore leap receives its correct value. The brackets are not necessary but
 make the code easier to understand.</p>
<p>Use your own judgment whether code needs to be commented, keeping in mind whether
 it is for your own or others benefit.</p>
<p class="heading">Variable Types<a name="types"></a> <a href="#top" class="linktop">(top)</a></p>
<p>Visual Basic has the following variable types:</p>
<table width="100%" border="1" class="text" cellspacing="0">
 <tr>
  <td><b>Type</b></td>
  <td><b>Description</b></td>
  <td><b>Example</b></td>
 </tr>
 <tr>
  <td>Date</td>
  <td>Stores a date/time combination</td>
  <td>1:46:32 AM 17-02-2002</td>
 </tr>
 <tr>
  <td>String</td>
  <td>Stores a &quot;string&quot; of text without formatting</td>
  <td>Hello Bob</td>
 </tr>
 <tr>
  <td>Integer</td>
  <td>A non decimal number -32768 to +32768</td>
  <td>199</td>
 </tr>
 <tr>
  <td>Byte</td>
  <td>An integer from 0 to 255, inclusive</td>
  <td>16</td>
 </tr>
 <tr>
  <td>Long</td>
  <td>An integer extending to billions</td>
  <td>1132434</td>
 </tr>
 <tr>
  <td>Double</td>
  <td>Stores decimal and large numbers</td>
  <td>1.0002</td>
 </tr>
 <tr>
  <td>Single</td>
  <td>Stores large numbers </td>
  <td>&nbsp;</td>
 </tr>
 <tr>
  <td>Boolean</td>
  <td>A single bit, stores True or False</td>
  <td>True</td>
 </tr>
 <tr>
  <td>Variant</td>
  <td>A variant should not be used, it can store values of any of the above
   types, but is very memory intensive and is bad programming practice to use
   them. </td>
  <td>&nbsp;</td>
 </tr>
</table>
<p>String, Integer and Boolean are the most common, but as your programs become
 more advanced you will need to use Long, Double and Date</p>
<p>A variable is declared as such:</p>
<p class="code"> Dim var as Boolean</p>
<p>If you wish to use this variable in other modules/forms declare it as public:</p>
<p class="code"> Public var as Boolean</p>
<p>A variable that is not given a type becomes a variant by default. Do not do
 this.</p>
<p class="code"> Dim var</p>
<p>To save space/time variables can be defined in a single line.</p>
<p class="code"> Dim var as Boolean, personName as String, age as Integer, siblings
 as Integer</p>
<p>Do not declare them in the following fashion</p>
<p class="code"> Dim var1, var2, var3 as Boolean</p>
<p><span class="code">var1</span> and <span class="code">var2</span> will become
 Variant type. Only <span class="code">var3</span> will be a Boolean type.</p>
<p class="heading">Static Variables<a name="static"></a> <a href="#top" class="linktop">(top)</a></p>
<p>If the keyword <span class="code">Static</span> is used in place of <span class="code">Dim</span>
 the variable will retain its value. For example:</p>
<p class="code"> Private Sub cmdCount_Click()<br>
 &nbsp;&nbsp;&nbsp;Static
 count as integer<br>
 <span class="comment">&nbsp;&nbsp;&nbsp;</span>count = count + 1<br>
 &nbsp;&nbsp;&nbsp;MsgBox
 count<br>
 End Sub </p>
<p>The first time the button is pressed (the procedure is run) a message box will
 display &quot;1&quot; the second time it would display &quot;2&quot; and so
 on. If Dim were used:</p>
<p class="comment"> Private Sub cmdCount_Click()<br>
 &nbsp;&nbsp;&nbsp;Dim
 count as integer<br>
 &nbsp;&nbsp;&nbsp;count
 = count + 1<br>
 &nbsp;&nbsp;&nbsp;MsgBox
 count<br>
 End Sub </p>
<p>The count would be re-created each time, and it would revert to 0, therefore
 pressing the button would only ever display 1. Static variables are rarely required
 and in most cases a global variable is easier to use and more suitable.</p>
<p>Static variables cannot be global, i.e. they should only exist within procedures.</p>
<p class="heading">Global/Local Variables<a name="global"></a> <a href="#top" class="linktop">(top)</a></p>
<p>A variable can be either declared locally or globally. A local variable is
 defined within a procedure and can only be accessed from within that procedure.
 A global variable is defined outside a procedure at the top of the page of code
 for that form or module. It can be accessed by all procedures in that module/form,
 and will obviously not lose its value once a procedure has finished executing.
 e.g.</p>
<p><b>frmMain:</b></p>
<p class="code"> Dim myName as Stirng &nbsp;&nbsp;<span class="comment">&nbsp;&nbsp;</span>&nbsp;<span class="comment">&nbsp;&nbsp;</span>&nbsp;<span class="comment">&nbsp;&nbsp;</span>&nbsp;&nbsp;<span class="comment">'Global
 variable - can be used by any procedure of frmMain</span><br>
 Dim myAge as Integer &nbsp;&nbsp;<span class="comment">&nbsp;&nbsp;</span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span class="comment">'Global
 variable - can be used by any procedure of frmMain</span><br>
 Public peopleCount as Integer &nbsp;&nbsp;&nbsp;<span class="comment">'Public
 global variable - notice it can be used by modChecks </span></p>
<hr width="75%" align="left" size="1" noshade>
<p class="code">Private Sub cmdRecord_Click()<br>
 <span class="comment">&nbsp;&nbsp;&nbsp;</span>Dim valid as Boolean &nbsp;&nbsp;<span class="comment">&nbsp;&nbsp;</span>&nbsp;<span class="comment">&nbsp;&nbsp;</span>&nbsp;&nbsp;<span class="comment">'A
 local variable - it can only be used in this sub</span><br>
 &nbsp;&nbsp;&nbsp;valid
 = chkValid.Value<br>
 &nbsp;&nbsp;&nbsp;If
 valid = true Then<br>
 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;myName
 = txtName.Text<br>
 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;myAge
 = txtAge.text<br>
 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;peopleCount
 = Val(txtCount.text)<br>
 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;message
 = txtMessage.text<br>
 &nbsp;&nbsp;&nbsp;End
 If<br>
 End Sub</p>
<hr width="75%" align="left" size="1" noshade>
<p class="code">Private Sub cmdDisplay_Click()<br>
 &nbsp;&nbsp;&nbsp;MsgBox
 &quot;Name &quot; &amp; myName<br>
 &nbsp;&nbsp;&nbsp;MsgBox
 &quot;Age &quot; &amp; myAge<br>
 End Sub</p>
<p class="code">&nbsp;</p>
<p><b>modChecks:</b></p>
<p class="code"> Public message as string &nbsp;&nbsp;<span class="comment">&nbsp;&nbsp;</span>&nbsp;&nbsp;<span class="comment">'Public
 global variable - notice it can be used by frmMain</span></p>
<hr width="75%" align="left" size="1" noshade>
<p class="code">Public Sub CheckMaximum()<br>
 &nbsp;&nbsp;&nbsp;If
 frmMain.peoplCount &gt; 5 Then<br>
 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;MsgBox
 message<br>
 &nbsp;&nbsp;&nbsp;End
 If<br>
 End Sub</p>
<p>Notice frmMain.peopleCount is used in CheckMaximum when referring to the variable.
 This applies when a form or module gets the value of a variable belonging to
 another form. If the variable belongs to a module then the form or module can
 directly refer to it directly, as is the case for the variable message. These
 variables are used as an example - it is up to you to decide whether the variable
 is declared in the module or form.</p>
<p class="heading">Public/Private Functions<a name="public"></a> <a href="#top" class="linktop">(top)</a></p>
<p>Similar rules to these apply to procedures. Private procedures can only be
 used within their own form or module. Public ones can be used by any form or
 module, but if they belong to a form the forms name must be written before the
 procedure name.</p>
<p>GENERAL RULE: If a variable, procedure is used in more than one form it should
 be placed in the module, if not it is best placed in the form it is used in.
</p>
<p class="heading">Arrays<a name="arrays"></a> <a href="#top" class="linktop">(top)</a></p>
<p>An array or variables is useful when you need to store values that belong in
 a list. The can be declared as follows:</p>
<p class="code"> Dim name(1 to 100) as String</p>
<p>An array can also be declared like so: </p>
<p class="code"> Dim name(100) as String</p>
<p>This would equate to:</p>
<p class="code"> Dim name(0 to 100) as String</p>
<p>In most cases when using arrays the list begins at 0, not 1.</p>
<p>The same rules regarding the Public keyword apply, except for two changes:<br>
 1. All arrays must be global (i.e. not declared within a procedure).<br>
 2. An array that is declared in a form cannot be Public.</p>
<p>It is possible to define an array without specifying a length:</p>
<p class="code"> Dim userName() as String</p>
<p>In this case the length must be specified lasted using ReDim. This will allow
 you decide on a length after some code has been executed. A ReDim statement
 can be called more than once for any array, however each time the values for
 each position in the array will be lost.</p>
<p class="code"> Dim userName() as String</p>
<hr width="75%" align="left" size="1" noshade>
<span class="code">Private Sub cmdSetLength_Click()<br>
&nbsp;&nbsp;&nbsp;Dim people as integer<br>
&nbsp;&nbsp;&nbsp;people = Val(txtPeopleCount.Text)<br>
&nbsp;&nbsp;&nbsp;ReDim userName(1 To people)<br>
End Sub </span>
<p>An array like those above represent a single dimension. It is possible to have
 as many dimensions as you like (although three is probably the maximum you will
 need). Note that if you create an array with many dimensions you may fun out
 of memory, as a 8 dimensional array like so <span class="code">userName(1 To
 10, 1 To 10, 1 to 10 ...) as String</span> is equivalent to 100,000,000 variables.</p>
<p>Multi-dimensional arrays are declared in the following fashion:</p>
<p class="code"> Dim userName(1 To 10, 1 To 10) as String</p>
<p>or</p>
<p class="code"> Dim userName() as String</p>
<hr width="75%" align="left" size="1" noshade>
<p class="code">Public Sub SetDimensions()<br>
 &nbsp;&nbsp;&nbsp;ReDim
 userName(1 To 10, 1 To 10)<br>
 End Sub</p>
<p class="heading">Constants<a name="constants"></a> <a href="#top" class="linktop">(top)</a></p>
<p>Constants provide an easy way to remember reoccurring numbers or text without
 creating a variable. Constants cannot be changed at run-time. Typical use may
 be:</p>
<p class="code"> Const PI = 3.141592653589</p>
<p>A constant can only be global. A public constant can only be declared in a
 module using the syntax:</p>
<p class="code"> Public Const PI = 3.141592653589</p>
<p>Visual Basic constants</p>
<p>Visual Basic has its own inbuilt constants. Each only begins with the prefix
 vb. Previously vbTab and vbCrLf were discussed. The most commonly used are for
 the basic colours and for key ASCII values (used in the KeyDown and KeyPress
 procedures). For a full list check your help file.</p>
<table width="30%" border="1" cellspacing="0">
 <tr>
  <td class="text" width="48%"><b>Example</b></td>
  <td class="text" width="52%"><b>Value</b></td>
 </tr>
 <tr>
  <td class="code" width="48%">vbKeyEscape</td>
  <td class="text" width="52%">27</td>
 </tr>
 <tr>
  <td class="code" width="48%">vbKeyRight</td>
  <td class="text" width="52%">39</td>
 </tr>
 <tr>
  <td class="code" width="48%">vbBlue</td>
  <td class="text" width="52%">16711680</td>
 </tr>
</table>
<p class="heading">Control Names<a name="control"></a> <a href="#top" class="linktop">(top)</a></p>
<p>Naming controls is an important part of programming. It should be done as each
 control is placed on the form not after the entire layout is designed. The prefix
 for each control should be standard and suggest what type of control it is.
 The generally accepted prefixes for the most common controls are:</p>
<table width="87%" border="1" cellspacing="0">
 <tr>
  <td class="text" width="92%"><b>Control</b></td>
  <td class="text" width="8%"><b>Prefix</b></td>
 </tr>
 <tr>
  <td class="text" width="92%">Command Button</td>
  <td class="code" width="8%">cmd</td>
 </tr>
 <tr>
  <td class="text" width="92%">
   <p>Label<br>
    <i>If a label is not used for input or output it is acceptable to leave
    its default name.</i></p>
   </td>
  <td class="code" width="8%">lbl</td>
 </tr>
 <tr>
  <td class="text" width="92%">Text Box</td>
  <td class="code" width="8%">txt</td>
 </tr>
 <tr>
  <td class="text" width="92%">Form</td>
  <td class="code" width="8%">frm</td>
 </tr>
 <tr>
  <td class="text" width="92%">Picture Box</td>
  <td class="code" width="8%">pic</td>
 </tr>
 <tr>
  <td class="text" width="92%">Image</td>
  <td class="code" width="8%">img</td>
 </tr>
 <tr>
  <td class="text" width="92%">Timer</td>
  <td class="code" width="8%">tmr</td>
 </tr>
 <tr>
  <td class="text" width="92%">Menu</td>
  <td class="code" width="8%">mnu</td>
 </tr>
 <tr>
  <td class="text" width="92%">Check Box</td>
  <td class="code" width="8%">chk</td>
 </tr>
 <tr>
  <td class="text" width="92%">Option Button</td>
  <td class="code" width="8%">opt</td>
 </tr>
 <tr>
  <td class="text" width="92%">etc.</td>
  <td class="text" width="8%">&nbsp;</td>
 </tr>
</table>
<p>By doing this your code will be far easier to interpret, a control named txtAge
 is far more descriptive that one named age, which could be a button, textbox
 or even a label. It also means that you don't run into trouble when you want
 (for example) a button and textbox to have the same name. Controls can also
 exist in arrays - if you copy and paste a control it will ask you if you would
 like to create an array. This functionality is rarely needed, but if it is the
 Index property represents its position in the array. Controls can only exist
 is one-dimensional arrays.</p>
<p class="heading">Variable/Constant/Procedure Names<a name="variable"></a> <a href="#top" class="linktop">(top)</a></p>
<p>The naming convention of variables/constants/procedures is not vital, but is
 good programming practice and is extremely helpful to someone examining your
 code. There is one rule for each type:<br>
 <br>
</p>
<table width="100%" border="1" class="text" cellspacing="0">
 <tr>
  <td width="9%"><b>Type</b></td>
  <td width="64%"><b>Rule</b></td>
  <td width="27%"><b>Example</b></td>
 </tr>
 <tr>
  <td width="9%">Variables</td>
  <td width="64%">The name has correct cases except for the first letter, which
   is always lower case.</td>
  <td class="code" width="27%">myName</td>
 </tr>
 <tr>
  <td width="9%">Constants</td>
  <td width="64%">The entire name is uppercase.</td>
  <td class="code" width="27%">DAYSPERYEAR or DAYS_PER_YEAR</td>
 </tr>
 <tr>
  <td width="9%">Procedures</td>
  <td width="64%">All letters are correct cases.</td>
  <td class="code" width="27%">CalculateAverage</td>
 </tr>
</table>
<p>If you have declared a variable then VB will automatically convert all instances
 of it in your code to the case you used when you declared it. This can be used
 if you are unsure what the name of the variable was, if you type is deliberately
 in the wrong case and it is correct the case will be changed to mach that of
 the declaration.</p>
<p>For example if you are unsure if the variable <span class="code">boxColour</span>
 was spelt <span class="code">boxColour</span> or <span class="code">boxColor</span>
 then type: <span class="code">BOXCOLOR</span> and press space/enter, because
 there is no variable by this name the case will remain unchanged. Therefore
 if you retype the text as to <span class="code">BOXCOLOUR</span> the case will
 automatically change to <span class="code">boxColour</span>.<br>
</p>

