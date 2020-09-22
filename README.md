<div align="center">

## Properties: Not Just For Classes Anymore


</div>

### Description

Shows a methodology for drastically reducing the amount of time needed to develop "similar but not quite" forms for a type of task.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Devin Watson](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/devin-watson.md)
**Level**          |Intermediate
**User Rating**    |4.8 (24 globes from 5 users)
**Compatibility**  |VB 6\.0
**Category**       |[Object Oriented Programming \(OOP\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/object-oriented-programming-oop__1-47.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/devin-watson-properties-not-just-for-classes-anymore__1-25707/archive/master.zip)





### Source Code

<p><strong><font face="Arial">Properties: Not Just For Classes Anymore</font></strong></p>
<p><font face="Arial">Often times, I've run into a situation where I say to myself &quot;I
wish I could've reused the same form for this purpose.&quot; Early on, it was just that; I
would copy a form's code and controls over to a new one, adjusting as necessary. Of
course, this led to bloated executables and more spaghetti code than I care to remember
(it still makes me shudder to think about it.)</font></p>
<p><font face="Arial">However, I did find a unique solution, whilst developing something
else entirely: a class module. When developing the properties of a particular class, I
wondered if it would be possible to actually do the same to a form. In essence, take a
standard VB form and then add my own custom properties. I had done it before using custom
methods, but would VB throw up on me for doing something I shouldn't?</font></p>
<p><font face="Arial">Surprisingly, VB didn't complain one bit, and I was able to reduce
four forms that were exactly the same except for a few &quot;under-the-hood&quot; changes.
These changes were easily be implemented as properties, and what would have taken me
several hours to create and debug, now only took 15 minutes. In the rest of this article,
I'll discuss and show how this methodology can be used to reduce development time even
more, using a simple but powerful example.</font></p>
<p><font face="Arial">Start up Visual Basic with a Standard EXE project. In the project,
add another, blank form. Inside this new form, open up the Code Editor window. In the
General Declarations area, type the following:</font></p>
<p><font face="Arial">Dim mstrTable As String<br>
<br>
And then type:<br>
<br>
Public Property Let Table(strTable As String)</font></p>
<p><font face="Arial">You'll pop into the new property's code, just as if you typed it in
a class module. Inside this new property, type the following:<br>
<br>
Public Property Let Table(strTable As String)<br>
<br>
mstrTable = Trim(strTable)<br>
<br>
If Not IsValid(mstrTable) Then<br>
&nbsp;&nbsp;&nbsp;&nbsp; Err.Raise 30000, &quot;Data Table Editor&quot;, &quot;Table name
appears to be invalid.&quot;<br>
&nbsp;&nbsp;&nbsp;&nbsp; Exit Property<br>
End If<br>
<br>
Me.Caption = &quot;Table Editor - &quot; &amp; mstrTable<br>
GetTableItems mstrTable<br>
Me.Visible = True<br>
Me.ZOrder 0<br>
<br>
End Property<br>
<br>
In the example above, setting the new Table property for the form performs several
functions. First, it removes leading and trailing spaces, then checks to see if the name
given is, in fact valid ( the function IsValid is not shown here, but rather used to point
out that your validation can remain within a form). If the table name is not valid, then
an error is raised. This way, the calling form would have something as follows:<br>
<br>
On Local Error Resume Next<br>
<br>
frmEdit.Table = &quot;monkeys&quot;<br>
If Err.Number = 30000 Then Exit Sub<br>
<br>
However, if the table name given is valid, then it goes on to change the caption to
reflect the new table name, and the items from that table are retrieved using another
private subroutine called GetTableItems. Once that has completed, then the form is made
visible, and moved to the forefront of the screen. <br>
<br>
<strong>Caveats<br>
</strong><br>
Using this method can save time by reducing the amount of code you type over and over.
However, it is advisable to use the Load statement to cache the form in memory during
program load time, so as to avoid any &quot;Object or With Variable Not Set&quot; errors.<br>
<br>
<strong>Is This The End Of Zombie Shakespeare?</strong><br>
<br>
Nope. You can use this method to infinitely extend a form. Just remember to make sure you
generalize it enough so that it encompasses the job you want to do, and that you cache it
up in memory during program load. <br>
<br>
Happy Programming!<br>
</font></p>

