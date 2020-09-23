<div align="center">

## msn messenger codes


</div>

### Description

This Is A Great Code Very nice Shows u how to do an offline online flood and get ur contact list to vb and change ur nickname ... Very easy. If u want to learn msn messenger coding start this way!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[mario achkar](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mario-achkar.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mario-achkar-msn-messenger-codes__1-41509/archive/master.zip)





### Source Code

```
SOURCE CODES! By MARIO ACHKAR
Special thanks to planet - source - code
Mario 's Source Codes by Mario Achkar:
1st of all if u would like to know how to do an online offline status flood on MSN (Microsoft service network) messenger
Note: 1st of all go to project, references, and check all 4 thing that begin with messenger like messenger type library or messenger API then make a command button in your form
Now Copy the Code below and just Paste It
HERE IS THE CODE:
Public Withevents MSN as msgrobject
Private Sub Form_Load()
Set MSN = new msgrobject
End Sub
Private Sub Command1_Click()
msn.localstate = mstate_invisible
msn.localstate = mstate_online
msn.localstate = mstate_invisible
msn.localstate = mstate_online
msn.localstate = mstate_invisible
msn.localstate = mstate_online
msn.localstate = mstate_invisible
msn.localstate = mstate_online
form1.caption = "MARIO HAS MADE THIS CODE"
End Sub
'NOW YOU HAVE AN ONLINE OFFLINE FLOOD EASY HAH?
HERE'S AN OTHER CODE THAT WILL ENABLE YOU TO CHANGE YOUR NICKNAME TO ANYTHING YOU WANT EXEPT FOR BAD WORDS ...
IN THE SAME PROJECT THAT U MADE BEFORE ADD AN OTHER COMMAND BUTTON AND A TEXTBOX.
NOW COPY AND PASTE THIS CODE :
Private Sub Command2_Click()
MSN.services.primaryservice.friendlyname = text1.text
End Sub
'Now run the program and put the nickname in the textbox and click the 2nd command button which u made now your nickname
will be just like the text box.
How To Get ur contact list into a list box.First Make A list Box then go to the code (don't forget to check the references) now put
Public Withevents MSN as msgrobject
dim user as imsgruser
now make a command button
now put
Private Sub Form_Load()
Set MSN = new msgrobject
End Sub
Private Sub Command1_click()
For each user in msn.list(mlist_contact)
List1.AddItem user.FriendlyName
Next
end sub
And There u Go!!!!!   By Mario
```

