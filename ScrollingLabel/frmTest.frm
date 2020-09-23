VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Test form"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   7935
   StartUpPosition =   2  'CenterScreen
   Begin ScrollingLabel.ctlScrollingLabel ctlScrollingLabel1 
      Height          =   2565
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2445
      _extentx        =   4313
      _extenty        =   4577
      forecolor       =   -2147483630
      scrollingintevral=   200
      font            =   "frmTest.frx":0000
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim strText1 As String
Dim strText2 As String
Dim strText3 As String
Dim strExampleText As String

'There is an example text from PSC, "About the Site" section...
'As you se, you have to use vbNewLine to get text in separate lines.

strText1 = "The idea for this site came to me back in 1997, when I" & vbNewLine & _
"began looking for Visual Basic source code on the Internet." & vbNewLine & _
"I was frustrated by the lack of quantity of source code," & vbNewLine & _
"as well as the amount of time I had to spend downloading" & vbNewLine & _
".zip files which were of questionable quality and relevance" & vbNewLine & _
"to what I was looking for. I thought it would be really" & vbNewLine & _
"useful if there was a  site that allowed me to actually see" & vbNewLine & _
"the code before I downloaded it, and maybe even let me copy" & vbNewLine & _
"and paste it from my browser to VB, so I didn't have to go" & vbNewLine & _
"through the hassle of unzipping it."

strText2 = "Thus www.Planet-Source-Code.com was born. Back then sites" & vbNewLine & _
"that had databases were VERY rare (maybe a handful of them" & vbNewLine & _
"existed). To put it in context, this was before most" & vbNewLine & _
"browsers supported frames or even tables.  The prevailing" & vbNewLine & _
"web scripting technology at the time was CGI and Perl which" & vbNewLine & _
"required alot of patience and time. However, when I heard" & vbNewLine & _
"of a strange new tool from Microsoft called" & vbNewLine & _
"Visual Interdev 1.0, I was intrigued enough to plunk down" & vbNewLine & _
"some cash for it at CompUSA."

strText3 = "Unlike today where you can't go to a book store without being" & vbNewLine & _
"inundated by web development books, there were no books on how" & vbNewLine & _
"to use Visual Interdev 1.0 at that time. Fortunately, the" & vbNewLine & _
"documentation was very good, and soon I learned how to tie a" & vbNewLine & _
"database into with a web site. After programming it in on my" & vbNewLine & _
"home PC in my spare time for about 3 months, I posted it to" & vbNewLine & _
"the Internet and started off the code database with about 2,000" & vbNewLine & _
"lines of my own code. Right away I was amazed and excited by the" & vbNewLine & _
"fact that 50 or so people would come consistently to the site" & vbNewLine & _
"every day.  As word of mouth spread the news about Planet Source" & vbNewLine & _
"Code, it began to grow bigger and bigger. Today, Planet Source" & vbNewLine & _
"Code has over five million lines of source code and averages a" & vbNewLine & _
"hit every second!"

strExampleText = strText1 & vbNewLine & vbNewLine & strText2 & vbNewLine & vbNewLine & strText3

ctlScrollingLabel1.Caption = strExampleText

frmTest.Width = ctlScrollingLabel1.Left + ctlScrollingLabel1.Width + 640
frmTest.Height = ctlScrollingLabel1.Top + ctlScrollingLabel1.Height + 640
End Sub
