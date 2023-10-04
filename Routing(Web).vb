Routing(Web)
If IOM.Info.IsDebug = False Then On Error Goto ErrorHandler

If IOM.Info.IsDebug then IOM.Info.MaxOpcodesExecuted = 15000

'********************************************************************************
'HOLDS SHELL VERSION
SHELL_VERSION = 15 'Increment as changes are made in the Shell.
'********************************************************************************

If RestartTime.Info.OffPathResponse = NULL then
	RestartCount = 0
	RestartTime = "Restart Time stamps: "
Else
	RestartCount = RestartCount.Info.OffPathResponse + 1
	If len(RestartTime.Info.OffPathResponse) < 1000 Then
		RestartTime = MakeString(RestartTime.Info.OffPathResponse , MakeString(NOW()," | "))
	End If
End If


'=======================CAI WEB SHELL VERSION 3.9.7==============================
'
'BEFORE YOU BEGIN SCRIPTING
'1) ENSURE COUNTRIES ARE INCLUDED (Use setcountry option at line 57)
'2) ENSURE LANGUAGES ARE INCLUDED (Use setlang option at line 63)
'3) ENSURE SAMPLE PROVIDERS ARE INCLUDED (Use setprovider option at line 76)
'==============================================================================
'==============================================================================
IOM.LayoutTemplate = "SurveyTemplate_DA.htm"


'-----------------------------------------------------------------------------
'Update Timings
If InterviewLength.Info.OffPathResponse <> Null then
	'Check if Time Duration is less than the OffPathResponse
	if Second(IOM.Info.ElapsedTime) < InterviewLength.Info.OffPathResponse then
		InterviewLength.Response = InterviewLength.Info.OffPathResponse
	end if
end if
If CurrentQuestionTime.Info.OffPathResponse <> Null then
	CurrentQuestionTime = CurrentQuestionTime.Info.OffPathResponse 
else
	CurrentQuestionTime.Response = Now()
end if
'-----------------------------------------------------------------------------

'INITIALISATION (DO NOT TOUCH)
StartInt(IOM)
'InitSamples(IOM)
ConfigButton(IOM)
SetHCSample(IOM)
InitMarker(IOM)
ConfigDisplay(IOM, "10")

'Function to Capture IP Address (Make sure this is placed after SetHCSample(IOM))

If SampleFields.gop.Info.OffPathResponse <> NULL Then
	SampleFields.gop.Response = SampleFields.gop.Info.OffPathResponse
End If
If Trim(IOM.SampleRecord.Item["gop"]) <> "" And IsEmpty(SampleFields.gop.Response) Then
	SampleFields.gop.Response = IOM.SampleRecord.Item["gop"]
End If
If Not(IsEmpty(SampleFields.gop.Response)) Then
	If Trim(IOM.SampleRecord.Item["gop"]) <> "" Then SampleFields.gop_last.Response = IOM.SampleRecord.Item["gop"]
End If

If IPADDRESSCAP.Info.OffPathResponse <> NULL Then
	IPADDRESSCAP.Response = IPADDRESSCAP.Info.OffPathResponse
End If

IPAddressCap(IOM)
'===========================================================================

If (IOM.Info.IsDebug = True) or (IOM.Info.IsTest = True) Then
	smode.Response = {TEST}
	If rstatus.AnswerCount() = 0 Then rstatus.Response = {TESTINCOMPLETE}
	dmUSER.Response = IOM.Info.User1
Else
	smode.Response = {LIVE}
	If rstatus.AnswerCount() = 0 Then rstatus.Response = {INCOMPLETE}
	If QSCREEN.AnswerCount() = 0 Then QSCREEN.Response = {SQUIT}	
End If

'===========================================================================

'GET SAMPLE FIELDS (CHANGE ACCORDINGLY)
If Trim(IOM.SampleRecord.Item["Id"]) <> "" Then SampleFields.ID.Response = IOM.SampleRecord.Item["Id"]
If Trim(IOM.SampleRecord.Item["SamID"]) <> "" Then SampleFields.SamID.Response = IOM.SampleRecord.Item["SamID"]
If Trim(IOM.SampleRecord.Item["country"]) <> "" Then SampleFields.CNT.Response = IOM.SampleRecord.Item["country"]
If Trim(IOM.SampleRecord.Item["IDTYPE"]) <> "" Then SampleFields.IDTYPE.Response = IOM.SampleRecord.Item["IDTYPE"]

'Create column “gop” in sample file with empty values for IP Address capturing.

'===========================================================================

iom.Title.Text = "<label style=""visibility:hidden"">LOGO</label>"

If IOM.Info.IsDebug or UCase(Trim(dmUSER.Response)) = "PROG" Then
	IOM.Navigations.Add(NavigationTypes.nvGoto)
	IOM.Navigations.Add(NavigationTypes.nvLast)
	IOM.Navigations.Add(NavigationTypes.nvStop)
End If


'1) SPECIFY COUNTRIES FOR THIS SURVEY (ADD COUNTRIES SEPARATED BY ":"). NOTE THAT COUNTRY CODES MUST CORRESPOND TO THE CODES IN country QUESTION
setcountry("_124", SampleFields.CNT.Response, IOM)

'DO NOT REMOVE
'With the current requirement this below line is not applicable any more.
'SampleFields.CNT.Response = country.Categories[country].Label 


'2) SET LANGUAGE FILTER (ADD COUNTRIES SEPARATED BY ":" AND THEN LANGUAGEAGES SEPARATED BY ":"). NOTE THAT COUNTRY CODES MUST CORRESPOND TO THE CODES IN country QUESTION
setlang(IOM)	


If langq.AnswerCount() = 0 Then langq.Ask() 'IF MORE THAN ONE LANGUAGE IN A COUNTRY, SET LANGUAGE ACCORDINGLY BEFORE SETTING IT BELOW

'SET OR ASK FOR SPECIFIC LANGUAGE IF MORE THAN ONE LANGUAGE IS TO BE SHOWN FOR A COUNTRY
IOM.Language = langq.Categories[langq].Name
'===========================================================================

'DO NOT TOUCH===============================================================
If smode.Response = {TEST} And UCase(Trim(dmUSER.Response)) = "CL" Then SampleFields.SamID.Response = samprovider.Categories[1].Name
'===========================================================================


'3) SET SAMPLE PROVIDER DETAILS
setprovider("", SampleFields.SamID.Response, IOM)


'samprovider.Response = CCategorical(SampleFields.SamID.Response)
If samprovider.AnswerCount() = 0 And smode.Response = {TEST} Then samprovider.Ask()


HeaderLogo.Response = samprovider.Response
HeaderText.Response = samprovider.Response
EmailText.Response = samprovider.Response

iom.Title.Text = PageTextAndLogo.Label

scrredir.Response = samprovider.Response
qfredir.Response = samprovider.Response
compredir.Response = samprovider.Response  

' IF NO ID PRESENT, THEN ASSIGN INTERVIEWER ID TO ID SO THAT REDIRECTS CAN BE TESTED WITH SOME IDS. DO NOT TOUCH
If Trim(IOM.SampleRecord.Item["Id"]) = "" And smode.Response = {TEST} Then IOM.SampleRecord.Item["Id"] = IOM.Info.InterviewerID


'----------------------------------------------------------------------------------------------------------------------------
'Added by Ugam.
'This line make all mutually exclusive responses unbold.
IOM.DefaultStyles.Categories[CategoryStyleTypes.csExclusive].label.Font.Effects = 0 
' IOM.Banners.AddNew("JSFile1","<script language=""javascript"" src=""https://images3.ipsosinteractive.com/images/UK/HCGOO/FIDP/FieldworkV4.js"" type=""text/javascript""></script>")
IOM.Banners.AddNew("Qname"," ")

DISP_SR.Label.Inserts[0] = IOM.Info.Serial
ShowTest(IOM,DISP_SR,"")
CountVar(IOM)
'End Ugam

'-------------------------------------------------------------------------------------------
INS_SValue = "XXXXX" 'Replace XXXXX with Nebu Project "S" Value.
INS_Redirect1 = {_1} ''Used for redirect text insertion
INS_Redirect2 = {_1} ''Used for redirect text insertion
'-------------------------------------------------------------------------------------------

'4) TEST REDIRECT LINKS. DO NOT TOUCH. SHOW IF TEST LINK IS FOR SAMPLE PROVIDER

If smode.Response = {TEST} and UCase(Trim(dmUSER.Response)) = "SP" Then
	redirtest.Response = {SCONT}
	redirtest.Ask() 
	
	Select Case redirtest.Response 
		Case {SCOMP}
			GoTo COMPLETE
		Case {STERM}
			GoTo TERMINATE
		Case {OVERQ}
			GoTo OVERQUOTA	
	End Select
End If

''''''''''''''''''''IOM.GlobalQuestions block (work in mrStudio 6.0.1)'''''''''''''''''''''''''''

IF IOM.Info.User1 = "LANG" THEN		
		IOM.Navigations.Remove(NavigationTypes.nvNext)
		IOM.Navigations.Remove(NavigationTypes.nvPrev)
		
		Dim Questions,SubQuestion, Ittr1, attributes, MyAttList
		FOR EACH Questions in IOM.Questions.selectrange("StartQuestions..EndQuestions")
			If FIND(UCASE(Questions.QuestionFullName),"PROG",,,) = -1 AND FIND(UCASE(Questions.QuestionFullName),"MRK",,,) = -1 then
				IF Questions.QuestionType = QuestionTypes.qtLoopCategorical then 'and Questions.QuestionFullName <> "KidAge_GRID" Then
					If Questions[0].Count > 1 Then
						MyAttList = null
						For each attributes in Questions.Categories
							MyAttList = MyAttList + attributes.Label + "<br/>"
						Next
						
						For Ittr1 = 0 to (Questions[0].Count) - 1															
							Questions[0].Item[Ittr1].Banners.AddNew(Questions[0].Item[Ittr1].QuestionName,"<hr/><font color='blue'>" + Questions[0].Item[Ittr1].QuestionFullName + "</font><p/>" + MyAttList + "<br/>")
							IOM.GlobalQuestions.Add(Questions[0].Item[Ittr1])
						Next					
					Else
						IOM.Questions[Questions.QuestionFullName].Banners.AddNew(Questions.QuestionFullName,"<hr/><font color='blue'>" + Questions.QuestionFullName + "</font>")								
						IOM.GlobalQuestions.Add(Questions)
					End If
				ElseIf Questions.QuestionType = QuestionTypes.qtCompound Then
					MyAttList = null
					For each attributes in Questions.Categories
						MyAttList = MyAttList + attributes.Label + "<br/>"
					Next
					
					For Ittr1 = 0 to (Questions.Count) - 1															
						Questions.Item[Ittr1].Banners.AddNew(Questions.Item[Ittr1].QuestionName,"<hr/><font color='blue'>" + Questions.Item[Ittr1].QuestionFullName + "</font><p/>" + MyAttList + "<br/>")
						IOM.GlobalQuestions.Add(Questions.Item[Ittr1])
					Next
				ELSE
					IOM.Questions[Questions.QuestionFullName].Banners.AddNew(Questions.QuestionFullName,"<hr/><font color='blue'>" + Questions.QuestionFullName + "</font>")								
					IOM.GlobalQuestions.Add(Questions)
				End If
			End If
		NEXT
		
		Concatenated_Page.Ask()
		EXIT
	END IF

''''''''''''''''''''END OF IOM.GlobalQuestions block'''''''''''''''''''''''''''

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'    SSSSSS      TTTTTTTTTT        AA        RRRRRRRRRR      TTTTTTTTTT
'  SS      SS        TT            AA        RR        RR        TT    
'  SS                TT          AA  AA      RR        RR        TT    
'  SS                TT          AA  AA      RR        RR        TT    
'    SSSSSS          TT        AA      AA    RRRRRRRRRR          TT    
'          SS        TT        AA      AA    RR        RR        TT    
'          SS        TT        AAAAAAAAAA    RR        RR        TT    
'  SS      SS        TT      AA          AA  RR        RR        TT    
'    SSSSSS          TT      AA          AA  RR        RR        TT  
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

VERSION = 1

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%	

If SampleFields.xclose.Response = 1 then goto OVERQUOTA




'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%	
'''''Test data code''''''''''''''''''''''''''''''''''''''''''
IF IOM.Info.Renderer = "xmlplayer" then
	Prog_Data = {Test}
else
	Prog_Data = {Live}
end if
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	

IOM.Banners.AddNew("JSFile2","<script type=""text/javascript"" src=""https://images3.ipsosinteractive.com/images/UK/HCGOO/FIDP/jquery.js""></script>")

'----------------------------------------------------------------------------------------------------------------------------
'General declaration for the quota related varaible
'Dim quota_pend_result
'----------------------------------------------------------------------------------------------------------------------------

''''''''''''''GDPR and privacy policies links variables'''''''''''''''''''''''''''''

SurveyName = "Le Page TITEFOAM" '--> Here we need to pass the survey/job name to be shown in the privacy policy link
JN = "{LePageTITEFOAM}"                        '--> Here we need to pass the job number to be shown in the privacy policy link
'Added 2022 privacy policy
ComplianceEmail = "ADD PROJECT SPECIFIC EMAIL" 'Added in 2022: please insert project Specific Email ID to show in Privacy policy.
If Trim(IOM.SampleRecord.Item["country"]) = "" Then SampleFields.CNT.Response = Mid(country.Categories[country].Name,1) 'If country going blank in Sample then punch it from Country question.

LANG_INS = langq.Categories[langq].Name    '--> This variable is used to pass back the survey language to show the privacy policy link in same language

'-->Below variable is used to create a unique ID and pass back the same to the privacy policy link
Merged_ID = IOM.ProjectName + "_" + CText(IOM.Info.Serial) + "_" + Samplefields.ID

If IOM.Info.IsDebug Or IOM.Info.IsTest Then
	QDUMYY_CUST = {TEST}
	QDUMYY_FI = {TEST}
Else
	QDUMYY_CUST = {LIVE}
	QDUMYY_FI = {LIVE}
End If

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''' ReCAPTCHA Setup Start - PLEASE UPDATE IF APPLICABLE
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

MRK_ReCAPTCHA_Flag = {no} 
' MRK_ReCAPTCHA_Flag = {yes} ' Please uncomment this line and comment the previous line for turning ReCAPTCHA on

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Please don't make any changes in the below code
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'Google ReCAPTCHA setup for validating participants and avoiding bots
If Containsany(MRK_ReCAPTCHA_Flag,{yes}) Then

	If TRIM(UCASE(SampleFields.IDTYPE.Response)) = "LIVE" Then
		If Containsany(MRK_RECAPTCHA_Lang.Categories,langq) then
			MRK_RECAPTCHA_Lang = langq
		Else
			MRK_RECAPTCHA_Lang = {ENG}
		End if
		MRK_ReCAPTCHA_Status.Label.Inserts["INS_RE_LANG"] = MRK_RECAPTCHA_Lang.Categories[MRK_RECAPTCHA_Lang].Label
		MRK_ReCAPTCHA_Status.Banners.AddNew("ReCAPTCHA_Banner1","<script type=""text/javascript"" src=""https://images3.ipsosinteractive.com/images/UK/HCGOO/Jayesh/FI/Template/v2.0/scripts/receiver.js"" ></script>")
		MRK_ReCAPTCHA_Status.Ask()
	End If

End If

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''' ReCAPTCHA Setup End
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'''Wave marker - PLEASE UPDATE IF APPLICABLE
QWave={Wave01}

'Added Country marker(DO NOT CHANGE LABEL IN OVERLAYS)
MRK_QCountry = CCategorical(country.Categories[country].Label)
ShowTest(IOM,MRK_QCountry,"")

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

If IOM.Info.IsDebug then
	'Goto	testC1	
End if

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'SCREENER QUESTIONS
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Start Time capturing for Screener
StartTimeCapture(QStartTime_Screener,IOM)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'Set up browser info collection
Browser[..].MustAnswer = False
IOM.Banners.AddNew("BrowserInfo1","<div style='display:none;'>&#160;<mrData Index='2'></mrData></div>")
IOM.Banners.AddNew("BrowserInfo2","<div style='display:none;'><script language=""javascript"">  jsver=""Generic"";  </script> <script language=""javascript1.1"">  jsver=""1.1"";  </script> <script language=""javascript1.2"">   jsver=""1.2"";  </script> <script language=""javascript1.3"">  jsver=""1.3"";  </script> <script language=""javascript1.4"">  jsver=""1.4"";  </script> <script language=""javascript1.5"">   jsver=""1.5"";  </script> <script language=""javascript1.6"">   jsver=""1.6"";  </script> <script language=""javascript1.7"">   jsver=""1.7"";  </script> <script language=""javascript1.8"">   jsver=""1.8"";  </script> <script language=""javascript1.9"">   jsver=""1.9"";  </script> <script language=""javascript2.0"">   jsver=""2.0"";  </script> <script position=""body"" type=""text/javascript"" src=""https://images3.ipsosinteractive.com/images/UK/HCGOO/FIDP/browserinfo.js""></script></div>")


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'!
Removed in shell Version 9
ShowQuestionNo(IOM,"")
QCOVID19.Ask()
!'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

ShowQuestionNo(IOM,"")
FirstScreen.Ask() ' Please ensure the first question shown in the survey is added to the page "FirstScreen"

IOM.Banners.Remove("BrowserInfo1")
IOM.Banners.Remove("BrowserInfo2")

' ShowTest(IOM,Browser,"")

getDeviceType(Browser,DeviceType)

' ShowTest(IOM,DeviceType,"")

ShowQuestionNo(IOM,"SVideo")
SVideo.Ask()
If Containsany(SVideo,{See,Hear}) then
			MRK_QSCREEN = MRK_QSCREEN + {SVideo}
			ShowTest(IOM,MRK_QSCREEN,"")
			If ContainsAny(QSCREEN,{SQUIT}) then QSCREEN.Response = {SVideo}
			Goto TERMINATE
 End If
 
ShowQuestionNo(IOM,"GRID_S1") 
If GRID_S1[0].S1_2.Response.Value.ContainsAny({_2023,_2022,_2021,_2020,_2019,_2018,_2017,_2016,_2015,_2014,_2013,_2012,_2011,_2010,_2009,_2008,_2007,_2006,_2005,_2004,_2003,_2002,_2001,_2000,_1999}) Then Dummy_Sage.Response={_1}
If GRID_S1[0].S1_2.Response.Value.ContainsAny({_1998,_1997,_1996,_1995,_1994,_1993}) Then Dummy_Sage.Response={_2}
If GRID_S1[0].S1_2.Response.Value.ContainsAny({_1992,_1991,_1990,_1989,_1988}) Then Dummy_Sage.Response={_3}
If GRID_S1[0].S1_2.Response.Value.ContainsAny({_1987,_1986,_1985,_1984,_1983}) Then Dummy_Sage.Response={_4}
If GRID_S1[0].S1_2.Response.Value.ContainsAny({_1982,_1981,_1980,_1979,_1978}) Then Dummy_Sage.Response={_5}
If GRID_S1[0].S1_2.Response.Value.ContainsAny({_1977,_1976,_1975,_1974,_1973}) Then Dummy_Sage.Response={_6}
If GRID_S1[0].S1_2.Response.Value.ContainsAny({_1972,_1971,_1970,_1969,_1968}) Then Dummy_Sage.Response={_7}
If GRID_S1[0].S1_2.Response.Value.ContainsAny({_1967,_1966,_1965,_1964,_1963,_1962,_1961,_1960,_1959,_1958,_1957,_1956,_1955,_1954,_1953,_1952,_1951,_1950}) Then Dummy_Sage.Response={_8}


ShowTest(IOM,Dummy_Sage,"Dummy_Sage")
If Containsany(Dummy_Sage,{_1,_8}) then
			MRK_QSCREEN = MRK_QSCREEN + {Dummy_Sage}
			ShowTest(IOM,MRK_QSCREEN,"")
			If ContainsAny(QSCREEN,{SQUIT}) then QSCREEN.Response = {Dummy_Sage}
			Goto TERMINATE
End If


ShowQuestionNo(IOM,"S2")
S2.Ask()
If Containsany(S2,{Female,Oth,PNTS}) then
			MRK_QSCREEN = MRK_QSCREEN + {S2}
			ShowTest(IOM,MRK_QSCREEN,"")
			If ContainsAny(QSCREEN,{SQUIT}) then QSCREEN.Response = {S2}
			Goto TERMINATE
End If

ShowQuestionNo(IOM,"S3")
S3.Ask()
If Containsany(S3,{somewhereelse}) then
			MRK_QSCREEN = MRK_QSCREEN + {S3}
			ShowTest(IOM,MRK_QSCREEN,"")
			If ContainsAny(QSCREEN,{SQUIT}) then QSCREEN.Response = {S3}
			Goto TERMINATE
 End If

ShowQuestionNo(IOM,"S4")
S4.Ask()
If Containsany(S4,{advertising,research}) then
			MRK_QSCREEN = MRK_QSCREEN + {S4}
			ShowTest(IOM,MRK_QSCREEN,"")
			If ContainsAny(QSCREEN,{SQUIT}) then QSCREEN.Response = {S4}
			Goto TERMINATE
 End If

ShowQuestionNo(IOM,"S5")
S5.Ask()

ShowQuestionNo(IOM,"S6")
S6.Ask()

ShowQuestionNo(IOM,"Grid_S7")
Grid_S7.Ask()

if Containsany(Grid_S7[{Improvements}].S7_inn,{Pastm,Past6,Pasty,Past2,Past5,Never}) then
	ShowQuestionNo(IOM,"S8")
	S8.Ask()
End if

If Containsany(S8,{Yourself})then Quota_Cell1.response = {Diy}
ShowTest(IOM,Quota_Cell1,"Quota_Cell1")

If Containsany(Quota_Cell1,{Diy})then
	ShowQuestionNo(IOM,"S11")
	S11.Ask()
End if

If Containsany(S6,{SelfEmployed,EmployedFull,EmployedPart,Retired})then
	ShowQuestionNo(IOM,"S9")
	S9.Ask()
End if

If Containsany(S9,{Contractor,Skilled})then Quota_Cell2.response = {Pro}
ShowTest(IOM,Quota_Cell2,"Quota_Cell2")

If not(Containsany(S9,{Contractor,Skilled}) and Containsany(S8,{Yourself}))then Fortermination.response = {Termt}
ShowTest(IOM,Fortermination,"Fortermination")

If Containsany(Quota_Cell2,{Pro})then
	ShowQuestionNo(IOM,"S10")
	S10.Ask()
End if

ShowQuestionNo(IOM,"infoSend")
infoSend.Ask()

If AnswerCount(MRK_QSCREEN) > 0 Then
	ShowTest(IOM,MRK_QSCREEN,"")
	GoTo TERMINATE
End If

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'End Time capturing for Screener
EndTimeCapture(QStartTime_Screener,QEndTime_Screener,QTotalTime_Screener,IOM)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'QUOTA CHECK logic

'QUOTA1 = {QUOTA1_CELL1}

If TRIM(UCASE(SampleFields.IDTYPE.Response)) = "TEST" Then
	''Ignore Quota pending for Ids which are TEST
Else
	if Not(PendQuota(IOM, "Quota_Cell1", "Quota_Cell1") = True) then goto OVERQUOTA
	if Not(PendQuota(IOM, "Quota_Cell2", "Quota_Cell2") = True) then goto OVERQUOTA
End If

'Parameter 1 = IOM
'Parameter 2 = Quota name in MQD
'Parameter 3 = Variables used for Quota (For 2D quota, pass 2 variables seperated by comma)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'REMOVE COMPLETE FROM QUOTA WHILE MOVING LIVE TO TEST
Dim strQuota, cell_name, QuotaVar
PROG_UserMove.Response = IOM.Info.User1

If IOM.Info.IsDebug then PROG_UserMove.Response = "move"
If PROG_UserMove.Response = "move" Then
	
	'Add IDs removal Code here.
End If
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

QUITSTAT(IOM)

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'MAIN SECTION
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Start Time capturing for Main
	StartTimeCapture(QStartTime_Main,IOM)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
testC1:
If not Containsany(S3,{Quebec}) then C1.Categories = C1.Categories - {Adfoam}
If Containsany(S3,{Quebec}) then C1.Categories = C1.Categories - {DAPtex}
ShowQuestionNo(IOM,"C1")
C1.Ask()
If Containsany(C1,{NOTA}) then
			MRK_QSCREEN = MRK_QSCREEN + {C1}
			ShowTest(IOM,MRK_QSCREEN,"")
			If ContainsAny(QSCREEN,{SQUIT}) then QSCREEN.Response = {C1}
			Goto TERMINATE
 End If

Grid_C2.Categories=C1.Response
ShowQuestionNo(IOM,"Grid_C2")
Grid_C2.Ask()

Grid_C3.Categories=C1.Response
ShowQuestionNo(IOM,"Grid_C3")
Grid_C3.Ask()

Grid_C4.Categories=C1.Response
ShowQuestionNo(IOM,"Grid_C4")
Grid_C4.Ask()

A1.Categories=C1.Response
ShowQuestionNo(IOM,"A1")
A1.Ask()

If not Containsany(A1,{NoneAbove}) then
	Grid_A2[..].A2_inn.Categories=A1.Response
	
	ShowQuestionNo(IOM,"Grid_A2")
	Grid_A2.Ask()
End if


ShowQuestionNo(IOM,"infoD")
infoD.Ask()

ShowQuestionNo(IOM,"AD1")
AD1.Ask()

ShowQuestionNo(IOM,"AD2")
AD2.Ask()

ShowQuestionNo(IOM,"AD3")
AD3.Ask()

ShowQuestionNo(IOM,"AD4")
AD4.Ask()

ShowQuestionNo(IOM,"QRecontact")
QRecontact.Ask()

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'End Time capturing for Main
EndTimeCapture(QStartTime_Main,QEndTime_Main,QTotalTime_Main,IOM)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

If ContainsAny(country,{_826,_840,_36,_250,_276,_380,_724}) Then
    ShowQuestionNo(IOM,"Q101")
    Q101_1.MustAnswer = False
    PageQ101.Ask()
End If

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'SURVEY COMPLETED

''''''Straight Liner Checks
'''E.g. straightline(IOM,GRID_Q2,"Q2",{Q2})

With AnswerSummary.Label.Inserts["INS_SUMMARY"]
  .Text = "<table border=""1"" width=""100%"">" + .Text + "</table>"
End With

ShowTest(IOM,AnswerSummary,"Answer Summary")


MRK_Straightline.Response = NULL

''''Write the Checks here

ShowTest(IOM,MRK_Straightline,"")

 GoTo COMPLETE
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'  FFFFFFFFFF  II    NN        NN    II      SSSSSS      HH        HH
'  FF          II    NNNN      NN    II    SS      SS    HH        HH
'  FF          II    NNNN      NN    II    SS            HH        HH
'  FF          II    NN  NN    NN    II    SS            HH        HH
'  FFFFFFFF    II    NN  NN    NN    II      SSSSSS      HHHHHHHHHHHH
'  FF          II    NN    NN  NN    II            SS    HH        HH
'  FF          II    NN      NNNN    II            SS    HH        HH
'  FF          II    NN      NNNN    II    SS      SS    HH        HH
'  FF          II    NN        NN    II      SSSSSS      HH        HH
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        	TERMINATED INTERVIEW							'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

COMPLETE:
	
	ShowQuestionNo(IOM,"")
	If MRK_Straightline.Categories.Count > 0 Then
		mark_StrtLine(IOM)
	End If	
	InterviewLengthMain = InterviewLength - InterviewLengthScreener
	
		
	'IOM.OffPathDataMode = OffPathDataModes.dmKeep
	
	SetInterviewVars(IOM)
	EndingInterviewInfo(IOM)
	IOM.Info.EstimatedProgress = 100
	EstimatedProgress(IOM)
	QSCREEN.Response = {SCOMP}
	COMP.Response = {Complete}
	
	If IOM.Questions["smode"].Response = {TEST} Then
		rstatus.Response = {TESTCOMPLETE}
	Elseif IOM.Questions["smode"].Response = {LIVE} Then
		rstatus.Response = {COMPLETE}
	End If
	
	If INTV.rPagesAsked > 0 then
		If Cdouble(InterviewLength/INTV.rPagesAsked) < Cdouble(8) Then
			Mrk_Racer = {yes}
		Else
			Mrk_Racer = {No}
		End If
	End if
	
'	ShowTest(IOM,Mrk_Racer,"Mrk_Racer")

	'===
	If InterviewLength_mins.Info.OffPathResponse = Null then
		InterviewLength_mins = CDouble((CDouble(InterviewLength) / 60))
		InterviewLength_mins = Round(InterviewLength_mins,2)
	Else
		InterviewLength_mins = InterviewLength_mins.Info.OffPathResponse
	End If
	'===
	
	if answercount(langq) = 0 then langq.Response = {ENG}
	
    Add2InterviewProperties(IOM, "RESULT", "COMPLETE")

	hCustomMsgs.Response = {MSG3}	
    'debug.MsgBox(hCustomMsgs.Response.Label + compredir.Response.Label)
    If UCase(Trim(dmUSER.Response)) <> "CL" Then
		IOM.Texts.EndOfInterview = hCustomMsgs.Response.Label + compredir.Response.Label
	Else
		IOM.Texts.EndOfInterview = hCustomMsgs.Response.Label
	End If
	
	If UCASE(TRIM(Samplefields.IDTYPE)) = "TEST" then IOM.Info.IsTest = TRUE
	
	Exit

TERMINATE:

	ShowQuestionNo(IOM,"")
	
	SetInterviewVars(IOM)
	EndingInterviewInfo(IOM)
	IOM.Info.EstimatedProgress = 100
	EstimatedProgress(IOM)
    COMP.Response = {Terminate}
	
	'IOM.OffPathDataMode = OffPathDataModes.dmKeep
	
	If IOM.Questions["smode"].Response = {TEST} Then
		rstatus.Response = {TESTSCREENOUT}
	Elseif IOM.Questions["smode"].Response = {LIVE} Then
		rstatus.Response = {SCREENOUT}
	End If
	
	If IOM.Questions["smode"].Response = {TEST} Then
   		IOM.Questions["QSCREEN"].show()
	End If

	Add2InterviewProperties(IOM, "RESULT", "TERM")

	hCustomMsgs.Response = {MSG1}
	
    If UCase(Trim(dmUSER.Response)) <> "CL" Then
		'If smode.Response = {TEST} Then Debug.MsgBox(hCustomMsgs.Response.Label + scrredir.Response.Label)		
		IOM.Texts.InterviewStopped = hCustomMsgs.Response.Label + scrredir.Response.Label
	Else
		IOM.Texts.InterviewStopped = hCustomMsgs.Response.Label
	End If	
	
	If UCASE(TRIM(Samplefields.IDTYPE)) = "TEST" then IOM.Info.IsTest = TRUE
	
	IOM.Terminate(Signals.sigStopped)

OVERQUOTA:

	ShowQuestionNo(IOM,"")
	TNUM.Response = {TNUM21}
	If SampleFields.xclose.Response =  "1" then TNUM.Response = {TNUM61}
	IOM.OffPathDataMode = OffPathDataModes.dmKeep
	SetInterviewVars(IOM)
	EndingInterviewInfo(IOM)
	IOM.Info.EstimatedProgress = 100
	EstimatedProgress(IOM)
	QSCREEN.Response = {OVERQ}
    COMP.Response = {OverQuota}
	
	If IOM.Questions["smode"].Response = {TEST} Then
		rstatus.Response = {TESTOVERQUOTA}
	Elseif IOM.Questions["smode"].Response = {LIVE} Then
		rstatus.Response = {OVERQUOTA}
	End If
	
    Add2InterviewProperties(IOM, "RESULT", "OQ")

	hCustomMsgs.Response = {MSG2}

    If UCase(Trim(dmUSER.Response)) <> "CL" Then
		'If smode.Response = {TEST} Then Debug.MsgBox(hCustomMsgs.Response.Label + scrredir.Response.Label)		
		IOM.Texts.InterviewStopped = hCustomMsgs.Response.Label + qfredir.Response.Label 		
	Else
		IOM.Texts.InterviewStopped = hCustomMsgs.Response.Label
	End If		
	
	If UCASE(TRIM(Samplefields.IDTYPE)) = "TEST" then IOM.Info.IsTest = TRUE
	
	IOM.Terminate(Signals.sigOverQuota)
	
ErrorHandler:
	IOM.Info.EstimatedProgress = IOM.Questions.Count
	IOM.Log(IOM.ProjectName + " ERROR: SERIAL (" + CText(IOM.Info.Serial) + "), ERROR: (" + CText(Err.Number) + ": " + Err.Description + ") - LAST QUESTION: " + IOM.Info.LastAsked, LogLevels.LOGLEVEL_ERROR)
	rstatus.Response = {ERRORINTV}
	QSCREEN.Response = {ERINV}
	InterviewErrorMessage.Response = CText(Err.Number) + " : " + Err.Description+ " : " + ctext(Err.LineNumber)
	if IOM.Info.IsDebug then debug.MsgBox(InterviewErrorMessage.Response)

	If Samplefields.IDTYPE.Response <> "" then
		SendMail(IOM,"healthcare-scripters@ipsos.com","ServerError@ipsos-research.com",IOM.ProjectName + ": Runtime error caught","Respondent ID:" + IOM.Questions["samplefields"].ID + "<P/>Serial No: " + CText(IOM.Info.Serial) + "<p/>Error Details: " + InterviewErrorMessage)
	End If
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	If Samplefields.IDTYPE.Response = "LIVE" then
		IntvErrorPage.Label.Inserts["INS"].Text = ""
	else
		IntvErrorPage.Label.Inserts["INS"].Text = Ctext("Error Details: " + InterviewErrorMessage)
	end if
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    IntvErrorPage.Show()

	IOM.Terminate(Signals.sigError)  	
  	''''''''''''''''''FUNCTIONS'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'		SET INTRO VARIABLES											'
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Sub InitSamples(IOM)
'	IOM.Questions["SampleFields"].Item["ID"].Response = ""
'	IOM.Questions["SampleFields"].Item["SamID"].Response = ""
'	IOM.Questions["SampleFields"].Item["masterid"].Response = ""	
'	IOM.Questions["SampleFields"].Item["CNT"].Response = ""
'	IOM.Questions["SampleFields"].Item["specialty"].Response = ""
'End Sub	
	
Sub SetInterviewVars(IOM)
	'Routine for setting introduction variables
	On Error Resume Next
	Dim Seconds
	IOM.Questions.Item["INTV"].Item["rElapsedTime"].Item["ElapsedTime"].Response = IOM.Info.ElapsedTime
	IOM.Questions.Item["INTV"].Item["rElapsedTime"].Item["Seconds"].Response = Second(IOM.Info.ElapsedTime)
	IOM.Questions.Item["INTV"].Item["rElapsedTime"].Item["ETDate"].Response = IOM.Info.ElapsedTime
	IOM.Questions.Item["INTV"].Item["rInterviewerID"].response = IOM.Info.InterviewerID
	IOM.Questions.Item["INTV"].Item["rIsDebug"].response = IOM.Info.IsDebug
	IOM.Questions.Item["INTV"].Item["rIsRestart"].response = IOM.Info.IsRestart  
	IOM.Questions.Item["INTV"].Item["rLastAsked"].response = IOM.Info.LastAsked 
	IOM.Questions.Item["INTV"].Item["rLastAskedTime"].response = IOM.Info.LastAskedTime 
	IOM.Questions.Item["INTV"].Item["rOrigin"].response = IOM.Info.Origin 
	IOM.Questions.Item["INTV"].Item["rOriginName"].response = IOM.Info.OriginName 
	IOM.Questions.Item["INTV"].Item["rPagesAnswered"].response = IOM.Info.PagesAnswered 
	IOM.Questions.Item["INTV"].Item["rPagesAsked"].response = IOM.Info.PagesAsked 
	IOM.Questions.Item["INTV"].Item["rRandomSeed"].response = IOM.Info.RandomSeed 
	IOM.Questions.Item["INTV"].Item["rReversalSeed"].response = IOM.Info.ReversalSeed 
	IOM.Questions.Item["INTV"].Item["rRotationSeed"].response = IOM.Info.RotationSeed 
	IOM.Questions.Item["INTV"].Item["rServerTime"].response = IOM.Info.ServerTime 
	IOM.Questions.Item["INTV"].Item["rServerTimeZone"].response = IOM.Info.ServerTimeZone 
	IOM.Questions.Item["INTV"].Item["rStartTime"].response = IOM.Info.StartTime 
	IOM.Questions.Item["INTV"].Item["rTimeouts"].response = IOM.Info.Timeouts 
	IOM.Questions.Item["INTV"].Item["rInterviewerTimeZone"].response = IOM.Info.InterviewerTimeZone 
	IOM.Questions.Item["INTV"].Item["rRespondentTimeZone"].response = IOM.Info.RespondentTimeZone 
	IOM.Questions.Item["INTV"].Item["rProjectName"].response = IOM.ProjectName 
	IOM.Questions.Item["INTV"].Item["rSessionToken"].response = IOM.SessionToken 
	IOM.Questions.Item["INTV"].Item["rVersion"].response = IOM.Version 
	IOM.Questions.Item["INTV"].Item["rLanguage"].response = IOM.Language 
	IOM.Questions.Item["INTV"].Item["rLocale"].response = IOM.Locale
End Sub
Sub InitMarker(IOM)
	On Error Resume Next
	Add2InterviewProperties(IOM, "RESULT", "IN_PROGRESS")
	'Current Interviewing Date
  	IOM.Questions.Item["INTVDATE"].Item["Current"].Item["CDate"].Response = now()
	IOM.Questions.Item["INTVDATE"].Item["Current"].Item["CDay"].Response = day(now())
	IOM.Questions.Item["INTVDATE"].Item["Current"].Item["CHour"].Response = hour(now())
	IOM.Questions.Item["INTVDATE"].Item["Current"].Item["CMonth"].Response = month(now())
	IOM.Questions.Item["INTVDATE"].Item["Current"].Item["CYear"].Response = year(now())
  	'Original Interviewing Date
  	if IOM.Questions.Item["INTVDATE"].Item["Initial"].Item["IDate"].Response  = null then
  		IOM.Questions.Item["INTVDATE"].Item["Initial"].Item["IDate"].Response = IOM.Questions.Item["INTVDATE"].Item["Current"].Item["CDate"].Response
		IOM.Questions.Item["INTVDATE"].Item["Initial"].Item["IDay"].Response = IOM.Questions.Item["INTVDATE"].Item["Current"].Item["CDay"].Response
		IOM.Questions.Item["INTVDATE"].Item["Initial"].Item["IHour"].Response = IOM.Questions.Item["INTVDATE"].Item["Current"].Item["CHour"].Response
		IOM.Questions.Item["INTVDATE"].Item["Initial"].Item["IMonth"].Response = IOM.Questions.Item["INTVDATE"].Item["Current"].Item["CMonth"].Response
		IOM.Questions.Item["INTVDATE"].Item["Initial"].Item["IYear"].Response = IOM.Questions.Item["INTVDATE"].Item["Current"].Item["CYear"].Response
  	end if
	IOM.Questions["COMP"].Response = {Quit}
	IOM.Questions["QUITSTAT"] = {QuitScreener}
 	IOM.Questions["StartTime"].Response = TimeNow()
  End Sub
  
Sub Add2InterviewProperties(IOM, Name, Value)
	'On Error Resume Next
	Dim MyProp
	'Check if Property already exists before setting value.  If not present, create the field.
	set MyProp = FindItem(IOM.Properties, Name)
	If Not(MyProp Is Null) then
		MyProp.Value = Value	
	Else
		Set MyProp = IOM.Properties.CreateProperty() 
		MyProp.Name = Name
		MyProp.Value = Value
		IOM.Properties.Add(MyProp)
	end if

End Sub
  
Sub EndingInterviewInfo(IOM)
	On Error Resume Next
	'Ending Interview Date Information
	IOM.Questions.Item["INTVDATE"].Item["Ending"].Item["EDate"].Response = now()
	IOM.Questions.Item["INTVDATE"].Item["Ending"].Item["EDay"].Response = day(now())
	IOM.Questions.Item["INTVDATE"].Item["Ending"].Item["EHour"].Response = hour(now())
	IOM.Questions.Item["INTVDATE"].Item["Ending"].Item["EMonth"].Response = month(now())
	IOM.Questions.Item["INTVDATE"].Item["Ending"].Item["EYear"].Response = year(now())
	IOM.Questions.Item["INTVDATE"].Item["Ending"].Item["ETotalTime"].Response = Second(IOM.Info.ElapsedTime)
End Sub
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'		LOOK OF SURVEY ITEMS											'
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ConfigButton(IOM)
	'SHOW BACK AND QUIT BUTTONS DURING TESTING ONLY 					'
	If Not(IOM.Info.IsTest) and Not (IOM.Info.IsDebug) Then
'	If Not (IOM.Info.IsDebug) Then
			'HIDE BACK BUTTON
		IOM.Navigations["Prev"].Style.Hidden = True
'		IOM.Navigations["Last"].Style.Hidden = True
		
    End if
    If Not (IOM.Info.IsDebug) Then
			'HIDE QUIT BUTTON
		IOM.Navigations["Stop"].Style.Hidden = True
  	End if
  	'Add following line since Workbench MDD overwrites width
  	IOM.Navigations[..].Style.Width="5em"   	
End Sub
Sub ConfigDisplay(IOM, SIZE)
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'		GRID DISPLAY													'
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'Head Column
  	With IOM.DefaultStyles.Grids[GridStyleTypes.gsColHeader]
		.VerticalAlign = VerticalAlignments.vaBottom
		.Align = Alignments.alCenter
		''.Cell.BgColor = "#009493" '"#339966"
		.Cell.BgColor = "#004360" '"#339966"
		.Color="#FFFFFF"
		.cell.Width = SIZE + "%"
		.cell.BorderColor="#000000"
		.Cell.borderStyle = BorderStyles.bsSolid
		.Cell.borderWidth = 1
		.Cell.Padding = 4
  	End With
		'Alternate Cells
'  	IOM.DefaultStyles.Grids[GridStyleTypes.gsAltRow].Cell.BgColor = "#E0E0E0"
'  	IOM.DefaultStyles.Grids[GridStyleTypes.gsAltRowHeader].Cell.BgColor = "#E0E0E0"
		'All Cells
  	With IOM.DefaultStyles.Grids[GridStyleTypes.gsCell]
		.cell.BorderColor="#000000"
		.Cell.borderStyle = BorderStyles.bsSolid
		.Cell.borderWidth = 1
		.Cell.Padding = 4
  	End With
		'Head Row
  	With IOM.DefaultStyles.Grids[GridStyleTypes.gsRowHeader]
		.cell.BorderColor="#000000"
   		.Cell.borderStyle = BorderStyles.bsSolid
   		.Cell.borderWidth = 1
   		.Cell.Padding = 4
  	End With
End Sub  
  	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'		QUOTAS      													'
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Function PendQuota(IOM, QuotaName)
'	On Error Resume Next
'	Dim Result
'	PendQuota = Null 'Leave as Null in case an error is returned
'	'Pend Quota
'	Result = QuotaEngine.QuotaGroups[QuotaName].Pend() 
'	If (IsSet(Result, QuotaResultConstants.qrWasPended)) then
'		PendQuota = True
'	else
'		PendQuota = False
'	end if
'End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'		INTERVIEW EVENTS     													'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub OnInterviewStart(IOM)

   IOM.Questions["AnswerSummary"].Label.Inserts["INS_SUMMARY"].Text = _
      BoldItalic("ANSWER SUMMARY")

End Sub

Sub OnBeforeQuestionAsk(Question, IOM)

	CheckExclusive(Question, IOM)

	IOM.Questions.Item["CurrentQuestionTime"].Response = now()
End Sub	

Sub OnAfterQuestionAsk(Question, IOM)
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'	   FOR PROGRESS BAR													'
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	EstimatedProgress(IOM)
	IOM.Info.EstimatedProgress = IOM.Info.EstimatedProgress + 1
	LogQValue(Question, IOM)  'Log Answers - Use for Troubleshooting only.
   	QuestionTimings(IOM)

   	With IOM.Questions["AnswerSummary"].Label.Inserts["INS_SUMMARY"]
		.Text = .Text + FormatQuestion(Question, IOM)
	End With

	Dim i, ban

	If IOM.Banners.Count > 0 Then
		For i  =  IOM.Banners.Count - 1 To 0 Step - 1
			If Left(IOM.Banners.Item[i].Name,14) = "exclusiveCheck" Then
				IOM.Banners.Remove(IOM.Banners.Item[i].Name)
			End if
		Next
	End if

End Sub	

Sub OnInterviewError(IOM, Error1)
	'Log following error.
	IOM.Log("Interview Serial(" + CText(IOM.Info.Serial) + ") exited with following error: " + CText(Error1), LogLevels.LOGLEVEL_ERROR)
End Sub    
'''''''''''''''''''''''''''''''''''''''''
'Miscellaneous Functions
'''''''''''''''''''''''''''''''''''''''''
'
'
'========================================================================================================
'========================================================================================================
'
'END SCRIPTING HERE. DO NOT CHANGE ANYTHING BELOW
'
'========================================================================================================
'========================================================================================================
Sub StartInt(IOM)
	If IOM.Questions["InterviewLength"].Info.OffPathResponse <> Null then
		'Check if Time Duration is less than the OffPathResponse
		if Second(IOM.Info.ElapsedTime) < IOM.Questions["InterviewLength"].Info.OffPathResponse then
			IOM.Questions["InterviewLength"].Response = IOM.Questions["InterviewLength"].Info.OffPathResponse
		end if	
	end if
	If IOM.Questions["CurrentQuestionTime"].Info.OffPathResponse <> Null then
		IOM.Questions["CurrentQuestionTime"] = IOM.Questions["CurrentQuestionTime"].Info.OffPathResponse 
	else
		IOM.Questions["CurrentQuestionTime"].Response = Now()
	end if
	if IOM.Info.IsDebug then ConfigButton(IOM)  'Remove Debug if you want respondent to see Previous/Stop Buttons
	'For Progress Bar.
	IOM.Info.EstimatedPages = 100
	IOM.Info.EstimatedProgress = 1
	EstimatedProgress(IOM)
End Sub

Sub setcountry(countries, SamCountry, IOM)
	On Error Resume Next
	Dim CNTCAT
	Dim arrCnt
	Dim arrCount
		
	if ascw(left(SamCountry,1)) >=  48 and ascw(left(SamCountry,1)) <=57 then
		For Each CNTCAT in IOM.Questions["country"].Categories
			If replace(trim(CNTCAT.Name),"_","") = UCase(Trim(SamCountry)) Then
				IOM.Questions["country"].Response = CNTCAT
				Exit Sub
			End if
		Next
	Else	
		For CNTCAT = 0 To IOM.Questions["country"].Categories.Count - 1
			If UCase(Trim(IOM.Questions["country"].Categories[CNTCAT].Label)) = UCase(Trim(SamCountry)) Then
				IOM.Questions["country"].Response = IOM.Questions["country"].Categories[CNTCAT]
				Exit Sub
			End If
		Next
	end if

	arrCnt = Split(countries, ":")
	IOM.Questions["country"].Categories.Filter = {}
	For arrCount = 0 to Ubound(arrCnt)
		IOM.Questions["country"].Categories.Filter = IOM.Questions["country"].Categories + CCategorical(Trim(arrCnt[arrCount]))
	Next
	If IOM.Questions["country"].Categories.Count > 1 Then
    	IOM.Questions["country"].Ask()
	Else
		IOM.Questions["country"].Response = CCategorical(IOM.Questions["country"].Categories[0].Name)
	End If	
End Sub

'Updated on 28Dec2017
Sub setprovider(allproviders, provider, IOM)
	If UCase(Trim(provider)) <> "FINEBU" then		
		If UCase(Trim(provider)) = "PILOT" then
			IOM.Questions["samprovider"].Response = {PILOT}
		Else
			'IOM.Questions["samprovider"].Response = CCategorical(UCase(Trim(provider)))
			IOM.Questions["samprovider"].Response = {UnBranded}
		End if
	Else
		IOM.Questions["samprovider"].Response = {FINEBU}
	End If
		
	Exit Sub
End Sub

Sub setlang(IOM)
	'On Error Resume Next
	
	Dim LangCode, langq
	Dim country
	Dim ErrLanguageSetting
	
	Set country = IOM.Questions["country"]
	Set langq = IOM.Questions["langq"]
	
	LangCode = trim(country.Categories[country.response].Properties["Lang"])
	
	if trim(LangCode) <> "" then
		'!
			PLEASE NOTE: If you get an run time error on the below line that means the 
							language code(s) mentioned at country question is not present at
							langq questions.
		!'						
		langq.Categories = CCategorical("{" + ctext(LangCode) + "}") + {}
		if langq.Categories.Count = 1 then	
			langq.Response = IOM.Questions["langq"].Categories
		Elseif langq.Categories.Count = 0 then	
			if IOM.Info.IsDebug then
				Set ErrLanguageSetting = IOM.Questions["ErrLanguageSetting"]
				ErrLanguageSetting.Label.Inserts["INS"] = "Language code (" + ctext(LangCode) + ") was not found in the langq question."
				IOM.Navigations["next"].Style.Hidden = True
				ErrLanguageSetting.Show()
			End if
		end if
	Else
		if IOM.Info.IsDebug then			
			Set ErrLanguageSetting = IOM.Questions["ErrLanguageSetting"]
			ErrLanguageSetting.Label.Inserts["INS"] = "Language Code not found in against the country (" + ctext(country.Response.Label) + ")."
			IOM.Navigations["next"].Style.Hidden = True
			ErrLanguageSetting.Show()
		End if
	End if
End Sub
'=======SOC(QID, ANSWERS, IOM)
'SOC STANDS FOR Screen Out Categorical
'THIS PROCEDURE CHECKS IF THE USER HAS GIVEN A RESPONSE OR RESPONSES
'THAT WILL ALLOW THEM TO SCREEN OUT. 
'----
'SYNTAX: SOC(IOM, QID, ANSWERS)
'QID: PROVIDE QUESTION NAME HERE (EX: S1)
'ANSWERS: PROVIDE ANSWERS FROM THE LIST SEPARATED BY ":" (EX: "NONEA:DKNOW")
'IOM: MUST ALWAYS BE THE LAST PARAMETER (EX: IOM)
Function SOC(QID, ANSWERS, IOM, TNUMID)
	IOM.Questions["QSCREEN"] = CCategorical("SQUIT")

	Dim arrANS
	Dim arrCount
	arrANS = ANSWERS.Split(":") 
	For arrCount = 0 to Ubound(arrANS)
		If QID.ContainsAny(arrANS[arrCount]) Then
'			IOM.Questions["QSCREEN"] = Union(IOM.Questions["QSCREEN"], QID.QuestionName)
			IOM.Questions["QSCREEN"] = CCategorical(QID.QuestionName)
			IOM.Questions["TNUM"] = CCategorical(TNUMID)

			If ContainsAny(IOM.Questions["smode"].Response,"TEST") Then
				IOM.Questions["rstatus"].Response = {TESTSCREENOUT}
			Else
				IOM.Questions["rstatus"].Response = {SCREENOUT}
			End If			
			Exit For
		End If
	Next
	SOC = False
	If IOM.Questions["QSCREEN"].AnswerCount() > 0 And ContainsAny(IOM.Questions["QSCREEN"].Response,"SQUIT") = False Then
		SOC = True
	End If	
End Function
'=======SOV(QID, TestValue, Operation, IOM)
'SOV STANDS FOR Screen Out Value
'THIS PROCEDURE CHECKS THE USER INPUT TestValueUE AGAINST THE GIVEN TestValueUE (TestValue) AND
'THE OPERATION. IF THE USER INPUT TestValueUE DOES NOT MEET THE CRITERIA FOR THE OPERATION
'THE RESPONDENT WILL BE SCREENED OUT
'----
'SYNTAX: SOV(QID, TestValue, Operation, IOM)
'QID: PROVIDE QUESTION NAME HERE (EX: S2). NOT OPTIONAL
'TestValue: PROVIDE A TestValueUE TO CHECK AGAINST (EX: 5). NOT OPTIONAL 
'Operation: SPECIFY WHICH OPERATION TO PERFORM: "=" or "<>" or "<" or ">" or "<=" or ">="
'IOM: MUST ALWAYS BE THE LAST PARAMETER (EX: IOM). NOT OPTIONAL
Function SOV(QID, TestValue, Operation, IOM, TNUMID)
	IOM.Questions["QSCREEN"] = CCategorical("SQUIT")

	SOV = False
	If SysCompare(QID.Response.Value, TestValue, Operation) = False Then
		'IOM.Questions["QSCREEN"] = Union(IOM.Questions["QSCREEN"], QID.QuestionName)
		IOM.Questions["QSCREEN"] = CCategorical(QID.QuestionName)
		IOM.Questions["TNUM"] = CCategorical(TNUMID)
		
		If ContainsAny(IOM.Questions["smode"].Response,"TEST") Then
			IOM.Questions["rstatus"].Response = {TESTSCREENOUT}
		Else
			IOM.Questions["rstatus"].Response = {SCREENOUT}
		End If		
		If IOM.Questions["QSCREEN"].AnswerCount() > 0 And ContainsAny(IOM.Questions["QSCREEN"].Response,"SQUIT") = False Then
			SOV = True
		End If	
	End If	
End Function
'=======SOR(QID, MIN, MAX, IOM)
'SOR STANDS FOR Screen Out Range
'THIS PROCEDURE CHECKS IF THE USER INPUT VALUE IS WITHIN MIN AND MAX VALUES SUPPLIED
'IF THE USER INPUT VALUE DOES NOT QUALIFY, THEY WILL BE SCREENED OUT
'----
'SYNTAX: SOR(QID, MIN, MAX, IOM)
'QID: PROVIDE QUESTION NAME HERE (EX: S2). NOT OPTIONAL
'MIN: SPECIFY MINIMUM VALUE TO CHECK AGAINST (EX: 5). NOT OPTIONAL 
'MAX: SPECIFY MAXIMUM VALUE TO CHECK AGAINST (EX: 25). NOT OPTIONAL 
'IOM: MUST ALWAYS BE THE LAST PARAMETER (EX: IOM). NOT OPTIONAL
Function SOR(QID, MIN, MAX, IOM, TNUMID)
	IOM.Questions["QSCREEN"] = CCategorical("SQUIT")

	SOR = False
	If QID.Response.Value < MIN OR QID.Response.Value > MAX Then
		'IOM.Questions["QSCREEN"] = Union(IOM.Questions["QSCREEN"], QID.QuestionName)
		IOM.Questions["QSCREEN"] = CCategorical(QID.QuestionName)
		IOM.Questions["TNUM"] = CCategorical(TNUMID)
		
		If ContainsAny(IOM.Questions["smode"].Response,"TEST") Then
			IOM.Questions["rstatus"].Response = {TESTSCREENOUT}
		Else
			IOM.Questions["rstatus"].Response = {SCREENOUT}
		End If	
		If IOM.Questions["QSCREEN"].AnswerCount() > 0 And ContainsAny(IOM.Questions["QSCREEN"].Response,"SQUIT") = False Then		
			SOR = True
		End If	
	End If
End Function
'=======SOT(GID, TestValue, Exclude, IOM)
'SOR STANDS FOR Screen Out Total
'THIS PROCEDURE CHECKS IF THE USER INPUT VALUES EQUAL THE VALUE (TestValue) SUPPLIED
'IF THE USER INPUT VALUE DOES NOT QUALIFY, THEY WILL BE SCREENED OUT
'----
'SYNTAX: SOT(QID, TestValue, Exclude, IOM)
'GID: SPECIFY GRID NAME HERE (EX: G2). NOT OPTIONAL
'TestValue: SPECIFY A VALUE TO CHECK AGAINST (EX: 100). NOT OPTIONAL 
'Operation: SPECIFY WHICH OPERATION TO PERFORM: "=" or "<>" or "<" or ">" or "<=" or ">="
'Exclude: SPECIFY PRECODES TO BE EXCLUDED FROM THE TOTAL SEPARATED BY ":" 
'IOM: MUST ALWAYS BE THE LAST PARAMETER (EX: IOM). NOT OPTIONAL
Function SOT(GID, TestValue, Operation, Exclude, IOM, TNUMID)
	IOM.Questions["QSCREEN"] = CCategorical("SQUIT")

	Dim arrANS
	Dim arrCount
	Dim selectedPrecodes
	Dim bFound
	Dim arrNew[]
	Dim nTotal
	Dim nCurrentValue
	nTotal = 0
	nCurrentValue = 0
	
	arrANS = Exclude.Split(":") 
	For Each selectedPrecodes In GID
		bFound = False
		For arrCount = 0 to Ubound(arrANS)
			If LCase(selectedPrecodes.QuestionName) = LCase(arrANS[arrCount]) Then
				bFound = True
			End If
		Next
		If bFound = False Then
			ReDim(arrNew, Ubound(arrNew) + 1, True)
			arrNew[Ubound(arrNew) + 1] = selectedPrecodes.QuestionName
		End If		
	Next
	For arrCount = 0 to Ubound(arrNew)
		nTotal = nTotal + GID.Item[arrCount].Item[0].Response
	Next	
	SOT = False
	If SysCompare(nTotal, TestValue, Operation) = False Then
		'IOM.Questions["QSCREEN"] = Union(IOM.Questions["QSCREEN"], GID.QuestionName)
		IOM.Questions["QSCREEN"] = CCategorical(GID.QuestionName)
		IOM.Questions["TNUM"] = CCategorical(TNUMID)
		
		If ContainsAny(IOM.Questions["smode"].Response,"TEST") Then
			IOM.Questions["rstatus"].Response = {TESTSCREENOUT}
		Else
			IOM.Questions["rstatus"].Response = {SCREENOUT}
		End If
		If IOM.Questions["QSCREEN"].AnswerCount() > 0 And ContainsAny(IOM.Questions["QSCREEN"].Response,"SQUIT") = False Then
			SOT = True			
		End If	
	End If
End Function

'=======CheckRank(Grid, QuestionName, MaxRank, Exclude, ErrorMsg, ShowDefaultErrorMsg, IOM)
'CheckRank IS A VERY USEFUL AND EASY FUNCTION TO VALIDATE RANKS IN A OPEN NUMERIC A GRID
'IF RANKING IS INCORRECT, THEN AN ERROR MESSAGE WILL BE GENERATED (DEFAULT, USER-DEFINED, BOTH, OR NONE)
'----
'SYNTAX: CheckRank(Grid, QuestionName, MaxRank, Exclude, ErrorMsg, ShowDefaultErrorMsg, IOM)
'Grid: SPECIFY GRID NAME HERE (EX: G2). NOT OPTIONAL
'QuestionName: SPECIFY THE COLUMN NAME IN A GRID WITHIN "". IF NO COLUMN IS PROVIDED, THEN EACH COLUMN/ROW WILL BE VALIDATED. OPTIONAL.
'MaxRank: SPECIFY THE MAXIMUM RANK. NOT OPTIONAL.
'Exclude:SPECIFY PRECODES TO BE EXCLUDED FROM THE TOTAL SEPARATED BY ":". IF NOTTHING SPECIFIED, THEN ALL PRECODES WILL BE EVALUATED. OPTIONAL.
'ErrorMsg: SPECIFY A PRECODE TO AN ERROR MESSAGE DEFINED IN THE "lCustomMsgs" LIST. IT MUST BE SUPPLIED WITHIN "". OPTIONAL.
'ShowDefaultErrorMsg: THIS IS A FLAG TO SWITCH ON/OFF SYSTEM ERROR MESSAGES RELATED TO ERROR TYPE. DEFAULT VALUE = TRUE.
'IOM: MUST ALWAYS BE THE LAST PARAMETER (EX: IOM). NOT OPTIONAL
Function CheckRank(Grid, QuestionName, MaxRank, Exclude, ErrorMsg, ShowDefaultErrorMsg, IOM)
On Error Resume Next
	Dim arrANS
	Dim arrCount
	Dim selectedPrecodes
	Dim bFound
	Dim arrNew[]
	Dim nTotal
	Dim ColIndex
	Dim oCol
	Dim TestValue
	Dim nCount
	Dim Operation
	Dim i
	Dim j
	Dim k
	Dim bRankFound
	ColIndex = 0
	nTotal = 0
	TestValue = 0
	Operation = "="
	'IF SOME PRECODES TO BE EXCLUDED, THEN CAPTURE THEM INTO AN ARRAY.
	arrANS = Exclude.Split(":") 
	'CALCULATE THE TOTAL OF RANKS FROM THE GIVEN MAXRANK.
	For nCount = 1 To MaxRank
		TestValue = TestValue + nCount	
	Next
	'IF ONLY ONE QUESTION WITHIN A GRID IS TO BE EVALUATED, THEN CAPTURE THE INDEX OF THAT QUESTION.
	If Trim(QuestionName) <> "" Then
		For Each oCol in Grid[0]
			If LCase(oCol.QuestionName) = LCase(Trim(QuestionName)) Then
				Exit For	
			End If
			ColIndex = ColIndex + 1
		Next
	End If
	'CAPTURE ONLY VALID PRECODES INTO AN ARRAY.
	For Each selectedPrecodes In Grid
		bFound = False
		For arrCount = 0 to Ubound(arrANS)
			If LCase(selectedPrecodes.QuestionName) = LCase(arrANS[arrCount]) Then
				bFound = True
			End If
		Next
		If bFound = False Then
			ReDim(arrNew, Ubound(arrNew) + 1, True)
			arrNew[Ubound(arrNew) + 1] = selectedPrecodes.QuestionName
		End If		
	Next
	If Trim(QuestionName) <> "" Then	'IF ONLY ONE QUESTION IN THE GRID TO BE EVALUATED
		nTotal = 0
		For arrCount = 0 to Ubound(arrNew)
			nTotal = nTotal + Grid.Item[arrCount].Item[ColIndex].Response
		Next
		If SysCompare(nTotal, TestValue, Operation) = False Then
			'CHECK IF RANKING IS SEQUENTIALLY GIVEN. EX: NOT GIVEN 1 AND 5 IN A TWO RANK QUESTION WHERE CORRECT RANKS SHOULD BE 1 AND 2.
			For i = 1 to MaxRank
				bRankFound = False
				For j = 0 to Ubound(arrNew)
					If i = Grid.Item[j].Item[ColIndex].Response Then
						bRankFound = True		
					End If
				Next
				If Not bRankFound Then
					CreateRankErrorBanner(Grid, QuestionName, MaxRank, ErrorMsg, ShowDefaultErrorMsg, IOM)
					CheckRank = True
					Exit Function				
				End If
			Next			
			'CHECK IF DUPLICATE VALUES HAVE BEEN ENTERED.
			For i = 0 to Ubound(arrNew) - 1
				k = i + 1
				For j = k to Ubound(arrNew)
					If Grid.Item[i].Item[ColIndex].Response = Grid.Item[j].Item[ColIndex].Response Then
						CreateRankErrorBanner(Grid, QuestionName, MaxRank, ErrorMsg, ShowDefaultErrorMsg, IOM)
						CheckRank = True
						Exit For
					End If
				Next
			Next				
			CheckRank = False
		Else
		    'TOTAL OF ALL RANKS DOES NOT MATCH THE CORRECT TOTAL.
			CreateRankErrorBanner(Grid, QuestionName, MaxRank, ErrorMsg, ShowDefaultErrorMsg, IOM)
			CheckRank = True
		End If
	Else	'ALL QUESTIONS IN THE GRID TO BE EVALUATED
		ColIndex = 0
		For Each oCol in Grid[0]
'			If LCase(oCol.QuestionName) = LCase(Trim(QuestionName)) Then
'				Exit For	
'			End If
'
			nTotal = 0
			For arrCount = 0 to Ubound(arrNew)
				nTotal = nTotal + Grid.Item[arrCount].Item[ColIndex].Response
			Next	              
			If SysCompare(nTotal, TestValue, Operation) = False Then
			'CHECK IF RANKING IS SEQUENTIALLY GIVEN. EX: NOT GIVEN 1 AND 5 IN A TWO RANK QUESTION WHERE CORRECT RANKS SHOULD BE 1 AND 2.
			For i = 1 to MaxRank
				bRankFound = False
				For j = 0 to Ubound(arrNew)
					If i = Grid.Item[j].Item[ColIndex].Response Then
						bRankFound = True		
					End If
				Next
				If Not bRankFound Then
					CreateRankErrorBanner(Grid, QuestionName, MaxRank, ErrorMsg, ShowDefaultErrorMsg, IOM)
					CheckRank = True
					Exit Function				
				End If
			Next			
			'CHECK IF DUPLICATE VALUES HAVE BEEN ENTERED.
			For i = 0 to Ubound(arrNew) - 1
				k = i + 1
				For j = k to Ubound(arrNew)
					If Grid.Item[i].Item[ColIndex].Response = Grid.Item[j].Item[ColIndex].Response Then
						CreateRankErrorBanner(Grid, QuestionName, MaxRank, ErrorMsg, ShowDefaultErrorMsg, IOM)
						CheckRank = True
						Exit For
					End If
				Next
			Next				
				CheckRank = False
			Else
		    	'TOTAL OF ALL RANKS DOES NOT MATCH THE CORRECT TOTAL.			
				CreateRankErrorBanner(Grid, Grid[0].Item[ColIndex].QuestionName, MaxRank, ErrorMsg, ShowDefaultErrorMsg, IOM)
				CheckRank = True				
				Exit For
			End If		
			ColIndex = ColIndex + 1
		Next	
	End If	
End Function
'=======PreFillGrid(Grid, QuestionName, TestValue, Exclude, IOM)
'PreFillGrid CAN BE USED TO POPULATE A GRID WITH DEFAULT VALUES BEFORE SHOWING IT TO RESPONDENTS.
'----
'SYNTAX: PreFillGrid(Grid, QuestionName, TestValue, Exclude, IOM)
'Grid: SPECIFY GRID NAME HERE (EX: G2). NOT OPTIONAL
'QuestionName: SPECIFY THE COLUMN NAME IN A GRID WITHIN "". IF NO COLUMN IS PROVIDED, THEN EACH COLUMN/ROW WILL BE VALIDATED. OPTIONAL.
'TestValue: SPECIFY THE VALUE TO CHECK AGAINST. IF NONE PROVIDED, 0 WILL BE ASSUMED.
'Exclude:SPECIFY PRECODES TO BE EXCLUDED FROM FILLING SEPARATED BY ":". IF NOTTHING SPECIFIED, THEN ALL PRECODES WILL BE EVALUATED. OPTIONAL.
'IOM: MUST ALWAYS BE THE LAST PARAMETER (EX: IOM). NOT OPTIONAL
Sub PreFillGrid(Grid, QuestionName, TestValue, Exclude, IOM)
On Error Resume Next
	Dim arrANS
	Dim arrCount
	Dim selectedPrecodes
	Dim bFound
	Dim arrNew[]
	Dim ColIndex
	Dim oCol
	ColIndex = 0
	'IF SOME PRECODES TO BE EXCLUDED, THEN CAPTURE THEM INTO AN ARRAY.	
	arrANS = Exclude.Split(":") 
	'IF NOT TESTVALUE SUPPIED, ASSUME 0
	If Trim(TestValue) = "" Then TestValue = "0"
	'IF ONLY A SPECIFIC QUESTION IS TO BE EVALUATED, THEN CAPTURE IT'S INDEX IN THE GRID
	If Trim(QuestionName) <> "" Then
		For Each oCol in Grid[0]
			If LCase(oCol.QuestionName) = LCase(Trim(QuestionName)) Then
				Exit For	
			End If
			ColIndex = ColIndex + 1
		Next
	End If
	'CAPTURE ALL RELEVANT PRECODES INTO AN ARRAY
	For Each selectedPrecodes In Grid
		bFound = False
		For arrCount = 0 to Ubound(arrANS)
			If LCase(selectedPrecodes.QuestionName) = LCase(arrANS[arrCount]) Then
				bFound = True
			End If
		Next
		If bFound = False Then
			ReDim(arrNew, Ubound(arrNew) + 1, True)
			arrNew[Ubound(arrNew) + 1] = selectedPrecodes.QuestionName
		End If		
	Next
	If Trim(QuestionName) <> "" Then 'PRE-FILL THE SELECTED QUESTION IN THE GRID
		For arrCount = 0 to Ubound(arrNew)
			Grid.Item[arrCount].Item[ColIndex].Response.Value = TestValue
		Next	
	Else ' 'PRE-FILL ALL QUESTIONS IN THE GRID
		ColIndex = 0
		For Each oCol in Grid[0]
			If LCase(oCol.QuestionName) = LCase(Trim(QuestionName)) Then
				Exit For	
			End If
			For arrCount = 0 to Ubound(arrNew)
				Grid.Item[arrCount].Item[ColIndex].Response.Value = TestValue
			Next	
			ColIndex = ColIndex + 1
		Next	
	End If	
End Sub
'=======CheckSum(Grid, QuestionName, TestValue, Operation, Exclude, ErrorMsg, ShowDefaultErrorMsg, IOM)
'CheckSum IS A VERY USEFUL AND EASY FUNCTION TO VALIDATE NUMERIC ANSWERS IN A GRID
'THIS FUNCTION CHECKS IF THE USER INPUT VALUE TOTAL MATCHES THE A GIVEN VALUE AGAINST THE SUPPLIED OPERATION.
'IF THE VALUE IS INCORRECT, THEN AN ERROR MESSAGE WILL BE GENERATED
'----
'SYNTAX: SumCheck(Grid, QuestionName, TestValue, Operation, Exclude, ErrorMsg, ShowDefaultErrorMsg, IOM)
'Grid: SPECIFY GRID NAME HERE (EX: G2). NOT OPTIONAL
'QuestionName: SPECIFY THE COLUMN NAME IN A GRID WITHIN "". IF NO COLUMN IS PROVIDED, THEN EACH COLUMN/ROW WILL BE VALIDATED. OPTIONAL.
'TestValue: SPECIFY THE VALUE TO CHECK AGAINST. NOT OPTIONAL.
'Operation: SPECIFY WHICH OPERATION TO PERFORM: "=" or or "<" or ">" or "<=" or ">=".
'Exclude:SPECIFY PRECODES TO BE EXCLUDED FROM THE TOTAL SEPARATED BY ":". IF NOTTHING SPECIFIED, THEN ALL PRECODES WILL BE EVALUATED. OPTIONAL.
'ErrorMsg: SPECIFY A PRECODE TO AN ERROR MESSAGE DEFINED IN THE "lCustomMsgs" LIST. IT MUST BE SUPPLIED WITHIN "". OPTIONAL.
'ShowDefaultErrorMsg: THIS IS A FLAG TO SWITCH ON/OFF SYSTEM ERROR MESSAGES RELATED TO ERROR TYPE. DEFAULT VALUE = TRUE.
'IOM: MUST ALWAYS BE THE LAST PARAMETER (EX: IOM). NOT OPTIONAL
Function CheckSum(Grid, QuestionName, TestValue, Operation, Exclude, ErrorMsg, ShowDefaultErrorMsg, IOM)
On Error Resume Next
	Dim arrANS
	Dim arrCount
	Dim selectedPrecodes
	Dim bFound
	Dim arrNew[]
	Dim nTotal
	Dim ColIndex
	Dim oCol
	Dim ErrMsg
	ColIndex = 0
	nTotal = 0
	arrANS = Exclude.Split(":") 
	'IF ONLY A SPECIFIC QUESTION IS TO BE EVALUATED, THEN CAPTURE IT'S INDEX IN THE GRID	
	If Trim(QuestionName) <> "" Then
		For Each oCol in Grid[0]
			If LCase(oCol.QuestionName) = LCase(Trim(QuestionName)) Then
				Exit For	
			End If
			ColIndex = ColIndex + 1
		Next
	End If
	'CAPTURE RELEVANT PRECODES INTO AN ARRAY
	For Each selectedPrecodes In Grid
		bFound = False
		For arrCount = 0 to Ubound(arrANS)
			If LCase(selectedPrecodes.QuestionName) = LCase(arrANS[arrCount]) Then
				bFound = True
			End If
		Next
		If bFound = False Then
			ReDim(arrNew, Ubound(arrNew) + 1, True)
			arrNew[Ubound(arrNew) + 1] = selectedPrecodes.QuestionName
		End If		
	Next
	'PERFORM SUM AGAINST ONE QUESTION IN THE GRID	
	If Trim(QuestionName) <> "" Then
		nTotal = 0
		For arrCount = 0 to Ubound(arrNew)
			nTotal = nTotal + Grid.Item[arrCount].Item[ColIndex].Response
		Next	
		If SysCompare(nTotal, TestValue, Operation) = False Then
			CheckSum = False
		Else
			CreateErrorBanner(Grid, QuestionName, nTotal, ErrorMsg, ShowDefaultErrorMsg, IOM, TestValue, Operation)
			CheckSum = True
		End If
	Else	'PERFORM SUM AGAINST ALL QUESTIONS IN THE GRID
		ColIndex = 0
		For Each oCol in Grid[0]
			If LCase(oCol.QuestionName) = LCase(Trim(QuestionName)) Then
				Exit For	
			End If
			nTotal = 0
			For arrCount = 0 to Ubound(arrNew)
				nTotal = nTotal + Grid.Item[arrCount].Item[ColIndex].Response
			Next	
			If SysCompare(nTotal, TestValue, Operation) = False Then
				CheckSum = False
			Else
				CreateErrorBanner(Grid, Grid[0].Item[ColIndex].QuestionName, nTotal, ErrorMsg, ShowDefaultErrorMsg, IOM, TestValue, Operation)
				CheckSum = True
				Exit For
			End If		
			ColIndex = ColIndex + 1
		Next	
	End If	
End Function
'=======ITERNAL SYSTEM FUNCTIONS (NOT TO BE CALLED BY THE USER)=======
'============================================================
'============================================================
'============================================================
'============================================================
'============================================================
'============================================================
Function SysCompare(VAL1, VAL2, OP)
	Dim bFlag
	bFlag = True
	If Trim(OP) = "" Then
		OP = "<"
	End if
	Select Case Trim(OP)
		Case "="
			If VAL1 = VAL2 Then
				bFlag = False
			End If
		Case "<>"
			If VAL1 <> VAL2 Then
				bFlag = False
			End If
		Case ">"
			If VAL1 > VAL2 Then
				bFlag = False
			End If
		Case "<"
			If VAL1 < VAL2 Then
				bFlag = False
			End If
		Case ">="
			If VAL1 >= VAL2 Then
				bFlag = False
			End If		
		Case "<="
			If VAL1 <= VAL2 Then
				bFlag = False
			End If
		Case Else
			'PROVIDE ERROR MESSAGE HERE
	End Select
	SysCompare = bFlag
End Function
Sub CreateErrorBanner(GID, QID, TOTAL, ERRORMSG, SHOW_DEFAULT_ERROR, IOM, TestValue, OP)
	Dim ErrMsg
	ErrMsg = ""
	If SHOW_DEFAULT_ERROR = True Then
		IOM.Questions["sysTemp1"].Response = GID[0].Item[QID].Label
		GID[0].Item[QID].Label.Style.Color = "red"
		IOM.Questions["sysTemp2"].Response = TOTAL
		IOM.Questions["sysTemp3"].Response = TestValue
		Select Case Trim(OP)
			Case "="
				IOM.Questions["sysErrorMessages"].Response.Value = {SYS_ERR_EQ}
			Case "<"
				IOM.Questions["sysErrorMessages"].Response.Value = {SYS_ERR_LT}		
			Case ">"
				IOM.Questions["sysErrorMessages"].Response.Value = {SYS_ERR_GT}	
			Case "<="
				IOM.Questions["sysErrorMessages"].Response.Value = {SYS_ERR_LTE}			
			Case ">="
				IOM.Questions["sysErrorMessages"].Response.Value = {SYS_ERR_GTE}		
			Case Else
				IOM.Questions["sysErrorMessages"].Response.Value = {SYS_ERR_LT}
		End Select
		ErrMsg = "{#sysErrorMessages}"
	End If
	IOM.Questions["hCustomMsgs"].Response = ERRORMSG
	ErrMsg = ErrMsg +  " {#hCustomMsgs}"
	GID.Banners.Clear() 
	GID.Banners.AddNew("ErrBanner", "<table bgcolor=""#FF0000""><tr><td><font color=""#FFFFFF""><b>" + ErrMsg + "</b></font></td></tr></table>")
End Sub
Sub CreateRankErrorBanner(GID, QID, MAX, ERRORMSG, SHOW_DEFAULT_ERROR, IOM)
	Dim ErrMsg
	ErrMsg = ""
	If SHOW_DEFAULT_ERROR = True Then
		IOM.Questions["sysTemp1"].Response = GID[0].Item[QID].Label
		GID[0].Item[QID].Label.Style.Color = "red"
		IOM.Questions["sysTemp2"].Response = MAX
		IOM.Questions["sysErrorMessages"].Response.Value = {SYS_ERR_RNK}
		ErrMsg = "{#sysErrorMessages}"
	End If
	IOM.Questions["hCustomMsgs"].Response = ERRORMSG
	ErrMsg = ErrMsg +  " {#\.hCustomMsgs}"
	'GID.Banners.Clear() 
	GID.Errors.AddNew("ErrBanner", "<table bgcolor=""#FF0000""><tr><td><font color=""#FFFFFF""><b>" + ErrMsg + "</b></font></td></tr></table>")
End Sub

SetHCSampleSub:
	'--- Start of IOM Script item SetHCSample ---
Sub SetHCSample(IOM)
Dim i
	On Error Resume Next
	'SET SAMPLE VARIABLES
	For each i in IOM.Questions.Item["SampleFields"]
'		If Trim(IOM.SampleRecord.Item[i.QuestionName] <> "") Then
		If Trim(IOM.SampleRecord.Item[i.QuestionName] <> "") And (i.QuestionName <> "gop") Then
			i.Response = IOM.SampleRecord.Item[i.QuestionName]
		End If
		IOM.Log(i.QuestionName + "=" + CText(i.Response))
	Next
	'Default to FS2 if no SamID is defined in the sample
	'if IOM.Questions.Item["SampleFields"].SamID.Response = Null then IOM.Questions.Item["SampleFields"].SamID.Response = "FI2"
	If Trim(IOM.SampleRecord.Item["country"]) <> "" Then IOM.Questions.Item["SampleFields"].CNT.Response = IOM.SampleRecord.Item["country"]
End Sub
	'--- End of IOM Script item SetHCSample ---

LogQValueSub:
	'--- Start of IOM Script item LogQValue ---
Sub LogQValue(MyQuestion, IOM)
	On Error Resume Next
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'  Routine to Log all answers to a question to the
	'  log file.
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Dim i
	select case MyQuestion.QuestionType
		case QuestionTypes.qtSimple 
			select case MyQuestion.QuestionDataType
				case DataTypeConstants.mtLevel 
					'Do Nothing
				case DataTypeConstants.mtNone 
					'Do Nothing
				case DataTypeConstants.mtObject 
					'Do Nothing
				Case Else
					if MyQuestion.Response is not null then 
						IOM.Log(MyQuestion.QuestionFullName + "=" + CText(MyQuestion.Response), LOGLEVELS.LOGLEVEL_INFO)
					end if
			end select
		case else
			For each i in MyQuestion
				LogQValue(i, IOM)
			Next
	end select
End Sub
	'--- End of IOM Script item LogQValue ---

Function PendQuota(IOM, QuotaIter, QuotaNames)
      On Error Resume Next
            
      
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '
      ' Example:  Pending Quota.  
      '
      '           if Not(PendQuota(IOM, 0) = True) then goto OVERQUOTA
      '           if Not(PendQuota(IOM, "COMPLETES") = True) then goto OVERQUOTA   
      '
      '        NOTE:  QuotaName can be string name or index value.
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      Dim Result
      Dim MyQuota
      Dim MyText, OQText
      Dim ArrayQuotas, objQName, strHyphen, NullQuota
      
      PendQuota = Null 'Leave as Null in case an error is returned
      
      'Force to be true for mrStudio
      if (IOM.Info.IsDebug) then 
            PendQuota = True
            exit function
      end if
      
      'Pend Quota
      Result = IOM.QuotaEngine.QuotaGroups[QuotaIter].Pend() 
      If (IsSet(Result, QuotaResultConstants.qrWasPended) and IsSet(result, QuotaResultConstants.qrBelowTarget)) then
            PendQuota = True
            MyText = "QUOTAINFO::Pending Passed For QuotaGroup["
            OQText = "No over-quota instance found."
      Else
            PendQuota = False
            MyText = "QUOTAINFO::Failed to Pend For QuotaGroup[" 
            OQText = "Over-quota occurred at quota table [" + IOM.QuotaEngine.QuotaGroups[QuotaIter].Name + "]. "
                                    
            ArrayQuotas = Split(QuotaNames,",")
			
			NullQuota = 0
			For each objQName in ArrayQuotas
				If IOM.Questions[objQName].Response = NULL Then NullQuota = 1
			Next
			
			If NullQuota = 0 Then
				OQText = OQText + "Quota cell:"
	            For each objQName in ArrayQuotas
	            	strHyphen = " - "
	            	OQText = OQText + strHyphen + IOM.Questions[objQName].Response.Label
	            Next
	        Else
	        	OQText = OQText + "One of the quota markers has missing values."
	        End If
      End If
      
      IOM.Questions["MRK_OverQuota"].Response = OQText
      
      'Log current counts to logs.
      'MyText = "QUOTAINFO::Failed to Pend For QuotaGroup[" 
      MyText = MyText + IOM.QuotaEngine.QuotaGroups[QuotaIter].Name 
      MyText = MyText + "]. - on Serial: " 
      MyText = MyText + CText(IOM.Info.Serial) 
      MyText = MyText + "  Current -"
      For Each MyQuota in IOM.QuotaEngine.QuotaGroups[QuotaIter].Quotas
            MyText = MyText + " Name=" 
            MyText = MyText + MyQuota.Name 
            MyText = MyText + "=(" 
            MyText = MyText + CText(MyQuota.Target) 
            MyText = MyText + "," 
            MyText = MyText + CText(MyQuota.Completed) 
            MyText = MyText + "," 
            MyText = MyText + CText(MyQuota.Pending) 
            MyText = MyText + ","
            MyText = MyText + CText(MyQuota.IsBelowQuota) 
            MyText = MyText + "," 
            MyText = MyText + CText(MyQuota.IsOverQuota) 
            MyText = MyText + ","  
            MyText = MyText + CText(MyQuota.WasPended) 
            MyText = MyText + ")" 
      Next
      
      IOM.Log(MyText, LogLevels.LOGLEVEL_INFO)

End Function

''''''OLD function''''''''''''
'Function PendQuota(IOM, QuotaIter)
'      On Error Resume Next
'      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'      '
'      ' Example:  Pending Quota.  
'      '
'      '           if Not(PendQuota(IOM, 0) = True) then goto OVERQUOTA
'      '           if Not(PendQuota(IOM, "COMPLETES") = True) then goto OVERQUOTA   
'      '
'      '        NOTE:  QuotaName can be string name or index value.
'      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'      Dim Result
'      Dim MyQuota
'      Dim MyText
'      PendQuota = Null 'Leave as Null in case an error is returned
'      
'      'Force to be true for mrStudio
'      if (IOM.Info.IsDebug) then 
'            PendQuota = True
'            exit function
'      end if
'      
'      'Pend Quota
'      Result = IOM.QuotaEngine.QuotaGroups[QuotaIter].Pend() 
'      If (IsSet(Result, QuotaResultConstants.qrWasPended) and IsSet(result, QuotaResultConstants.qrBelowTarget)) then
'            PendQuota = True
'            MyText = "QUOTAINFO::Pending Passed For QuotaGroup[" 
'      Else
'            PendQuota = False
'            MyText = "QUOTAINFO::Failed to Pend For QuotaGroup[" 
'      End If
'      'Log current counts to logs.
'      'MyText = "QUOTAINFO::Failed to Pend For QuotaGroup[" 
'      MyText = MyText + IOM.QuotaEngine.QuotaGroups[QuotaIter].Name  
'      MyText = MyText + "]. - on Serial: " 
'      MyText = MyText + CText(IOM.Info.Serial) 
'      MyText = MyText + "  Current -"
'      For Each MyQuota in IOM.QuotaEngine.QuotaGroups[QuotaIter].Quotas
'            MyText = MyText + " Name=" 
'            MyText = MyText + MyQuota.Name 
'            MyText = MyText + "=(" 
'            MyText = MyText + CText(MyQuota.Target) 
'            MyText = MyText + "," 
'            MyText = MyText + CText(MyQuota.Completed) 
'            MyText = MyText + "," 
'            MyText = MyText + CText(MyQuota.Pending) 
'            MyText = MyText + ","
'            MyText = MyText + CText(MyQuota.IsBelowQuota) 
'            MyText = MyText + "," 
'            MyText = MyText + CText(MyQuota.IsOverQuota) 
'            MyText = MyText + ","  
'            MyText = MyText + CText(MyQuota.WasPended) 
'            MyText = MyText + ")"
'      Next
'      IOM.Log(MyText, LogLevels.LOGLEVEL_INFO)
'
'End Function


'Function PendQuota(IOM, QuotaIter)
'	On Error Resume Next
'	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'	'
'	' Example:  Pending Quota.  
'	'
'	'           if Not(PendQuota(IOM, 0) = True) then goto OVERQUOTA
'	'           if Not(PendQuota(IOM, "COMPLETES") = True) then goto OVERQUOTA	
'	'
'	'        NOTE:  QuotaName can be string name or index value.
'	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'	Dim Result
'	Dim MyQuota
'	Dim MyText
'	PendQuota = Null 'Leave as Null in case an error is returned
'	
'	'Force to be true for mrStudio
'	if IOM.Info.IsDebug then 
'		PendQuota = True
'		exit function
'	end if
'	
'	'Pend Quota
'	Result = IOM.QuotaEngine.QuotaGroups[QuotaIter].Pend() 
'	If (IsSet(Result, QuotaResultConstants.qrWasPended) and IsSet(result, QuotaResultConstants.qrBelowTarget)) then
'		PendQuota = True
'	Else
'		PendQuota = False
'		'If unable to pend, send current counts to logs.
'		MyText = "QUOTAINFO::Failed to Pend For QuotaGroup[" 
'		MyText = MyText + IOM.QuotaEngine.QuotaGroups[QuotaIter].Name  
'		MyText = MyText + "]. - on Serial: " 
'		MyText = MyText + CText(IOM.Info.Serial) 
'		MyText = MyText + "  Current -"
'		For Each MyQuota in IOM.QuotaEngine.QuotaGroups[QuotaIter].Quotas
'			MyText = MyText + " Name=" 
'			MyText = MyText + MyQuota.Name 
'			MyText = MyText + "=(" 
'			MyText = MyText + CText(MyQuota.Target) 
'			MyText = MyText + "," 
'			MyText = MyText + CText(MyQuota.Completed) 
'			MyText = MyText + "," 
'			MyText = MyText + CText(MyQuota.Pending) 
'			MyText = MyText + ","
'			MyText = MyText + CText(MyQuota.IsBelowQuota) 
'			MyText = MyText + "," 
'			MyText = MyText + CText(MyQuota.IsOverQuota) 
'			MyText = MyText + ","  
'			MyText = MyText + CText(MyQuota.WasPended) 
'			MyText = MyText + ")"
'		Next
'		IOM.Log(MyText, LogLevels.LOGLEVEL_INFO)
'	End If
'
'End Function

Sub QuestionTimings(IOM)
	Dim NewTime
	NewTime = Now()
	'Check if Null
	if IOM.Questions.Item["CurrentQuestionTime"].Response = Null then 
		IOM.Questions.Item["InterviewLength"].Response = Second(IOM.Info.ElapsedTime)
	else
		'If interview gets restarted, use current elapsed time since it'll be when the interview restarted.
		if DateDiff(IOM.Questions.Item["CurrentQuestionTime"].Response, NewTime, "s") < 300 then
			IOM.Questions.Item["InterviewLength"].Response = IOM.Questions.Item["InterviewLength"].Response + DateDiff(IOM.Questions.Item["CurrentQuestionTime"].Response, NewTime, "s")
		end if
	end if			
	''IOM.Questions.Item["CurrentQuestionTime"].Response = NewTime
End Sub

'----------------------------------------------------------------------------------------------------------------------------
'Added by Ugam
'----------------------------------------------------------------------------------------------------------------------------
sub ShowTest(IOM, QName, TestText)
	
	If ContainsAny(IOM.Questions["SMODE"],{test}) Then
		RemoveBanner(IOM,QName,"TestScreen")
		
		IOM.Info.EstimatedProgress = IOM.Info.EstimatedProgress - 1
		QName.Banners.AddNew("TestScreen","<B><font color=""red"">DISPLAYED DURING TESTING ONLY</font><p/>" + ctext(TestText) + "<br/></B>")
		QName.Show()		
	End If

end sub

sub AskTest(IOM, QName, TestText)

	If ContainsAny(IOM.Questions["SMODE"],{test}) Then
		IOM.Info.EstimatedProgress = IOM.Info.EstimatedProgress - 1
		QName.Banners.AddNew("TestScreen","<B><font color=""red"">ASKED DURING TESTING ONLY</font><p/>" + ctext(TestText) + "<br/></B>")
		QName.Ask()	
	End If

end sub

Sub ShowOther(Qname)	
	Qname.QuestionTemplate = "Hide_Question.htm"
	Qname.Style.Hidden = True	
	Qname.Label.Style.Hidden = True
	Qname.Style.Control.Type = ControlTypes.ctSingleLineEdit	
	Qname.mustanswer = false
End Sub

Sub FormatOtherSpecify(IOM,Ques)
	if IOM.Info.IsDebug = false then	
		Ques.Style.Hidden = true
	end if
	Ques.MustAnswer = false	
	Ques.Style.Control.Type = ControlTypes.ctSingleLineEdit
	Ques.Style.Control.ReadOnly = true
	Ques.Style.BgColor = "#d2d2d2"
End Sub

function jsGridOtherSpecify(IOM,Qname,OtherIds)
	Dim DSjsGridOtherSpecify	
	
	DSjsGridOtherSpecify = "<script>addEvent (window,""load"",function () &#123;GridOtherSpecify(""" + OtherIds + """);  &#125;)</script>"
		
	Qname.Banners.AddNew("jsGridOtherSpecify",DSjsGridOtherSpecify)
end function

Sub jsShowHeader (IOM, Qname, Heading1, Heading2,PageItem)
	'signoff
	
	RemoveIOMBannerIfExists(IOM,"jsShowHeader")
	Dim jsShowHeader_Banner

	jsShowHeader_Banner = "<script>addEvent (window,""load"",function () &#123;ShowRowHeader (""{Heading1}"", ""{Heading2}"",{PageItem}) &#125;); </script>"
	
	jsShowHeader_Banner = replace(jsShowHeader_Banner,"{Heading1}",IOM.Questions["prog_jsBanner"].Categories[Heading1].Label)
	jsShowHeader_Banner = replace(jsShowHeader_Banner,"{Heading2}",IOM.Questions["prog_jsBanner"].Categories[Heading2].Label)
	jsShowHeader_Banner = replace(jsShowHeader_Banner,"{PageItem}",PageItem)	
	
	qname.Banners.AddNew("jsShowHeader",jsShowHeader_Banner)	
End Sub

Sub jsShowNumericHeaders (IOM, qname, NoOfHeaders, NumHead1, NumHead2, NumHead3,PageItem)
	RemoveIOMBannerIfExists(IOM,"jsNumericHeaders")
	Dim jsNumericHeaders_Banner
	
	jsNumericHeaders_Banner = "<script>addEvent (window,""load"",function () &#123;ShowNumericHeaders({NoOfHeaders}, ""{NumHead1}"", ""{NumHead2}"", ""{NumHead3}"",{PageItem}) &#125;);</script>"
	
	jsNumericHeaders_Banner = Replace(jsNumericHeaders_Banner,"{NoOfHeaders}",NoOfHeaders)	
	jsNumericHeaders_Banner = Replace(jsNumericHeaders_Banner,"{NumHead1}",IOM.Questions["prog_jsBanner"].Categories[NumHead1].Label)	
	jsNumericHeaders_Banner = Replace(jsNumericHeaders_Banner,"{NumHead2}",IOM.Questions["prog_jsBanner"].Categories[NumHead2].Label)
	jsNumericHeaders_Banner = Replace(jsNumericHeaders_Banner,"{NumHead3}",IOM.Questions["prog_jsBanner"].Categories[NumHead3].Label)
	jsNumericHeaders_Banner = Replace(jsNumericHeaders_Banner,"{PageItem}",PageItem)
			
	qname.Banners.AddNew("jsNumericHeaders",jsNumericHeaders_Banner)
End Sub

function jsRightText(IOM,Qname,Columns,RightText,IgnoreTBs, PageItem)	
	RemoveIOMBannerIfExists(IOM,"jsRightText")
	
	Dim DSjsRightText, prog_righttext

	Set prog_righttext = IOM.Questions["prog_righttext"]
	prog_righttext.Response = RightText

	DSjsRightText ="<SCRIPT>addEvent (window,""load"",function () &#123;RightText ({Columns},""{rtext}"",""{IgnoreTbs}"",{PageItem}) &#125;); </SCRIPT>"

	DSjsRightText = Replace(DSjsRightText,"{Columns}",Columns)
	DSjsRightText = Replace(DSjsRightText,"{rtext}",prog_righttext.Response.Label)
	DSjsRightText = Replace(DSjsRightText,"{IgnoreTbs}",IgnoreTBs)
	DSjsRightText = Replace(DSjsRightText,"{PageItem}",PageItem)
			
	Qname.Banners.AddNew("jsRightText",DSjsRightText)
end function

function jsLeftText(IOM,Qname,Columns,LeftText,IgnoreTBs, PageItem)
	RemoveIOMBannerIfExists(IOM,"jsLeftText")
	
	Dim DSjsLeftText, prog_lefttext
	
	DSjsLeftText = "<SCRIPT>addEvent (window,""load"",function () &#123;LeftText ({Columns},""{ltext}"",""{IgnoreTbs}"",""{PageItem}"") &#125;); </SCRIPT>"
	
	Set prog_lefttext = IOM.Questions["prog_lefttext"]
	prog_lefttext = CCategorical(LeftText)
	
	DSjsLeftText = Replace(DSjsLeftText,"{Columns}",Columns)	
	DSjsLeftText = Replace(DSjsLeftText,"{ltext}",prog_lefttext.Response.Label)
	DSjsLeftText = Replace(DSjsLeftText,"{IgnoreTbs}",IgnoreTBs)	
	DSjsLeftText = Replace(DSjsLeftText,"{PageItem}",PageItem)	
	
	Qname.Banners.AddNew("jsLeftText",DSjsLeftText)	
end function

Sub HeadStyle (qname)
	'!
	When we have multiple questions in the fields section the columns header do not have proper 
	formatting. This functions aims at doing so.
	
	the only parameter this function requires is the reference to the inner question. The function 
	needs to be called as many times the number of questions in the field section.
	
	e.g.
	HeadStyle(GRID_Q52[..].Q52A)
	!'
	qname.Label.Style.Align = Alignments.alCenter
	qname.Label.Style.Width = "10em"
	qname.Label.Style.Cell.BorderColor = "#000000"
	qname.Label.Style.Cell.BorderStyle = BorderStyles.bsSolid
	qname.Label.Style.Cell.BorderWidth = 1
	qname.Label.Style.Cell.Padding = 8
	'qname.Label.Style.Cell.BgColor = "#009493"
	qname.Label.Style.Cell.BgColor = "#004360"	
	qname.Label.Style.Color = "#ffffff"
End Sub

SUB GRID_DROPLIST(IOM,QNAME,INN_Q) ' DROPLIST IN A GRID QUESTION
	QNAME[..].ITEM[INN_Q].Style.Control = controltypes.ctDropList 
	QNAME[..].ITEM[INN_Q].Categories[..].Label.Style.BgColor = "#ffffff"
	QNAME[..].ITEM[INN_Q].Categories[..].Label.Style.Color = "#000000"
	 
	QNAME[..].ITEM[INN_Q].Categories[..].Categories[..].Label.Style.BgColor = "#ffffff"
	QNAME[..].ITEM[INN_Q].Categories[..].Categories[..].Label.Style.Color = "#000000"
END SUB


sub jsShowRunTotalRowsWise(IOM,Qname,TotalText, NoOfColumns, NoOfRows,HasColumnHeader, RowsToHide, RepeatHeader, IgnoreTBs, PageItem)
	RemoveIOMBannerIfExists(IOM,"jsShowRunTotalRowsWise")
	
	Dim DSjsShowRunTotalRowsWise

	DSjsShowRunTotalRowsWise = "<SCRIPT>RunTotalRowsWiseAddEvents (""{TotalText}"",{NoOfColumns}, {NoOfRows}, {HasColumnHeader}, ""{RowsToHide}"", {RepeatHeader} ,""{IgnoreTBs}"", {PageItem})</SCRIPT>"
	
	DSjsShowRunTotalRowsWise = Replace(DSjsShowRunTotalRowsWise,"{TotalText}",IOM.Questions["prog_jsBanner"].Categories[TotalText].Label)
	DSjsShowRunTotalRowsWise = Replace(DSjsShowRunTotalRowsWise,"{NoOfColumns}",NoOfColumns)
	DSjsShowRunTotalRowsWise = Replace(DSjsShowRunTotalRowsWise,"{NoOfRows}",NoOfRows)
	DSjsShowRunTotalRowsWise = Replace(DSjsShowRunTotalRowsWise,"{IgnoreTBs}",IgnoreTBs)
	DSjsShowRunTotalRowsWise = Replace(DSjsShowRunTotalRowsWise,"{PageItem}",PageItem)
	DSjsShowRunTotalRowsWise = Replace(DSjsShowRunTotalRowsWise,"{HasColumnHeader}",lcase(HasColumnHeader))
	DSjsShowRunTotalRowsWise = Replace(DSjsShowRunTotalRowsWise,"{RepeatHeader}",lcase(RepeatHeader))
	DSjsShowRunTotalRowsWise = Replace(DSjsShowRunTotalRowsWise,"{RowsToHide}",RowsToHide)
	
	Qname.Banners.AddNew("jsShowRunTotalRowsWise",DSjsShowRunTotalRowsWise)
end sub

'Custom function for Q536/Q537
sub jsShowRunTotalColumnsWise_Custom(IOM, Qname, TotalText, NoOfColumns, NoOfRows, HasRowHeader, ColumnsToHide, HasOtherColumn, Start_Column, Left_Text, Right_Text, IgnoreTBs, PageItem)
'!
The Javascript function that would be called will add a row at the end of table starting with the 2nd column and running total would be shown 

Explantion on how to use the function "jsShowRunTotalColumnsWise_Custom" for showing the running total
for columns
1. IOM			: Interview Object Model

2. Qname		: Question name for which we need to to show the totals.

3. TotalText	: The text that needs to shown for the total row. Here we need to pass the ID of question
					prog_jsBanner which has the text that needs to be shown for Totla row.

4. NoOfColumns	: No. of columns shown. if the columns shown are dynamically changing then it has
					to be handled accordingly i.e. by passing category count or question count.
					We would also need to take care if the grid is a column grid or a row grid.
					
5. NoOfRows		: No. of rows shown. if the rows shown are dynamically changing then it has
					to be handled accordingly i.e. by passing category count or question count.
					We would also need to take care if the grid is a column grid or a row grid.
					
6. HasRowHeader	: If the rows shown has headers then we need to pass the value "true" 
					else the value should "false". This is a string parameter.
					
7. ColumnsToHide: If there are any columns that doesnot need to shown the running total the we need to
					pass it as a parameter. If there are more than one columns then they need to be
					passed comma seperated. This is a string parameter.

8. HasOtherColumn : Indicates whether grid has a column for "Other Specify" or not. Should be "true" or "false".

9. Start_Column: Indicates from which column the total row should be started. If by default you want 
					to have Total row to start with 1st row then specify 1, else specify the relevant 
					column number from where you want the Total row should begin.

10. Left_Text:	Specify any left text to be shown along the total figure at the bottom, if nothing then just
				specify {_}. To be added via prog_lefttext

11. Right_Text:	Specify any right text to be shown along the total figure at the bottom if nothing then just
				specify {_}. To be added via prog_Righttext

12. IgnoreTBs	: If there are any other specify in the grid then the name(s) of those textboxes 
					need to be passed. If more then one other specify then they need to be passed
					comma seperated. 
				  This normally used to ignore other specify text boxes used in the Grids. However it 
				  can be used to ignore one full row of textbox by specifying their names in this 
				  parameter. however only ignoring only one of the textboxes will cause problems.

13. PageItem		: If there are other questions shown along with the question [Question where running 
					total need to be shown] then we need to specify at which location is the question
					shown on the screen. If they questions are randomised then we would need to devide 
					a method to be able to pass this value [this may depend upon study requirements].
				  If there only question then we need to pass 1 (one).
	
!'
	RemoveIOMBannerIfExists(IOM,"jsShowRunTotalByColumns")

	Dim DSjsShowRunTotalByColumns
		
	DSjsShowRunTotalByColumns = "<SCRIPT>RunTotalColumnsWiseAddEvents_Custom (""{TotalText}"",{NoOfColumns}, {NoOfRows}, {HasRowHeader}, ""{ColumnsToHide}"", {HasOtherColumn}, {Start_Column}, ""{Left_Text}"", ""{Right_Text}"", ""{IgnoreTBs}"", {PageItem})</SCRIPT>"
		
	DSjsShowRunTotalByColumns = Replace(DSjsShowRunTotalByColumns, "{TotalText}", IOM.Questions["prog_jsBanner"].Categories[TotalText].Label)
	DSjsShowRunTotalByColumns = Replace(DSjsShowRunTotalByColumns, "{NoOfColumns}", NoOfColumns)
	DSjsShowRunTotalByColumns = Replace(DSjsShowRunTotalByColumns, "{NoOfRows}", NoOfRows)
	DSjsShowRunTotalByColumns = Replace(DSjsShowRunTotalByColumns, "{HasRowHeader}", lcase(HasRowHeader))
	DSjsShowRunTotalByColumns = Replace(DSjsShowRunTotalByColumns, "{ColumnsToHide}", ColumnsToHide)
	DSjsShowRunTotalByColumns = Replace(DSjsShowRunTotalByColumns, "{HasOtherColumn}", lcase(HasOtherColumn))
	DSjsShowRunTotalByColumns = Replace(DSjsShowRunTotalByColumns, "{Start_Column}", Start_Column)
	DSjsShowRunTotalByColumns = Replace(DSjsShowRunTotalByColumns, "{Left_Text}", IOM.Questions["prog_lefttext"].Categories[Left_Text].Label)
	DSjsShowRunTotalByColumns = Replace(DSjsShowRunTotalByColumns, "{Right_Text}", IOM.Questions["prog_righttext"].Categories[Right_Text].Label)
	DSjsShowRunTotalByColumns = Replace(DSjsShowRunTotalByColumns, "{IgnoreTBs}", IgnoreTBs)
	DSjsShowRunTotalByColumns = Replace(DSjsShowRunTotalByColumns, "{PageItem}", PageItem)
	
	Qname.Banners.AddNew("jsShowRunTotalByColumns",DSjsShowRunTotalByColumns)
end sub

sub jsShowRunTotalColumnsWiseP(IOM, Qname, TotalText, NoOfColumns, NoOfRows, HasRowHeader, ColumnsToHide, IgnoreTBs, PageItem, EndText)
'!
The Javascript function that would be called will add a row at the end of table and running total would be shown 

Explantion on how to use the function "jsShowRunTotalColumnsWise" for showing the running total
for columns
1. IOM			: Interview Object Model

2. Qname		: Question name for which we need to to show the totals.

3. TotalText	: The text that needs to shown for the total row. Here we need to pass the ID of question
					prog_jsBanner which has the text that needs to be shown for Totla row.

4. NoOfColumns	: No. of columns shown. if the columns shown are dynamically changing then it has
					to be handled accordingly i.e. by passing category count or question count.
					We would also need to take care if the grid is a column grid or a row grid.
					
5. NoOfRows		: No. of rows shown. if the rows shown are dynamically changing then it has
					to be handled accordingly i.e. by passing category count or question count.
					We would also need to take care if the grid is a column grid or a row grid.
					
6. HasRowHeader	: If the rows shown has headers then we need to pass the value "true" 
					else the value should "false". This is a string parameter.
					
7. ColumnsToHide: If there are any columns that doesnot need to shown the running total the we need to
					pass it as a parameter. If there are more than one columns then they need to be
					passed comma seperated. This is a string parameter.
					
8. IgnoreTBs	: If there are any other specify in the grid then the name(s) of those textboxes 
					need to be passed. If more then one other specify then they need to be passed
					comma seperated. 
				  This normally used to ignore other specify text boxes used in the Grids. However it 
				  can be used to ignore one full row of textbox by specifying their names in this 
				  parameter. however only ignoring only one of the textboxes will cause problems.

9. PageItem		: If there are other questions shown along with the question [Question where running 
					total need to be shown] then we need to specify at which location is the question
					shown on the screen. If they questions are randomised then we would need to devide 
					a method to be able to pass this value [this may depend upon study requirements].
				  If there only question then we need to pass 1 (one).
!'
	RemoveIOMBannerIfExists(IOM,"jsShowRunTotalByColumnsP")

	Dim DSjsShowRunTotalByColumnsP
		
	DSjsShowRunTotalByColumnsP = "<SCRIPT>RunTotalColumnsWiseAddEventsP (""{TotalText}"",{NoOfColumns}, {NoOfRows}, {HasRowHeader}, ""{ColumnsToHide}"", ""{IgnoreTBs}"", {PageItem}, ""{EndText}"")</SCRIPT>"
		
	DSjsShowRunTotalByColumnsP = Replace(DSjsShowRunTotalByColumnsP, "{TotalText}", IOM.Questions["prog_jsBanner"].Categories[TotalText].Label)
	DSjsShowRunTotalByColumnsP = Replace(DSjsShowRunTotalByColumnsP, "{NoOfColumns}", NoOfColumns)
	DSjsShowRunTotalByColumnsP = Replace(DSjsShowRunTotalByColumnsP, "{NoOfRows}", NoOfRows)
	DSjsShowRunTotalByColumnsP = Replace(DSjsShowRunTotalByColumnsP, "{IgnoreTBs}", IgnoreTBs)
	DSjsShowRunTotalByColumnsP = Replace(DSjsShowRunTotalByColumnsP, "{PageItem}", PageItem)
	DSjsShowRunTotalByColumnsP = Replace(DSjsShowRunTotalByColumnsP, "{HasRowHeader}", lcase(HasRowHeader))
	DSjsShowRunTotalByColumnsP = Replace(DSjsShowRunTotalByColumnsP, "{ColumnsToHide}", ColumnsToHide)
	DSjsShowRunTotalByColumnsP = Replace(DSjsShowRunTotalByColumnsP, "{EndText}", IOM.Questions["prog_jsBanner"].Categories[EndText].Label)
	
	Qname.Banners.AddNew("jsShowRunTotalByColumnsP",DSjsShowRunTotalByColumnsP)
end sub


Sub ShowQuestionNo (IOM,Qno)
      Dim smode
      Set smode = IOM.Questions["smode"]
      
      RemoveIOMBannerIfExists(IOM,"Qname")
      If ContainsAny(smode,{TEST}) Then
            IOM.Banners.AddNew("Qname",Qno)
      Else
            IOM.Banners.AddNew("Qname"," ")
      End If
End Sub     

 
 sub RemoveIOMBannerIfExists(IOM,BannerName)
      Dim i
      for i=0 to IOM.Banners.Count - 1          
            if ucase(IOM.Banners.Item[i].Name) = ucase(trim(BannerName)) then
                  IOM.Banners.Remove(BannerName)
                  exit for
            end if
      next
end sub


Sub jsFocusTabDown (IOM, qname, NoOfColumns, NoOfRows, IgnoreTBs, OtherTBs, PageItem)
	RemoveIOMBannerIfExists(IOM,"jsTabDown")
	
	Dim jsTabDown
	jsTabDown = "<script>addEvent (window,""load"",function () &#123;FocusTabDown({NoOfColumns}, {NoOfRows}, ""{IgnoreTBs}"", ""{OtherTBs}"", {PageItem}) &#125;);</script>"
	
	jsTabDown = Replace(jsTabDown,"{NoOfColumns}", NoOfColumns)
	jsTabDown = Replace(jsTabDown,"{NoOfRows}", NoOfRows)
	jsTabDown = Replace(jsTabDown,"{IgnoreTBs}", IgnoreTBs)
	jsTabDown = Replace(jsTabDown,"{OtherTBs}", OtherTBs)
	jsTabDown = Replace(jsTabDown,"{PageItem}", PageItem)
	
	qname.Banners.AddNew("jsTabDown",jsTabDown)			
End Sub

Function RunTotal(Question, IOM, Attempt)
		
		Dim Total,QCat, TotalRequired, QTOTALError
		
		Set QTOTALError = IOM.Questions["hCustomMsgs"].Categories.SYS_ERR_EQ		
		TotalRequired = IOM.Questions["DUMTotal"].Response		
		Total = 0
				
		For QCat = 1 To Question.Count
			if lcase(Question.Item[QCat-1].QuestionName) <> "total" then
				Total = Total + Question.Item[QCat-1].Item[0].Response
			End If
		Next
		
		If Total <> TotalRequired Then
			'Insert total entered by user in the error msg
	        QTOTALError.Label.Inserts["UTOTAL"] = Total
	        
	        'Insert the desired total
	        QTOTALError.Label.Inserts["TOTAL"] = TotalRequired
	        Question.Errors.AddNew("TotalError",QTOTALError.Label)
	        'If attempt is zero that means it has been called without validation function 
			'and hence progress bar has to be manually adjusted
	        if Attempt = 0 Then IOM.Info.EstimatedProgress = IOM.Info.EstimatedProgress - 1
	        RunTotal = False
		Else
			RunTotal = True
		End If

End Function

Sub ShowRT (Qname,RT)
		Qname.Codes[RT].Style.Hidden = True
		Qname.Codes[RT].Style.ElementAlign = ElementAlignments.eaRight
		Qname.Codes[RT].Style.Indent = -1
		Qname.Codes[RT].Label.Style.Font.Effects = 0
		Qname.Codes[RT].Label.Style.Cell.BorderStyle = BorderStyles.bsNone
		Qname.Codes[RT].Label.Style.Cell.BgColor = ""
		Qname.Codes[RT].Label.Style.COLOR = "#000000"
End Sub

Sub Show_BgColDk (Qname,RT)
	Qname.Codes[RT].Style.ElementAlign=ElementAlignments.eaRight
	Qname.Codes[RT].Label.Style.Cell.BorderStyle = BorderStyles.bsNone
	Qname.Codes[RT].Label.Style.Cell.BgColor = ""	
	Qname.Codes[RT].Label.Style.Cell.BorderStyle = BorderStyles.bsNone
	Qname.Codes[RT].Label.Style.Cell.BgColor = ""
	Qname.Codes[RT].Label.Style.COLOR = "#000000"
End Sub

'function CheckQuota(QuotaName,IOM)
'	dim osmode,dmUSER
'	
'	Set osmode = iom.Questions["smode"]
'	set dmUSER = IOM.Questions["dmUSER"]
'	
'	If osmode.Response = {LIVE} or UCase(Trim(dmUSER.Response)) = "PROG" Then
'		Dim quota_pend_result
'		quota_pend_result = IOM.QuotaEngine.QuotaGroups[QuotaName].Pend() 'Pend Quota
'		If Not (IsSet(quota_pend_result, QuotaResultConstants.qrWasPended)) Then
'				CheckQuota = 1
'		Else
'			CheckQuota = 0
'		End If	
'	Else
'		CheckQuota = 0
'	End If
'end function

Sub AddProgress(IOM,Amount)
  IOM.Info.EstimatedProgress = IOM.Info.EstimatedProgress + Amount
  
  EstimatedProgress(IOM) ' Do not comment this line
End Sub

Sub EstimatedProgress(IOM)
	IOM.Questions["InterviewProgress"] = clong((cdouble(IOM.Info.EstimatedProgress) / cdouble(IOM.Info.EstimatedPages)) * 100)
End Sub

Sub OtherSpecifyInGrid(Question,OtherCats,IOM)
'!Syntax: OtherSpecifyInGrid (Question,OtherCats,IOM)
      This Function does the following.  
•     Convert the OtherBox to SingleLineEdit box.
•     Set MustAnswer false for all textboxes.
•     Hide the TextBox for all category except for Other Categories
•     Removes the border between category and textbox.
!'
 
'Convert the OtherBox to SingleLineEdit box.
Question[..].item[0].Style.Control.Type = controltypes.ctSingleLineEdit
 
'Set MustAnswer false for all textboxes.
Question[..].item[0].MustAnswer = False

 'Hide the TextBox for all category except for Other Categories

Dim iItem
      For iItem = 0 to Question.count - 1       
            If Not(ContainsAny(CCategorical(Question[iItem].QuestionName),OtherCats)) Then
                  Question[iItem].item[0].Style.Hidden = True
            End If
      Next

'Removes the border between category and textbox.

'-----------------------------------------------------------------------------------------------------------------
Dim Cat, SubCat
 
For Each Cat in Question.Categories
      if Cat.CategoryType = CategoryTypes.ctCategoryList then
            for Each SubCat in Question.Categories[Cat.name].Categories
                  'For UK
                  Question.Categories[Cat.name].Categories[SubCat].Label.Style.Cell.BorderWidth = 0
                  Question.Categories[Cat.name].Categories[SubCat].Label.Style.Cell.BorderStyle = BorderStyles.bsNone
                  Question.Categories[Cat.name].Categories[SubCat].Label.Style.Cell.BorderColor = mr.None
            
                  Question.Categories[Cat.name].Categories[SubCat].Label.Style.Cell.BorderleftWidth = 1
                  Question.Categories[Cat.name].Categories[SubCat].Label.Style.Cell.BorderleftStyle = BorderStyles.bsSolid
                  Question.Categories[Cat.name].Categories[SubCat].Label.Style.Cell.BorderleftColor = mr.black
                
                  Question.Categories[Cat.name].Categories[SubCat].Label.Style.Cell.BorderTopWidth = 1
                  Question.Categories[Cat.name].Categories[SubCat].Label.Style.Cell.BorderTopStyle = BorderStyles.bsSolid
                  Question.Categories[Cat.name].Categories[SubCat].Label.Style.Cell.BorderTopColor = mr.black
                 
                  Question.Categories[Cat.name].Categories[SubCat].Label.Style.Cell.BorderBottomWidth = 1
                  Question.Categories[Cat.name].Categories[SubCat].Label.Style.Cell.BorderBottomStyle = BorderStyles.bsSolid
                  Question.Categories[Cat.name].Categories[SubCat].Label.Style.Cell.BorderBottomColor = mr.black
                  
                  'For US
                  Question.Categories[Cat.name].Categories[SubCat].Label.Style.Cell.BorderRightColor =  mr.None
                  Question.Categories[Cat.name].Categories[SubCat].Label.Style.Cell.BorderRightwidth =  0
                  Question.Categories[Cat.name].Categories[SubCat].Label.Style.Cell.BorderRightStyle = borderstyles.bsNone
            next
      Else
            'For UK
            Question.Categories[Cat].Label.Style.Cell.BorderWidth = 0
            Question.Categories[Cat].Label.Style.Cell.BorderStyle = BorderStyles.bsNone
            Question.Categories[Cat].Label.Style.Cell.BorderColor = mr.None
            
            Question.Categories[Cat].Label.Style.Cell.BorderleftWidth = 1
            Question.Categories[Cat].Label.Style.Cell.BorderleftStyle = BorderStyles.bsSolid
            Question.Categories[Cat].Label.Style.Cell.BorderleftColor = mr.black
            Question.Categories[Cat].Label.Style.Cell.BorderTopWidth = 1
            Question.Categories[Cat].Label.Style.Cell.BorderTopStyle = BorderStyles.bsSolid
            Question.Categories[Cat].Label.Style.Cell.BorderTopColor = mr.black
            Question.Categories[Cat].Label.Style.Cell.BorderBottomWidth = 1
            Question.Categories[Cat].Label.Style.Cell.BorderBottomStyle = BorderStyles.bsSolid
            Question.Categories[Cat].Label.Style.Cell.BorderBottomColor = mr.black
            'For US
            Question.Categories[Cat].Label.Style.Cell.BorderRightColor =  mr.None
            Question.Categories[Cat].Label.Style.Cell.BorderRightwidth =  0
            Question.Categories[Cat].Label.Style.Cell.BorderRightStyle = borderstyles.bsNone
      end if
Next

'For UK     
Question[..].item[0].Style.Cell.BorderWidth = 0
Question[..].item[0].Style.Cell.BorderStyle = borderstyles.bsNone
Question[..].item[0].Style.Cell.BorderColor = mr.None
Question[..].item[0].Style.Cell.BorderRightWidth = 1
Question[..].item[0].Style.Cell.BorderRightStyle = borderstyles.bsSolid
Question[..].item[0].Style.Cell.BorderRightColor = mr.black       
Question[..].item[0].Style.Cell.BorderBottomWidth = 1
Question[..].item[0].Style.Cell.BorderBottomStyle = borderstyles.bsSolid
Question[..].item[0].Style.Cell.BorderBottomColor = mr.black      
Question[..].item[0].Style.Cell.BorderTopWidth = 1
Question[..].item[0].Style.Cell.BorderTopStyle = borderstyles.bsSolid
Question[..].item[0].Style.Cell.BorderTopColor = mr.black
'-----------------------------------------------------------------------------------------------------------------
'For US
Question[..].Item[0].Style.Cell.BorderLeftColor = mr.None
Question[..].Item[0].Style.Cell.BorderLeftWidth = 0
Question[..].Item[0].Style.Cell.BorderLeftStyle = borderstyles.bsNone

'-----------------------------------------------------------------------------------------------------------------
End Sub

Function LeastFullCell(QuotaQuestion, IOM, Punched)
	'LeastFullCell function returns the categorical name of the cell with the lowest
	'completion rate (complete count divided by target count) using a provided Quota question as
	'its basis. It is configurable for it to return the lowest cell from the entire Quota
	'category list, or only those that are punched in the response list (for the latter, Punched = True).
	'If more than 1 cell has the lowest percentage, then one from the list is randomly returned.
	'Makes use of separate GetCellName function

	Dim LowPercent	'Stores the 'benchmark' percent to which each cell compares its percent to. Type=Double
	Dim LowCell		'Stores the lowest cell (or cells, when multiple cells have the lowest value). Type=Categorical
	Dim obj			'Quota cell object used in iteration. Type=Quota
	Dim objName		'Categorical name of the cell. Type=Categorical
	Dim objPercent	'The COMPLETED percent of the current cell being evaluated. Type=Double
	Dim objTarget	'The target value of the cell being evaluated. Type=Long	
	
	'Set Low Percent to a high value that is usually unused
	LowPercent = cDouble(9999)
	
	'For each obj in the specified Quota Group, check its percentage value.
	'If it is lower than LowPercent, then set LowPercent to the new value and store which cell
	'it is. After all iterations, you will be left with the cell with lowest percentage
	If NOT (IOM.Info.IsDebug) Then
		For each obj in IOM.QuotaEngine.QuotaGroups[QuotaQuestion.QuestionName].Quotas
			'Ensure the Target of the quota cell is not 0
			if obj.Target=0 then
				objTarget=1
			else
				objTarget=obj.Target
			end if
			'Calculate the Completed percent by dividing the completed count by the target count
			objPercent = cDouble(obj.Completed)/cDouble(objTarget)
		
			If objPercent < LowPercent Then
				' Convert the current obj.name to its original categorical name (See GetCellName Function for How-To)
				objName = cCategorical(GetCellName(obj.Name))
			
				'Punched determines whether the function finds the lowest cell overall, or for only those
				'punched/marked in the response list of the quota question marker
				If Punched Then
					If QuotaQuestion.Response.ContainsAny(objName) Then
						LowPercent = objPercent
						LowCell = objName
					End If					
				Else
					LowPercent = objPercent
					LowCell = objName
				End If
				
			ElseIf objPercent = LowPercent Then
				' Convert the current obj.name to its original categorical name (See GetCellName Function for How-To)
				objName = cCategorical(GetCellName(obj.Name))
			
				'Punched determines whether the function finds the lowest cell overall, or for only those
				'punched/marked in the response list of the quota question marker
				If Punched Then
					If QuotaQuestion.Response.ContainsAny(objName) Then
						LowPercent = objPercent
						LowCell = LowCell + objName
					End If					
				Else
					LowPercent = objPercent
					LowCell = LowCell + objName
				End If
			End If
		Next
	Else
		If Punched Then
			LowCell = GetAnswer(IOM.Questions[QuotaQuestion.QuestionName].Response, 0)
		Else
			LowCell = GetAnswer(IOM.Questions[QuotaQuestion.QuestionName].Categories, 0)
		End If
	End If
	'Return Lowest Cell, if more than 1 cell had the lowest value - randomly return 1 of those values
	LeastFullCell = LowCell.Ran(1)		
End Function


 
Function GetCellName(objName)
	'GetCelName is used to convert the full name of a quota cell (from the quota engine)
	'into the categorical response off which it was originally based.
	'This is done by splitting it into an array. Take the last element of the
	'array and return all the text excluding the final character (which is a closing brace).
	'This is not a stand alone function - it is used in conjunction with other functions...
	Dim str
	
	str = Split(CText(objName), ".")
	str = str[Len(str) - 1]
	str = Left(str, Len(str) - 1)
	
	GetCellName = str
End Function

Function LeastFullCell_2D(QuotaQuestion_V, QuotaQuestion_H, QuotaName, IOM, Punched)


	Dim LowPercent	'Stores the 'benchmark' percent to which each cell compares its percent to. Type=Double
	Dim LowCell		'Stores the lowest cell (or cells, when multiple cells have the lowest value). Type=Categorical
	Dim obj			'Quota cell object used in iteration. Type=Quota
	Dim objName		'Categorical name of the cell. Type=Categorical
	Dim objPercent	'The COMPLETED percent of the current cell being evaluated. Type=Double
	Dim objTarget	'The target value of the cell being evaluated. Type=Long	
	
	'Set Low Percent to a high value that is usually unused
	LowPercent = cDouble(9999)
	
		
	'For each obj in the specified Quota Group, check its percentage value.
	'If it is lower than LowPercent, then set LowPercent to the new value and store which cell
	'it is. After all iterations, you will be left with the cell with lowest percentage
	If NOT (IOM.Info.IsDebug) Then
		For each obj in IOM.QuotaEngine.QuotaGroups[QuotaName].Quotas
		
			
			if ContainsAny(QuotaQuestion_H,cCategorical(GetCellName_H(obj.Name))) then
			
				'Ensure the Target of the quota cell is not 0
				if obj.Target=0 then
					objTarget=1
				else
					objTarget=obj.Target
				end if
				'Calculate the Completed percent by dividing the completed count by the target count
				objPercent = cDouble(obj.Completed)/cDouble(objTarget)
			
				'test1 = test1 + ctext(
			
				If objPercent < LowPercent Then
					' Convert the current obj.name to its original categorical name (See GetCellName Function for How-To)
					objName = cCategorical(GetCellName_V(obj.Name))
				
					'Punched determines whether the function finds the lowest cell overall, or for only those
					'punched/marked in the response list of the quota question marker
					If Punched Then
						If QuotaQuestion_V.Response.ContainsAny(objName) Then
							LowPercent = objPercent
							LowCell = objName
						End If					
					Else
						LowPercent = objPercent
						LowCell = objName
					End If
					
				ElseIf objPercent = LowPercent Then
					' Convert the current obj.name to its original categorical name (See GetCellName Function for How-To)
					objName = cCategorical(GetCellName_V(obj.Name))
				
					'Punched determines whether the function finds the lowest cell overall, or for only those
					'punched/marked in the response list of the quota question marker
					If Punched Then
						If QuotaQuestion_V.Response.ContainsAny(objName) Then
							LowPercent = objPercent
							LowCell = LowCell + objName
						End If					
					Else
						LowPercent = objPercent
						LowCell = LowCell + objName
					End If
				End If
			End If
		Next
	Else
		If Punched Then
			LowCell = GetAnswer(IOM.Questions[QuotaQuestion_V.QuestionName].Response, 0)
		Else
			LowCell = GetAnswer(IOM.Questions[QuotaQuestion_V.QuestionName].Categories, 0)
		End If
	End If
	'Return Lowest Cell, if more than 1 cell had the lowest value - randomly return 1 of those values
	LeastFullCell_2D = LowCell.Ran(1)		
End Function


Function GetCellName_H(objName)
	Dim str
	
	str = Split(CText(objName), ".")
	str = str[9]
	str = Left(str, Len(str) - 1)
	
	GetCellName_H = str
End Function

Function GetCellName_V(objName)
	Dim str
	
	str = Split(CText(objName), ".")
	str = str[4]
	str = Left(str, Len(str) - 1)
	
	GetCellName_V = str
End Function

function Grid_Att_Ran (GridQ,Seed,MRK_ORD,IOM)
	if Seed.Info.OffPathResponse = Null then
		If ISEMPTY(Seed) then Seed.Response = CText(GetRandomSeed())
	Else
		Seed.Response = Seed.Info.OffPathResponse
	End If
	
	SetRandomSeed(CLong(Seed.Response))
	GridQ.Categories.Order = OrderConstants.oRandomize
	
	Dim Cat, i
	i=0
	For Each Cat in GridQ.Categories
		i=i+1
		MRK_ORD[cat].item[0].response = i
	Next

	ShowTest(IOM,MRK_ORD,"")

end function

Function CatQ_Ran (Qname,order_var, Seed,IOM)
	if Seed.Info.OffPathResponse = Null then
		If ISEMPTY(Seed) then Seed.Response = CText(GetRandomSeed())
	Else
		Seed.Response = Seed.Info.OffPathResponse
	End If
	
	SetRandomSeed(CLong(Seed.Response))
	Qname.Categories.Order = OrderConstants.oRandomize
	
	Dim Cat, Cat2
	For Each Cat in Qname.Categories				
		if cat.CategoryType = CategoryTypes.ctCategoryList then
			For Each Cat2 in cat.Categories
				order_var = order_var + "," + Cat2.name	
			next
		Else
			order_var = order_var + "," + Cat.name	
		end if
	Next
	
 	ShowTest(IOM,order_var,"")
 	
end function

function Grid_Att_Rot (GridQ,Seed,MRK_ORD,IOM)
	if Seed.Info.OffPathResponse = Null then
		If ISEMPTY(Seed) then Seed.Response = CText(GetRotationSeed())
	Else
		Seed.Response = Seed.Info.OffPathResponse
	End If
	
	SetRotationSeed(CLong(Seed.Response))
	GridQ.Categories.Order = OrderConstants.oRotate
	
	Dim Cat, i
	i=0
	For Each Cat in GridQ.Categories
		i=i+1
		MRK_ORD[cat].item[0].response = i
	Next
	
	ShowTest(IOM,MRK_ORD,"")

end function

Function CatQ_Rot (Qname,order_var, Seed,IOM)
	if Seed.Info.OffPathResponse = Null then
		If ISEMPTY(Seed) then Seed.Response = CText(GetRotationSeed())
	Else
		Seed.Response = Seed.Info.OffPathResponse
	End If
	
	SetRotationSeed(CLong(Seed.Response))
	Qname.Categories.Order = OrderConstants.oRotate
	
	Dim Cat, Cat2
	For Each Cat in Qname.Categories				
		if cat.CategoryType = CategoryTypes.ctCategoryList then
			For Each Cat2 in cat.Categories
				order_var = order_var + "," + Cat2.name	
			next
		Else
			order_var = order_var + "," + Cat.name	
		end if
	Next
	
 	ShowTest(IOM,order_var,"")

end function

Function InnerRanCapture(Question,Order,Seed,IOM)
      
    if Seed.Info.OffPathResponse = Null then
	  Seed.Response = CText(GetRandomSeed())
    Else
	  Seed.Response = Seed.Info.OffPathResponse
    End If
    
    Dim cat
    
    For each cat in Question.Categories
	  SetRandomSeed(CLong(Seed.Response))
	  Question[cat].item[0].Categories.Order = OrderConstants.oRandomize
    Next
    
    Dim i
    i=0
    For Each Cat in Question[0].item[0].Categories
	  i=i+1
	  Order[cat].item[0].response = i
    Next

    ShowTest(IOM,Order,"")
    
End Function

Function InnerRotCapture(Question,Order,Seed,IOM)
      
    if Seed.Info.OffPathResponse = Null then
	  Seed.Response = CText(GetRotationSeed())
    Else
	  Seed.Response = Seed.Info.OffPathResponse
    End If
    
    Dim cat
    
    For each cat in Question.Categories
	  SetRotationseed(CLong(Seed.Response))
	  Question[cat].item[0].Categories.Order = OrderConstants.oRotate
    Next
    
    Dim i
    i=0
    For Each Cat in Question[0].item[0].Categories
	  i=i+1
	  Order[cat].item[0].response = i
    Next

    ShowTest(IOM,Order,"")
    
End Function

'--- Start CountVar Function

Sub CountVar(IOM)      

      Dim count, strMsg

      count =  IOM.MDM.Variables.count

      strMsg = "Your MDD currently exceeds the 10,000 instance variable limit. It contains XCOUNT Instance variables - It is imperative that programming implement the necessary modifications to lower instance variable count."      

      if count >= 10000 then

            strMsg = Replace(strMsg, "XCOUNT", count)

            IOM.Texts.InterviewStopped = strMsg

            if (IOM.Info.IsDebug) then Debug.MsgBox(strMsg)

            IOM.Texts.InterviewStopped = strMsg

            IOM.Terminate(Signals.sigStopped)

      elseif (count >= 7000 AND count < 10000) then

                  strMsg = "WARNING: Your MDD is closing approaching the 10,000 instance variable limit. It contains XCOUNT Instance variables - It is imperative that programming take the necessary steps to keep variable instance under 10,000."

            strMsg = Replace(strMsg, "XCOUNT", count)            

            if (IOM.Info.IsDebug) then Debug.MsgBox(strMsg)

      end if

 End Sub

'--- End CountVar Function

Function Insertion(Ques1,Ques2,IOM)

'Eg: Insertion(question1,question2)
' Where,
'	question1 = input question (Responses for pop-in will be fetch from this question)
'	question2 = Insertion question (Responses will be inserted in this question's label)
 
	Dim strFmt 
	
	strFmt = Ques1.Response.Label 
	
	strFmt = Replace(strFmt,",",", ") 
	
	strFmt = Replace(strFmt,",",IOM.Questions["prog_jsBanner"].Categories.ANDINS.Label,find(strFmt,",",,True),1) 
	
	Ques2.Label.Inserts[0] = strFmt 

End Function

'Function OtherValSetup(Question,value)
'	'EG. OtherValSetup(GRID_Q1,true)
'	'Value = True: This indicates, respondent has to enter/select response in all the sub questions 
'	'Value = False: This indicates, respondent can enter/select response in any of one sub question
'
'	If value = true then
'		Question.Validation.Function = "OtherValidation1"
'	Else
'		Question.Validation.Function = "OtherValidation2"
'	End If
'	
'End Function
'
'Function OtherValidation1(Question,IOM,Attempts)
'		Dim i
'		For i = 1 to Question[0].count - 1
'		
'			If Question[{_other}].item[i].Response <> null AND Question[{_other}].item[0].Response = null then
'				Question.Errors.AddNew("Error3", IOM.Questions.hCustomMsgs.Categories[{ERR13}].label)
'				OtherValidation1 = false
'				Exit function
'			End If
'		
'			If Question[{_other}].item[i].Response = null AND Question[{_other}].item[0].Response <> null then
'				Question.Errors.AddNew("Error4", IOM.Questions.hCustomMsgs.Categories[{ERR14}].label)
'				OtherValidation1 = false
'				Exit function
'			End If
'			
'		Next
'End Function
''
'Function OtherValidation2(Question,IOM,Attempts)
'		Dim i,count1
'		count1 = 0
'		
'		For i = 1 to Question[0].count - 1
'		
'			If Question[{_other}].item[i].Response = null AND Question[{_other}].item[0].Response <> null then ' other response blank
'			
'			Else
'				count1 = 1
'			End If
'			
'			If Question[{_other}].item[i].Response <> null AND Question[{_other}].item[0].Response = null then ' other text box blank
'				Question.Errors.AddNew("Error4", IOM.Questions.hCustomMsgs.Categories[{ERR13}].label)
'				OtherValidation2 = false
'				Exit function
'			End If
'						
'		Next
'		
'		If count1 = 0 then
'			Question.Errors.AddNew("error3", IOM.Questions.hCustomMsgs.Categories[{ERR14}].label)
'			OtherValidation2 = false
'			Exit function
'		End If
'		
'End Function

'--- Start CountVar Function

Function CatFilter(Question, Value)
  
    On Error Goto ErrorMsg
    
    dim Iter
    CatFilter = {}
    select case LCase(VarTypeName(Value))
        case "categorical", "text", "object"        
        case else
            debug.MsgBox("Error in CatFilter.  Invalid filter type, must be categorical or text categeorical expression. Question: " + Question.questionName + ". Value: " + CText(Value) + ". VarTypeName: " + VarTypeName(Value))
            CatFilter = null
            Exit Function
    end select

    for each Iter in Question       
        if Iter[0].ContainsAny(Value) then CatFilter = CatFilter + CCategorical(Iter.QuestionName)
    next
    Exit Function
ErrorMsg:
    debug.MsgBox(MakeString("Error in CatFilter.  Error Desc: ",  err.Description, ". Line: ",err.LineNumber, ". Parameter values: ", Question.QuestionName, ", ",Value))
    CatFilter = null
End Function

Function GRID_Formatting(IOM,Question,OverALLWidth,FirstColumnWidth,ColumnGRID)
'Question = GRID ID
'OverALLWidth = Over all width that needs to set for the question
'FirstColumnWidth = Width that needs to set for 1st column in display
'ColumnGRID = If the question is Column grid then "TRUE" else it should be "FALSE"
'Example for Column Orientation: GRID_Formatting(IOM,Question,"100%","20%",true)
'Example for Row Orientation: GRID_Formatting(IOM,Question,"100%","20%",false)

    If ColumnGRID = FALSE then
	    Question.Style.Width = OverALLWidth
	    Question.Categories[..].Label.Style.Cell.Width = FirstColumnWidth
	End If
	
	If ColumnGRID = TRUE then
	    Question.Style.Width = OverALLWidth
	    Question[..].item[0].Categories[..].Label.Style.Cell.Width = FirstColumnWidth
	End If
End Function

Function HC_NestedCategories(IOM,Question)
	Dim catQues,catQues1,NestedCat
	NestedCat = {}
	
	For each catQues in Question.categories
		If catQues.CategoryType = categorytypes.ctCategoryList then
			NestedCat = NestedCat + HC_NestedCategories(IOM,catQues)
		Else
			NestedCat = NestedCat + catQues
		End If
	Next
	
	HC_NestedCategories = CCategorical(NestedCat)

End Function


'------------------------End Of Function By Ugam-------------------------------------------------
'----------Functions proposed by IIS team (Ro competence)--------------------------------------------------------------------------------------
Sub SF_GridPrevAnsw(IOM, Grid1, Grid1Item, Grid2, Grid2Item)
      On Error Goto ErrorHandler
      Dim Element, QuestionCount, Iter
      QuestionCount = Grid2[0].Count 
      For each Element in Grid1.Categories
            'Update each Previous Answer(Grid1) into Current Grid(Grid2)
            If Grid2.Categories.ContainsAny(Element) = True Then
                  Grid2[Element].Item[Grid2Item].Response = Grid1[Element].Item[Grid1Item].Response
                  Grid2[Element].Item[Grid2Item].Style.Control.ReadOnly = TRUE            
            End If
      Next
      For each Element in Grid2.Categories
            'Update Width For All Elements
            For Iter = 0 to QuestionCount - 1
                  Grid2[Element].Item[Iter].Style.Width = "3em"
            Next
            'Hide those Previous Answers that do not get displayed
            If Grid1.Categories.ContainsAny(Element) = False Then
                  Grid2[Element].Item[Grid2Item].Style.Hidden = True
                  'Set MustAnswer to false to avoid error if left unanswered
                  Grid2[Element].Item[Grid2Item].MustAnswer = False
            End If
      Next
      Exit Sub

ErrorHandler:
      'Output error to logs
      IOM.Log("SF_GridPrevAnsw::Line: " + CText(Err.LineNumber) + " ERROR: " + CText(Err.Number) + " - " + CText(Err.Description), LOGLEVELS.LOGLEVEL_INFO)
      IOM.Questions["InterviewErrorMessage"].Response = CText(Err.Number) + " : " + Err.Description+ " : " + ctext(Err.LineNumber)
End Sub

Function SF_GridAnsw(qid, AnswerList)
	Dim Cat, CatArr, CatAnswerList
	
	CatArr = {}
	CatAnswerList = CCategorical("{" + AnswerList + "}")
	
	For Each Cat in qid.Categories
		If qid[Cat].Item[0].Response.ContainsAny(CatAnswerList) Then
			CatArr = CatArr + Cat
		End If
	Next
	
	SF_GridAnsw = CatArr
End Function

'**Copy the below entire section and paste at the respective location. Then update "Number_Of_Rows" and "Question_ID" details.
'Section SF_GridSplit
'	Dim startpos, NumRows, qid
'	
'	startpos = 0
'	NumRows = Number_Of_Rows 'How many rows to display per page
'	Set qid = Question_ID 'Set the name of the grid question here
'	
'	While (startpos < qid.Categories.Count)
'	    qid.Categories.Filter = qid.Categories.Mid(startpos, NumRows)
'	    qid.Ask()
'
'	    startpos = startpos + NumRows
'	    qid.Categories.Filter = NULL
'	End While
'End Section	

Function EMAILVAL(QEMAIL,IOM, Attempts)
	If SF_CheckPhone(QEMAIL, IOM) = FALSE then
		QEMAIL.Errors.AddNew("Error","Invalid PHONE NO")
		EMAILVAL = false
	end If
End Function

Function SF_CheckPhone(phone, IOM)
	On Error Goto ErrorHandler
	Dim isphone
	isphone = false
	If IsEmpty(phone) = True or (IsDBNull(phone) or phone = Null) Then
		SF_CheckPhone = False
		Exit Function
	End If	
	if Validate(phone,,,"\d{10}|\d{3}-\d{3}-\d{4}") then
		isphone = true
	end if
	SF_CheckPhone = isphone
	Exit Function
	
ErrorHandler:
	'Output error to logs
	SF_CheckPhone = Null
	IOM.Log("SF_CheckPhone::Line: " + CText(Err.LineNumber) + " ERROR: " + CText(Err.Number) + " - " + CText(Err.Description), LOGLEVELS.LOGLEVEL_INFO)
	
End Function

Function SF_CheckEmail(email, IOM)
	On Error Goto ErrorHandler
	Dim isemail
	isemail = false
	If IsEmpty(email) = True or (IsDBNull(email) or email = Null) Then
		SF_CheckEmail = False
		Exit Function
	End If
	if Validate(email,,,".+\@.+\..+") then
		isemail = true
	end if
	SF_CheckEmail = isemail
	Exit Function
	
ErrorHandler:
	'Output error to logs
	SF_CheckEmail = Null
	IOM.Log("SF_CheckEmail::Line: " + CText(Err.LineNumber) + " ERROR: " + CText(Err.Number) + " - " + CText(Err.Description), LOGLEVELS.LOGLEVEL_INFO)
	
End Function

Function SF_Rand(val, IOM)
     On Error Goto ErrorHandler
      dim min,max,dif,def
      if Validate(val,,,"\d{1,10}\.\.\d{1,10}|\.\.\d{1,10}|\d{1,10}\.\.|\.\.") then
            val = val.split("..")
            if val[0] <> null then
                  min = clong(val[0])
            else
                  min = 0
            end if
            if val[1] <> null then
                  max = clong(val[1])
            else
                  max = 1000000
            end if
            if min > max then
                  dif = min
                  min = max
                  max = dif
            end if
            If Max = Min Then
                  SF_Rand = min
            Else
                  dif = max - min
                  def = clong(Rnd()*dif*1000)
                  def = def mod dif
                  def = def + min
                  SF_Rand = def                 
            End If
      else
            SF_Rand = -1
      end if
      Exit Function

ErrorHandler:
      'Output error to logs
      IOM.Log("SF_Rand::Line: " + CText(Err.LineNumber) + " ERROR: " + CText(Err.Number) + " - " + CText(Err.Description), LOGLEVELS.LOGLEVEL_INFO)
      SF_Rand = Null    
End Function 

'---------------------End if IIS functions

''''''''''''''''''''' Start - Jayesh - Functions added for checking '''''''''''''''''''''

Function FormatQuestion(Question, IOM)
Dim Html, Iteration, SubQuestion


   Select Case Question.QuestionType
      Case QuestionTypes.qtSimple
         ' A simple question. Show question text and response
         If Question.Info.IsOnPath Then
            Html = Html + "<tr>"
            Html = Html + "<td>" + BoldItalic(Question.QuestionName) + ": " + Italicise(Question.Label.Text) + "</td>"
            Html = Html + FormatResponse(Question)
            Html = Html + "</tr>"
         End If

      Case QuestionTypes.qtLoopCategorical, QuestionTypes.qtLoopNumeric
         ' Question is a loop. Format all iterations of the loop
         For Each Iteration In Question
            Html = Html + FormatQuestion(Iteration, IOM)
         Next

         ' If there are any responses to the questions of the
         ' loop then show the loop heading.  Also show a break
         ' before loop responses (in order to group them)
         If Html <> "" Then
            Html = "<tr><td>" + BoldItalic(Question.QuestionName) + ": " + Italicise(Question.Label.Text) + "</td></tr>" + Html
         End If

      Case QuestionTypes.qtBlock, QuestionTypes.qtCompound, QuestionTypes.qtPage
         ' Build up responses for all questions in the block/page/compound question.
         For Each SubQuestion In Question
            Html = Html + FormatQuestion(SubQuestion, IOM)
         Next
   End Select

   FormatQuestion = Html
End Function

Function FormatResponse(Question)
   ' Format the response for a simple question
   If (Question.QuestionDataType = DataTypeConstants.mtCategorical) Then
      FormatResponse = "<td>" + Question.Response.Label + "</td>"
   ElseIf (Question.QuestionDataType <> DataTypeConstants.mtNone) Then
      FormatResponse = "<td>" + CText(Question.Response.Value) + "</td>"
 End If
End Function

Function Italicise(qlabel)
   Italicise = "<i>" + qlabel + "</i>" 
End Function

Function BoldItalic(qlabel)
   BoldItalic = "<b><i>" + qlabel + "</i></b>" 
End Function

Function CheckExclusive(Question, IOM)

	Dim Iteration, SubQuestion

	Select Case Question.QuestionType
      Case QuestionTypes.qtSimple
			If (Question.QuestionDataType = DataTypeConstants.mtCategorical) Then
				AddExclusiveCheck(Question, IOM)
			End If

      Case QuestionTypes.qtLoopCategorical, QuestionTypes.qtLoopNumeric

	         For Each Iteration In Question
				CheckExclusive(Iteration, IOM)
	         Next

		Case QuestionTypes.qtBlock, QuestionTypes.qtCompound, QuestionTypes.qtPage
			For Each SubQuestion In Question
				CheckExclusive(SubQuestion, IOM)
''				If (SubQuestion.QuestionDataType = DataTypeConstants.mtCategorical) Then
'					CheckExclusive(IOM,SubQuestion)
''				End If
         	Next
	End Select
	
End Function

Function AddExclusiveCheck(Question, IOM)

	Dim Cat, counter, ban, code

	counter = 1
	For Each Cat in Question.Categories

		ban = "exclusiveCheck" + CText(counter)
		RemoveBanner(IOM, IOM, ban)

		If BitAnd(Cat.Attributes, CategoryAttributes.caExclusive) Then
			IOM.Banners.AddNew(ban ,"<SCRIPT>$(document).ready(function(){$.updateExclusive(""" +cat.Name+ """);})</SCRIPT>")
		End if

		counter = counter + 1

	Next
	 
End Function

Sub RemoveBanner(IOM, Q, BannerName)
      Dim i
      For i=0 to Q.Banners.Count - 1          
            If Ucase(Q.Banners.Item[i].Name) = Ucase(trim(BannerName)) Then
                  Q.Banners.Remove(BannerName)
                  Exit For
            End if
      Next
End Sub

''''''''''''''''''''' End - Jayesh - Functions added for checking '''''''''''''''''''''


Sub SendMail(IOM,toAddress,fromAddress,SubjectMessage,Message)
	If NOT IOM.Info.IsDebug then
		Dim myMail
		set myMail = CreateObject("CDO.Message")
		myMail.From = fromAddress
		'myMail.Sender= fromAddress
		myMail.To = toAddress
		myMail.BodyPart.Charset = "utf-8"
		myMail.Subject  = SubjectMessage
		myMail.HTMLBody = Message
		myMail.Fields.Update()
		myMail.Configuration.Fields.Item["http://schemas.microsoft.com/cdo/configuration/smtpserver"] = "pmta.extranet.iext" 'This is Power MTA IP
		myMail.Configuration.Fields.Item["http://schemas.microsoft.com/cdo/configuration/sendusing"] = 2
		
		myMail.Configuration.Fields.Item["http://schemas.microsoft.com/cdo/configuration/smtpserverport"] = 25
		myMail.Configuration.Fields.Item["http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout"] = 60
		
		myMail.Configuration.Fields.Update()
		myMail.Send()
		set myMail = null
	End If
End sub

Sub QUITSTAT(IOM)
	
	IOM.Questions.QUITSTAT.Response = {RegularQuit}
	
	If IOM.Questions.InterviewLengthScreener.Info.OffPathResponse = NULL then
		IOM.Questions.InterviewLengthScreener.Response = IOM.Questions.InterviewLength.Response
	Else
		IOM.Questions.InterviewLengthScreener.Response = IOM.Questions.InterviewLengthScreener.Info.OffPathResponse
	End If

End Sub

function jsCategoricalHeader(Question,header1,header2,Iscategorical,IOM)
	
	''Iscategorical = true for categorical grid questions and false for numeric grid questions.
	
	
	
	dim fetch,insert,mainheader
	
	if Iscategorical = true then
	
		mainheader = "<script src='https://ajax.googleapis.com/ajax/libs/jquery/1.10.2/jquery.min.js'></script><script> $(document).ready(function(){ 	$('.mrQuestionTable').each(function(index){	$(this).find('td:first').addClass('ugHead');if (index == 0){$(this).find('td:first').append('{Heading1}');}if (index == 1) { 	$(this).find('td:first').append('{Heading2}');}});}); </script>"
		
		mainheader = replace(mainheader,"{Heading1}",IOM.Questions["prog_jsBanner"].Categories[header1].Label)
		mainheader = replace(mainheader,"{Heading2}",IOM.Questions["prog_jsBanner"].Categories[header2].Label)
		
	end if
	
	if Iscategorical = false then
		mainheader = "<script src='https://ajax.googleapis.com/ajax/libs/jquery/1.10.2/jquery.min.js'></script> <script> $(document).ready(function(){ var addrow = '<tr><td></td><td></td></tr>'; $('.mrQuestionTable').find('tr:first').before(addrow); $('.mrQuestionTable').find('td:first').addClass('ugHead'); $('.mrQuestionTable').find('td:first').append('{Heading1}'); $('.mrQuestionTable').find('td:first').next().addClass('ugHead'); $('.mrQuestionTable').find('td:first').next().append('{Heading2}'); }); </script>"
		
		mainheader = replace(mainheader,"{Heading1}",IOM.Questions["prog_jsBanner"].Categories[header1].Label)
		mainheader = replace(mainheader,"{Heading2}",IOM.Questions["prog_jsBanner"].Categories[header2].Label)
	end if
		
		
	Question.Banners.AddNew("updatedheader",mainheader)	
	
end function 


Sub straightline (IOM,Qname,innQname,IDinMRK)

	'Qname - pass the question name 
	'innQname - pass the inner question name between "" as "Q54"
	'IDinMRK - Pass the category id created in marker MRK_Straightline for e.g. {Q54}
	
		Dim catcount, categorylist, MRK_Straightline, cntans, QAllCatCount
		
		Set MRK_Straightline = IOM.Questions["MRK_Straightline"]
		catcount = 0
		cntans = 0
		QAllCatCount = 0
		For each categorylist in Qname
			If Answercount(categorylist.Item[innQname]) > 0  Then
				cntans = cntans + 1
			End if
		Next
		
		If cntans > 0 Then
			For each categorylist in Qname
				If Qname[0].Item[innQname] = categorylist.Item[innQname] Then
					catcount = catcount + 1
				End if
				QAllCatCount = QAllCatCount + 1
			Next
			if catcount = QAllCatCount And QAllCatCount > 2 Then
				MRK_Straightline.Response = MRK_Straightline.Response + IDinMRK
			End if
			
			If IOM.Info.IsTest Then
				AddProgress(IOM,-1)
				'MRK_Straightline.Show()
			End if
		Else
			MRK_Straightline.Categories = MRK_Straightline.Categories - IDinMRK
		End if
	
      
End Sub


Sub mark_StrtLine (IOM)

	Dim Mrk_SLineFail, MRK_Straightline
	
	Set MRK_Straightline = IOM.Questions["MRK_Straightline"]
	Set Mrk_SLineFail = IOM.Questions["Mrk_SLineFail"]
	
	If CDouble(((Answercount(MRK_Straightline) * 100)/MRK_Straightline.Categories.Count)) >= 66 Then
		Mrk_SLineFail = {StraightFail}
	Else
		Mrk_SLineFail = {StraightPass}
	End if
	
	ShowTest(IOM,Mrk_SLineFail,"")
      
End Sub  

Sub IPAddressCap(IOM)

	Dim RestartCount, EncryptedFrist, Encryptedlast, i
	Set RestartCount = IOM.Questions["RestartCount"]
	
	IF IOM.Questions["SampleFields"].gop.Response <> "" then
		
		EncryptedFrist = "X"
		For i = 0 to Len(IOM.Questions["SampleFields"].gop.Response) - 1
			EncryptedFrist = EncryptedFrist + Format(AscW(Mid(IOM.Questions["SampleFields"].gop.Response,i,1)),"d3")
		Next

		EncryptedLast = "X"
		For i = 0 to Len(IOM.Questions["SampleFields"].gop_last.Response) - 1
			EncryptedLast = EncryptedLast + Format(AscW(Mid(IOM.Questions["SampleFields"].gop_last.Response,i,1)),"d3")
		Next

		If RestartCount = 0 Then
			IOM.Questions["IPADDRESSCAP"].Response = EncryptedFrist
		Else
			IOM.Questions["IPADDRESSCAP_Last"].Response = EncryptedLast
		End If
	End If
	
End Sub

Sub getDeviceType(theBrows,theCatQ)    

      if lcase(theBrows.platform) = "ipad" then
            theCatQ = {iPad}
      elseif lcase(theBrows.platform) = "ipod" then
            theCatQ = {iPod}
      elseif lcase(theBrows.platform) = "iphone" then
            theCatQ = {iPhone}
      elseif lcase(theBrows.platform) = "win32" or lcase(theBrows.platform) = "win64" then
            if find(lcase(theBrows.fullua),"iemobile") <> -1 then
                  if find(lcase(theBrows.fullua),"phone") <> -1 then                
                        theCatQ = {WinPhone}
                  else
                       theCatQ = {WinMobileOther}
                  end if            
            elseif find(lcase(theBrows.fullua),"mobile") <> -1 then
                  theCatQ = {WinMobileOther}
            else
                  theCatQ = {WinPC}
            end if
      elseif lcase(theBrows.platform) = "macppc" or lcase(theBrows.platform) = "macintel" then
            theCatQ = {Mac}
      elseif lcase(theBrows.platform) = "blackberry" then
            theCatQ = {BlackBerry}
      else
            if find(lcase(theBrows.fullua),"silk") <> -1 or find(lcase(theBrows.Version),"silk") <> -1 then
                  theCatQ = {KindleFireTablet}
            elseif find(lcase(theBrows.fullua),"nook") <> -1 then
                  theCatQ = {NookTablet}
            elseif find(lcase(theBrows.fullua),"tablet") <> -1 then
                  theCatQ = {TabletOther}
            elseif find(lcase(theBrows.fullua),"android") <> -1 then
                  if find(lcase(theBrows.fullua),"mobile") <> -1 then
                       theCatQ = {AndroidPhone}
                  else
                        theCatQ = {AndroidTablet}
                  end if            
            else
                  theCatQ = {others}
            end if           
      end if

End Sub

Function GetCategoriesFromGrid(iGrid, Response)
	Dim Ques, SubCat	
	GetCategoriesFromGrid={}
	For Each Ques In iGrid
		If Ques[0].ContainsAny(CCategorical(Response)) Then
			GetCategoriesFromGrid = GetCategoriesFromGrid + CCategorical(Ques.QuestionName)
		End If
	Next
End Function

Sub MoveCounts(IOM,strQuota,cell_name)
		
	IOM.QuotaEngine.QuotaGroups[strQuota].Quotas[cell_name].Completed = IOM.QuotaEngine.QuotaGroups[strQuota].Quotas[cell_name].Completed - 1
	
End Sub

Function FnCheckColGridTotal(Question, IOM, Attempt)
    Dim tot1,i,cat1,cat2, DUMTOTAL, innerQ
    
    Set DUMTOTAL = IOM.Questions["DUMTOTAL"]
    Set innerQ = Question[0].Item[0]
    
    i=1
    For each cat1 in Question
        i=i+1
        tot1 = 0
        for each cat2 in cat1.Item[0]
            tot1 = tot1 + cat2.Item[0]
        next
                
        If tot1 <> DUMTOTAL Then
            IOM.Questions["sysTemp2"].Response = tot1
            IOM.Questions["sysTemp1"].Response = Question.Categories[CCategorical(cat1.QuestionName)].Label
            IOM.Questions["sysTemp3"].Response = DUMTOTAL
            Question.Errors.AddNew("err"+CText(i),IOM.Questions["sysErrorMessages"].Categories.SYS_ERR_EQ.Label)
            FnCheckColGridTotal = FALSE
            'Exit Function
        End If
        
    Next

End Function

'==========================================
'Start Time and End Time Capturing Code
Sub StartTimeCapture(QStartTime,IOM)
	If QStartTime.Info.OffPathResponse <> Null then
	  QStartTime = QStartTime.Info.OffPathResponse
	Else 
	  QStartTime.Response = Now()
	End if
End Sub

Sub EndTimeCapture(QStartTime,QEndTime,QTotalTime,IOM)
	If QTotalTime.Info.OffPathResponse <> Null then 
		QEndTime = QEndTime.Info.OffPathResponse 
	Else
		QEndTime.Response = Now()
	End if  
	QTotalTime = Clong(DateDiff(QStartTime,QEndTime,"n"))
	ShowTest(IOM,QTotalTime,"")
End Sub
'==========================================

'==========================================
'IDs removal Sub procedure
Sub IDsRemoval(IOM,QuotaName,QuotaType,Variable1,Variable2,Variable3,Variable4)

	'QuotaName --> Quota name used in MQD
	'QuotaType --> Type of quota as given below
					'"1:0" --> Only 1 Variable at left side
					'"0:1" --> Only 1 Variable at Top side
					'"1:1" --> 1 variable at Left and 1 Variable at Top
					'"2:1" --> 2 variable at Left and 1 Variable at Top
					'"1:2" --> 1 variable at Left and 2 Variable at Top
					'"2:2" --> 2 variable at Left and 2 Variable at Top
	'Variable1/Variable2/Variable3/Variable4 --> This will be depends on QuotaType which you have created, Always mention left position variables then top position variables
	

	Dim cell_name,Var1Position,iVar1,Var2Position,iVar2,Var3Position,iVar3,Var4Position,iVar4
	
	'Variable1 Category position picking
	If Variable1 <> "" Then
		Var1Position = 0
		Variable1.Categories.Order = OrderConstants.oNormal
		For each iVar1 in Variable1.Categories
			If Containsany(Variable1,iVar1) then Exit For
			Var1Position = Var1Position + 1
		Next
	End if
	
	'Variable2 Category position picking
	If Variable2 <> "" Then
		Var2Position = 0
		Variable2.Categories.Order = OrderConstants.oNormal
		For each iVar2 in Variable2.Categories
			If Containsany(Variable2,iVar2) then Exit For
			Var2Position = Var2Position + 1	
		Next
	End if
	
	'Variable3 Category position picking
	If Variable3 <> "" Then
		Var3Position = 0
		Variable3.Categories.Order = OrderConstants.oNormal
		For each iVar3 in Variable3.Categories
			If Containsany(Variable3,iVar3) then Exit For
			Var3Position = Var3Position + 1	
		Next
	End if
	
	'Variable4 Category position picking
	If Variable4 <> "" Then
		Var4Position = 0
		Variable4.Categories.Order = OrderConstants.oNormal
		For each iVar4 in Variable4.Categories
			If Containsany(Variable4,iVar4) then Exit For
			Var4Position = Var4Position + 1
		Next
	End if
	
	''Single Variable Quota (1-Left)
	If QuotaType = "1:0" Then
		cell_name = Makestring("Side.(0.",Variable1.QuestionName,").(",Var1Position,".",Variable1.Categories[Variable1].Name,")")
	End if
	
	''Single Variable Quota (1-Top)
	If QuotaType = "0:1" Then
		cell_name = Makestring("Top.(0.",Variable1.QuestionName,").(",Var1Position,".",Variable1.Categories[Variable1].Name,")")
	End if
	
	''Table Quota (1-Left and 1-Top)
	If QuotaType = "1:1" Then
		cell_name = MakeString("Side.(0.",Variable1.QuestionName,").(",Var1Position,".",Variable1.Categories[Variable1].Name,").Top.(0.",Variable2.QuestionName,").(",Var2Position,".",Variable2.Categories[Variable2].Name,")")
	End if
	
	''Table Quota (1-Left and 2-Top)
	If QuotaType = "1:2" Then
		cell_name = MakeString("Side.(0.",Variable1.QuestionName,").(",Var1Position,".",Variable1.Categories[Variable1].Name,").Top.(0.",Variable2.QuestionName,").(",Var2Position,".",Variable2.Categories[Variable2].Name,")",".(0.",Variable3.QuestionName,").(",Var3Position,".",Variable3.Categories[Variable3].Name,")")
	End if
	
	''Table Quota (2-Left and 1-Top)
	If QuotaType = "2:1" Then
		cell_name = MakeString("Side.(0.",Variable1.QuestionName,").(",Var1Position,".",Variable1.Categories[Variable1].Name,").(0.",Variable2.QuestionName,").(",Var2Position,".",Variable2.Categories[Variable2].Name,").Top.(0.",Variable3.QuestionName,").(",Var3Position,".",Variable3.Categories[Variable3].Name,")")
	End if
	
	''Table Quota (2-Left and 2-Top)
	If QuotaType = "2:2" Then
		cell_name = MakeString("Side.(0.",Variable1.QuestionName,").(",Var1Position,".",Variable1.Categories[Variable1].Name,").(0.",Variable2.QuestionName,").(",Var2Position,".",Variable2.Categories[Variable2].Name,").Top.(0.",Variable3.QuestionName,").(",Var3Position,".",Variable3.Categories[Variable3].Name,").(0.",Variable4.QuestionName,").(",Var4Position,".",Variable4.Categories[Variable4].Name,")")
	End if
	
	'debug.MsgBox(cell_name)
	
	If NOT IOM.Info.IsDebug Then
		MoveCounts(IOM,QuotaName,cell_name)
	End if

End Sub
'==========================================

function WordCount(Ques, IOM)
      
      'Works with Comma, Full stop, Space, Horizontal Tab, Line Feed,Enter
      
      dim CountChar, count
      count=0
      
      For CountChar = 0 to len(Ques)
            If ((ascw(mid(Ques,CountChar))=44 or ascw(mid(Ques,CountChar))=46 or ascw(mid(Ques,CountChar))=32 or ascw(mid(Ques,CountChar))=9 or ascw(mid(Ques,CountChar))=10 or (mid(Ques,CountChar))=null) and _
            ((ascw(mid(Ques,CountChar-1))>=48 and ascw(mid(Ques,CountChar-1))<=57) or _
            (ascw(mid(Ques,CountChar-1))>=65 and ascw(mid(Ques,CountChar-1))<=90) or _
            (ascw(mid(Ques,CountChar-1))>=97 and ascw(mid(Ques,CountChar-1))<=122))) then
                  Count=Count+1
            End If
      Next
      
      WordCount = Count
      
end function
'==========================================

End Routing