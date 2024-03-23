' // creator of the script is MEZYKPA
' // in case of any malfunction, please refer with the script name to MEZYKPA and the error snapshot
' // 
' // the script serves the purpose of download flat files from SAP via scripting tool built in SAP
' // the flat files are later used in Power BI report 'Monitoring 2024'

Dim todaysDate, targetDate, dayOfWeek, strUserName, connection, application, SapGuiAuto, objNetwork, objShell, session, objExcel, objWorkbook

Set objNetwork = CreateObject("WScript.Network")
Set objShell = CreateObject("WScript.Shell")

' Get the current time
currentTime = Time()

' Define the start and end time of PA1-PA6 LT and ST
startTime1 = #06:00:00 AM#
endTime1 = #11:28:59 PM#
'ST 3wk
startTime2 = #11:30:00 AM#
endTime2 = #12:43:00 PM#
'job check starting time
startTime3 = #06:00:00 AM#
endTime3 = #11:28:59 PM#

'job check starting time
startTime4 = #06:00:00 AM#
endTime4 = #12:43:00 PM#

strUserName = objNetwork.UserName
strPathwayPA1 = "\OneDrive - Danone\General - EU-IT&DATA HUB D2D - PLA\50 Support team\60. Monitoring\90. New Job Checks\PA1 Job Checks"
strPathwayPA3 = "\OneDrive - Danone\General - EU-IT&DATA HUB D2D - PLA\50 Support team\60. Monitoring\90. New Job Checks\PA3 Job Checks"
strPathwayPA6 = "\OneDrive - Danone\General - EU-IT&DATA HUB D2D - PLA\50 Support team\60. Monitoring\90. New Job Checks\PA6 Job Checks"
strPathwayP02 = "\OneDrive - Danone\General - EU-IT&DATA HUB D2D - PLA\50 Support team\60. Monitoring\90. New Job Checks\P02 Job Checks"
strPathwayP06 = "\OneDrive - Danone\General - EU-IT&DATA HUB D2D - PLA\50 Support team\60. Monitoring\90. New Job Checks\P06 Job Checks"
strPathwayP01 = "\OneDrive - Danone\General - EU-IT&DATA HUB D2D - PLA\50 Support team\60. Monitoring\90. New Job Checks\P01 Job Checks"
strPathwayFcst = "\OneDrive - Danone\General - EU-IT&DATA HUB D2D - PLA\50 Support team\60. Monitoring\90. New Job Checks\Forecast Check"

objShell.Run """C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"""

WScript.Sleep 1500

If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If

WScript.Sleep 1500

' Get today's date
todaysDate = Date

' Determine the day of the week (Sunday = 1, Monday = 2, etc.)
dayOfWeek = Weekday(todaysDate)

If dayOfWeek = 2 Then ' If it's Monday
    ' Set to the previous Friday
    targetDate = DateAdd("d", -3, todaysDate)
Else
    ' Set to the previous day
    targetDate = DateAdd("d", -1, todaysDate)
End If


VariousDate = Day(targetDate) & "." & Month(targetDate) & "." & Year(targetDate)

if (currentTime >= startTime4) And (currentTime <= endTime4) Then 'if time between 6AM and 1 PM then opens PA1
   Set connection = application.OpenConnection("APO Prod PA1 (EMEA & RUCIS EDP)", True)
   If Not IsObject(connection) Then
      Set connection = application.Children(0)
   End If
   If Not IsObject(session) Then
      Set session    = connection.Children(0)
   End If
   If IsObject(WScript) Then
      WScript.ConnectObject session,     "on"
      WScript.ConnectObject application, "on"
   End If
   session.findById("wnd[0]").maximize

         ' fcst check PA1 long term and short term morning if current time is between 6 AM and 11 AM
   If (currentTime >= startTime1) And (currentTime <= endTime1) Then
         session.findById("wnd[0]/tbar[0]/btn[3]").press
         session.findById("wnd[0]/tbar[0]/okcd").text = "listcube"
         session.findById("wnd[0]").sendVKey 0
         session.findById("wnd[0]/usr/ctxtP_DTA").text = "ibplflr"
         session.findById("wnd[0]/usr/ctxtP_DTA").caretPosition = 7
         session.findById("wnd[0]").sendVKey 0
         session.findById("wnd[0]").sendVKey 8
         session.findById("wnd[0]/usr/chkL_AG").selected = false
         session.findById("wnd[0]/usr/chkL_NO").selected = true
         session.findById("wnd[0]/usr/txtL_MX").text = ""
         session.findById("wnd[0]/usr/txtL_MX").setFocus
         session.findById("wnd[0]/usr/txtL_MX").caretPosition = 11
         session.findById("wnd[0]/tbar[1]/btn[25]").press
         session.findById("wnd[0]/usr/chkS003").selected = true
         session.findById("wnd[0]/usr/chkS006").selected = true
         session.findById("wnd[0]/usr/chkS008").selected = true
         session.findById("wnd[0]/usr/chkS008").setFocus
         session.findById("wnd[0]/tbar[1]/btn[8]").press
         session.findById("wnd[0]/tbar[1]/btn[8]").press
         session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").select
         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "LT PA1 Fcst.txt"
         session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\" & strUserName & strPathwayFcst
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         session.findById("wnd[1]/tbar[0]/btn[11]").press
         session.findById("wnd[0]/tbar[0]/btn[3]").press
         session.findById("wnd[0]/tbar[0]/btn[3]").press
         session.findById("wnd[0]/usr/ctxtP_DTA").text = "ibplfsr"
         session.findById("wnd[0]").sendVKey 0
         session.findById("wnd[0]/tbar[1]/btn[8]").press
         session.findById("wnd[0]/usr/chkL_AG").selected = false
         session.findById("wnd[0]/usr/chkL_NO").selected = true
         session.findById("wnd[0]/usr/txtL_MX").text = ""
         session.findById("wnd[0]/usr/chkL_NO").setFocus
         session.findById("wnd[0]/tbar[1]/btn[25]").press
         session.findById("wnd[0]/usr/chkS002").selected = true
         session.findById("wnd[0]/usr/chkS004").selected = true
         session.findById("wnd[0]/usr/chkS007").selected = true
         session.findById("wnd[0]/usr/chkS009").selected = true
         session.findById("wnd[0]/usr/chkS009").setFocus
         session.findById("wnd[0]/tbar[1]/btn[8]").press
         session.findById("wnd[0]/tbar[1]/btn[8]").press
         session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").select
         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ST PA1 Fcst.txt"
         session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\" & strUserName & strPathwayFcst
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         session.findById("wnd[1]/tbar[0]/btn[11]").press
         session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
         session.findById("wnd[0]").sendVKey 0
   End If
         ' fcst check PA1 short term midday if current time is between 11:30 AM and 12:40 PM (1h10m after morning release time)
   If (currentTime >= startTime2) And (currentTime <= endTime2) Then
         session.findById("wnd[0]/tbar[0]/btn[3]").press
         session.findById("wnd[0]/tbar[0]/okcd").text = "listcube"
         session.findById("wnd[0]").sendVKey 0
         session.findById("wnd[0]/usr/ctxtP_DTA").text = "ibplfsr"
         session.findById("wnd[0]").sendVKey 0
         session.findById("wnd[0]/tbar[1]/btn[8]").press
         session.findById("wnd[0]/usr/chkL_AG").selected = false
         session.findById("wnd[0]/usr/chkL_NO").selected = true
         session.findById("wnd[0]/usr/txtL_MX").text = ""
         session.findById("wnd[0]/usr/chkL_NO").setFocus
         session.findById("wnd[0]/tbar[1]/btn[25]").press
         session.findById("wnd[0]/usr/chkS002").selected = true
         session.findById("wnd[0]/usr/chkS004").selected = true
         session.findById("wnd[0]/usr/chkS007").selected = true
         session.findById("wnd[0]/usr/chkS009").selected = true
         session.findById("wnd[0]/usr/chkS009").setFocus
         session.findById("wnd[0]/tbar[1]/btn[8]").press
         session.findById("wnd[0]/tbar[1]/btn[8]").press
         session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").select
         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ST PA1 3W Fcst.txt"
         session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\" & strUserName & strPathwayFcst
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         session.findById("wnd[1]/tbar[0]/btn[11]").press
         session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
         session.findById("wnd[0]").sendVKey 0
   End if
         'job run in PA1
   if (currentTime >= startTime3) And (currentTime <= endTime3) Then
         session.findById("wnd[0]/tbar[0]/okcd").text = "sm37"
         session.findById("wnd[0]").sendVKey 0
         session.findById("wnd[0]/usr/txtBTCH2170-JOBNAME").text = "*SCM*"
         session.findById("wnd[0]/usr/txtBTCH2170-USERNAME").text = "*"
         session.findById("wnd[0]/usr/ctxtBTCH2170-FROM_DATE").text = VariousDate
         session.findById("wnd[0]/usr/txtBTCH2170-ABAPNAME").setFocus
         session.findById("wnd[0]/usr/txtBTCH2170-ABAPNAME").caretPosition = 0
         session.findById("wnd[0]/tbar[1]/btn[8]").press
         session.findById("wnd[0]/mbar/menu[3]/menu[0]/menu[2]").select
         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "JOBS PA1 SCM.txt"
         session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\" & strUserName & strPathwayPA1
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         session.findById("wnd[1]/tbar[0]/btn[11]").press
         session.findById("wnd[0]/tbar[0]/btn[3]").press
         session.findById("wnd[0]/usr/txtBTCH2170-JOBNAME").text = "PLOB_DELTA_SYNC_D0002_PO"
         session.findById("wnd[0]/usr/txtBTCH2170-USERNAME").text = "*"
         session.findById("wnd[0]/usr/txtBTCH2170-JOBNAME").caretPosition = 4
         session.findById("wnd[0]").sendVKey 8
         session.findById("wnd[0]/mbar/menu[3]/menu[0]/menu[2]").select
         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "JOBS PA1 PLOB_DELTA_SYNC_D0002_PO.txt"
         session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\" & strUserName & strPathwayPA1
         session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 12
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         session.findById("wnd[1]/tbar[0]/btn[11]").press
         session.findById("wnd[0]/tbar[0]/btn[3]").press
         session.findById("wnd[0]/usr/txtBTCH2170-JOBNAME").text = "*APODP*00*-*"
         session.findById("wnd[0]/usr/txtBTCH2170-USERNAME").text = "*"
         session.findById("wnd[0]/usr/txtBTCH2170-JOBNAME").caretPosition = 4
         session.findById("wnd[0]").sendVKey 8
         session.findById("wnd[0]/mbar/menu[3]/menu[0]/menu[2]").select
         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "JOBS PA1 APODP 00 -.txt"
         session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\" & strUserName & strPathwayPA1
         session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 12
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         session.findById("wnd[1]/tbar[0]/btn[11]").press
         session.findById("wnd[0]/tbar[0]/btn[3]").press
         session.findById("wnd[0]/usr/txtBTCH2170-JOBNAME").text = "*PL1SCMDEXTRCT-02*"
         session.findById("wnd[0]/usr/txtBTCH2170-USERNAME").text = "*"
         session.findById("wnd[0]/usr/txtBTCH2170-JOBNAME").caretPosition = 4
         session.findById("wnd[0]").sendVKey 8
         session.findById("wnd[0]/mbar/menu[3]/menu[0]/menu[2]").select
         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "JOBS PA1 PL1SCMDEXTRCT-02.txt"
         session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\" & strUserName & strPathwayPA1
         session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 12
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         session.findById("wnd[1]/tbar[0]/btn[11]").press
         session.findById("wnd[0]/tbar[0]/btn[3]").press

         If dayOfWeek = 2 Then
               ' Code to run only on Mondays
               session.findById("wnd[0]/usr/txtBTCH2170-JOBNAME").text = "FR1*TVARV*"
               session.findById("wnd[0]/usr/txtBTCH2170-USERNAME").text = "*"
               session.findById("wnd[0]/usr/txtBTCH2170-JOBNAME").caretPosition = 4
               session.findById("wnd[0]").sendVKey 8
               session.findById("wnd[0]/mbar/menu[3]/menu[0]/menu[2]").select
               session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
               session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
               session.findById("wnd[1]/tbar[0]/btn[0]").press
               session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "JOBS PA1 FR1 TVARV.txt"
               session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\" & strUserName & strPathwayPA1
               session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 12
               session.findById("wnd[1]/tbar[0]/btn[0]").press
               session.findById("wnd[1]/tbar[0]/btn[11]").press
               session.findById("wnd[0]/tbar[0]/btn[3]").press
         End If
   End if
   session.findById("wnd[0]/tbar[0]/btn[3]").press
   session.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
   session.findById("wnd[0]").sendVKey 0
End if

'go to PA3
WScript.Sleep 1500

If (currentTime >= startTime1) And (currentTime <= endTime1) Then 'if time is 6 am to 11:30 then it will be run
   Set connection = application.OpenConnection("APO Prod PA3 (waters)", True)
   Set session = connection.Children(0)

   WScript.Sleep 1500

   if dayOfWeek = 2 Then
      ' fcst check PA3 long term morning if current time is between 6 AM and 11:30 AM and it's monday
      If (currentTime >= startTime1) And (currentTime <= endTime1) Then
         session.findById("wnd[0]/tbar[0]/btn[3]").press
         session.findById("wnd[0]/tbar[0]/okcd").text = "listcube"
         session.findById("wnd[0]").sendVKey 0
         session.findById("wnd[0]/usr/ctxtP_DTA").text = "ibplflr"
         session.findById("wnd[0]/usr/ctxtP_DTA").caretPosition = 7
         session.findById("wnd[0]").sendVKey 0
         session.findById("wnd[0]").sendVKey 8
         session.findById("wnd[0]/usr/chkL_AG").selected = false
         session.findById("wnd[0]/usr/chkL_NO").selected = true
         session.findById("wnd[0]/usr/txtL_MX").text = ""
         session.findById("wnd[0]/usr/txtL_MX").setFocus
         session.findById("wnd[0]/usr/txtL_MX").caretPosition = 11
         session.findById("wnd[0]/tbar[1]/btn[25]").press
         session.findById("wnd[0]/usr/chkS003").selected = true
         session.findById("wnd[0]/usr/chkS006").selected = true
         session.findById("wnd[0]/usr/chkS008").selected = true
         session.findById("wnd[0]/usr/chkS008").setFocus
         session.findById("wnd[0]/tbar[1]/btn[8]").press
         session.findById("wnd[0]/tbar[1]/btn[8]").press
         session.findById("wnd[0]/tbar[1]/btn[33]").press
         session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").setCurrentCell -1,"TEXT"
         session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectColumn "TEXT"
         session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").contextMenu
         session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectContextMenuItem "&FILTER"
         session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "*monitoring 2024*"
         session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 17
         session.findById("wnd[2]").sendVKey 0
         session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "0"
         session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell
         session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").select
         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "LT PA3 Fcst.txt"
         session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\" & strUserName & strPathwayFcst
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         session.findById("wnd[1]/tbar[0]/btn[11]").press
         session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
         session.findById("wnd[0]").sendVKey 0
      End If
   End if
   'job run in PA3
   if (currentTime >= startTime3) And (currentTime <= endTime3) Then
      session.findById("wnd[0]/tbar[0]/okcd").text = "sm37"
      session.findById("wnd[0]").sendVKey 0
      session.findById("wnd[0]/usr/txtBTCH2170-JOBNAME").text = "GB3*"
      session.findById("wnd[0]/usr/txtBTCH2170-USERNAME").text = "*"
      session.findById("wnd[0]/usr/ctxtBTCH2170-FROM_DATE").text = VariousDate
      session.findById("wnd[0]/usr/txtBTCH2170-ABAPNAME").setFocus
      session.findById("wnd[0]/usr/txtBTCH2170-ABAPNAME").caretPosition = 0
      session.findById("wnd[0]/tbar[1]/btn[8]").press
      session.findById("wnd[0]/mbar/menu[3]/menu[0]/menu[2]").select
      session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
      session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
      session.findById("wnd[1]/tbar[0]/btn[0]").press
      session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "JOBS PA3 GB3.txt"
      session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\" & strUserName & strPathwayPA3
      session.findById("wnd[1]/tbar[0]/btn[0]").press
      session.findById("wnd[1]/tbar[0]/btn[11]").press
      session.findById("wnd[0]/tbar[0]/btn[3]").press
      session.findById("wnd[0]/usr/txtBTCH2170-JOBNAME").text = "FR2*"
      session.findById("wnd[0]/usr/txtBTCH2170-USERNAME").text = "*"
      session.findById("wnd[0]/usr/ctxtBTCH2170-FROM_DATE").text = VariousDate
      session.findById("wnd[0]/usr/txtBTCH2170-ABAPNAME").setFocus
      session.findById("wnd[0]/usr/txtBTCH2170-ABAPNAME").caretPosition = 0
      session.findById("wnd[0]/tbar[1]/btn[8]").press
      session.findById("wnd[0]/mbar/menu[3]/menu[0]/menu[2]").select
      session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
      session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
      session.findById("wnd[1]/tbar[0]/btn[0]").press
      session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "JOBS PA3 FR2.txt"
      session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\" & strUserName & strPathwayPA3
      session.findById("wnd[1]/tbar[0]/btn[0]").press
      session.findById("wnd[1]/tbar[0]/btn[11]").press
      session.findById("wnd[0]/tbar[0]/btn[3]").press
      session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
      session.findById("wnd[0]").sendVKey 0
   End if
   session.findById("wnd[0]/tbar[0]/btn[3]").press
   session.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
   session.findById("wnd[0]").sendVKey 0
End if

if (currentTime >= startTime3) And (currentTime <= endTime3) Then 'if time 6AM-11:30AM then it will be run
   'PA6 fcst
   WScript.Sleep 1500

   Set connection = application.OpenConnection("APO Prod PA6 (Specialized Nutrition)", True)
   Set session = connection.Children(0)

   WScript.Sleep 1500

   if dayOfWeek = 2 Then
      ' fcst check PA6 long term morning if current time is between 6 AM and 11:30 AM and it's monday
      If (currentTime >= startTime1) And (currentTime <= endTime1) Then
         session.findById("wnd[0]/tbar[0]/btn[3]").press
         session.findById("wnd[0]/tbar[0]/okcd").text = "listcube"
         session.findById("wnd[0]").sendVKey 0
         session.findById("wnd[0]/usr/ctxtP_DTA").text = "ibplflr"
         session.findById("wnd[0]/usr/ctxtP_DTA").caretPosition = 7
         session.findById("wnd[0]").sendVKey 0
         session.findById("wnd[0]").sendVKey 8
         session.findById("wnd[0]/usr/chkL_AG").selected = false
         session.findById("wnd[0]/usr/chkL_NO").selected = true
         session.findById("wnd[0]/usr/txtL_MX").text = ""
         session.findById("wnd[0]/usr/txtL_MX").setFocus
         session.findById("wnd[0]/usr/txtL_MX").caretPosition = 11
         session.findById("wnd[0]/tbar[1]/btn[25]").press
         session.findById("wnd[0]/usr/chkS003").selected = true
         session.findById("wnd[0]/usr/chkS006").selected = true
         session.findById("wnd[0]/usr/chkS008").selected = true
         session.findById("wnd[0]/usr/chkS008").setFocus
         session.findById("wnd[0]/tbar[1]/btn[8]").press
         session.findById("wnd[0]/tbar[1]/btn[8]").press
         session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").select
         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "LT PA6 Fcst.txt"
         session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\" & strUserName & strPathwayFcst
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         session.findById("wnd[1]/tbar[0]/btn[11]").press
         session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
         session.findById("wnd[0]").sendVKey 0
      End If
   End if

   'job run in PA6
   if (currentTime >= startTime3) And (currentTime <= endTime3) Then
      session.findById("wnd[0]/tbar[0]/okcd").text = "sm37"
      session.findById("wnd[0]").sendVKey 0
      session.findById("wnd[0]/usr/txtBTCH2170-JOBNAME").text = "*APO*"
      session.findById("wnd[0]/usr/txtBTCH2170-USERNAME").text = "*"
      session.findById("wnd[0]/usr/ctxtBTCH2170-FROM_DATE").text = VariousDate
      session.findById("wnd[0]/usr/txtBTCH2170-ABAPNAME").setFocus
      session.findById("wnd[0]/usr/txtBTCH2170-ABAPNAME").caretPosition = 0
      session.findById("wnd[0]/tbar[1]/btn[8]").press
      session.findById("wnd[0]/mbar/menu[3]/menu[0]/menu[2]").select
      session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
      session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
      session.findById("wnd[1]/tbar[0]/btn[0]").press
      session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "JOBS PA6 APO.txt"
      session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\" & strUserName & strPathwayPA6
      session.findById("wnd[1]/tbar[0]/btn[0]").press
      session.findById("wnd[1]/tbar[0]/btn[11]").press
      session.findById("wnd[0]/tbar[0]/btn[3]").press
      session.findById("wnd[0]/usr/txtBTCH2170-JOBNAME").text = "*FR8 CCR*"
      session.findById("wnd[0]/usr/txtBTCH2170-USERNAME").text = "*"
      session.findById("wnd[0]/usr/ctxtBTCH2170-FROM_DATE").text = VariousDate
      session.findById("wnd[0]/usr/txtBTCH2170-ABAPNAME").setFocus
      session.findById("wnd[0]/usr/txtBTCH2170-ABAPNAME").caretPosition = 0
      session.findById("wnd[0]/tbar[1]/btn[8]").press
      session.findById("wnd[0]/mbar/menu[3]/menu[0]/menu[2]").select
      session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
      session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
      session.findById("wnd[1]/tbar[0]/btn[0]").press
      session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "JOBS PA6 FR8CCR.txt"
      session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\" & strUserName & strPathwayPA6
      session.findById("wnd[1]/tbar[0]/btn[0]").press
      session.findById("wnd[1]/tbar[0]/btn[11]").press
      session.findById("wnd[0]/tbar[0]/btn[3]").press
         ' Check if today is Monday (2)
         If dayOfWeek = 2 Then
            ' Code to run only on Mondays
               session.findById("wnd[0]/usr/txtBTCH2170-JOBNAME").text = "FR8_DELETE_PAST_FC_VILLEFRANCHE"
               session.findById("wnd[0]/usr/txtBTCH2170-USERNAME").text = "*"
               session.findById("wnd[0]/usr/ctxtBTCH2170-FROM_DATE").text = VariousDate
               session.findById("wnd[0]/usr/txtBTCH2170-ABAPNAME").setFocus
               session.findById("wnd[0]/usr/txtBTCH2170-ABAPNAME").caretPosition = 0
               session.findById("wnd[0]/tbar[1]/btn[8]").press
               session.findById("wnd[0]/mbar/menu[3]/menu[0]/menu[2]").select
               session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
               session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
               session.findById("wnd[1]/tbar[0]/btn[0]").press
               session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "JOBS PA6 VILLEFRANCHE.txt"
               session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\" & strUserName & strPathwayPA6
               session.findById("wnd[1]/tbar[0]/btn[0]").press
               session.findById("wnd[1]/tbar[0]/btn[11]").press
               session.findById("wnd[0]/tbar[0]/btn[3]").press
         End if
      session.findById("wnd[0]/usr/txtBTCH2170-JOBNAME").text = "*SCM*"
      session.findById("wnd[0]/usr/txtBTCH2170-USERNAME").text = "*"
      session.findById("wnd[0]/usr/ctxtBTCH2170-FROM_DATE").text = VariousDate
      session.findById("wnd[0]/usr/txtBTCH2170-ABAPNAME").setFocus
      session.findById("wnd[0]/usr/txtBTCH2170-ABAPNAME").caretPosition = 0
      session.findById("wnd[0]/tbar[1]/btn[8]").press
      session.findById("wnd[0]/mbar/menu[3]/menu[0]/menu[2]").select
      session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
      session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
      session.findById("wnd[1]/tbar[0]/btn[0]").press
      session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "JOBS PA6 SCM.txt"
      session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\" & strUserName & strPathwayPA6
      session.findById("wnd[1]/tbar[0]/btn[0]").press
      session.findById("wnd[1]/tbar[0]/btn[11]").press
      session.findById("wnd[0]/tbar[0]/btn[3]").press
         ' Check if today is Tuesday (3)
         If dayOfWeek = 3 Then
            ' Code to run only on Tuesdays
               session.findById("wnd[0]/usr/txtBTCH2170-JOBNAME").text = "GB8PPM_MODE_PRIORITY"
               session.findById("wnd[0]/usr/txtBTCH2170-USERNAME").text = "*"
               session.findById("wnd[0]/usr/ctxtBTCH2170-FROM_DATE").text = VariousDate
               session.findById("wnd[0]/usr/txtBTCH2170-ABAPNAME").setFocus
               session.findById("wnd[0]/usr/txtBTCH2170-ABAPNAME").caretPosition = 0
               session.findById("wnd[0]/tbar[1]/btn[8]").press
               session.findById("wnd[0]/mbar/menu[3]/menu[0]/menu[2]").select
               session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
               session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
               session.findById("wnd[1]/tbar[0]/btn[0]").press
               session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "JOBS PA6 GB8PPM.txt"
               session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\" & strUserName & strPathwayPA6
               session.findById("wnd[1]/tbar[0]/btn[0]").press
               session.findById("wnd[1]/tbar[0]/btn[11]").press
               session.findById("wnd[0]/tbar[0]/btn[3]").press
               session.findById("wnd[0]/tbar[0]/btn[3]").press
         End If
   End if

   session.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
   session.findById("wnd[0]").sendVKey 0
End if

if (currentTime >= startTime3) And (currentTime <= endTime3) Then 'if 6AM-11:30AM, then it will be run
   'P01
   Set connection = application.OpenConnection("Themis Prod P01 (EMEA: West Europe)", True)
   Set session = connection.Children(0)

   WScript.Sleep 1500

   'job run in P01
   if (currentTime >= startTime3) And (currentTime <= endTime3) Then
         session.findById("wnd[0]/tbar[0]/okcd").text = "sm37"
         session.findById("wnd[0]").sendVKey 0
         session.findById("wnd[0]/usr/txtBTCH2170-JOBNAME").text = "BE1SAPCI*"
         session.findById("wnd[0]/usr/txtBTCH2170-USERNAME").text = "*"
         session.findById("wnd[0]/usr/ctxtBTCH2170-FROM_DATE").text = VariousDate
         session.findById("wnd[0]/usr/txtBTCH2170-ABAPNAME").setFocus
         session.findById("wnd[0]/usr/txtBTCH2170-ABAPNAME").caretPosition = 0
         session.findById("wnd[0]/tbar[1]/btn[8]").press
         session.findById("wnd[0]/mbar/menu[3]/menu[0]/menu[2]").select
         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "JOBS P01 BE1SAPCI.txt"
         session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\" & strUserName & strPathwayP01
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         session.findById("wnd[1]/tbar[0]/btn[11]").press
         session.findById("wnd[0]/tbar[0]/btn[3]").press
   End if
   session.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
   session.findById("wnd[0]").sendVKey 0
End if


if (currentTime >= startTime3) And (currentTime <= endTime3) Then 'if 6AM-11:30AM, then it will be run
   'P02
   Set connection = application.OpenConnection("Themis Prod P02 (EMEA: East/North Europe)", True)
   Set session = connection.Children(0)

   WScript.Sleep 1500

   'job run in P02
   if (currentTime >= startTime3) And (currentTime <= endTime3) Then
         session.findById("wnd[0]/tbar[0]/okcd").text = "sm37"
         session.findById("wnd[0]").sendVKey 0
         session.findById("wnd[0]/usr/txtBTCH2170-JOBNAME").text = "RO*CI*"
         session.findById("wnd[0]/usr/txtBTCH2170-USERNAME").text = "*"
         session.findById("wnd[0]/usr/ctxtBTCH2170-FROM_DATE").text = VariousDate
         session.findById("wnd[0]/usr/txtBTCH2170-ABAPNAME").setFocus
         session.findById("wnd[0]/usr/txtBTCH2170-ABAPNAME").caretPosition = 0
         session.findById("wnd[0]/tbar[1]/btn[8]").press
         session.findById("wnd[0]/mbar/menu[3]/menu[0]/menu[2]").select
         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "JOBS P02 RO CI.txt"
         session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\" & strUserName & strPathwayP02
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         session.findById("wnd[1]/tbar[0]/btn[11]").press
         session.findById("wnd[0]/tbar[0]/btn[3]").press
         session.findById("wnd[0]/usr/txtBTCH2170-JOBNAME").text = "BG*CI*"
         session.findById("wnd[0]/usr/txtBTCH2170-USERNAME").text = "*"
         session.findById("wnd[0]/usr/ctxtBTCH2170-FROM_DATE").text = VariousDate
         session.findById("wnd[0]/usr/txtBTCH2170-ABAPNAME").setFocus
         session.findById("wnd[0]/usr/txtBTCH2170-ABAPNAME").caretPosition = 0
         session.findById("wnd[0]/tbar[1]/btn[8]").press
         session.findById("wnd[0]/mbar/menu[3]/menu[0]/menu[2]").select
         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "JOBS P02 BG CI.txt"
         session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\" & strUserName & strPathwayP02
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         session.findById("wnd[1]/tbar[0]/btn[11]").press
         session.findById("wnd[0]/tbar[0]/btn[3]").press
         session.findById("wnd[0]/usr/txtBTCH2170-JOBNAME").text = "HU*CI*"
         session.findById("wnd[0]/usr/txtBTCH2170-USERNAME").text = "*"
         session.findById("wnd[0]/usr/ctxtBTCH2170-FROM_DATE").text = VariousDate
         session.findById("wnd[0]/usr/txtBTCH2170-ABAPNAME").setFocus
         session.findById("wnd[0]/usr/txtBTCH2170-ABAPNAME").caretPosition = 0
         session.findById("wnd[0]/tbar[1]/btn[8]").press
         session.findById("wnd[0]/mbar/menu[3]/menu[0]/menu[2]").select
         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "JOBS P02 HU CI.txt"
         session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\" & strUserName & strPathwayP02
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         session.findById("wnd[1]/tbar[0]/btn[11]").press
         session.findById("wnd[0]/tbar[0]/btn[3]").press
         session.findById("wnd[0]/usr/txtBTCH2170-JOBNAME").text = "CZ*CI*"
         session.findById("wnd[0]/usr/txtBTCH2170-USERNAME").text = "*"
         session.findById("wnd[0]/usr/ctxtBTCH2170-FROM_DATE").text = VariousDate
         session.findById("wnd[0]/usr/txtBTCH2170-ABAPNAME").setFocus
         session.findById("wnd[0]/usr/txtBTCH2170-ABAPNAME").caretPosition = 0
         session.findById("wnd[0]/tbar[1]/btn[8]").press
         session.findById("wnd[0]/mbar/menu[3]/menu[0]/menu[2]").select
         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "JOBS P02 CZ CI.txt"
         session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\" & strUserName & strPathwayP02
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         session.findById("wnd[1]/tbar[0]/btn[11]").press
         session.findById("wnd[0]/tbar[0]/btn[3]").press
         session.findById("wnd[0]/usr/txtBTCH2170-JOBNAME").text = "SE1*CI*"
         session.findById("wnd[0]/usr/txtBTCH2170-USERNAME").text = "*"
         session.findById("wnd[0]/usr/ctxtBTCH2170-FROM_DATE").text = VariousDate
         session.findById("wnd[0]/usr/txtBTCH2170-ABAPNAME").setFocus
         session.findById("wnd[0]/usr/txtBTCH2170-ABAPNAME").caretPosition = 0
         session.findById("wnd[0]/tbar[1]/btn[8]").press
         session.findById("wnd[0]/mbar/menu[3]/menu[0]/menu[2]").select
         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "JOBS P02 SE1 CI.txt"
         session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\" & strUserName & strPathwayP02
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         session.findById("wnd[1]/tbar[0]/btn[11]").press
         session.findById("wnd[0]/tbar[0]/btn[3]").press
         session.findById("wnd[0]/usr/txtBTCH2170-JOBNAME").text = "PL*CI*"
         session.findById("wnd[0]/usr/txtBTCH2170-USERNAME").text = "*"
         session.findById("wnd[0]/usr/ctxtBTCH2170-FROM_DATE").text = VariousDate
         session.findById("wnd[0]/usr/txtBTCH2170-ABAPNAME").setFocus
         session.findById("wnd[0]/usr/txtBTCH2170-ABAPNAME").caretPosition = 0
         session.findById("wnd[0]/tbar[1]/btn[8]").press
         session.findById("wnd[0]/mbar/menu[3]/menu[0]/menu[2]").select
         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "JOBS P02 PL CI.txt"
         session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\" & strUserName & strPathwayP02
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         session.findById("wnd[1]/tbar[0]/btn[11]").press
         session.findById("wnd[0]/tbar[0]/btn[3]").press
         session.findById("wnd[0]/usr/txtBTCH2170-JOBNAME").text = "NO1*CI*"
         session.findById("wnd[0]/usr/txtBTCH2170-USERNAME").text = "*"
         session.findById("wnd[0]/usr/ctxtBTCH2170-FROM_DATE").text = VariousDate
         session.findById("wnd[0]/usr/txtBTCH2170-ABAPNAME").setFocus
         session.findById("wnd[0]/usr/txtBTCH2170-ABAPNAME").caretPosition = 0
         session.findById("wnd[0]/tbar[1]/btn[8]").press
         session.findById("wnd[0]/mbar/menu[3]/menu[0]/menu[2]").select
         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "JOBS P02 NO1 CI.txt"
         session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\" & strUserName & strPathwayP02
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         session.findById("wnd[1]/tbar[0]/btn[11]").press
         session.findById("wnd[0]/tbar[0]/btn[3]").press
         session.findById("wnd[0]/usr/txtBTCH2170-JOBNAME").text = "LV2*SAPCIF001*"
         session.findById("wnd[0]/usr/txtBTCH2170-USERNAME").text = "*"
         session.findById("wnd[0]/usr/ctxtBTCH2170-FROM_DATE").text = VariousDate
         session.findById("wnd[0]/usr/txtBTCH2170-ABAPNAME").setFocus
         session.findById("wnd[0]/usr/txtBTCH2170-ABAPNAME").caretPosition = 0
         session.findById("wnd[0]/tbar[1]/btn[8]").press
         session.findById("wnd[0]/mbar/menu[3]/menu[0]/menu[2]").select
         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "JOBS P02 LV2 SAPCIF001.txt"
         session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\" & strUserName & strPathwayP02
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         session.findById("wnd[1]/tbar[0]/btn[11]").press
         session.findById("wnd[0]/tbar[0]/btn[3]").press
   End if
   session.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
   session.findById("wnd[0]").sendVKey 0
End if

if (currentTime >= startTime3) And (currentTime <= endTime3) Then 'if 6AM-11:30AM, then it will be run
   'P06
   Set connection = application.OpenConnection("Themis Prod P06 (EMEA)", True)
   Set session = connection.Children(0)

   WScript.Sleep 1500

   'job run in P06
   if (currentTime >= startTime3) And (currentTime <= endTime3) Then
         session.findById("wnd[0]/tbar[0]/okcd").text = "sm37"
         session.findById("wnd[0]").sendVKey 0
         session.findById("wnd[0]/usr/txtBTCH2170-JOBNAME").text = "*CI*"
         session.findById("wnd[0]/usr/txtBTCH2170-USERNAME").text = "*"
         session.findById("wnd[0]/usr/ctxtBTCH2170-FROM_DATE").text = VariousDate
         session.findById("wnd[0]/usr/txtBTCH2170-ABAPNAME").setFocus
         session.findById("wnd[0]/usr/txtBTCH2170-ABAPNAME").caretPosition = 0
         session.findById("wnd[0]/tbar[1]/btn[8]").press
         session.findById("wnd[0]/mbar/menu[3]/menu[0]/menu[2]").select
         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
         session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "JOBS P06 CI.txt"
         session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\" & strUserName & strPathwayP06
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         session.findById("wnd[1]/tbar[0]/btn[11]").press
         session.findById("wnd[0]/tbar[0]/btn[3]").press
   End if
   session.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
   session.findById("wnd[0]").sendVKey 0
End if

Dim strFilePath2

' Define the path to the secondary VBScript
strFilePath2 = "C:\Users\" & strUserName & "\OneDrive - Danone\General - EU-IT&DATA HUB D2D - PLA\50 Support team\60. Monitoring\[DO NOT USE] Monitoring prompt message.vbs"

' Run the secondary VBScript
objShell.Run """" & strFilePath2 & """"