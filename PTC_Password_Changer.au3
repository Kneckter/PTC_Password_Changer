;*********** Use **************
; Auther:    Kneckter
; Date:      February 19, 2018
; Name:      PTC Password Changer Version 2
;
; Press Esc to terminate script.
; Tested using Internet Explorer 11 on Windows 7 and 10.
; Might add functionality to change emails too.
; Password changer only: Changes passwords in 10 seconds per account, 1 account at a time.
; That is alittle more than 2:45 hours for 1000 accounts.
; Add accounts to a file as ptc,username,oldpass,newpass,newemail
; If you are only changing passwords use ptc,username,oldpass,newpass,
; If you are only changing emails use ptc,username,pass,,newemail
;******************************

#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <Misc.au3>
#Include <WinAPI.au3>
#include <File.au3>
#include <IE.au3>

   Opt('MustDeclareVars', 1) ; Makes it so variables have to be declared before used.
   HotKeySet("{ESC}", "Stop") ; Sets ESC to go to the Stop function

   Global $Button_1, $Button_2, $Button_3, $Button_4, $Lbl_1, $Lbl_2, $CSVLoc, $CSVTB, $oIE

   ; Form's variables go Text, Width, Height
   GUICreate("PTC Password Changer v2", 250, 125) ; Will create a dialog box
   GUISetBkColor(0x000000) ; Sets the BG color of the box.
   Opt("GUICoordMode", 1) ; Absolute coordinates (default) still relative to the dialog box.

   ; Lables's variables go Text, Left, Top, Width, Height, Style
   $Lbl_1 = GUICtrlCreateLabel("Pokemon Trainer Club Options", 0, 5, 250, 15, $SS_CENTER)
   $Lbl_2 = GUICtrlCreateLabel("Select an Account CSV File", 0, 80, 250, 15, $SS_CENTER)

   ; Button's variables go Text, Left, Top, Width, Height
   $Button_1 = GUICtrlCreateButton("Change Password", 5, 20, 115, 25) ; Sets to a varriable and creates a button
   $Button_2 = GUICtrlCreateButton("Change Email", 125, 20, 115, 25) ; Sets to a varriable and creates a button
   $Button_3 = GUICtrlCreateButton("Change Password and Email", 5, 50, 235, 25) ; Sets to a varriable and creates a button
   $Button_4 = GUICtrlCreateButton("Browse", 5, 95, 115, 25) ; Sets to a varriable and creates a button

   ; Textbox's variables go Text, Left, Top, Width, Height, style
   $CSVTB = GUICtrlCreateInput("",125,95,115,25) ; Sets to a varriable and creates a textbox

   ; Label text colors
   GUICtrlSetColor($Lbl_1, 0xFFF200) ; Sets the label text color to yellow.
   GUICtrlSetColor($Lbl_2, 0xFFF200) ; Sets the label text color to yellow.

   ; Buttons colors
   GUICtrlSetBkColor($Button_1, 0xFFF200) ; Sets the button color to yellow.
   GUICtrlSetBkColor($Button_2, 0xFFF200) ; Sets the button color to yellow.
   GUICtrlSetBkColor($Button_3, 0xFFF200) ; Sets the button color to yellow.
   GUICtrlSetBkColor($Button_4, 0xFFF200) ; Sets the button color to yellow.

   ; Inputbox color
   GUICtrlSetBkColor($CSVTB, 0xFFF200)

   HomeBox() ; Head to the HomeBox funtion when opening

Func HomeBox() ; Funtion to hold some information about the dialog box
   Local $msg ; Used to hold what happens to the GUI.

   GUISetState(@SW_SHOW)   ; will display the dialog box

   ; Run the GUI until the dialog is closed
   While 1
      $msg = GUIGetMsg() ; Sets any events that happen in the box to msg
      Select ; Controls the box depending on what happens
         Case $msg = $GUI_EVENT_CLOSE
            Exit 0
         Case $msg = $Button_1
            But_1()
         Case $msg = $Button_2
            But_2()
         Case $msg = $Button_3
            But_3()
         Case $msg = $Button_4
            But_4()
      EndSelect
   WEnd
EndFunc

func But_1()
   ToolTip("Running password changer **Checking account file** Hit ESC to abort.",0,0)
   GUISetState(@SW_HIDE) ; Hides the dialog box while the program runs.

   Local $RowNub, $TRow, $CSV[1], $SS, $Count
   $RowNub = 1 ; Set the row number
   $Count = 1 ; Used to count the number accounts and restart IE

   ; Read lines into an array
   FileOpen($CSVLoc, 0) ; Open the text file
   _FileReadToArray($CSVLoc,$CSV, Default, ",")

   ; Check for errors with the file handling
   If FileOpen($CSVLoc, 0) = -1 Then
     MsgBox($MB_ICONERROR, "Account File Error", "An error occurred when reading the account file. Ensure the correct file has been selected.")
      _FileWriteLog(@WorkingDir & "\PTC_Password_Changer_Errors.log", "An error occurred when reading the account file. Ensure the correct file has been selected.")
     Stop()
   ElseIf UBound($CSV, $UBOUND_COLUMNS) < 5 Then
     MsgBox($MB_ICONERROR, "Account File Error", "An error occurred when reading the account file. Ensure the correct file has the right columns.")
      _FileWriteLog(@WorkingDir & "\PTC_Password_Changer_Errors.log", "An error occurred when reading the account file. Ensure the correct file has the right columns.")
     Stop()
   EndIf

   ; Count the rows in the file
   $TRow = _FileCountLines ($CSVLoc) ; Set the total rows to the variable

   ; Start the logs before loopping
   _FileWriteLog(@WorkingDir & "\PTC_Password_Changer_Errors.log", "There are " & $TRow & " accounts in the file.")
   _FileWriteLog(@WorkingDir & "\PTC_Password_Changer_Errors.log", "New batch process started.")
   _FileWriteLog(@WorkingDir & "\PTC_Password_Changer_Successes.log", "There are " & $TRow & " accounts in the file.")
   _FileWriteLog(@WorkingDir & "\PTC_Password_Changer_Successes.log", "New batch process started.")

   while $RowNub <= $TRow
      ToolTip("Running password changer for account " & $CSV[$RowNub][1] & ". There are " & $TRow - $RowNub & " left. Hit ESC to abort.",0,0)

      ; Put the link in the address bar
      $oIE = _IECreate ("https://club.pokemon.com/us/pokemon-trainer-club/edit-profile/",0,0,0,0)  ; Create a new IE tab in a new window and make it invisible
      _IELoadWait($oIE, 0, 500000) ; Use this to make the script wait instead of the above (it has problems)

      ; Get the form locations and fill them in
      Local $o_form1 = _IEFormGetObjByName ($oIE, "login-form")
      Local $o_login = _IEFormElementGetObjByName ($o_form1, "username")
      Local $o_password = _IEFormElementGetObjByName ($o_form1, "password")
      _IEFormElementSetValue ($o_login, $CSV[$RowNub][1])
      _IEFormElementSetValue ($o_password, $CSV[$RowNub][2])
      _IEFormSubmit($o_form1, 0) ; Submit the form
      _IELoadWait($oIE, 0, 500000) ; Use this to make the script wait instead of the above (it has problems)

      ; Check if the page says wrong password
      Local $sText = _IEBodyReadText($oIE)
      Local $sSearch = StringInStr ( $sText, "Change Password")

      If $sSearch = 0 Then
         ; If the pw is wrong, write a log and move to the next account
         _FileWriteLog(@WorkingDir & "\PTC_Password_Changer_Errors.log", "Account #" & $RowNub & " " & $CSV[$RowNub][1] & " reported: Your username/password is incorrect or the program glitched.")
         $RowNub = $RowNub + 1
      ElseIf $sSearch > 0 Then
         ; If the pw is right and "Change Password" is located on the page, navigate to the pw changer
         _IENavigate($oIE, "https://club.pokemon.com/us/pokemon-trainer-club/my-password",0)
         _IELoadWait($oIE, 0, 500000) ; Use this to make the script wait instead of the above (it has problems)

         ; Get the form locations and fill them in
         Local $o_form2 = _IEFormGetObjByName ($oIE, "account")
         Local $o_current = _IEFormElementGetObjByName ($o_form2, "current_password")
         Local $o_npassword = _IEFormElementGetObjByName ($o_form2, "password")
         Local $o_confirm = _IEFormElementGetObjByName ($o_form2, "confirm_password")
         _IEFormElementSetValue ($o_current, $CSV[$RowNub][2])
         _IEFormElementSetValue ($o_npassword, $CSV[$RowNub][3])
         _IEFormElementSetValue ($o_confirm, $CSV[$RowNub][3])
         _IEFormSubmit($o_form2, 0) ; Submit the form and don't wait. Also, use the submit functio because the button doesn't have a form name
         _IELoadWait($oIE, 0, 500000) ; Use this to make the script wait instead of the above (it has problems)

         ; Check for errors and write logs
         Local $sText = _IEBodyReadText($oIE)
         Local $sSearch = StringInStr ( $sText, "Your password has been updated")
         If $sSearch = 0 Then
            ; If the text is not found, the password was not changed
            _FileWriteLog(@WorkingDir & "\PTC_Password_Changer_Errors.log", "Account #" & $RowNub & " " & $CSV[$RowNub][1] & " reported: Your password has NOT been updated. Check the new password requirements or IE glitched.")
         ElseIf $sSearch > 0 Then
            ; If the text was found, the password was changed
            _FileWriteLog(@WorkingDir & "\PTC_Password_Changer_Successes.log", "Account #" & $RowNub & " " & $CSV[$RowNub][1] & " reported: Your password has been updated.")
         ElseIf @error Then
            ; Catch any error calls
            _FileWriteLog(@WorkingDir & "\PTC_Password_Changer_Errors.log", "Account #" & $RowNub & " " & $CSV[$RowNub][1] & " reported error: " & @error)
         EndIf

         ; Put the link in the address bar
         _IENavigate($oIE, "https://club.pokemon.com/us/pokemon-trainer-club/logout",0)
         _IELoadWait($oIE, 0, 500000) ; Use this to make the script wait instead of the above (it has problems)

         ; Next line in the file
         $RowNub = $RowNub + 1
      ElseIf @error Then
         _FileWriteLog(@WorkingDir & "\PTC_Password_Changer_Errors.log", "Account #" & $RowNub & " " & $CSV[$RowNub][1] & " reported error: " & @error)
      EndIf

      ; Use this to check the PID of the IE window and close it because IE hogs RAM
      Local $hIE = _IEPropertyGet($oIE,"hwnd")
      Local $iPID = ""
      _WinAPI_GetWindowThreadProcessId($hIE,$iPID)
	  _IEQuit($oIE) ; Close the IE window so we can start a new one
	  While ProcessExists ($iPID) <>  0  ; Sleep until the IE window is closed
		 Sleep(100)
	  WEnd
   WEnd
   If @error Then
      _FileWriteLog(@WorkingDir & "\PTC_Password_Changer_Errors.log", "Account #" & $RowNub & " " & $CSV[$RowNub][1] & " reported error at the end of the loop: " & @error)
   EndIf
   FileClose($CSVLoc)
   Stop()
EndFunc

func But_2()
   ToolTip("Running email changer **Checking account file** Hit ESC to abort.",0,0)
   GUISetState(@SW_HIDE) ; Hides the dialog box while the program runs.

   Local $RowNub, $TRow, $CSV[1], $SS, $Count
   $RowNub = 1 ; Set the row number
   $Count = 1 ; Used to count the number accounts and restart IE

   ; Read lines into an array
   FileOpen($CSVLoc, 0) ; Open the text file
   _FileReadToArray($CSVLoc,$CSV, Default, ",")

   ; Check for errors with the file handling
   If FileOpen($CSVLoc, 0) = -1 Then
     MsgBox($MB_ICONERROR, "Account File Error", "An error occurred when reading the account file. Ensure the correct file has been selected.")
      _FileWriteLog(@WorkingDir & "\PTC_Password_Changer_Errors.log", "An error occurred when reading the account file. Ensure the correct file has been selected.")
     Stop()
   ElseIf UBound($CSV, $UBOUND_COLUMNS) < 5 Then
     MsgBox($MB_ICONERROR, "Account File Error", "An error occurred when reading the account file. Ensure the correct file has the right columns.")
      _FileWriteLog(@WorkingDir & "\PTC_Password_Changer_Errors.log", "An error occurred when reading the account file. Ensure the correct file has the right columns.")
     Stop()
   EndIf

   ; Count the rows in the file
   $TRow = _FileCountLines ($CSVLoc) ; Set the total rows to the variable

   ; Start the logs before loopping
   _FileWriteLog(@WorkingDir & "\PTC_Password_Changer_Errors.log", "There are " & $TRow & " accounts in the file.")
   _FileWriteLog(@WorkingDir & "\PTC_Password_Changer_Errors.log", "New batch process started.")
   _FileWriteLog(@WorkingDir & "\PTC_Password_Changer_Successes.log", "There are " & $TRow & " accounts in the file.")
   _FileWriteLog(@WorkingDir & "\PTC_Password_Changer_Successes.log", "New batch process started.")

   while $RowNub <= $TRow
      ToolTip("Running email changer for account " & $CSV[$RowNub][1] & ". There are " & $TRow - $RowNub & " left. Hit ESC to abort.",0,0)

      ; Put the link in the address bar
      $oIE = _IECreate ("https://club.pokemon.com/us/pokemon-trainer-club/edit-profile/",0,0,0,0)  ; Create a new IE tab in a new window and make it invisible
      _IELoadWait($oIE, 0, 500000) ; Use this to make the script wait instead of the above (it has problems)

      ; Get the form locations and fill them in
      Local $o_form1 = _IEFormGetObjByName ($oIE, "login-form")
      Local $o_login = _IEFormElementGetObjByName ($o_form1, "username")
      Local $o_password = _IEFormElementGetObjByName ($o_form1, "password")
      _IEFormElementSetValue ($o_login, $CSV[$RowNub][1])
      _IEFormElementSetValue ($o_password, $CSV[$RowNub][2])
      _IEFormSubmit($o_form1, 0) ; Submit the form
      _IELoadWait($oIE, 0, 500000) ; Use this to make the script wait instead of the above (it has problems)

      ; Check if the page says wrong password
      Local $sText = _IEBodyReadText($oIE)
      Local $sSearch = StringInStr ( $sText, "Change Password")

      If $sSearch = 0 Then
         ; If the pw is wrong, write a log and move to the next account
         _FileWriteLog(@WorkingDir & "\PTC_Password_Changer_Errors.log", "Account #" & $RowNub & " " & $CSV[$RowNub][1] & " reported: Your username/password is incorrect or the program glitched.")
         $RowNub = $RowNub + 1
      ElseIf $sSearch > 0 Then
         ; If the pw is right and "Change Password" is located on the page, navigate to the email changer
         _IENavigate($oIE, "https://club.pokemon.com/us/pokemon-trainer-club/my-email",0)
         _IELoadWait($oIE, 0, 500000) ; Use this to make the script wait instead of the above (it has problems)

         ; Get the form locations and fill them in
         Local $o_form2 = _IEFormGetObjByName ($oIE, "account")
         Local $o_current = _IEFormElementGetObjByName ($o_form2, "current_password")
         Local $o_email = _IEFormElementGetObjByName ($o_form2, "email")
         Local $o_confirm = _IEFormElementGetObjByName ($o_form2, "confirm_email")
         _IEFormElementSetValue ($o_current, $CSV[$RowNub][2])
         _IEFormElementSetValue ($o_email, $CSV[$RowNub][4])
         _IEFormElementSetValue ($o_confirm, $CSV[$RowNub][4])
         _IEFormSubmit($o_form2, 0) ; Submit the form and don't wait. Also, use the submit functio because the button doesn't have a form name
         _IELoadWait($oIE, 0, 500000) ; Use this to make the script wait instead of the above (it has problems)

         ; Check for errors and write logs
         Local $sText = _IEBodyReadText($oIE)
         Local $sSearch = StringInStr ( $sText, "Your email address has been updated")
         If $sSearch = 0 Then
            ; If the text is not found, the password was not changed
            _FileWriteLog(@WorkingDir & "\PTC_Password_Changer_Errors.log", "Account #" & $RowNub & " " & $CSV[$RowNub][1] & " reported: Your email has NOT been updated. Check the new password requirements or IE glitched.")
         ElseIf $sSearch > 0 Then
            ; If the text was found, the password was changed
            _FileWriteLog(@WorkingDir & "\PTC_Password_Changer_Successes.log", "Account #" & $RowNub & " " & $CSV[$RowNub][1] & " reported: Your email has been updated.")
         ElseIf @error Then
            ; Catch any error calls
            _FileWriteLog(@WorkingDir & "\PTC_Password_Changer_Errors.log", "Account #" & $RowNub & " " & $CSV[$RowNub][1] & " reported error: " & @error)
         EndIf

         ; Put the link in the address bar
         _IENavigate($oIE, "https://club.pokemon.com/us/pokemon-trainer-club/logout",0)
         _IELoadWait($oIE, 0, 500000) ; Use this to make the script wait instead of the above (it has problems)

         ; Next line in the file
         $RowNub = $RowNub + 1
      ElseIf @error Then
         _FileWriteLog(@WorkingDir & "\PTC_Password_Changer_Errors.log", "Account #" & $RowNub & " " & $CSV[$RowNub][1] & " reported error: " & @error)
      EndIf

      ; Use this to check the PID of the IE window and close it because IE hogs RAM
      Local $hIE = _IEPropertyGet($oIE,"hwnd")
      Local $iPID = ""
      _WinAPI_GetWindowThreadProcessId($hIE,$iPID)
	  _IEQuit($oIE) ; Close the IE window so we can start a new one
	  While ProcessExists ($iPID) <>  0  ; Sleep until the IE window is closed
		 Sleep(100)
	  WEnd
   WEnd
   If @error Then
      _FileWriteLog(@WorkingDir & "\PTC_Password_Changer_Errors.log", "Account #" & $RowNub & " " & $CSV[$RowNub][1] & " reported error at the end of the loop: " & @error)
   EndIf
   FileClose($CSVLoc)
   Stop()
EndFunc

func But_3()
   ToolTip("Running password and email changer **Checking account file** Hit ESC to abort.",0,0)
   GUISetState(@SW_HIDE) ; Hides the dialog box while the program runs.

   Local $RowNub, $TRow, $CSV[1], $SS, $Count
   $RowNub = 1 ; Set the row number
   $Count = 1 ; Used to count the number accounts and restart IE

   ; Read lines into an array
   FileOpen($CSVLoc, 0) ; Open the text file
   _FileReadToArray($CSVLoc,$CSV, Default, ",")

   ; Check for errors with the file handling
   If FileOpen($CSVLoc, 0) = -1 Then
     MsgBox($MB_ICONERROR, "Account File Error", "An error occurred when reading the account file. Ensure the correct file has been selected.")
      _FileWriteLog(@WorkingDir & "\PTC_Password_Changer_Errors.log", "An error occurred when reading the account file. Ensure the correct file has been selected.")
     Stop()
   ElseIf UBound($CSV, $UBOUND_COLUMNS) < 5 Then
     MsgBox($MB_ICONERROR, "Account File Error", "An error occurred when reading the account file. Ensure the correct file has the right columns.")
      _FileWriteLog(@WorkingDir & "\PTC_Password_Changer_Errors.log", "An error occurred when reading the account file. Ensure the correct file has the right columns.")
     Stop()
   EndIf

   ; Count the rows in the file
   $TRow = _FileCountLines ($CSVLoc) ; Set the total rows to the variable

   ; Start the logs before loopping
   _FileWriteLog(@WorkingDir & "\PTC_Password_Changer_Errors.log", "There are " & $TRow & " accounts in the file.")
   _FileWriteLog(@WorkingDir & "\PTC_Password_Changer_Errors.log", "New batch process started.")
   _FileWriteLog(@WorkingDir & "\PTC_Password_Changer_Successes.log", "There are " & $TRow & " accounts in the file.")
   _FileWriteLog(@WorkingDir & "\PTC_Password_Changer_Successes.log", "New batch process started.")

   while $RowNub <= $TRow
      ToolTip("Running password changer for account " & $CSV[$RowNub][1] & ". There are " & $TRow - $RowNub & " left. Hit ESC to abort.",0,0)

      ; Put the link in the address bar
      $oIE = _IECreate ("https://club.pokemon.com/us/pokemon-trainer-club/edit-profile/",0,0,0,0)  ; Create a new IE tab in a new window and make it invisible
      _IELoadWait($oIE, 0, 500000) ; Use this to make the script wait instead of the above (it has problems)

      ; Get the form locations and fill them in
      Local $o_form1 = _IEFormGetObjByName ($oIE, "login-form")
      Local $o_login = _IEFormElementGetObjByName ($o_form1, "username")
      Local $o_password = _IEFormElementGetObjByName ($o_form1, "password")
      _IEFormElementSetValue ($o_login, $CSV[$RowNub][1])
      _IEFormElementSetValue ($o_password, $CSV[$RowNub][2])
      _IEFormSubmit($o_form1, 0) ; Submit the form
      _IELoadWait($oIE, 0, 500000) ; Use this to make the script wait instead of the above (it has problems)

      ; Check if the page says wrong password
      Local $sText = _IEBodyReadText($oIE)
      Local $sSearch = StringInStr ( $sText, "Change Password")

      If $sSearch = 0 Then
         ; If the pw is wrong, write a log and move to the next account
         _FileWriteLog(@WorkingDir & "\PTC_Password_Changer_Errors.log", "Account #" & $RowNub & " " & $CSV[$RowNub][1] & " reported: Your username/password is incorrect or the program glitched.")
         $RowNub = $RowNub + 1
      ElseIf $sSearch > 0 Then
         ; If the pw is right and "Change Password" is located on the page, navigate to the pw changer
         _IENavigate($oIE, "https://club.pokemon.com/us/pokemon-trainer-club/my-password",0)
         _IELoadWait($oIE, 0, 500000) ; Use this to make the script wait instead of the above (it has problems)

         ; Get the form locations and fill them in
         Local $o_form2 = _IEFormGetObjByName ($oIE, "account")
         Local $o_current = _IEFormElementGetObjByName ($o_form2, "current_password")
         Local $o_npassword = _IEFormElementGetObjByName ($o_form2, "password")
         Local $o_confirm = _IEFormElementGetObjByName ($o_form2, "confirm_password")
         _IEFormElementSetValue ($o_current, $CSV[$RowNub][2])
         _IEFormElementSetValue ($o_npassword, $CSV[$RowNub][3])
         _IEFormElementSetValue ($o_confirm, $CSV[$RowNub][3])
         _IEFormSubmit($o_form2, 0) ; Submit the form and don't wait. Also, use the submit functio because the button doesn't have a form name
         _IELoadWait($oIE, 0, 500000) ; Use this to make the script wait instead of the above (it has problems)

         ; Check for errors and write logs
         Local $sText = _IEBodyReadText($oIE)
         Local $sSearch = StringInStr ( $sText, "Your password has been updated")
         If $sSearch = 0 Then
            ; If the text is not found, the password was not changed
            _FileWriteLog(@WorkingDir & "\PTC_Password_Changer_Errors.log", "Account #" & $RowNub & " " & $CSV[$RowNub][1] & " reported: Your password has NOT been updated. Check the new password requirements or IE glitched.")
         ElseIf $sSearch > 0 Then
            ; If the text was found, the password was changed
            _FileWriteLog(@WorkingDir & "\PTC_Password_Changer_Successes.log", "Account #" & $RowNub & " " & $CSV[$RowNub][1] & " reported: Your password has been updated.")
         ElseIf @error Then
            ; Catch any error calls
            _FileWriteLog(@WorkingDir & "\PTC_Password_Changer_Errors.log", "Account #" & $RowNub & " " & $CSV[$RowNub][1] & " reported error: " & @error)
         EndIf

         ToolTip("Running email changer for account " & $CSV[$RowNub][1] & ". There are " & $TRow - $RowNub & " left. Hit ESC to abort.",0,0)

         ; If the pw is right and "Change Password" is located on the page, navigate to the email changer
         _IENavigate($oIE, "https://club.pokemon.com/us/pokemon-trainer-club/my-email",0)
         _IELoadWait($oIE, 0, 500000) ; Use this to make the script wait instead of the above (it has problems)

         ; Get the form locations and fill them in
         Local $o_form2 = _IEFormGetObjByName ($oIE, "account")
         Local $o_current = _IEFormElementGetObjByName ($o_form2, "current_password")
         Local $o_email = _IEFormElementGetObjByName ($o_form2, "email")
         Local $o_confirm = _IEFormElementGetObjByName ($o_form2, "confirm_email")
         _IEFormElementSetValue ($o_current, $CSV[$RowNub][3])
         _IEFormElementSetValue ($o_email, $CSV[$RowNub][4])
         _IEFormElementSetValue ($o_confirm, $CSV[$RowNub][4])
         _IEFormSubmit($o_form2, 0) ; Submit the form and don't wait. Also, use the submit functio because the button doesn't have a form name
         _IELoadWait($oIE, 0, 500000) ; Use this to make the script wait instead of the above (it has problems)

         ; Check for errors and write logs
         Local $sText = _IEBodyReadText($oIE)
         Local $sSearch = StringInStr ( $sText, "Your email address has been updated")
         If $sSearch = 0 Then
            ; If the text is not found, the password was not changed
            _FileWriteLog(@WorkingDir & "\PTC_Password_Changer_Errors.log", "Account #" & $RowNub & " " & $CSV[$RowNub][1] & " reported: Your email has NOT been updated. Check the new password requirements or IE glitched.")
         ElseIf $sSearch > 0 Then
            ; If the text was found, the password was changed
            _FileWriteLog(@WorkingDir & "\PTC_Password_Changer_Successes.log", "Account #" & $RowNub & " " & $CSV[$RowNub][1] & " reported: Your email has been updated.")
         ElseIf @error Then
            ; Catch any error calls
            _FileWriteLog(@WorkingDir & "\PTC_Password_Changer_Errors.log", "Account #" & $RowNub & " " & $CSV[$RowNub][1] & " reported error: " & @error)
         EndIf

		 ; Put the link in the address bar
         _IENavigate($oIE, "https://club.pokemon.com/us/pokemon-trainer-club/logout",0)
         _IELoadWait($oIE, 0, 500000) ; Use this to make the script wait instead of the above (it has problems)

         ; Next line in the file
         $RowNub = $RowNub + 1
      ElseIf @error Then
         _FileWriteLog(@WorkingDir & "\PTC_Password_Changer_Errors.log", "Account #" & $RowNub & " " & $CSV[$RowNub][1] & " reported error: " & @error)
      EndIf

      ; Use this to check the PID of the IE window and close it because IE hogs RAM
      Local $hIE = _IEPropertyGet($oIE,"hwnd")
      Local $iPID = ""
      _WinAPI_GetWindowThreadProcessId($hIE,$iPID)
	  _IEQuit($oIE) ; Close the IE window so we can start a new one
	  While ProcessExists ($iPID) <>  0  ; Sleep until the IE window is closed
		 Sleep(100)
	  WEnd
   WEnd
   If @error Then
      _FileWriteLog(@WorkingDir & "\PTC_Password_Changer_Errors.log", "Account #" & $RowNub & " " & $CSV[$RowNub][1] & " reported error at the end of the loop: " & @error)
   EndIf
   FileClose($CSVLoc)
   Stop()
EndFunc

func But_4()
   ; Display an open dialog to select a file.
   Local $sFileSelectFolder = FileOpenDialog("Select a folder", "", "Text or CSV files (*.txt;*.csv)")
   GUICtrlSetData($CSVTB, $sFileSelectFolder)
   $CSVLoc = GUICtrlRead($CSVTB)
   Stop()
EndFunc

Func Stop()
   ToolTip("",0,0)

   If $oIE = "" Then
      Sleep(10)
   Else
      _IEQuit($oIE)
   EndIf

   If $CSVLoc = "" Then
      Sleep(10)
   Else
      FileClose($CSVLoc)
   EndIf

   HomeBox() ; Returns to the box.
EndFunc
