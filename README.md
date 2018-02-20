
# PTC_Password_Changer

A program  to bulk change Pokemon Trainer Club passwords and emails that was written with AutoIt for Windows.

## Features:

* Change only passwords for multiple PTC accounts
* Change only emails associated to PTC accounts
* Change both passwords and emails for multiple PTC accounts
* Accounts read from TXT or CSV files

## Information

PTC_Password_Changer was created to help protect all of your PTC accounts by allowing you to mass change your PTC account passwords and emails without having to manually copy/paste all of the information into the browser.

## Requirements

* Internet Explorer (tested with version 11 but should work with other versions)
* Windows OS (tested with version 7 and 10 but should work with other versions)
* Access to the PTC website (this will not work if your IP address is blocked)
* Any text editor (for editing/viewing the accounts file and logs)

## Installation

No installation is required if you are using the executable from the Releases. Download the .zip file and extract the .exe file. 

## Compile Code Yourself

If you would like to clone and compile your own code, you can use AutoIt's installation package for the needed files or SciTE-Lite Version 3.5.4. If you do not know how to use AutoIt or SciTE-Lite, please download the released zip file for the compiled program.

## Usage

To use PTC_Password_Changer after it has been extracted or compiled, you will need to make an accounts file (see below). Once the accounts file is created, you can click on *Browse* button to select the file. Choose any of the first 3 buttons to start the process:

* Change Password - Changes the password for each account in the accounts file.
* Change Email - Changes the email for each account in the accounts file.
* Change Password and Email - Changes both the password and email for each account in the accounts file.

*Note that the location of the accounts file will be used to save the log files to.*

To create the accounts file, you can either use a Text file (.txt) or Comma-Separated Values file (.csv). Either way, the file must contain 5 values per line and at least one line. This file needs the following information in the following format: 

* ptc,Username,OldPassword,NewPassword,NewEmail

*Note that there are no spaces in the file as those would be included as part of the values.*

If you are only changing passwords, you can omit the email address but keep the value empty like this:

* ptc,Username,OldPassword,NewPassword,

If you are only changing emails, you can omit the new password but keep the value empty like this:

* ptc,Username,OldPassword,,NewEmail
