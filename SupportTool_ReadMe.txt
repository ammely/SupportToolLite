The tool consist of two parts:

1- Software removal 
	To remove any software that releated to Tobii Dynavox, you need first to select the software from the list below and then press Start. 
	You will be asked to confirm, press Yes!
A. "Remove Progressive Sweet": First you will get a list of all installed progressive sw such as: Talk, Browse, Control, Phone, Switcher. Select sw that you want to remove and press OK
B. "Remove PCEye5 Bundle": 	Will remove all components that are: Control, ETSettings, Experience SW for Windows, UN and Switcher.
C. "Remove all ET SW": 		Will first backup Calibration profiles into %temp% folder. 
							Then uninstall all ET SW for I-Series*, WC2, GP users. i.e. Experience sw for Windows and Tobii Eye Tracking Core
							For I-Series device, it will run "before uninstall" script that has been provided to tech support. 
							Then it will uninstall all ET drivers and remove all services 
							*For I-Series, you need to install both Experience SW for Windows and ETSettings.
D. "Remove WC&GP Bundle": 	Will remove all WC and GP software that are: WC2/ GP, Tobii EyeTracker Core, UN, Virtual Remote
E. "Remove VC++": 			Will list out all installed Redist VC++, select any one that need to uninstall and press OK. 
F. "Remove PCEye Package": 	Will remove all old PCEye Package that are: TGIS*, ET Browser, PCEye Configuration Guide, Gaze HID, UN, GS and GS language pack.
							*Not recommended to use it on I-Series+ 
G. "Remove Communicator": 	Will only remove Communicator SW.
H. "Remove Compass":		Will only remove Compass SW.
I. "Remove TGIS only":		Will remove TGIS*, GS, and GS language pack.
							*Can be used on I-Series+
J. "Remove TGIS profile calibrations" 
K. "Remove all users C5":	Will remove all C5 users. 

Also you can:
A. "Backup Gaze Interaction", 
B. "Copy License"



2- List of useful tools, by press on one of following, it will display info on output:
A. All versions:				List all Tobii related software installed. List of all active processes, services, Eye Tracker PID and Tobii drivers
						List of useful info and store it in "infofolder" such as Device info, Battery log and Hardware info. 
B. Get Services:				List all active services and processes
C. Restart Services:			Will restart all ET services and processes
D. Firmware v / Upgrade:		List of current ET firmware (for all ET) and upgrade it (only for IS4)
E. Show/Hide Track Status:		Will launch Track Status 
F. WCF:					Checking if there are any other SW that blocking connection between ET sw and Communication sw
G. Fix Partner Window:			Install FTDI driver for Partner Window for I-Series
H. Delete Emails in C5:			Launch app to delete emails in C5
I. SMBIOS:					Will only launch getSMBIOSvalues.cmd


Support Tool 2.0.7 (June - 2022)
- Added Partner Window driver
- Added Delet Emails in C5
- Bug fixe

Support Tool 2.0.6 (April - 2022)
- Combined all software and drivers into one button
- Save SW version into local folder
- Bug fixes


Support Tool 2.0.4 (December - 2021)
- Redesign, improvements and bug fixs 
- Removed unnecessary functiones such as: Reset TETC, BeforeUninstallGG,DriverSetupGG, and UninstallGG
- Meraged some function into one function such as: Process/PID/Drivers, Firmware v / Upgrade


Support Tool 2.0.3 (November - 2021)
New features:
- Remove Progressive sweet: Listing available progressive sweet apps and remove it by selecting from the list. Progressive sweet apps are: Control, Switcher, Browse, Talk and Phone.
- Remove VC++: Listing all available VC++ redist. Installed and remove it by select from the list. 
- Driver Setup GG: Only for Gibbon Gaze, run it if 4.49.0.4000 won’t be removed.
- Uninstall GG: Only for Gibbon Gaze, run it if 4.49.0.4000 won’t be removed.
- HW Info: Collect HW info included battery info and save it in same folder as Support Tool file is. 
- Install PDK GG: only for Gibbon Gaze, if PDK is missing, install PDK.
Bug fixes



Support Tool 2.0.1 (June - 2021)
- New updated version of Support Too
