# TenShillings-RC6-Widget-VB6

An example of a moveable/resizable single-image, VB6 / TwinBasic program using the RC6 widget framework, complete with right click menu, help, about and licence forms and help CHM.

The RC6 version here operates on Vista+ to Windows 11+. There is another version on github in my repo. that operates with RC5, the earlier version of RichClient. That change allows the widget to run under Windows XP.

The program does very little but place some eye-candy on the desktop in the manner of the old Yahoo Widget/Konfabulator engine. You can add more logic as you require or replace the main image with your own. 
This is merely a template for other single-image VB6 / TwinBasic RC6 programs that I knocked up late at night whilst drinking gin and tonic.

<img width="829" height="520" alt="TenShillings" src="https://github.com/user-attachments/assets/07f63cb6-c0fe-4f70-ad02-b57a9ab18b58" />

If you want a template for a more complex program with multi-function clickable areas then try this program, 
https://github.com/yereverluvinunclebert/Panzer-CPU-Gauge-VB6

BUILD: The program runs without any Microsoft plugins.
 
 Built using: VB6, MZ-TOOLS 3.0, VBAdvance, CodeHelp Core IDE Extender
 Framework 2.2 & Rubberduck 2.4.1, RichClient 6

  ![vb6-logo-200](https://github.com/yereverluvinunclebert/Panzer-CPU-Gauge-VB6/assets/2788342/bf00fa3d-f1d4-417b-bc50-9446f2c3e674)
 
 Links:
 
	https://www.vbrichclient.com/#/en/About/
	MZ-TOOLS https://www.mztools.com/  
	CodeHelp http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=62468&lngWId=1  
	Rubberduck http://rubberduckvba.com/  
	Rocketdock https://punklabs.com/  
	Registry code ALLAPI.COM  
	La Volpe http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=67466&lngWId=1  
	PrivateExtractIcons code http://www.activevb.de/rubriken/  
	Persistent debug code http://www.vbforums.com/member.php?234143-Elroy  
	Open File common dialog code without dependent OCX - http://forums.codeguru.com/member.php?92278-rxbagain  
	VBAdvance  
 
 <img width="800" height="491" alt="tenShillingsHelp" src="https://github.com/user-attachments/assets/2e366c76-90c3-4992-9614-94b65eb559e2" />

 Tested on :
 
	Windows 7 Professional 32bit on Intel    
	Windows 7 Ultimate 64bit on Intel    
	Windows 7 Professional 64bit on Intel    
	Windows XP SP3 32bit on Intel    
	Windows 10 Home 64bit on Intel    
	Windows 10 Home 64bit on AMD    
	Windows 11 64bit on Intel  
   
 CREDITS:
 
 I have really tried to maintain the credits as the project has progressed. If I 
 have made a mistake and left someone out then do forgive me. I will make amends 
 if anyone points out my mistake in leaving someone out.
 
 MicroSoft in the 90s - MS built good, lean and useful tools in the late 90s and 
 early 2000s. Thanks for VB6.
 
 Olaf Schmidt - This tool was built using the RichClient RC5 Cairo wrapper for 
 VB6. Specifically the components using transparency and reading images directly 
 from PSD. Thanks for the massive effort Olaf in creating Cairo counterparts for 
 all VB6 native controls and giving us access to advanced features on controls 
 such as transparency.
 
 Shuja Ali @ codeguru for his settings.ini code.
 
 ALLAPI.COM        For the registry reading code.
 
 Rxbagain on codeguru for his Open File common dialog code without a dependent 
 OCX - http://forums.codeguru.com/member.php?92278-rxbagain
 
 si_the_geek       for his special folder code
 
 Elroy on VB forums for the balloon tooltips
 
 Harry Whitfield for his quality testing, brain stimulation and being an 
 unwitting source of inspiration. 
 
 Dependencies:
 
 o A windows-alike o/s such as Windows XP, 7-11 or Apple Mac OSX 11. 
 
 o Microsoft VB6 IDE installed with its runtime components. The program runs 
 without any additional Microsoft OCX components, just the basic controls that 
 ship with VB6.  
 
 
 * Uses the latest version of the RC6 Cairo framework from Olaf Schmidt.
 
 During development the RC6 components need to be registered. These scripts are 
 used to register. Run each by double-clicking on them.
 
	RegisterRC6inPlace.vbs
	RegisterRC6WidgetsInPlace.vbs
 
 During runtime on the users system, the RC6 components are dynamically 
 referenced using modRC6regfree.bas which is compiled into the binary.	
 
![prefsGeneral](https://github.com/user-attachments/assets/8a228ec6-89a9-459a-98bd-d1d17ce9728d)

 Requires a PzCPU Gauge folder in C:\Users\<user>\AppData\Roaming\ 
 eg: C:\Users\<user>\AppData\Roaming\PzCPU Gauge
 Requires a settings.ini file to exist in C:\Users\<user>\AppData\Roaming\PzCPU Gauge
 The above will be created automatically by the compiled program when run for the 
 first time.
 
o Krool's replacement for the Microsoft Windows Common Controls found in
mscomctl.ocx (slider) are replicated by the addition of one
dedicated OCX file that are shipped with this package.

During development only, this must be copied to C:\windows\syswow64 and should be registered.

- CCRSlider.ocx

Register this using regsvr32, ie. in a CMD window with administrator privileges.
	
	c:                          ! set device to boot drive with Windows
	cd \windows\syswow64s	    ! change default folder to syswow64
	regsvr32 CCRSlider.ocx	! register the ocx

This will allow the custom controls to be accessible to the VB6 IDE
at design time and the sliders will function as intended (if this ocx is
not registered correctly then the relevant controls will be replaced by picture boxes).

The above is only for development, for ordinary users, during runtime there is no need to do the above. The OCX will reside in the program folder. The program reference to this OCX is contained within the supplied resource file, Panzer CPU Gauge.RES. The reference to this file is already compiled into the binary. As long as the OCX is in the same folder as the binary the program will run without the need to register the OCX manually.

 <img width="451" height="406" alt="components" src="https://github.com/user-attachments/assets/0dad6c34-439c-4fbd-8522-23036269a89e" />
 
 * OLEGuids.tlb
 
 This is a type library that defines types, object interfaces, and more specific 
 API definitions needed for COM interop / marshalling. It is only used at design 
 time (IDE). This is a Krool-modified version of the original .tlb from the 
 vbaccelerator website. The .tlb is compiled into the executable.
 For the compiled .exe this is NOT a dependency, only during design time.
 
 From the command line, copy the tlb to a central location (system32 or wow64 
 folder) and register it.
 
 COPY OLEGUIDS.TLB %SystemRoot%\System32\
 REGTLIB %SystemRoot%\System32\OLEGUIDS.TLB
 
 In the VB6 IDE - project - references - browse - select the OLEGuids.tlb


 * SETUP.EXE - The program is currently distributed using setup2go, a very useful 
 and comprehensive installer program that builds a .exe installer. Youll have to 
 find a copy of setup2go on the web as it is now abandonware. Contact me
 directly for a copy. The file "install PzCPU Gauge 0.1.0.s2g" is the configuration 
 file for setup2go. When you build it will report any errors in the build.
 
 * HELP.CHM - the program documentation is built using the NVU HTML editor and 
 compiled using the Microsoft supplied CHM builder tools (HTMLHelp Workshop) and 
 the HTM2CHM tool from Yaroslav Kirillov. Both are abandonware but still do
 the job admirably. The HTML files exist alongside the compiled CHM file in the 
 HELP folder.
 
  Project References:

	VisualBasic for Applications  
	VisualBasic Runtime Objects and Procedures  
	VisualBasic Objects and Procedures  
	OLE Automation  
	vbRichClient6  
 
 <img width="445" height="359" alt="projectReferences" src="https://github.com/user-attachments/assets/51a33785-2e3c-4b32-9fd0-6c4b0ebb8531" />
 
 LICENCE AGREEMENTS:
 
 Copyright © 2023 Dean Beedell
 
 In addition to the GNU General Public Licence please be aware that you may use 
 any of my own imagery in your own creations but commercially only with my 
 permission. In all other non-commercial cases I require a credit to the 
 original artist using my name or one of my pseudonyms and a link to my site. 
 With regard to the commercial use of incorporated images, permission and a 
 licence would need to be obtained from the original owner and creator, ie. me.

 
<img width="800" height="491" alt="about" src="https://github.com/user-attachments/assets/3f543f52-928e-4e38-911a-9eaa185cafdb" />
