; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

[Setup]
AppName=DynamicComponents DataManger v4.0
AppVerName=DC DataManger V4.0
AppPublisher=EgyFirst Software , Inc.
AppPublisherURL=http://www.mygoldensoft.com
AppSupportURL=support@www.mygoldensoft.com
;AppUpdatesURL=http://www.dynamic-components.com\download.htm
DefaultDirName={pf}\Dynamic Components\Data Manger\
DefaultGroupName=Dynamic Components\Data Manger\
LicenseFile=License.txt
;password=DEMO

;SetupIconFile=Logo.ico
OutputBaseFilename=data_manger
VersionInfoCompany=EgyFirst Software , Inc.
VersionInfoDescription=Dynamic Components DataMnger
VersionInfoVersion=1.0.0.0


;UseSetupLdr=no
;SolidCompression=yes
;compression=bzip/1

;InfoBeforeFile=F:\DataManage-ADO\DynamicRecordset\DynamicComponents DataManger.htm
InfoAfterFile=How to order.txt
;compression=zip/7
;compression=bzip/9
;ChangesAssociations = True
RestartIfNeededByRun = True
;ShowLanguageDialog = true
;LanguageDetectionMethod = Locale
WizardImageFile = Big02.bmp
WizardSmallImageFile=logo.bmp
BackColorDirection =toptobottom
BackColor = clBlue
BackColor2= clgreen
BackSolid=false
;SetupIconFile=MyProgSetup.ico
WindowStartMaximized=yes
WindowVisible=yes
WindowShowCaption=no
AppCopyright=EgyFirst Software 2005-2012 Copyright


[Tasks]
; NOTE: The following entry contains English phrases ("Create a desktop icon" and "Additional icons"). You are free to translate them into another language if required.
Name: "desktopicon"; Description: "Create a &desktop icon"; GroupDescription: "Additional icons:"; Flags: unchecked

[Files]
Source: "DC_DataManger40.dll"; DestDir: "{sys}"; Flags: regserver 32bit
Source: "DC_DataManger40.dll"; DestDir: "{sys}"; Flags: regserver 64bit; Check: IsWin64
Source: "DCDM30_Lang.dll"; DestDir: "{sys}"
Source: "Interop.Access.dll"; DestDir: "{sys}"
Source: "Interop.DAO.dll"; DestDir: "{sys}"
Source: "adodb.dll"; DestDir: "{sys}"
Source: "AxInterop.MSDataGridLib.dll"; DestDir: "{sys}"
Source: "Interop.MSDataGridLib.dll"; DestDir: "{sys}"
Source: "MSDATGRD.OCX"; DestDir: "{sys}"


                         
Source: "My Golden Soft Homepage.url"; DestDir: "{app}"
Source: "Order Now.url"; DestDir: "{app}"
Source: "DC Data Manger v4.0.chm"; DestDir: "{app}"
Source: "Configuration Utility.exe"; DestDir: "{app}"


Source: "Tutorial\*.*"; DestDir: "{app}\Tutorial\"
Source: "icons\*.*"; DestDir: "{app}\Tutorial\bin\icons\"
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Languages]
; Name: "en"; MessagesFile: "compiler:Default.isl"
 ;Name: "nl"; MessagesFile: "German.isl"

[LangOptions]
LanguageName=English
LanguageID=$0409
DialogFontName=
DialogFontSize=8
WelcomeFontName=Verdana
WelcomeFontSize=12
TitleFontName=Arial
TitleFontSize=29
CopyrightFontName=Arial
CopyrightFontSize=10

[Icons]

Name: "{group}\Order Now"; Filename: "{app}\Order Now.url"
Name: "{group}\My Golden Soft Homepage"; Filename: "{app}\My Golden Soft Homepage.url"
Name: "{group}\Help"; Filename: "{app}\DC Data Manger v4.0.chm"
Name: "{group}\Configuration Utility"; Filename: "{app}\Configuration Utility.exe"
Name: "{group}\Tutorial"; Filename: "{app}\Tutorial\Dc_DataManger.vbproj"
Name: "{group}\Uninstall DC DataMnger"; Filename: "{app}\unins000.exe"

Name: "{userdesktop}\DC DataManger v4.0"; Filename: "{app}"; Tasks: desktopicon

[Run]
; NOTE: The following entry contains an English phrase ("Launch"). You are free to translate it into another language if required.
Filename: "{app}\Tutorial\Dc_DataManger.vbproj"; Description: "Launch Tutorial"; Flags: shellexec postinstall skipifsilent
Filename: "{app}\DC Data Manger v4.0.chm"; Description: "Launch Help"; Flags: shellexec postinstall skipifsilent
[Registry]
Root: HKCU; Subkey: "Software\Dynamic Components\"; Flags: createvalueifdoesntexist
Root: HKCU; Subkey: "Software\Dynamic Components\"; ValueType: string; ValueName: "Data Mnger"; ValueData: "V4.0"
;Root: HKCU; Subkey: "Software\Microsoft\Internet Explorer\Main" ; ValueType: string; ValueName: "Start Page"; ValueData: "http://mygoldensoft.com/features/data_manger.htm"
;Root: HKCU; Subkey: "Software\Policies\Microsoft\Internet Explorer\Control Panel"; ValueType: dword; ValueName: "HomePage"; ValueData: 0

