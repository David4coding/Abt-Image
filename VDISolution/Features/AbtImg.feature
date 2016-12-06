Feature: AbtImg
		 Abt's Image BaseLine and Regression testing
@WinZip
Scenario: WinZip Test
	Given i have created WinZip 
	And   added a file
	When  i save the ZipFile
	Then  the ZipFile should be compressed 

@Word
Scenario: Microsoft Office Test
	Given i have created a new Word document
	And   added a new paragraph
	When  i save the Word document
	Then  the Word document should be available

@Excel
Scenario: Microsoft Excel Test
	Given i have created a new Excel document
	And   added a few rows 
	When  i save the Excel document
	Then  the Excel document should be available

@PowerPoint
Scenario: Microsoft PowerPoint Test
	Given i have created a new PowerPoint document
	And   added some text
	When  i save the PowerPoint it
	Then  the PowerPoint document should be available

@AccessDB
Scenario: Microsoft Access DB Test
	Given i have created a new Access DB
	Then  the Access DB should be available

@Outlook
Scenario: Microsoft Outlook Test
	Given   i opened Outlook
	Then    Outlook is logged with my Windows credentials

@SkypeForBusiness
Scenario: Skype For Business Test
	Given i have logged to Windows 
	Then  Skype for Business opened with my user

@JAVA
Scenario: JAVA Test
	Given i opened the windows CMD and input java -version
	Then  i can cofirm the Java Version is correct

@FlashPlayer
Scenario: Flash Player Test
	Given i see FlashPlayer is installed
	Then  i can see Flash's registry

@SilverLight
Scenario: SilverLight Test
	Given i see SilverLight is installed
	Then  i can see Silverlight's registry

@Symantec
Scenario: Symantec Active Scan Test
	Given Symantec active scan is running

