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
