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
