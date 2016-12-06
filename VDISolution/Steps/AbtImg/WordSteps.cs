using System;
using TechTalk.SpecFlow;
using ImgDataModel;
using Xunit;


namespace VDISolution.Steps.AbtImg
{
    [Binding]
    public class WordSteps
    {
        [Given(@"i have created a new Word document")]
        public void GivenIHaveCreatedANewWordDocument()
        {
            Office.WordWrapper.CreateWord();
        }
        
        [Given(@"added a new paragraph")]
        public void GivenAddedANewParagraph()
        {
            Office.WordWrapper.addParagraph();
        }

        [When(@"i save the Word document")]
        public void WhenISaveIt()
        {
            Office.WordWrapper.saveFile();
        }

        [Then(@"the Word document should be available")]
        public void ThenTheWordDocumentShouldBeAvailable()
        {
            Assert.True(Office.WordWrapper.assertWordDoc());
        }
    }
}
