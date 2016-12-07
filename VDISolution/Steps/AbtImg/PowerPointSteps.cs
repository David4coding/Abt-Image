using System;
using TechTalk.SpecFlow;
using ImgDataModel;
using Xunit;

namespace AbtImg.Steps.AbtImg
{
    [Binding]
    public class PowerPointSteps
    {
        [Given(@"i have created a new PowerPoint document")]
        public void GivenIHaveCreatedANewPowerPointDocument()
        {
            Office.PowerPointWrapper.CreatePowerPoint();
        }
        
        [Given(@"added some text")]
        public void GivenAddedSomeText()
        {
            Office.PowerPointWrapper.addText();
        }
        
        [When(@"i save the PowerPoint it")]
        public void WhenISaveThePowerPointIt()
        {
            Office.PowerPointWrapper.savePowerPoint();
        }
        
        [Then(@"the PowerPoint document should be available")]
        public void ThenThePowerPointDocumentShouldBeAvailable()
        {
            Assert.True(Office.PowerPointWrapper.assertPowerPoint());
        }
    }
}
