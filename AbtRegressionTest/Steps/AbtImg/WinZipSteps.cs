using System;
using TechTalk.SpecFlow;
using ImgDataModel;
using Xunit;

namespace AbtRegressionTest.Steps.AbtImg
{
    [Binding]
    public class WinZipSteps
    {
        private bool result = false;

        [Given(@"i have created WinZip")]
        public void GivenIHaveCreatedWinZip()
        {
            WinZip.WritteBinaryFile();
        }
        
        [Given(@"added a file")]
        public void GivenAddedAFile()
        {
            WinZip.addToZip();
        }
        
        [When(@"i save the ZipFile")]
        public void WhenISaveIt()
        {
            result = WinZip.CheckZip(WinZip.path + "VDIZip.zip");
        }
        
        [Then(@"the ZipFile should be compressed")]
        public void ThenTheZipFileShouldBeCompressed()
        {
            Assert.True(result);
        }
    }
}
