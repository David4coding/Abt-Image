using System;
using TechTalk.SpecFlow;
using ImgDataModel;
using Xunit;

namespace AbtImg.Steps.AbtImg
{
    [Binding]
    public class OutlookSteps
    {
        [Given(@"i opened Outlook")]
        public void GivenIOpenedOutlook()
        {
            Office.OutlookWrapper.hasActiveDirectoryCredentials();
        }
        
        [Then(@"Outlook is logged with my Windows credentials")]
        public void ThenOutlookIsLoggedWithMyWindowsCredentials()
        {
            Assert.True(Office.OutlookWrapper.assertOutlook());
        }
    }
}
