using System;
using TechTalk.SpecFlow;
using ImgDataModel;
using Xunit;

namespace VDISolution.Steps.AbtImg
{
    [Binding]
    public class AccessDBSteps
    {
        [Given(@"i have created a new Access DB")]
        public void GivenIHaveCreatedANewAccessDB()
        {
            Office.AccessWrapper.CreateAccess();
        }
        
        [Then(@"the Access DB should be available")]
        public void ThenTheAccessDBShouldBeAvailable()
        {
          Assert.True(Office.AccessWrapper.assertAccessDB());
        }
    }
}
