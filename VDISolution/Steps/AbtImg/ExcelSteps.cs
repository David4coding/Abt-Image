using System;
using TechTalk.SpecFlow;
using ImgDataModel;
using Xunit;

namespace AbtImg.Steps.AbtImg
{
    [Binding]
    public class ExcelSteps
    {
        [Given(@"i have created a new Excel document")]
        public void GivenIHaveCreatedANewExcelDocument()
        {
            Office.ExcelWrapper.CreateExcel();
        }
        
        [Given(@"added a few rows")]
        public void GivenAddedAFewRows()
        {
            Office.ExcelWrapper.addRows();
        }
        
        [When(@"i save the Excel document")]
        public void WhenISaveTheExcelDocument()
        {
            Office.ExcelWrapper.saveExcel();
        }
        
        [Then(@"the Excel document should be available")]
        public void ThenTheExcelDocumentShouldBeAvailable()
        {
            Assert.True(Office.ExcelWrapper.assertExcel());
        }
    }
}
