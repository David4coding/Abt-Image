using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VDISolution
{
    [TestClass]
    public class VDI
    {
        [TestMethod]
        public void WinZipTest()
        {
            WinZip zip = new WinZip();
            Assert.IsTrue( zip.isAvailable());
        }

        [TestMethod]
        public void JAVATest()
        {
            JAVA version = new JAVA();
            Assert.IsTrue(version.CheckVersion());
        }

        [TestMethod]
        public void SymantecTest()
        {
            Symantec version = new Symantec();
            Assert.IsTrue(version.isSymantecActive());
        }
    }
}
