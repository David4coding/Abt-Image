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
            Symantec scan = new Symantec();
            Assert.IsTrue(scan.isSymantecActive());
        }

        [TestMethod]
        public void SilverLightTest()
        {
            SilverLight registry = new SilverLight();
            Assert.IsTrue(registry.isInstalled());
        }

        [TestMethod]
        public void FlashPlayerTest()
        {
            FlashPlayer registry = new FlashPlayer();
            Assert.IsTrue(registry.isInstalled());
        }
    }
}
