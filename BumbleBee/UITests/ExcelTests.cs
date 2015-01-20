using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Excel = Microsoft.Office.Interop.Excel;

namespace UITests
{
    [TestClass]
    public class ExcelTests
    {
        private static Excel.Application excel;

        private Excel.Workbook wb;

        [ClassInitialize()]
        public static void ClassInit(TestContext context)
        {
            excel = new Excel.Application {
                Visible = false
            };

        }

        [TestInitialize()]
        public void Initialize() { }

        [TestCleanup()]
        public void Cleanup()
        {
            if (wb != null)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(wb);
                wb = null;
            }
        }

        [ClassCleanup()]
        public static void ClassCleanup()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excel);
        }

        [TestMethod]
        public void TestMethod1()
        {
        }

        [TestMethod]
        [DeploymentItem("Testfiles/EmptyBook.xlsx")]
        public void TestInitializeBB()
        {
            
        }
    }
}
