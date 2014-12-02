using System;
using System.IO;
using Infotron.Parsing;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExpectorTests
{
    [TestClass]
    public class ExpectorTests
    {
        ExcelFormulaParser P;

        [TestInitialize]
        public void setUp()
        {
            P = new ExcelFormulaParser();
        }         


        [TestMethod]
        public void TestWithOneBranch()
        {
            string Formula = "IF(B5<20,\"OK\")";
            Assert.AreEqual(true,  P.IsTestFormula(Formula));   
        }

        [TestMethod]
        public void TestWithTwoBranches()
        {
            string Formula = "IF(B5<20,\"OK\", \"NOT OK\")";
            Assert.AreEqual(true, P.IsTestFormula(Formula));
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void NotATestFormulaShoulThrowException()
        {
            string Formula = "SUM(A1:A7)";
            int result = P.GetPassingBranch(Formula);
        }




        [TestMethod]
        public void FirstBranchIsPassing()
        {
            string Formula = "IF(B5<20,\"OK\")";
            Assert.AreEqual(0, P.GetPassingBranch(Formula));
        }

        [TestMethod]
        public void FirstBranchIsNotPassing()
        {
            string Formula = "IF(B5<20,\"ERROR\",\"OK\")";
            Assert.AreEqual(1, P.GetPassingBranch(Formula));
        }

        [TestMethod]
        public void FormulaAsResultIsPassingBranch()
        {
            string Formula = "IF(SUM(K21:M21)=SUM(O17:O19), SUM(O17:O19), \"error\")";
            Assert.AreEqual(0, P.GetPassingBranch(Formula));
        }

        [TestMethod]
        public void FirstBranchIsNotPassingComplexerFormula()
        {
            string Formula = "IF(SUM(B5:B12)<20,\"ERROR\",\"OK\")";
            Assert.AreEqual(1, P.GetPassingBranch(Formula));
        }

        [TestMethod]
        public void RetrieveCondition()
        {
            string Formula = "IF(B5<20,\"ERROR\",\"FOUT\")";
            Assert.AreEqual("B5<20", P.GetCondition(Formula));
        }

        [TestMethod]
        public void RetrieveRangeCondition()
        {
            string Formula = "IF(SUM(B5:B12)<20,\"ERROR\",\"FOUT\")";
            Assert.AreEqual("SUM(B5:B12)<20", P.GetCondition(Formula));
        }

        [TestMethod]
        public void SplitCondition()
        {
            string Condition = "SUM(B5:B12)<20";
            Assert.AreEqual("SUM(B5:B12)", P.Split(Condition)[0]);
            Assert.AreEqual("<", P.Split(Condition)[1]);
            Assert.AreEqual("20", P.Split(Condition)[2]);
        }

        [TestMethod]
        public void SplitSingelItem()
        {
            string Condition = "A6";
            Assert.AreEqual("A6", P.Split(Condition)[0]);
        }



        //--------------------- Moving tests
        [TestMethod]
        public void MoveTestFormula()
        {
            string Formula = "IF(SUM(B5:B12)<Sheet2!A5,\"ERROR\",\"FOUT\")";
            Assert.AreEqual("SUM(Sheet1!B5:B12)<Sheet2!A5", P.GetCondition(Formula, "Sheet1"));
        }

    }
}
