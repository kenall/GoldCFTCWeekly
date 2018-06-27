using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using GoldCFTCWeekly;
namespace UnitTestFetchData
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
        }

        [TestMethod]
        public void TestFetchData()
        {
            DateTime date = new DateTime(2015, 1, 7);
            DataFetch df = new DataFetch();
            bool ret = df.FetchData(out List<int> lst, date);
            Assert.IsFalse(ret);
        }
    }
}
