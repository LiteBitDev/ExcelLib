using System;
using System.Data;
using System.Diagnostics;
using System.Text;
using System.Threading;
using LiteBit.ExcelLib.ExcelModuleCaller;
using LiteBit.ExcelLib.ExcelModuleCaller.Enums;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UnitTester
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            string file = @"D:\LiteBit\Tut\ExpenseReport.xlsx";
            ExcelReader reader = new ExcelReader();

            DataSet data;
            ExcelParseError error;
            bool result = reader.TryRead(file, out data, out error);

            Assert.AreEqual(true, result);
        }

        [TestMethod]
        public void TestMethod2()
        {
            string file = @"D:\LiteBit\Tut\ExpenseReport.xlsx";
            ExcelReader reader = new ExcelReader();
            bool timer = false;

            DateTime start = DateTime.Now;
            DataSet data = null;
            reader.TryReadAsnyc(file, Encoding.UTF8, (b, set, arg3) =>
            {
                timer = true;
                data = set;
                Debug.Print("File {0}: {1}, {2}", file, b, arg3);
            });

            while (timer == false)
            {
                Thread.Sleep(100);
                Debug.Write(".");    
            }

            Debug.Print("Took: {0}", (DateTime.Now - start).TotalSeconds);
            if (data == null)
            {
                return;
            }

            //print all the data
            foreach (DataTable dataTable in data.Tables)
            {
                foreach (DataRow row in dataTable.Rows)
                {
                    StringBuilder builder = new StringBuilder();
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        object v = row[i];
                        string value = (string.IsNullOrEmpty(v.ToString()) == true) ? "NA" : v.ToString();
                        builder.Append(string.Format("{0}\t", value));
                    }

                    Debug.Print(builder.ToString());
                }
            }
        }
    }
}
