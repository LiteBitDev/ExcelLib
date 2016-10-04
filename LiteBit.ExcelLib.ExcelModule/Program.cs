using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using ExcelCom = Microsoft.Office.Interop.Excel;

namespace LiteBit.ExcelLib.ExcelModule
{
    class Program
    {
        #region Fields

        /// <summary>
        /// The m encoding
        /// </summary>
        private static Encoding mEncoding = Encoding.UTF8;

        #endregion Fields

        static void Main(string[] args)
        {
            //args:
            // 0 - encoding
            // 1 - excel path
            // 2 - save path
            // 3 - header prefix.

            //exit codes:
            //1 - invalid args.
            //2 - args do not contain valid information.
            //3 - workbook empty.
            //10 - overall exception
            //0 - success
            ExcelCom.Application xlApp = null;
            ExcelCom.Workbook xlWorkBook = null;
            try
            {
                if (args == null || args.Length < 4)
                {
                    Environment.ExitCode = 1;
                    return;
                }

                string encoding = args[0];
                if (string.IsNullOrEmpty(encoding) == false)
                {
                    mEncoding = Encoding.GetEncoding(encoding);
                }

                string path = args[1];
                string savePath = args[2];
                byte headerPrefix;
                if (string.IsNullOrEmpty(path) == true || File.Exists(path) == false || string.IsNullOrEmpty(savePath) == true ||
                    byte.TryParse(args[3], out headerPrefix) == false)
                {
                    Environment.ExitCode = 2;
                    return;
                }

                xlApp = new ExcelCom.Application();
                xlWorkBook = xlApp.Workbooks.Open(path, 0, true, 5, "", "", true, ExcelCom.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                if (xlWorkBook.Sheets == null || xlWorkBook.Sheets.Count == 0)
                {
                    Environment.ExitCode = 3;
                    return;
                }

                //create main header
                List<byte> rawData = new List<byte>();
                rawData.Add(headerPrefix);
                //encoding name and length
                byte[] encodingName = Encoding.UTF8.GetBytes(mEncoding.WebName);
                rawData.Add((byte)encodingName.Length);
                rawData.AddRange(encodingName);
                //workbook name and length
                byte[] wbName = mEncoding.GetBytes(xlWorkBook.Name);
                rawData.Add((byte)wbName.Length);
                rawData.AddRange(wbName);
                //main header end.

                for (int i = 1; i <= xlWorkBook.Sheets.Count; i++)
                {
                    ExcelCom.Worksheet xlWorkSheet = (ExcelCom.Worksheet)xlWorkBook.Worksheets[i];
                    ExcelCom.Range range = xlWorkSheet.UsedRange;
                    if (range == null)
                    {
                        //releasing objects...
                        ReleaseComObjects(xlWorkSheet);
                        continue;
                    }

                    //worksheet header
                    List<byte> wsRawData = new List<byte>();
                    byte[] wsName = mEncoding.GetBytes(xlWorkSheet.Name);
                    wsRawData.Add((byte)wsName.Length);
                    wsRawData.AddRange(wsName);
                    //rows and columns amount
                    wsRawData.AddRange(BitConverter.GetBytes((ushort)range.Columns.Count));
                    //data
                    for (int rowIndex = 1; rowIndex <= range.Rows.Count; rowIndex++)
                    {
                        List<byte> rowRawData = new List<byte>();
                        for (int columnIndex = 1; columnIndex <= range.Columns.Count; columnIndex++)
                        {
                            object cell = range.Cells[rowIndex, columnIndex];
                            string valueToSave = (cell == null || cell is ExcelCom.Range == false || ((ExcelCom.Range)cell).Value2 == null)
                            ? string.Empty
                            : ((ExcelCom.Range)cell).Value2.ToString();

                            byte[] cellData = mEncoding.GetBytes(valueToSave);
                            rowRawData.Add((byte)cellData.Length);
                            rowRawData.AddRange(cellData);
                        }

                        wsRawData.AddRange(BitConverter.GetBytes((uint)rowRawData.Count));
                        wsRawData.AddRange(rowRawData);
                    }

                    //print all the data...
                    rawData.AddRange(BitConverter.GetBytes((uint)wsRawData.Count));
                    rawData.AddRange(wsRawData);
                    ReleaseComObjects(xlWorkSheet);
                    ReleaseComObjects(range);
                }

                File.WriteAllBytes(savePath, rawData.ToArray());
            }
            catch (Exception)
            {
                Environment.ExitCode = 10;
            }
            finally
            {
                TryCloseExcel(xlWorkBook);
                TryCloseExcelApp(xlApp);
                ReleaseComObjects(xlWorkBook);
                ReleaseComObjects(xlApp);
                GC.Collect();
            }
        }

        #region Private Static Methods

        /// <summary>
        /// Releases the COM objects.
        /// </summary>
        /// <param name="obj">The object.</param>
        private static void ReleaseComObjects(object obj)
        {
            try
            {
                if (obj == null)
                {
                    return;
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
            }
            catch (Exception)
            {
                //just for protection.
            }
            finally
            {
                obj = null;
                GC.Collect();
            }
        }

        /// <summary>
        /// Tries the close excel.
        /// </summary>
        /// <param name="xlWorkBook">The xl work book.</param>
        private static void TryCloseExcel(ExcelCom.Workbook xlWorkBook)
        {
            try
            {
                if (xlWorkBook == null)
                {
                    return;
                }

                xlWorkBook.Close(false);
            }
            catch (Exception)
            {
                //just for protection.
            }
        }

        /// <summary>
        /// Tries the close excel application.
        /// </summary>
        /// <param name="xlApp">The xl application.</param>
        private static void TryCloseExcelApp(ExcelCom.Application xlApp)
        {
            try
            {
                if (xlApp == null)
                {
                    return;
                }

                xlApp.Quit();
                //wait for it...
                while (xlApp.Quitting == true)
                {
                    Thread.Sleep(1);
                }

            }
            catch (Exception)
            {
                //just for protection.
            }
        }

        #endregion Private Static Methods
    }
}
