using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Resources;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using LiteBit.ExcelLib.ExcelModuleCaller.Enums;
using LiteBit.ExcelLib.ExcelModuleCaller.Helpers;
using LiteBit.ExcelLib.ExcelModuleCaller.Parser;
using LiteBit.ExcelLib.ExcelModuleCaller.Properties;

namespace LiteBit.ExcelLib.ExcelModuleCaller
{
    public class ExcelReader
    {
        #region Fields

        /// <summary>
        /// The execute name
        /// </summary>
        private readonly string mExecName = "EM.exe";

        /// <summary>
        /// The result file
        /// </summary>
        private readonly string mResultFile = "Output.dat";

        /// <summary>
        /// The m parser
        /// </summary>
        private readonly ExcelRawDataParser mParser;

        /// <summary>
        /// The excel_ parser
        /// </summary>
        private readonly byte[] EXCEL_PARSER_RAW;

        #endregion Fields

        #region C'tor

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelReader"/> class.
        /// </summary>
        public ExcelReader()
        {
            mParser = new ExcelRawDataParser();
            EXCEL_PARSER_RAW = Resources.LiteBit_ExcelLib_ExcelModule;
        }

        #endregion C'tor
			
        #region Public Methods

        /// <summary>
        /// Reads the specified file name.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="result">The result.</param>
        /// <returns></returns>
        public bool TryRead(string fileName, out DataSet result)
        {
            ExcelParseError error;
            return TryRead(fileName, out result, out error);
        }

        /// <summary>
        /// Reads the specified file name.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="result">The result.</param>
        /// <param name="error">The error.</param>
        /// <returns></returns>
        public bool TryRead(string fileName, out DataSet result, out ExcelParseError error)
        {
            Encoding enc = Encoding.UTF8;
            return TryRead(fileName, enc, out result, out error);
        }

        /// <summary>
        /// Reads the specified file name.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="encoding">The encoding.</param>
        /// <param name="result">The result.</param>
        /// <param name="error">The error.</param>
        /// <returns></returns>
        public bool TryRead(string fileName, Encoding encoding, out DataSet result, out ExcelParseError error)
        {
            result = null;
            error = ExcelParseError.Unknown;
            string tempFolder = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            try
            {
                Directory.CreateDirectory(tempFolder);


                //save it temporarly to a temp location.
                string tempEmPath = Path.Combine(tempFolder, mExecName);
                string tempOutputPath = Path.Combine(tempFolder, mResultFile);
                //write file
                File.WriteAllBytes(tempEmPath, EXCEL_PARSER_RAW);
                //run the file
                //args:
                // 0 - encoding
                // 1 - excel path
                // 2 - save path
                // 3 - header prefix.
                ProcessStartInfo startInfo = new ProcessStartInfo(tempEmPath)
                {
                    Arguments = string.Format("{0} \"{1}\" {2} {3}",
                    encoding.WebName, fileName, tempOutputPath, ExcelRawDataParser.mHEADER_PREFIX),
                    CreateNoWindow = true,
                    RedirectStandardError = true,
                    UseShellExecute = false
                };

                Process excelModuleProcess = Process.Start(startInfo);
                if (excelModuleProcess == null)
                {
                    error = ExcelParseError.ExternalModuleFail;
                    return false;
                }

                excelModuleProcess.WaitForExit();
                while (excelModuleProcess.HasExited == false)
                {
                    //wait for it...
                    Thread.Sleep(1);
                }

                if (excelModuleProcess.ExitCode != 0)
                {
                    error = ExitCodeInterperter.Interpret(excelModuleProcess.ExitCode);
                    return false;
                }

                //parse the data.
                byte[] resultData = File.ReadAllBytes(tempOutputPath);
                if (mParser.ConvertDataToDataSet(resultData, out result) == false)
                {
                    error = ExcelParseError.OutputDataInvalid;
                    return false;
                }

                return true;
            }
            catch (Exception)
            {
                error = ExcelParseError.Exception;
                return false;
            }
            finally
            {
                //delete the temp folder
                SafeFolderDelete(tempFolder);
            }
        }

        /// <summary>
        /// Reads the asnyc.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="encoding">The encoding.</param>
        /// <param name="callBack">The call back.</param>
        public void TryReadAsnyc(string fileName, Encoding encoding, Action<bool, DataSet, ExcelParseError> callBack)
        {
            Task.Factory.StartNew(() =>
            {
                DataSet result;
                ExcelParseError error;
                bool readResult = TryRead(fileName, encoding, out result, out error);

                if (callBack == null)
                {
                    return;
                }

                callBack(readResult, result, error);
            });
        }

        #endregion Public Methods

        #region Private Methods

        /// <summary>
        /// Safes the folder delete.
        /// </summary>
        /// <param name="folderPath">The folder path.</param>
        private void SafeFolderDelete(string folderPath)
        {
            try
            {
                if (Directory.Exists(folderPath) == false)
                {
                    return;
                }

                Directory.Delete(folderPath, true);
            }
            catch 
            {
                //just for protection.
            }
        }

        #endregion Private Methods
			
			
    }
}
