using System;
using System.Data;
using System.Text;

namespace LiteBit.ExcelLib.ExcelModuleCaller.Parser
{
    internal class ExcelRawDataParser
    {
        #region Fields

        /// <summary>
        /// Used for version testing.
        /// </summary>
        public const byte mHEADER_PREFIX = 0xA1;

        #endregion Fields

        #region Public Methods

        /// <summary>
        /// Converts the data to data set.
        /// </summary>
        /// <param name="data">The data.</param>
        /// <param name="result">The result.</param>
        /// <returns></returns>
        public bool ConvertDataToDataSet(byte[] data, out DataSet result)
        {
            result = null;
            try
            {
                //check if the data prefix is right.
                int offset = 0;
                byte headerPrefix = data[offset];
                if (headerPrefix != mHEADER_PREFIX)
                {
                    //File version not supported.
                    return false;
                }

                offset++;
                //get encoding name and type.
                byte encodingLength = data[offset];
                offset++;
                string encodingSaveName = Encoding.UTF8.GetString(data, offset, encodingLength);
                Console.WriteLine(encodingSaveName);
                Encoding encoder = Encoding.GetEncoding(encodingSaveName);
                offset += encodingLength;
                byte wbNameLength = data[offset];
                offset++;
                string wbNameSaved = encoder.GetString(data, offset, wbNameLength);
                offset += wbNameLength;
                result = new DataSet(wbNameSaved);
                //run on all the data...
                while (offset < data.Length)
                {
                    //ws length
                    //ws data
                    //get the entire ws len
                    uint wsDataLen = BitConverter.ToUInt32(data, offset);
                    offset += sizeof(uint);

                    byte[] innerData = new byte[wsDataLen];
                    Buffer.BlockCopy(data, offset, innerData, 0, (int)wsDataLen);
                    offset += (int)wsDataLen;
                    int innerOffset = 0;

                    //get ws name
                    byte wsNameLen = innerData[innerOffset];
                    innerOffset++;
                    string wsNameSaved = encoder.GetString(innerData, innerOffset, wsNameLen);
                    innerOffset += wsNameLen;
                    //get number of columns to create.
                    ushort wsColumnCount = BitConverter.ToUInt16(innerData, innerOffset);
                    innerOffset += sizeof(ushort);
                    //create data table
                    DataTable table = new DataTable(wsNameSaved);
                    for (int c = 0; c < wsColumnCount; c++)
                    {
                        table.Columns.Add(new DataColumn(c.ToString()));
                    }

                    //run on the ws data
                    while (innerOffset < wsDataLen)
                    {
                        //row length
                        //row data
                        uint rowLen = BitConverter.ToUInt32(innerData, innerOffset);
                        innerOffset += sizeof(uint);
                        int tempRowLimit = innerOffset + (int)rowLen;
                        DataRow row = table.NewRow();
                        int columnIndex = 0;
                        while (innerOffset < tempRowLimit)
                        {
                            //cell length
                            //cell encoder data - string.
                            byte cellLen = innerData[innerOffset];
                            innerOffset++;
                            string cellValue = encoder.GetString(innerData, innerOffset, cellLen);
                            innerOffset += cellLen;
                            row[columnIndex] = cellValue;
                            columnIndex++;
                        }

                        table.Rows.Add(row);
                    }

                    result.Tables.Add(table);
                }

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        #endregion Public Methods
    }
}
