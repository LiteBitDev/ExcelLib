using LiteBit.ExcelLib.ExcelModuleCaller.Enums;

namespace LiteBit.ExcelLib.ExcelModuleCaller.Helpers
{
    public static class ExitCodeInterperter
    {
        /// <summary>
        /// Interprets the specified exit code.
        /// </summary>
        /// <param name="exitCode">The exit code.</param>
        /// <returns></returns>
        public static ExcelParseError Interpret(int exitCode)
        {
            //exit codes:
            //1 - invalid args.
            //2 - args do not contain valid information.
            //3 - workbook empty.
            //10 - overall exception
            //0 - success

            switch (exitCode)
            {
                case 1:
                    return ExcelParseError.InvalidArgs;
                case 2:
                    return ExcelParseError.InvalidArgsInfo;
                case 3:
                    return ExcelParseError.FileInvalid;
                case 10:
                    return ExcelParseError.Exception;
                default:
                    return ExcelParseError.Unknown;
            }            
        }
    }
}
