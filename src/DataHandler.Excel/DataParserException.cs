using System;

namespace DataHandler.Excel
{
    public class DataParserException
        : Exception
    {
        public DataParserException(string message)
            : base(message)
        { }
    }
}