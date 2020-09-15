using System;

namespace Excel.DataTable
{
    public class DataParserException
        : Exception
    {
        public DataParserException(string message)
            : base(message)
        { }
    }
}