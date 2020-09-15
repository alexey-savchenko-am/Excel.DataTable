using Excel.DataTable.Implementation;

namespace Excel.DataTable
{
    public class DataParser<T>
        : ExcelDataParser<T>
        where T : class, new()
    {
        public DataParser()
            : base(new OpenXmlDataObtainer(), new OpenXmlDataWriter())
        { }
        
        public DataParser(IDataObtainer dataObtainer, IDataWriter dataWriter)
            : base(dataObtainer, dataWriter)
        { }
    }
}