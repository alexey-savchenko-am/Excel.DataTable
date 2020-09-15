using System;

namespace Excel.DataTable.Implementation
{
    public interface IExcelDataWorker
    {
        IDataParser<TModel> SafeExecute<TModel>(Func<IDataParser<TModel>> action)
            where TModel : class, new();
    }
    
    public class ExcelDataWorker
        : IExcelDataWorker
    {

        public IDataParser<TModel> SafeExecute<TModel>(Func<IDataParser<TModel>> action)
            where TModel : class, new()
        {
            IDataParser<TModel> dataParser = null;
            
            try
            {
                dataParser = action.Invoke();
            }
            finally
            {
                dataParser?.Clear();
            }

            return dataParser;
        }
        
    }
}