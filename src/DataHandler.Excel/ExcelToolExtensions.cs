using DataHandler.Excel.Implementation;
using Microsoft.Extensions.DependencyInjection;

namespace DataHandler.Excel
{
    public static class ExcelToolExtensions
    {
        public static void RegisterExcelTool(this IServiceCollection services)
        {
            services.AddTransient<IDataObtainer, OpenXmlDataObtainer>();
            services.AddTransient<IDataWriter, OpenXmlDataWriter>();
            services.AddTransient(typeof(IDataParser<>), typeof(ExcelDataParser<>));
        }
    }
}