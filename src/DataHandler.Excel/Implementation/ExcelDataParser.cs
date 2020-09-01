using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using DataHandler.Excel.Models;
using DocumentFormat.OpenXml.Packaging;
using Tools.Excel;

namespace DataHandler.Excel.Implementation
{
    public class ExcelDataParser<TModel>
        : IDataParser<TModel>, IDisposable
        where TModel : class, new()
    {

        private IDataObtainer _dataObtainer;
        private readonly IDataWriter _dataWriter;

        private List<Stream> _fileStreamList 
           = new List<Stream>(); 
       
        private List<TModel> _result;

        public List<TModel> Result => _result;

        public ExcelDataParser(
            IDataObtainer dataObtainer, 
            IDataWriter dataWriter)
        {
            _dataObtainer = dataObtainer ?? throw new InvalidOperationException($"dataObtainer is null");
            _dataWriter = dataWriter ?? throw new InvalidOperationException($"dataWriter is null");
        }
        
        
        public IDataWriter Writer => _dataWriter;

        /// <summary>
        /// Связывание IDataParser<TModel> с физическим .xlsx файлом на диске
        /// Возможно связать IDataParser<TModel> с несколькими файлами, имеющими одинаковый формат столбцов
        /// Для этого необходимо вызвать метод Bind несколько раз для одного экземпляра данного класса
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        /// <exception cref="FileNotFoundException"></exception>
        public IDataParser<TModel> Bind(string filePath, bool openToWrite = false)
        {
            var fileInfo = new FileInfo(filePath);
            
            if(!fileInfo.Exists)
                throw new FileNotFoundException($"{filePath} was not found");

            this._fileStreamList.Add(
                openToWrite 
                    ? fileInfo.Open(FileMode.Open, FileAccess.ReadWrite) 
                    : fileInfo.OpenRead());

            return this;
        }
        
        
        public IDataParser<TModel> BindCopy(string filePath, string copyPath, bool openToWrite = false)
        {
            if(!File.Exists(filePath))
                throw new FileNotFoundException($"{filePath} was not found");
            
            if(File.Exists(copyPath) && openToWrite) File.Delete(copyPath);
            File.Copy(filePath, copyPath);
            
            var fileInfo = new FileInfo(copyPath);
            
            
            this._fileStreamList.Add(
                openToWrite 
                    ? fileInfo.Open(FileMode.Open, FileAccess.ReadWrite) 
                    : fileInfo.OpenRead());

            return this;
        }

        
        
        public IDataParser<TModel> Bind(Stream stream)
        {
            
            if(!this._fileStreamList.Contains(stream))
                this._fileStreamList.Add(stream);
            
            return this;
        }

        
        public IDataParser<TModel> ExtractData(bool disposeAfterReading = true, string sheetName = "") => ExtractData(null, disposeAfterReading, sheetName);
        
        
        /// <summary>
        /// Получить данные из .xlsx файла с возможностью фильтрации на выходе
        /// </summary>
        /// <param name="filter">Коллбэк метод, используемый для фильтрации данных на выходе</param>
        /// <returns></returns>
        /// <exception cref="NullReferenceException"></exception>
        public IDataParser<TModel> ExtractData(Func<TModel, bool> filter, bool disposeAfterReading = true, string sheetName = "")
        {
            if(!this._fileStreamList.Any())
                throw new DataParserException("Use Bind method before extracting data from file or stream");
                
            var filterValues = ExtractFilterValues(typeof(TModel));
            
            var resultDataTable = new DataTable();
            
            
            Stopwatch sw = new Stopwatch();
            sw.Start();
            
            foreach (var stream in this._fileStreamList)
            {
                var dataTableTask = _dataObtainer
                    .ObtainTable(
                        stream, 
                        filterValues, 
                        isEditable:false, 
                        //disposeStreamAfterReading: true,
                        sheetName: sheetName);
                
               dataTableTask.Combine(ref resultDataTable);
                
            }
            sw.Stop();
            var time = sw.ElapsedMilliseconds;
            Console.WriteLine("ExtractData " + time + " ms passed");
            sw.Reset();
            
            /*if(disposeAfterReading)
                _fileStreamList.ForEach(ClearStream);*/
            
            var resultModelList = resultDataTable.DataRows
                .AsParallel()
                .Select(ProjectDataRowToModel);
                
            if(filter != null)
                resultModelList = resultModelList.Where(filter);
            
            this._result =  resultModelList.ToList();

            return this;
        }
        
        public IDataParser<TModel> ExtractDataParallel(string sheetName = "") => ExtractDataParallel(null, sheetName);
        
        public IDataParser<TModel> ExtractDataParallel(Func<TModel, bool> filter, string sheetName = "")
        {
            if(!this._fileStreamList.Any())
                throw new DataParserException("Use Bind method before extracting data from file or stream");
                
            var filterValues = ExtractFilterValues(typeof(TModel));
            
            var resultDataTables = new BlockingCollection<DataTable>();
            
            Stopwatch sw = new Stopwatch();
            sw.Start();
            
            Parallel.ForEach(this._fileStreamList, async stream =>
            {
                var dataTable = await _dataObtainer
                    .ObtainTableAsync(
                        stream, 
                        filterValues, 
                        isEditable:false,
                        sheetName: sheetName);
                
                resultDataTables.TryAdd(dataTable);
            });
            
            while (!resultDataTables.Any()) { }
            
            //---------Clear Streams------------
            this.ClearStreamList();
            //----------------------------------
            
            
            sw.Stop();
            var time = sw.ElapsedMilliseconds;
            Console.WriteLine("ExtractDataParallel " + time + " ms passed");

            
            var dataRows = resultDataTables.SelectMany(dt => dt.DataRows);
            
            var resultModelList = dataRows
                .AsParallel()
                .Select(ProjectDataRowToModel);
                
            if(filter != null)
                resultModelList = resultModelList.Where(filter);
            
            this._result =  resultModelList.ToList();

            return this;
        }
        

        public IDataParser<TModel> WriteData(IEnumerable<TModel> data, RowStyles rowStyle = RowStyles.Simple, bool closeDocumentsAfterWriting = true, string sheetName = "")
        {
            if(!this._fileStreamList.Any())
                throw new DataParserException("Use Bind method before writing data to file or stream");

            var filterSets = GetFilterSetList(data);

            foreach (var stream in this._fileStreamList)
            {
                var dataTable = _dataWriter
                    .WriteToTable(
                        stream,
                        filterSets,
                        isEditable: true,
                        rowStyle, 
                        sheetName);
            }
            
/*            if(closeDocumentsAfterWriting)
                _fileStreamList.ForEach(ClearStream);*/
            
            return this;
        }


        public IDataParser<TModel> UpdateCells(IEnumerable<CellTemplate> cellTemplates, bool closeDocumentsAfterWriting = true)
        {
            if(!this._fileStreamList.Any())
                throw new DataParserException("Use Bind method before extracting data from file or stream");
            
            foreach (var stream in this._fileStreamList)
            {
                
                using (var document = SpreadsheetDocument.Open(stream, false))
                {
                    _dataWriter.UpdateCells(stream, cellTemplates);
                }
            }
            
            /*if(closeDocumentsAfterWriting)
                _fileStreamList.ForEach(ClearStream);*/
            
            return this;

        }
        
        private List<FilterSet> GetFilterSetList(IEnumerable<TModel> data)
        {
            var filterSet = new List<FilterSet>();
            
            var headerMapValues = ExtractFilterValues(typeof(TModel));
            
            filterSet.Add(new FilterSet
            {
                FilterValues = headerMapValues
            });
            
            foreach (var dataRow in data)
            {
                var filterValues = new List<FilterValue>();
                
                foreach (var fv in headerMapValues)
                {
                    var value = fv.PropertyInfo.GetValue(dataRow);
                    
                    filterValues.Add(new FilterValue
                    {
                        Label = value?.ToString() ?? string.Empty,
                        PropertyInfo = fv.PropertyInfo,
                        Type = fv.Type
                    });
                }

                filterSet.Add(
                    new FilterSet
                    {
                        FilterValues = filterValues
                    });

            }

            return filterSet;
        }

        public List<TResultModel> Each<TResultModel>(Func<TModel, TResultModel> callback)
        {
            if(!this._result.Any())
                throw new DataParserException($"Result is not evaluated. Use ExtractData before calling Each method");

            var result = this._result.Select(callback);

            return result.ToList();
        }
        
/*        public List<TResultModel> EachAsync<TResultModel>(Func<Task<TModel>, TResultModel> callback)
        {
            if(!this._result.Any())
                throw new DataParserException($"Result is not evaluated. Use ExtractData before calling Each method");


            foreach (var r in this._result)
            {
                callback(r)
            }
            
            var result = this._result.Select(callback);

            return result.ToList();
        }*/

        public async Task<List<TResultModel>> EachAsync<TResultModel>(Func<TModel, TResultModel> callback)
        {
            if(!this._result.Any())
                throw new DataParserException($"Result is not evaluated. Use ExtractData before calling Each method");

            var taskList = this._result.Select(model => 
                Task.Run(() => callback(model)));

            var result = await Task.WhenAll(taskList);

            return result.ToList();
        }
        
        private TModel ProjectDataRowToModel(DataRow row)
        {
            var model = new TModel();

            row.DataCells.ForEach(cell => cell.PropertyInfo.SetValue(model, cell.Value));
            
            return model;
        }

        private List<FilterValue> ExtractFilterValues(Type dataModelType)
        {
            var result = new List<FilterValue>();

            PropertyInfo[] props = dataModelType
                .GetProperties();
            
            foreach (PropertyInfo prop in props)
            {
                object[] attrs = prop.GetCustomAttributes(true);
                
                foreach (var attr in attrs)
                {
                    var dataColumnAttr = attr as DataColumnAttribute;
                    
                    if (dataColumnAttr != null)
                    {
                        result.Add(new FilterValue
                        {
                            PropertyInfo = prop,
                            Label = dataColumnAttr.ColumnName
                        });
                    }
                }
            }
            return result;
        }


        public void Dispose()
        {
            Clear();
        }
        
        public IDataParser<TModel> Clear()
        {
            ClearStreamList();
            _result = null;
            return this;
        }
        
        private void ClearStreamList()
        {
            foreach (var stream in this._fileStreamList.ToList())
            {
                if (stream != null)
                {
                    stream.Close();
                    stream.Dispose();
                }
                
                _fileStreamList.Remove(stream);
            }
        }

    }
}