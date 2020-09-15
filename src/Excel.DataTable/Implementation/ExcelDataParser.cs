using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using Excel.DataTable.Models;

namespace Excel.DataTable.Implementation
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
            _dataObtainer = dataObtainer ?? throw new ArgumentNullException($"dataObtainer is null");
            _dataWriter = dataWriter ?? throw new ArgumentNullException($"dataWriter is null");
        }
        
        
        public IDataWriter Writer => _dataWriter;


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

        
        public IDataParser<TModel> ExtractData(string sheetName = "") => ExtractData(null, sheetName);
        
        public IDataParser<TModel> ExtractData(Func<TModel, bool> filter, string sheetName = "")
        {
            if(!this._fileStreamList.Any())
                throw new DataParserException("Use Bind method before extracting data from file or stream");
                
            var filterValues = ExtractFilterValues(typeof(TModel));
            
            var resultDataTables = new BlockingCollection<Models.DataTable>();
            var spin = new SpinWait();
            
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
            
            while (resultDataTables.Count < _fileStreamList.Count) spin.SpinOnce();
            
            sw.Stop();
            var time = sw.ElapsedMilliseconds;
            Console.WriteLine("ExtractDataParallel " + time + " ms elapsed");
            
            var dataRows = resultDataTables.SelectMany(dt => dt.DataRows);
            
            var resultModelList = dataRows
                .AsParallel()
                .Select(ProjectDataRowToModel);
                
            if(filter != null)
                resultModelList = resultModelList.Where(filter);
            
            this._result =  resultModelList.ToList();

            return this;
        }
        

        public IDataParser<TModel> WriteData(IEnumerable<TModel> data, RowStyles rowStyle = RowStyles.Simple, bool keepDocumentsOpen = false, string sheetName = "")
        {
            if(!this._fileStreamList.Any())
                throw new DataParserException("Use Bind method before writing data to file or stream");
            
            if(!this._fileStreamList.All(s => s.CanWrite))
                throw new DataParserException("Can not write data to excel table. Use Bind method with flag openToWrite = true");
            
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
            
            //---------Clear Streams------------
            if(!keepDocumentsOpen)
                this.ClearStreamList();
            //----------------------------------
            
            return this;
        }


        public IDataParser<TModel> UpdateCells(IEnumerable<CellTemplate> cellTemplates, bool keepDocumentsOpen = false)
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