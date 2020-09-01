﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using DataHandler.Excel.Models;

namespace DataHandler.Excel
{
    public interface IDataParser<TModel>
        where  TModel : class, new()
    {
        List<TModel> Result { get; }
        
        IDataWriter Writer { get; }

        IDataParser<TModel> Bind(string filePath, bool openToWrite = false);

        IDataParser<TModel> Bind(Stream stream);

        IDataParser<TModel> BindCopy(string filePath, string copyPath, bool openToWrite = false);

        IDataParser<TModel>  ExtractData(bool disposeAfterReading = true, string sheetName = "");
        
        IDataParser<TModel>  ExtractData(Func<TModel, bool> filter, bool disposeAfterReading = true, string sheetName = "");

        
        IDataParser<TModel> ExtractDataParallel(string sheetName = "");
        IDataParser<TModel> ExtractDataParallel(Func<TModel, bool> filter, string sheetName = "");

        IDataParser<TModel> WriteData(IEnumerable<TModel> data, RowStyles rowStyle = RowStyles.Simple,
            bool closeDocumentsAfterWriting = true, string sheetName = "");

        IDataParser<TModel> UpdateCells(IEnumerable<CellTemplate> cellTemplates, bool closeDocumentsAfterWriting = true);
        
        List<TResultModel> Each<TResultModel>(Func<TModel, TResultModel> callback);

        Task<List<TResultModel>> EachAsync<TResultModel>(Func<TModel, TResultModel> callback);
        
        IDataParser<TModel> Clear();
    }
}