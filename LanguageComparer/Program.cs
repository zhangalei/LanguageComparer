﻿using Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LanguageComparer
{
    class Program
    {
        static void Main(string[] args)
        {
            var options = new CmdOptions();
            var isValid = CommandLine.Parser.Default.ParseArgumentsStrict(args, options);
            char[] trimChar = new char[] { ' ', '_', '\t' };
            if (isValid)
            {
                Dictionary<string, Item> correctDictionary = new Dictionary<string, Item>();
                List<CheckItem> checkList;

                var fileName = string.Format($"{Directory.GetCurrentDirectory()}\\{options.TestFile}");

                using (var stream = File.Open(fileName, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateOpenXmlReader(stream))
                    {
                        var result = reader.AsDataSet().Tables[0].AsEnumerable();
                        checkList = result.Select(
                            row => new CheckItem
                            {
                                ModuleName = row.Field<object>(0).ToString(),
                                EnglishPhrase = row.Field<object>(1).ToString(),
                                EnglishPhraseKey = row.Field<object>(1).ToString().Trim(trimChar).ToUpper(),
                                ForeignLanguagePhrase = row.Field<object>(2).ToString().Trim()
                            }
                            ).ToList();
                    }
                }

                fileName = string.Format($"{Directory.GetCurrentDirectory()}\\{options.DictionaryFile}");
                using (var stream = File.Open(fileName, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateOpenXmlReader(stream))
                    {
                        var expected = reader.AsDataSet().Tables[0].AsEnumerable().Where(r => r.ItemArray.All(v => v != null && v != DBNull.Value))
                            .GroupBy(x => x.Field<object>(0).ToString().Trim() + "|||"+ x.Field<object>(1).ToString().Trim(trimChar))
                            .Select(g => g.First());
                        correctDictionary = expected.ToDictionary<DataRow, string, Item>(
                            row => row.Field<object>(0).ToString().Trim() + "|||" + row.Field<object>(1).ToString().Trim(trimChar).ToUpper(),
                            row => new Item
                            {
                                GroupName = row.Field<object>(0).ToString(),
                                VariableName = row.Field<object>(1).ToString(),
                                ForeignLanguagePhrase = row.Field<object>(2).ToString().Trim(),
                                EnglishPhrase = row.Field<object>(3).ToString()
                            });

                    }
                }

                foreach (CheckItem word in checkList)
                {
                    if (correctDictionary.ContainsKey(word.ModuleName + "|||" + word.EnglishPhraseKey))
                    {
                        if (correctDictionary[word.ModuleName + "|||" + word.EnglishPhraseKey]?.ForeignLanguagePhrase == word.ForeignLanguagePhrase)
                        {
                            word.Result = Result.Match;
                            word.DictionaryItem = correctDictionary[word.ModuleName + "|||" + word.EnglishPhraseKey];
                            //Console.WriteLine($"{word.Result}: {word.ModuleName}, {word.EnglishPhrase}, {word.ArabicPhrase}");
                        }
                        else
                        {
                            word.Result = Result.NotMatch;
                            word.DictionaryItem = correctDictionary[word.ModuleName + "|||" + word.EnglishPhraseKey];
                            //Console.WriteLine($"{word.Result}: {word.ModuleName}, {word.EnglishPhrase}, {word.ArabicPhrase}");
                        }
                    }
                    else
                    {
                        word.Result = Result.NotFound;
                        //Console.WriteLine($"{word.Result}: {word.ModuleName}, {word.EnglishPhrase}, {word.ArabicPhrase}");
                    }
                }

                StringBuilder sb = new StringBuilder();

                var checkListOutput = from r in checkList
                                      select new
                                      {
                                          ModuleName = r.ModuleName,
                                          EnglishPhrase = r.EnglishPhrase,
                                          ForeignLanguagePhrase = r.ForeignLanguagePhrase,
                                          Result = Enum.GetName(typeof(Result), r.Result),
                                          ForeignLanguagePhraseInDictionary = r.DictionaryItem?.ForeignLanguagePhrase,
                                          GroupNameInDictionary = r.DictionaryItem?.GroupName,
                                          VariableNameInDictionary = r.DictionaryItem?.VariableName,
                                          EnglishPhraseInDictionary = r.DictionaryItem?.EnglishPhrase 
                                      };

                DataTable dt = checkListOutput.ConvertToDataTable();
                IEnumerable<string> columnNames = dt.Columns.Cast<DataColumn>().
                                                  Select(column => column.ColumnName);
                sb.AppendLine(string.Join(";", columnNames));

                foreach (DataRow row in dt.Rows)
                {
                    IEnumerable<string> fields = row.ItemArray.Select(field => field.ToString());
                    sb.AppendLine(string.Join(";", fields));
                }

                File.WriteAllText($"result_{DateTime.Now.ToString("yyyyMMdd_HH.mm.ss.fff", CultureInfo.InvariantCulture)}.csv", sb.ToString());
            }
            else
            {
                // Display the default usage information
                Console.WriteLine(options.GetUsage());
            }
        }

    }

    public class Item
    {
        public string GroupName { get; set; }
        public string VariableName { get; set; }
        public string ForeignLanguagePhrase { get; set; }
        public string EnglishPhrase { get; set; }
    }

    public class CheckItem
    {
        public string ModuleName { get; set; }
        public string EnglishPhrase { get; set; }
        public string ForeignLanguagePhrase { get; set; }
        public Result Result { get; set; } = Result.Unknown;
        public string EnglishPhraseKey { get; set; }
        public Item DictionaryItem { get; set; }
    }

    public enum Result
    {
        Unknown = 0,
        Match = 1,
        NotMatch =2,
        NotFound =3
    }
}
