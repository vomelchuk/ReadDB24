using Microsoft.Data.Sqlite;
using ReadDB24.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ReadDB24.Services
{
    internal class SqliteDBService
    {
        private readonly string dbPath;
        private readonly string[] ExcelLetters = new[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };

        public SqliteDBService(string dbPath)
        {
            this.dbPath = dbPath;
        }

        public async Task<(BindingList<OutputTableModel>, string)> GetRecordsByName(string name, string startDate, string endDate)
        {
            var errorMessage = string.Empty;
            var data = new BindingList<OutputTableModel>();

            var sqlite_conn = CreateSqliteConnection(this.dbPath);

            try
            {
                //sqlite_conn.Open();

                using (var sqlite_cmd = new SqliteCommand())
                {
                    sqlite_cmd.Connection = sqlite_conn;
                    sqlite_cmd.CommandText = $"SELECT * FROM GetAllRecords WHERE person LIKE '%{name}%' AND RecordDate BETWEEN '{startDate}' AND '{endDate}'";
                    //sqlite_cmd.CommandText = $"select substr(m.dateOfRecord, 7,4)||'-'||substr(m.dateOfRecord, 4,2)||'-'||substr(m.dateOfRecord, 1,2) as RecordDate, sheetName.value as sheetName, f1.value as f1, f2.value as f2, f3.value as f3, f4.value as f4, orderNumber.value as orderNumber, militaryPosition.value as militaryPosition, militaryProfession.value as militaryProfession, militaryRank.value as militaryRank, realMilitaryRank.value as realMilitaryRank, person.value as person, yearOfBirthEnlistment.value as yearOfBirthEnlistment, category.value as category, department.value as department, note.value as note, rvkAndOther.value as rvkAndOther, dateOfEnlistment.value as dateOfEnlistment, reason.value as reason, emptyField1.value as emptyField1, decreeNumber.value as decreeNumber, dateOfOut.value as dateOfOut, whereIs.value as whereIs, dateOfIn.value as dateOfIn, hospitalAto.value as hospitalAto, decreeAto.value as decreeAto from (select * from main where person in (select id from person where value like '{name}')) as m left join sheetName on sheetName.id = m.sheetName left join field1 as f1 on f1.id = m.field1 left join field2 as f2 on f2.id = m.field2 left join field3 as f3 on f3.id = m.field3 left join field4 as f4 on f4.id = m.field4 left join orderNumber on orderNumber.id = m.orderNumber left join militaryPosition on militaryPosition.id = m.militaryPosition left join militaryProfession on militaryProfession.id = m.militaryProfession left join militaryRank on militaryRank.id = m.militaryRank left join militaryRank as realMilitaryRank on realMilitaryRank.id = m.realMilitaryRank left join person on person.id = m.person left join yearOfBirthEnlistment on yearOfBirthEnlistment.id = m.yearOfBirthEnlistment left join category on category.id = m.category left join department on department.id = m.department left join note on note.id = m.note left join rvkAndOther on rvkAndOther.id = m.rvkAndOther left join dateOfEnlistment on dateOfEnlistment.id = m.dateOfEnlistment left join reason on reason.id = m.reason left join emptyField1 on emptyField1.id = m.emptyField1 left join decreeNumber on decreeNumber.id = m.decreeNumber left join dateOfOut on dateOfOut.id = m.dateOfOut left join whereIs on whereIs.id = m.whereIs left join dateOfIn on dateOfIn.id = m.dateOfIn left join hospitalAto on hospitalAto.id = m.hospitalAto left join decreeAto on decreeAto.id = m.decreeAto where RecordDate between '{startDate}' and '{endDate}' order by 1";
                    sqlite_cmd.CommandType = CommandType.Text;
                    
                    sqlite_conn.Open();

                    using (var reader = await sqlite_cmd.ExecuteReaderAsync())
                    {

                        while (reader.Read())
                        {
                            data.Add(new OutputTableModel
                            {
                                RecordDate = ConvertIfDateFormat(reader, 0),
                                SheetName = reader[1].ToString(),
                                Field1 = reader[2].ToString(),
                                Field2 = reader[3].ToString(),
                                Field3 = reader[4].ToString(),
                                Field4 = reader[5].ToString(),
                                OrderNumber = reader[6].ToString(),
                                MilitaryPosition = reader[7].ToString(),
                                MilitaryRank = reader[9].ToString(),
                                RealMilitaryRank = reader[10].ToString(),
                                FullName = reader[11].ToString(),
                                DateOfBirth = reader[12].ToString(),
                                Department = reader[14].ToString(),
                                Note = reader[15].ToString(),
                                RvkAndOther = reader[16].ToString(),
                                EnlistmentDate = ConvertIfDateFormat(reader, 17),
                                Reason = reader[18].ToString(),
                                DecreeNumber = reader[20].ToString(),
                                DateOfOut = ConvertIfDateFormat(reader, 21),
                                WhereIs = reader[22].ToString(),
                                DateOfIn = ConvertIfDateFormat(reader, 23),
                                HospitalAto = reader[24].ToString(),
                                DecreeAto = reader[25].ToString()
                            });
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
            }
            finally
            {
                CloseSqliteConnection(sqlite_conn);
            }
           
            return (data, errorMessage);
        }

        public string ImportExcel(string excelFile, ConfigModel? config)
        {
            var message = string.Empty;

            var regex = new Regex(@".*\\таблиця (\d{2}\.\d{2}\.\d{4})\.xlsx");

            var results = regex.Match(excelFile);
            if (!results.Success)
            {
                return $"'{excelFile}' файл повинен бути у форматі 'таблиця 01.01.2000.xlsx'.";
            }

            var fileDate = results.Groups[1].Value;
            var sw = new Stopwatch();
            sw.Start();

            var iw = new Stopwatch();

            iw.Start();
            message = message + PopulateData(excelFile, "таблиця", fileDate, ConvertJsonToMappedObject(config.Main));
            iw.Stop();
            message = $"{message}Обробка аркушу 'таблиця' склала '{iw.ElapsedMilliseconds / 1000}' сек.";
            iw.Reset();

            iw.Start();
            message = message + PopulateData(excelFile, "прикоманд", fileDate, ConvertJsonToMappedObject(config.Attached));
            iw.Stop();
            message = $"{message}\nОбробка аркушу 'прикоманд' склала '{iw.ElapsedMilliseconds / 1000}' сек.";
            iw.Reset();

            iw.Start();
            message = message + PopulateData(excelFile, "запаснарота", fileDate, ConvertJsonToMappedObject(config.Reserved));
            iw.Stop();
            message = $"{message}\nОбробка аркушу 'запаснарота' склала '{iw.ElapsedMilliseconds / 1000}' сек.";
            iw.Reset();

            sw.Stop();

            message = $"{message}\n\nВесь процес зайняв {sw.ElapsedMilliseconds / 1000} сек.";
            return message;
        }

        private Dictionary<string, int> ConvertJsonToMappedObject(List<MappingExcelModel> mappingExcelModels)
        {
            var columns = new Dictionary<string, int>();

            foreach (var mappingExcelModel in mappingExcelModels)
            {
                if(mappingExcelModel.DbField == null)
                {
                    continue;
                }

                columns.Add(mappingExcelModel.DbField, mappingExcelModel.ExcelColumn != null ? Array.IndexOf(ExcelLetters, mappingExcelModel.ExcelColumn.ToUpper()) : -1);
            }

            return columns;
        }

        private static string? ConvertIfDateFormat(SqliteDataReader reader, int index) =>
            DateTime.TryParse(reader[index].ToString(), out var dateTime)
            ? dateTime.ToShortDateString()
            : reader[index].ToString();

        private string PopulateData(string excelFileName, string sheetName, string dateOfRecords, Dictionary<string, int> columnIndex)
        {
            var message = string.Empty;

            // Get to know if records with date exist
            var sqlite_conn = CreateSqliteConnection(dbPath);
            var areRecordsExist = false;
            //return;
            try
            {
                sqlite_conn.Open();

                using (var sqlite_cmd = sqlite_conn.CreateCommand())
                {
                    sqlite_cmd.CommandText = $"SELECT m.dateOfRecord, sh.value from main as m left join sheetName as sh on sh.id=m.sheetName where dateOfRecord like '%{dateOfRecords}%' and sh.value='{sheetName}'";
                    sqlite_cmd.CommandType = CommandType.Text;
                    var reader = sqlite_cmd.ExecuteReader();
                    if (reader.HasRows)
                    {
                        areRecordsExist = true;
                    }
                }
            }
            catch (Exception ex)
            {
                message = $"Помилка: {ex.Message}";
            }
            finally
            {
                CloseSqliteConnection(sqlite_conn);
            }

            if (areRecordsExist)
            {
                return $"Записи з аркушу '{sheetName}' на '{dateOfRecords}' вже існують в базі даних.";
            }

            // 1. read data from Excel sheet
            var rows = GetDataFromExcel(excelFileName, sheetName, out var errorMessage);
            if (errorMessage.Length > 0)
            {
                return errorMessage;
            }

            // 2. select data from DB for single tables
            var sheetNameDB = SelectSingleData("sheetName");
            var field1DB = SelectSingleData("field1");
            var field2DB = SelectSingleData("field2");
            var field3DB = SelectSingleData("field3");
            var field4DB = SelectSingleData("field4");
            var orderNumberDB = SelectSingleData("orderNumber");
            var militaryPositionDB = SelectSingleData("militaryPosition");
            var militaryProfessionDB = SelectSingleData("militaryProfession");
            var militaryRankDB = SelectSingleData("militaryRank");
            var personDB = SelectSingleData("person");
            var yearOfBirthEnlistmentDB = SelectSingleData("yearOfBirthEnlistment");
            var categoryDB = SelectSingleData("category");
            var departmentDB = SelectSingleData("department");
            var noteDB = SelectSingleData("note");
            var rvkAndOtherDB = SelectSingleData("rvkAndOther");
            var dateOfEnlistmentDB = SelectSingleData("dateOfEnlistment");
            var reasonDB = SelectSingleData("reason");
            var emptyField1DB = SelectSingleData("emptyField1");
            var decreeNumberDB = SelectSingleData("decreeNumber");
            var dateOfOutDB = SelectSingleData("dateOfOut");
            var whereIsDB = SelectSingleData("whereIs");
            var dateOfInDB = SelectSingleData("dateOfIn");
            var hospitalAtoDB = SelectSingleData("hospitalAto");
            var decreeAtoDB = SelectSingleData("decreeAto");

            // 3. insert new values into DB for single tables
            InsertSheetNameValue("sheetName", sheetName, sheetNameDB);
            InsertSingleValues("field1", rows, field1DB, columnIndex["field1"]);
            InsertSingleValues("field2", rows, field2DB, columnIndex["field2"]);
            InsertSingleValues("field3", rows, field3DB, columnIndex["field3"]);
            InsertSingleValues("field4", rows, field4DB, columnIndex["field4"]);
            InsertSingleValues("orderNumber", rows, orderNumberDB, columnIndex["orderNumber"]);
            InsertSingleValues("militaryPosition", rows, militaryPositionDB, columnIndex["militaryPosition"], true);
            if (columnIndex.ContainsKey("militaryProfession"))
            {
                InsertSingleValues("militaryProfession", rows, militaryProfessionDB, columnIndex["militaryProfession"], true);
            }
            
            InsertSingleValues("militaryRank", rows, militaryRankDB, columnIndex["militaryRank"]);
            militaryRankDB = SelectSingleData("militaryRank");
            InsertSingleValues("militaryRank", rows, militaryRankDB, columnIndex["realMilitaryRank"]);
            InsertSingleValues("person", rows, personDB, columnIndex["person"]);
            InsertSingleValues("yearOfBirthEnlistment", rows, yearOfBirthEnlistmentDB, columnIndex["yearOfBirthEnlistment"]);
            InsertSingleValues("category", rows, categoryDB, columnIndex["category"]);
            InsertSingleValues("department", rows, departmentDB, columnIndex["department"]);
            InsertSingleValues("note", rows, noteDB, columnIndex["note"]);
            InsertSingleValues("rvkAndOther", rows, rvkAndOtherDB, columnIndex["rvkAndOther"]);
            InsertSingleValues("dateOfEnlistment", rows, dateOfEnlistmentDB, columnIndex["dateOfEnlistment"]);
            InsertSingleValues("reason", rows, reasonDB, columnIndex["reason"]);
            InsertSingleValues("emptyField1", rows, emptyField1DB, columnIndex["emptyField1"]);
            InsertSingleValues("decreeNumber", rows, decreeNumberDB, columnIndex["decreeNumber"]);
            InsertSingleValues("dateOfOut", rows, dateOfOutDB, columnIndex["dateOfOut"]);
            InsertSingleValues("whereIs", rows, whereIsDB, columnIndex["whereIs"]);
            InsertSingleValues("dateOfIn", rows, dateOfInDB, columnIndex["dateOfIn"]);
            InsertSingleValues("hospitalAto", rows, hospitalAtoDB, columnIndex["hospitalAto"]);
            InsertSingleValues("decreeAto", rows, decreeAtoDB, columnIndex["decreeAto"]);

            // 4. select data from DB after new values inserted
            sheetNameDB = SelectSingleData("sheetName");
            field1DB = SelectSingleData("field1");
            field2DB = SelectSingleData("field2");
            field3DB = SelectSingleData("field3");
            field4DB = SelectSingleData("field4");
            orderNumberDB = SelectSingleData("orderNumber");
            militaryPositionDB = SelectSingleData("militaryPosition");
            militaryProfessionDB = SelectSingleData("militaryProfession");
            militaryRankDB = SelectSingleData("militaryRank");
            personDB = SelectSingleData("person");
            yearOfBirthEnlistmentDB = SelectSingleData("yearOfBirthEnlistment");
            categoryDB = SelectSingleData("category");
            departmentDB = SelectSingleData("department");
            noteDB = SelectSingleData("note");
            rvkAndOtherDB = SelectSingleData("rvkAndOther");
            dateOfEnlistmentDB = SelectSingleData("dateOfEnlistment");
            reasonDB = SelectSingleData("reason");
            emptyField1DB = SelectSingleData("emptyField1");
            decreeNumberDB = SelectSingleData("decreeNumber");
            dateOfOutDB = SelectSingleData("dateOfOut");
            whereIsDB = SelectSingleData("whereIs");
            dateOfInDB = SelectSingleData("dateOfIn");
            hospitalAtoDB = SelectSingleData("hospitalAto");
            decreeAtoDB = SelectSingleData("decreeAto");

            // 5. insert records into db

            foreach (var row in rows)
            {
                var sName = sheetNameDB.Where(x => string.Equals(sheetName, x.Value,StringComparison.Ordinal)).FirstOrDefault().Key;
                var field1 = ParseValue(field1DB, row, columnIndex["field1"]);
                var field2 = ParseValue(field2DB, row, columnIndex["field2"]);
                var field3 = ParseValue(field3DB, row, columnIndex["field3"]);
                var field4 = ParseValue(field4DB, row, columnIndex["field4"]);
                var orderNumber = ParseValue(orderNumberDB, row, columnIndex["orderNumber"]);
                var militaryPosition = ParseValue(militaryPositionDB, row, columnIndex["militaryPosition"], true);
                var militaryProfession = columnIndex.ContainsKey("militaryProfession") ? ParseValue(militaryProfessionDB, row, columnIndex["militaryProfession"], true) : null;
                var militaryRank = ParseValue(militaryRankDB, row, columnIndex["militaryRank"]);
                var realMilitaryRank = ParseValue(militaryRankDB, row, columnIndex["realMilitaryRank"]);
                var person = ParseValue(personDB, row, columnIndex["person"]); 
                var yearOfBirthEnlistment = ParseValue(yearOfBirthEnlistmentDB, row, columnIndex["yearOfBirthEnlistment"]);
                var category = ParseValue(categoryDB, row, columnIndex["category"]);
                var department = ParseValue(departmentDB, row, columnIndex["department"]);
                var note = ParseValue(noteDB, row, columnIndex["note"]);
                var rvkAndOther = ParseValue(rvkAndOtherDB, row, columnIndex["rvkAndOther"]);
                var dateOfEnlistment = ParseValue(dateOfEnlistmentDB, row, columnIndex["dateOfEnlistment"]);
                var reason = ParseValue(reasonDB, row, columnIndex["reason"]);
                var emptyField1 = ParseValue(emptyField1DB, row, columnIndex["emptyField1"]);
                var decreeNumber = ParseValue(decreeNumberDB, row, columnIndex["decreeNumber"]);
                var dateOfOut = ParseValue(dateOfOutDB, row, columnIndex["dateOfOut"]);
                var whereIs = ParseValue(whereIsDB, row, columnIndex["whereIs"]);
                var dateOfIn = ParseValue(dateOfInDB, row, columnIndex["dateOfIn"]);
                var hospitalAto = ParseValue(hospitalAtoDB, row, columnIndex["hospitalAto"]);
                var decreeAto = ParseValue(decreeAtoDB, row, columnIndex["decreeAto"]);

                sqlite_conn = CreateSqliteConnection(this.dbPath);
                try
                {
                    sqlite_conn.Open();
                    SqliteCommand sqlite_cmd = new SqliteCommand();
                    sqlite_cmd = sqlite_conn.CreateCommand();
                    sqlite_cmd.CommandText = "INSERT INTO main (dateOfRecord, sheetName, field1, field2, field3, field4, orderNumber, militaryPosition, militaryProfession, militaryRank, realMilitaryRank, person, yearOfBirthEnlistment, category, department, note, rvkAndOther, dateOfEnlistment, reason, emptyField1, decreeNumber, dateOfOut, whereIs, dateOfIn, hospitalAto, decreeAto) values (@dateOfRecord, @sheetName, @field1, @field2, @field3, @field4, @orderNumber, @militaryPosition, @militaryProfession, @militaryRank, @realMilitaryRank, @person, @yearOfBirthEnlistment, @category, @department, @note, @rvkAndOther, @dateOfEnlistment, @reason, @emptyField1, @decreeNumber, @dateOfOut, @whereIs, @dateOfIn, @hospitalAto, @decreeAto)";

                    sqlite_cmd.Parameters.AddWithValue("@dateOfRecord", dateOfRecords);
                    sqlite_cmd.Parameters.AddWithValue("@sheetName", sName);
                    sqlite_cmd.Parameters.AddWithValue("@field1", field1 == null ? DBNull.Value : long.Parse(field1));
                    sqlite_cmd.Parameters.AddWithValue("@field2", field2 == null ? DBNull.Value : long.Parse(field2));
                    sqlite_cmd.Parameters.AddWithValue("@field3", field3 == null ? DBNull.Value : long.Parse(field3));
                    sqlite_cmd.Parameters.AddWithValue("@field4", field4 == null ? DBNull.Value : long.Parse(field4));
                    sqlite_cmd.Parameters.AddWithValue("@orderNumber", orderNumber == null ? DBNull.Value : long.Parse(orderNumber));
                    sqlite_cmd.Parameters.AddWithValue("@militaryPosition", militaryPosition == null ? DBNull.Value : long.Parse(militaryPosition));
                    sqlite_cmd.Parameters.AddWithValue("@militaryProfession", militaryProfession == null ? DBNull.Value : long.Parse(militaryProfession));
                    sqlite_cmd.Parameters.AddWithValue("@militaryRank", militaryRank == null ? DBNull.Value : long.Parse(militaryRank));
                    sqlite_cmd.Parameters.AddWithValue("@realMilitaryRank", realMilitaryRank == null ? DBNull.Value : long.Parse(realMilitaryRank));
                    sqlite_cmd.Parameters.AddWithValue("@person", person == null ? DBNull.Value : long.Parse(person));
                    sqlite_cmd.Parameters.AddWithValue("@yearOfBirthEnlistment", yearOfBirthEnlistment == null ? DBNull.Value : long.Parse(yearOfBirthEnlistment));
                    sqlite_cmd.Parameters.AddWithValue("@category", category == null ? DBNull.Value : long.Parse(category));
                    sqlite_cmd.Parameters.AddWithValue("@department", department == null ? DBNull.Value : long.Parse(department));
                    sqlite_cmd.Parameters.AddWithValue("@note", note == null ? DBNull.Value : long.Parse(note));
                    sqlite_cmd.Parameters.AddWithValue("@rvkAndOther", rvkAndOther == null ? DBNull.Value : long.Parse(rvkAndOther));
                    sqlite_cmd.Parameters.AddWithValue("@dateOfEnlistment", dateOfEnlistment == null ? DBNull.Value : long.Parse(dateOfEnlistment));
                    sqlite_cmd.Parameters.AddWithValue("@reason", reason == null ? DBNull.Value : long.Parse(reason));
                    sqlite_cmd.Parameters.AddWithValue("@emptyField1", emptyField1 == null ? DBNull.Value : emptyField1);
                    sqlite_cmd.Parameters.AddWithValue("@decreeNumber", decreeNumber == null ? DBNull.Value : long.Parse(decreeNumber));
                    sqlite_cmd.Parameters.AddWithValue("@dateOfOut", dateOfOut == null ? DBNull.Value : dateOfOut);
                    sqlite_cmd.Parameters.AddWithValue("@whereIs", whereIs == null ? DBNull.Value : long.Parse(whereIs));
                    sqlite_cmd.Parameters.AddWithValue("@dateOfIn", dateOfIn == null ? DBNull.Value : dateOfIn);
                    sqlite_cmd.Parameters.AddWithValue("@hospitalAto", hospitalAto == null ? DBNull.Value : long.Parse(hospitalAto));
                    sqlite_cmd.Parameters.AddWithValue("@decreeAto", decreeAto == null ? DBNull.Value : long.Parse(decreeAto));

                    sqlite_cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    errorMessage = $"Помилка: {ex.Message}";
                }
                finally
                {
                    CloseSqliteConnection(sqlite_conn);
                }
            }

            return message;
        }

        private static string? ParseValue(Dictionary<long, string> db, List<string> row, int index, bool removeDataAfterEnter = false)
        {
            if (index == -1)
            {
                return null;
            }

            var value = row[index].Trim();
            if (removeDataAfterEnter && value.Contains('\n'))
            {
                value = value[..value.IndexOf('\n')];
            }

            return db.ContainsValue(value) ?
            db.First(x => string.Equals(x.Value, value, StringComparison.Ordinal)).Key.ToString()
            : null;

        }

        private void InsertSheetNameValue(string tableName, string value, Dictionary<long, string> db)
        {
            var doesValueExist = db.Where(x => string.Equals(x.Value, value, StringComparison.Ordinal)).Any();
            if(doesValueExist)
            {
                return;
            }

            var sqlite_conn = CreateSqliteConnection(this.dbPath);
            try
            {
                sqlite_conn.Open();
                SqliteCommand sqlite_cmd = new SqliteCommand();
                sqlite_cmd = sqlite_conn.CreateCommand();
                sqlite_cmd.CommandText = $"INSERT INTO {tableName} (value) VALUES (@value)";
                sqlite_cmd.Parameters.AddWithValue("@value", value);
                sqlite_cmd.ExecuteNonQuery();
                sqlite_conn.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                CloseSqliteConnection(sqlite_conn);
            }
        }

        private void InsertSingleValues(string tableName, List<List<string>> excelRows, Dictionary<long, string> db, int index, bool removeDataAfterEnter = false)
        {
            if (index < 0)
            {
                return;
            }

            var listOfValues = new List<string>();
            foreach (var excelRow in excelRows)
            {
                var item = excelRow.ElementAt(index).Trim();
                if (item.Length == 0)
                {
                    continue;
                }

                if (removeDataAfterEnter && item.Contains('\n'))
                {
                    item = item[..item.IndexOf('\n')];
                }

                var isItemExist = db.Where(x => string.Equals(x.Value, item, StringComparison.Ordinal)).Any();
                if (isItemExist)
                {
                    continue;
                }

                if (!listOfValues.Contains(item))
                {
                    listOfValues.Add(item);
                }
            }

            if (!listOfValues.Any())
            {
                return;
            }


            foreach (var item in listOfValues)
            {
                var sqlite_conn = CreateSqliteConnection(this.dbPath);
                try
                {
                    sqlite_conn.Open();
                    SqliteCommand sqlite_cmd = new SqliteCommand();
                    sqlite_cmd = sqlite_conn.CreateCommand();
                    sqlite_cmd.CommandText = $"INSERT INTO {tableName} (value) VALUES (@value)";
                    sqlite_cmd.Parameters.AddWithValue("@value", item);
                    sqlite_cmd.ExecuteNonQuery();
                    sqlite_conn.Close();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                finally
                {
                    CloseSqliteConnection(sqlite_conn);
                }
            }
        }

        private Dictionary<long, string> SelectSingleData(string tableName)
        {
            var data = new Dictionary<long, string>();

            var sqlite_conn = CreateSqliteConnection(this.dbPath);

            try
            {
                sqlite_conn.Open();
                
                using var sqlite_cmd = sqlite_conn.CreateCommand();
                sqlite_cmd.CommandText = $"SELECT * from {tableName}";
                sqlite_cmd.CommandType = CommandType.Text;
                
                var reader = sqlite_cmd.ExecuteReader();
                while (reader.Read())
                {
                    data.Add((long)reader[0], reader[1].ToString());
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                CloseSqliteConnection(sqlite_conn);
            }

            return data;
        }

        private List<List<string>> GetDataFromExcel(string fileName, string sheetName, out string errorMessage)
        {
            errorMessage = string.Empty;
            var data = new List<List<string>>();

            var connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + fileName + "';Extended Properties=\"Excel 12.0;HDR=No;IMEX=0\"";
    
            using (var connection = new OleDbConnection(connectionString))
            {
                var command = new OleDbCommand($"SELECT * FROM [{sheetName}$] where [F2] <> ''", connection);
                try
                {
                    connection.Open();
                    var reader = command.ExecuteReader();
                    try
                    {
                        while (reader.Read())
                        {
                            var row = new List<string>();

                            for (var i = 0; i < reader.FieldCount; i++)
                            {
                                row.Add(reader.GetValue(i).ToString());
                            }

                            data.Add(row);
                        }
                    }
                    catch (Exception ex)
                    {
                        errorMessage = ex.Message;
                    }
                    finally
                    {
                        reader.Close();
                    }
                }
                catch (Exception ex)
                {
                    errorMessage = $"Помилка: {ex.Message}";
                }
            }

            return data;
        }

        private SqliteConnection CreateSqliteConnection(string dbPath)
        {
            SqliteConnection sqlite_conn;
            // Create a new database connection:
            sqlite_conn = new SqliteConnection($"Data Source={dbPath};");
            SQLitePCL.raw.SetProvider(new SQLitePCL.SQLite3Provider_e_sqlite3());
            return sqlite_conn;
        }

        private void CloseSqliteConnection(SqliteConnection sqlite_conn)
        {
            sqlite_conn.Close();
            sqlite_conn.Dispose();
            SqliteConnection.ClearAllPools();
        }
    }
}
