using System;
using System.Data;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;

namespace Project_parser
{
    public class DB
    {
        public SqlConnection Connection = null;

        /// <summary>
        /// Инициализация класса
        /// </summary>
        /// <param name="connectionString">Строка подключения</param>
        public DB(string connectionString)
        {
            Connection = new SqlConnection(connectionString);
        }

        /// <summary>
        /// Подключаемся к бд
        /// </summary>
        /// <returns>Если удалось подключиться, то возвращает true. Иначе false</returns>
        public bool Connect()
        {
            try
            {
                Connection.Open();
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Отключаемся от базы данных и говорим уборщику мусора чтоб убрал все созданные объекты внутри экземпляра класса
        /// </summary>
        public void Close()
        {
            if (Connection != null)
                Connection.Close();
        }
        public string GetAdvTypeByInt(int id)
        {
            if (Connection != null)
            {
                using (var sql = new SqlCommand("SELECT \"type\" FROM Nedvizh.dbo.adv_types WHERE \"id\" = " + id + ";", Connection))
                {
                    var dataReader = sql.ExecuteReader();
                    if (dataReader.Read())
                    {
                        return dataReader.GetString(0);
                    }
                    dataReader.Close();
                }
            }
            return null;
        }

        public Advertisement GetAdvPdata(int id)
        {
            var result = new Advertisement();
            if (Connection != null)
            {
                using (var sql = new SqlCommand("SELECT \"adv_cost\", \"adv_area\" FROM Nedvizh.dbo.adv_pdata WHERE \"id\" = " + id + ";", Connection))
                {
                    var dataReader = sql.ExecuteReader();
                    if (dataReader.Read())
                    {
                        result.Price = dataReader.GetString(0);
                        result.Square = dataReader.GetString(1);
                    }
                    dataReader.Close();
                }
            }
            return result;
        }

        public string GetAdvType(int id)
        {
            if (Connection != null)
            {

                using (var sql = new SqlCommand("SELECT \"adv_type\" FROM Nedvizh.dbo.adv_pdata WHERE \"id\" = " + id + ";", Connection))
                {
                    byte type = 0;
                    var dataReader = sql.ExecuteReader();
                    if (dataReader.Read())
                    {
                        type = dataReader.GetByte(0);
                    }
                    dataReader.Close();
                    return GetAdvTypeByInt(type);
                }
            }
            return string.Empty;
        }

        public string GetUrlById(int id)
        {
            var result = string.Empty;
            if (Connection != null)
            {

                using (var sql = new SqlCommand("SELECT \"adv_link\" FROM Nedvizh.dbo.adv_links WHERE \"id\" = " + id + ";", Connection))
                {
                    var dataReader = sql.ExecuteReader();
                    if (dataReader.Read())
                    {
                        result = dataReader.GetString(0);
                    }
                    dataReader.Close();
                }
            }
            return result;
        }

        public bool UpdateDate(int id, DateTime date, string dateString)
        {
            if (Connection != null)
            {
                using (var sql = new SqlCommand("UPDATE Nedvizh.dbo.main_table SET \"date_act\" = @date_act, \"date_act_str\" = @date_act_str WHERE \"id\" = " + id + ";", Connection))
                {
                    var dateParam = new SqlParameter("@date_act", SqlDbType.DateTime, 255) { Value = date };
                    var dateSParam = new SqlParameter("@date_act_str", SqlDbType.VarChar, 255) { Value = dateString };
                    sql.Parameters.Add(dateParam);
                    sql.Parameters.Add(dateSParam);
                    sql.Prepare();
                    sql.ExecuteNonQuery();
                }
                return true;
            }
            return false;
        }
        public bool UpdateStatus(int id, char status, int source)
        {
            if (Connection != null)
            {
                using (var sql = new SqlCommand("UPDATE Nedvizh.dbo.main_table SET \"adv_status\" = @adv_status, \"source\" = @source WHERE \"id\" = " + id + ";", Connection))
                {                    
                    var statusParam = new SqlParameter("@adv_status", SqlDbType.Char, 255) { Value = status };
                    var sourceParam = new SqlParameter("@source", SqlDbType.TinyInt, 255) { Value = source };
                    sql.Parameters.Add(statusParam);
                    sql.Parameters.Add(sourceParam);
                    sql.Prepare();
                    sql.ExecuteNonQuery();
                }
                return true;
            }
            return false;
        }

        public bool UpdateComType(string num, string comType)
        {
            if (Connection != null)
            {
                using (var sql = new SqlCommand("UPDATE Nedvizh.dbo.adv_pdata SET \"adv_comType\" = @adv_comType WHERE \"id\" in (SELECT id FROM Nedvizh.dbo.adv_links WHERE \"adv_num\" = @adv_num );", Connection))
                {
                    var typeParam = new SqlParameter("@adv_comType", SqlDbType.VarChar, 255) { Value = comType };
                    var numParam = new SqlParameter("@adv_num", SqlDbType.VarChar, 255) { Value = num };
                    sql.Parameters.Add(typeParam);
                    sql.Parameters.Add(numParam);
                    sql.Prepare();
                    sql.ExecuteNonQuery();
                }
                return true;
            }
            return false;
        }
        public bool UpdateComTypeSec(string num, string comType)
        {
            if (Connection != null)
            {
                using (var sql = new SqlCommand("UPDATE Nedvizh.dbo.adv_pdata SET \"adv_comTypeSec\" = @adv_comType WHERE \"id\" in (SELECT id FROM Nedvizh.dbo.adv_links WHERE \"adv_num\" = @adv_num );", Connection))
                {
                    var typeParam = new SqlParameter("@adv_comType", SqlDbType.VarChar, 255) { Value = comType };
                    var numParam = new SqlParameter("@adv_num", SqlDbType.VarChar, 255) { Value = num };
                    sql.Parameters.Add(typeParam);
                    sql.Parameters.Add(numParam);
                    sql.Prepare();
                    sql.ExecuteNonQuery();
                }
                return true;
            }
            return false;
        }

        public int GetRoomTypeByString(string type)
        {
            if (Connection != null)
            {
                using (var sql = new SqlCommand("SELECT \"id\" FROM Nedvizh.dbo.adv_rooms_type WHERE \"room_type\" = '" + type + "';", Connection))
                {
                    byte ret = 0;
                    var dataReader = sql.ExecuteReader();
                    if (dataReader.Read())
                    {
                        ret = dataReader.GetByte(0);
                    }
                    dataReader.Close();
                    return ret;
                }
            }
            return 0;
        }


        public int GetTownByName(string name)
        {
            if (Connection != null)
            {
                using (var sql = new SqlCommand("SELECT \"id\" FROM Nedvizh.dbo.town_names_table WHERE \"town_name\" = '" + name + "';", Connection))
                {
                    byte ret = 0;
                    var dataReader = sql.ExecuteReader();
                    if (dataReader.Read())
                    {
                        ret = dataReader.GetByte(0);
                    }
                    dataReader.Close();
                    return ret;
                }
            }
            return 0;
        }
        public int GetTownByMainId(int id)
        {
            if (Connection != null)
            {
                using (var sql = new SqlCommand("SELECT \"addr02\" FROM Nedvizh.dbo.addr_table WHERE \"id\" = '" + id + "';", Connection))
                {
                    int ret = -1;
                    var dataReader2 = sql.ExecuteReader();
                    if (dataReader2.Read())
                    {
                        if (!dataReader2.IsDBNull(0))
                            ret = dataReader2.GetByte(0);
                    }
                    dataReader2.Close();
                    return ret;
                }
            }
            return -2;
        }

        public string GetTownById(int id)
        {
            if (Connection != null)
            {
                using (var sql = new SqlCommand("SELECT \"town_name\" FROM Nedvizh.dbo.town_names_table WHERE \"id\" = '" + id + "';", Connection))
                {
                    string ret = "";
                    var dataReader2 = sql.ExecuteReader();
                    if (dataReader2.Read())
                    {
                        ret = dataReader2.GetString(0);
                    }
                    dataReader2.Close();
                    return ret;
                }
            }
            return "";
        }

        public string GetRoomType(int id)
        {
            if (Connection != null)
            {
                using (var sql = new SqlCommand("SELECT \"room_type\" FROM Nedvizh.dbo.adv_rooms_type WHERE \"id\" = '" + id + "';", Connection))
                {
                    string ret = "";
                    var dataReader2 = sql.ExecuteReader();
                    if (dataReader2.Read())
                    {
                        ret = dataReader2.GetString(0);
                    }
                    dataReader2.Close();
                    return ret;
                }
            }
            return "";
        }

        public int GetAdvTypeByString(string type)
        {
            if (Connection != null)
            {
                using (var sql = new SqlCommand("SELECT \"id\" FROM Nedvizh.dbo.adv_types WHERE \"type\" = '" + type + "';", Connection))
                {
                    byte ret = 0;
                    var dataReader = sql.ExecuteReader();
                    if (dataReader.Read())
                    {
                        ret = dataReader.GetByte(0);
                    }
                    dataReader.Close();
                    return ret;
                }
            }
            return -1;
        }

        public string GetAdvTypeById(int id)
        {
            if (Connection != null)
            {
                using (var sql = new SqlCommand("SELECT \"type\" FROM Nedvizh.dbo.adv_types WHERE \"id\" = " + id + ";", Connection))
                {
                    var dataReader = sql.ExecuteReader();
                    string type = "?";
                    if (dataReader.Read())
                    {
                        type = dataReader.GetString(0);
                    }
                    dataReader.Close();
                    return type;
                }
            }
            return "?";
        }
        /// <summary>
        /// Возвращает ID послденего объявления в БД (Int32)
        /// </summary>
        public int GetLastID()
        {
            if (Connection != null)
            {
                using (var sql = new SqlCommand("SELECT TOP 1 id FROM Nedvizh.dbo.main_table ORDER BY id DESC;", Connection))
                {
                    var last = 0;
                    var dataReader = sql.ExecuteReader();
                    if (dataReader.Read())
                    {
                        last = dataReader.GetInt32(0);
                    }
                    dataReader.Close();
                    return last;
                }
            }
            return -1;
        }
        public int GetFirstIdByNumber(string number)
        {
            if (Connection != null)
            {
                using (var sql = new SqlCommand("SELECT \"id\" FROM Nedvizh.dbo.adv_links WHERE \"adv_num\" = @adv_num ORDER BY id", Connection))
                {
                    var numParam = new SqlParameter("@adv_num", SqlDbType.VarChar, 255) { Value = number };
                    sql.Parameters.Add(numParam);
                    sql.Prepare();
                    var id = -1;
                    var dataReader = sql.ExecuteReader();
                    if (dataReader.Read())
                    {
                        id = dataReader.GetInt32(0);
                    }
                    dataReader.Close();
                    return id;
                }
            }
            return -1;
        }

        public int GetIdByNumber(string number)
        {
            if (Connection != null)
            {
                using (var sql = new SqlCommand("SELECT \"id\" FROM Nedvizh.dbo.adv_links WHERE \"adv_num\" = @adv_num ORDER BY id desc;", Connection))
                {
                    var numParam = new SqlParameter("@adv_num", SqlDbType.VarChar, 255) { Value = number };
                    sql.Parameters.Add(numParam);
                    sql.Prepare();
                    var id = -1;
                    var dataReader = sql.ExecuteReader();
                    if (dataReader.Read())
                    {
                        id = dataReader.GetInt32(0);
                    }
                    dataReader.Close();
                    return id;
                }
            }
            return -1;
        }

        public string GetLinkById(int id)
        {
            if (Connection == null)
            {
                return string.Empty;
            }
            using (var sql = new SqlCommand("SELECT \"adv_link\" FROM Nedvizh.dbo.adv_links WHERE \"id\" = " + id, Connection)) //  adv_nlink
            {
                var result = string.Empty;
                var reader = sql.ExecuteReader();
                if (reader.Read())
                {
                    result = reader["adv_link"].ToString();
                }
                reader.Close();
                return result;
            }
        }

        public bool ChangeAdvNLink(int id, string link)
        {
            using (var sql = new SqlCommand("UPDATE Nedvizh.dbo.adv_links SET \"adv_nlink\" = @adv_nlink WHERE \"id\" = @id", Connection))
            {
                var nlinkParam = new SqlParameter("@adv_nlink", SqlDbType.VarChar, 255) { Value = link };
                var idParam = new SqlParameter("@id", SqlDbType.Int, 255) { Value = id };
                sql.Parameters.Add(nlinkParam);
                sql.Parameters.Add(idParam);
                sql.Prepare();
                return sql.ExecuteNonQuery() == 1;
            }
        }

        public string GetNumberById(int id)
        {
            if (Connection == null)
            {
                return string.Empty;
            }
            using (var sql = new SqlCommand("SELECT \"adv_num\" FROM Nedvizh.dbo.adv_links WHERE \"id\" = " + id, Connection))
            {
                var result = string.Empty;
                var reader = sql.ExecuteReader();
                if (reader.Read())
                {
                    result = reader["adv_num"].ToString();
                }
                reader.Close();
                return result;
            }
        }
        /// <summary>
        /// Вставляем объяву в БД
        /// </summary>
        /// <returns>Если удалось вставить, то возвращаем ID. Иначе -1</returns>
        public int InsertRecord()
        {
            if (Connection == null)
            {
                return -1;
            }
            var inserted = -1;
            using (var sql = new SqlCommand("INSERT INTO Nedvizh.dbo.main_table (\"date_ins\", \"time_ins\", \"Date_act\", \"Date_ev_save\") VALUES (CONVERT(DATETIME, '" + DateTime.Now.ToString("yyyy.dd.MM hh:mm:ss") + "', 104), CONVERT(TIME, '" + DateTime.Now.ToString("hh:mm:ss") + "', 104), CONVERT(DATETIME, '1970.01.01 00:00:00', 104), CONVERT(DATETIME, '1970.01.01 00:00:00', 104)); SELECT SCOPE_IDENTITY();", Connection))
            {
                try
                {
                    inserted = Convert.ToInt32(sql.ExecuteScalar());
                }
                catch (Exception ex)
                {
                    General.WriteLog(ex.ToString());
                    return -1;
                }
            }
            using (var sql = new SqlCommand("INSERT INTO Nedvizh.dbo.adv_links (\"id\", \"adv_link\", \"adv_rlink\", \"adv_title\", \"adv_num\") VALUES (" + inserted + ", '', '', '', 0);", Connection))
            {
                try
                {
                    sql.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    General.WriteLog(ex.ToString());
                    return -1;
                }
            }
            using (var sql = new SqlCommand("INSERT INTO Nedvizh.dbo.adv_pdata (\"id\", \"adv_cost\", \"adv_area\", \"adv_type\") VALUES (" + inserted + ", 0, 0, 0);", Connection))
            {
                try
                {
                    sql.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    General.WriteLog(ex.ToString());
                    return -1;
                }
            }
            using (var sql = new SqlCommand("INSERT INTO Nedvizh.dbo.addr_table (\"id\", \"addr_coord\") VALUES (" + inserted + ", '');", Connection))
            {
                try
                {
                    sql.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    General.WriteLog(ex.ToString());
                    return -1;
                }
            }
            using (var sql = new SqlCommand("INSERT INTO Nedvizh.dbo.adv_users_info (\"id\", \"adv_user_num_of_views\") VALUES (" + inserted + ", 0);", Connection))
            {
                try
                {
                    sql.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    General.WriteLog(ex.ToString());
                    return -1;
                }
            }
            return inserted;
        }

        public bool UpdateAddr(int id, string addr, string hood, string street, string house, string flat, string cadastre, string town)
        {
            if (Connection == null)
            {
                return false;
            }
            var townId = GetTownByName(town);
            using (var sql = new SqlCommand("UPDATE Nedvizh.dbo.addr_table SET \"addr\" = @addr, \"addr02\" = " + townId + ",\"addr03\" = @addr03, \"addr04\" = @addr04, \"addr05\" = @addr05, \"addr06\" = @addr06, \"addr_cadastre\" = @addr_cadastre WHERE \"id\" = " + id, Connection))
            {
                var addrParam = new SqlParameter("@addr", SqlDbType.VarChar, 128) { Value = addr };
                var addr03Param = new SqlParameter("@addr03", SqlDbType.VarChar, 32) { Value = hood };
                var addr04Param = new SqlParameter("@addr04", SqlDbType.VarChar, 32) { Value = street };
                var addr05Param = new SqlParameter("@addr05", SqlDbType.VarChar, 32) { Value = house };
                var addr06Param = new SqlParameter("@addr06", SqlDbType.VarChar, 32) { Value = flat };
                var addrCadastreParam = new SqlParameter("@addr_cadastre", SqlDbType.VarChar, 16) { Value = cadastre };
                sql.Parameters.Add(addrParam);
                sql.Parameters.Add(addr03Param);
                sql.Parameters.Add(addr04Param);
                sql.Parameters.Add(addr05Param);
                sql.Parameters.Add(addr06Param);
                sql.Parameters.Add(addrCadastreParam);
                sql.Prepare();
                try
                {
                    sql.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    General.WriteLog(ex.ToString());
                    return false;
                }
            }
            return true;
        }

        public bool InsertPageInfo(string addr, int type, int first, int last, string date)
        {
            if (Connection == null)
            {
                return false;
            }
            using (var sql = new SqlCommand("INSERT INTO Nedvizh.dbo.l_parse_hist (\"l_date\", \"l_addr\", \"l_type\", \"l_first\", \"l_last\")" +
                "VALUES ( CONVERT(DATETIME, '" + date + "', 104), @l_addr, @l_type, @l_first, @l_last);", Connection))
            {
                var addrParam = new SqlParameter("@l_addr", SqlDbType.VarChar, 255) { Value = addr };
                var typeParam = new SqlParameter("@l_type", SqlDbType.Int, 255) { Value = type };
                var firstParam = new SqlParameter("@l_first", SqlDbType.Int, 255) { Value = first };
                var lastParam = new SqlParameter("@l_last", SqlDbType.Int, 255) { Value = last };
                sql.Parameters.Add(addrParam);
                sql.Parameters.Add(typeParam);
                sql.Parameters.Add(firstParam);
                sql.Parameters.Add(lastParam);
                sql.Prepare();
                sql.ExecuteNonQuery();
            }
            return true;
        }
        public bool UpdateAdvData(int id, string const_, string purp, string offer, string rooms, string text)
        {
            if (Connection == null)
            {
                return false;
            }

            using (var sql = new SqlCommand(null, Connection)
            {
                CommandText = "UPDATE Nedvizh.dbo.adv_pdata SET \"adv_const\" = @adv_const, \"adv_purp\" = @adv_purp, \"adv_offer\" = @adv_offer, \"adv_rooms\" = @adv_rooms, \"adv_text\" = @adv_text WHERE \"id\" = @id;"
            })
            {
                var roomsInt = GetRoomTypeByString(rooms);
                var constParam = new SqlParameter("@adv_const", SqlDbType.VarChar, 255) { Value = const_ };
                var purpParam = new SqlParameter("@adv_purp", SqlDbType.VarChar, 255) { Value = purp };
                var offerParam = new SqlParameter("@adv_offer", SqlDbType.VarChar, 255) { Value = offer };
                var roomsParam = new SqlParameter("@adv_rooms", SqlDbType.TinyInt) { Value = roomsInt };
                var textParam = new SqlParameter("@adv_text", SqlDbType.VarChar, 512) { Value = text };
                var idParam = new SqlParameter("@id", SqlDbType.Int) { Value = id };
                sql.Parameters.Add(constParam);
                sql.Parameters.Add(purpParam);
                sql.Parameters.Add(offerParam);
                sql.Parameters.Add(roomsParam);
                sql.Parameters.Add(textParam);
                sql.Parameters.Add(idParam);
                sql.Prepare();
                try
                {
                    sql.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    General.WriteLog(ex.ToString());
                    return false;
                }
            }
            return true;
        }

        public void ExportToExcelFull(int idFirst, int idLast, string fileName, string Type)
        {
            string JOIN, WHERE, ORDER;
            int TYPE, COUNT;
            if (Type == "null")
            {
                ORDER = "";
                JOIN = "";
                WHERE = " WHERE b.id BETWEEN " + idFirst + " AND " + idLast;
                COUNT = idLast - idFirst + 1;
            }
            else
            {
                TYPE = GetAdvTypeByString(Type);
                WHERE = " WHERE pd.adv_type = " + TYPE;
                ORDER = " ORDER BY b.id";
                COUNT = GetCount(WHERE);
                JOIN = " JOIN Nedvizh.dbo.adv_pdata AS pd ON (b.id = pd.id)";
            }
            var ObjWorkExcel = new Application();
            var ObjWorkBook = ObjWorkExcel.Workbooks.Open(fileName);
            var count = COUNT;
            var res1 = new string[count][];
            var res2 = new string[count][];
            var res3 = new string[count][];

            using (var sql = new SqlCommand("SELECT b.adv_const, b.adv_rooms, b.adv_offer, b.adv_text, b.adv_purp, b.adv_comType, b.adv_comTypeSec FROM Nedvizh.dbo.adv_pdata as b" + JOIN + " " + WHERE + ORDER +";", Connection))
            {
                var dataReader = sql.ExecuteReader();
                var j = 0;
                while (dataReader.Read())
                {
                    res1[j] = new string[7];
                    for (int i = 0; i < 7; i++)
                    {
                        if (dataReader[i] != DBNull.Value)
                            res1[j][i] = dataReader[i].ToString();
                        else
                            res1[j][i] = "";
                    }
                    j++;
                }
                dataReader.Close();
            }
            for (int j = 0; j < count; j++)
            {
                try { res1[j][1] = GetRoomType(Int16.Parse(res1[j][1])); } catch { }
            }

            using (var sql = new SqlCommand("SELECT adv_user_name, adv_user_phone, adv_user_mail, adv_user_num_of_advs, adv_user_num_of_views FROM Nedvizh.dbo.adv_users_info as b" + JOIN + " " + WHERE + ORDER + ";", Connection))
            {
                var dataReader = sql.ExecuteReader();
                var j = 0;
                while (dataReader.Read())
                {
                    res2[j] = new string[5];
                    for (int i = 0; i < 5; i++)
                    {
                        if (dataReader[i] != DBNull.Value)
                            res2[j][i] = dataReader[i].ToString();
                        else
                            res2[j][i] = "";
                    }
                    j++;
                }
                dataReader.Close();
            }

            using (var sql = new SqlCommand("SELECT addr, addr03, addr04, addr05, addr06, addr_cadastre FROM Nedvizh.dbo.addr_table as b" + JOIN + " " + WHERE + ORDER + ";", Connection))
            {
                var dataReader = sql.ExecuteReader();
                var j = 0;
                while (dataReader.Read())
                {
                    res3[j] = new string[6];
                    for (int i = 0; i < 6; i++)
                    {
                        if (dataReader[i] != DBNull.Value)
                            res3[j][i] = dataReader[i].ToString();
                        else
                            res3[j][i] = "";
                    }
                    j++;
                }
                dataReader.Close();
            }

            var xlSht = ObjWorkBook.Sheets[1];
            string[] titles = { "Ограничения", "Комнаты", "От кого", "Текст", "Назначение", "ТипКН_Лента", "ТипКН_Объява", "UserName", "UserPhone", "UserMail", "UserКолОбъяв", "UserКолПросмотров", "Адрес полный", "Микрорайон", "Улица", "№ Дома", "№ Квартиры", "Кадастровый номер" };
            xlSht.Cells[1, 24].Resize[1, titles.GetUpperBound(0) + 1].Value = titles;
            xlSht.Cells[1, 24].Resize[1, titles.GetUpperBound(0) + 1].EntireRow.Font.Bold = true;

            for (var i = 0; i < count; i++)
            {
                xlSht.Cells[i + 2, 24].Resize[1, res1[i].GetUpperBound(0) + 1].Value = res1[i];
                xlSht.Cells[i + 2, 31].Resize[1, res2[i].GetUpperBound(0) + 1].Value = res2[i];
                xlSht.Cells[i + 2, 36].Resize[1, res3[i].GetUpperBound(0) + 1].Value = res3[i];

            }

            //xlSht.Range("B2").WrapText = false;
            xlSht.Columns[27].WrapText = false;
            ObjWorkExcel.UserControl = true;
            ObjWorkBook.Save();
            ObjWorkBook.Close(0);
            ObjWorkExcel.Quit();
        }

        public void ExportToExcel(int idFirst, int idLast, string fileName, string Type)
        {
            string JOIN, WHERE, ORDER;
            int TYPE, COUNT;
            if (Type == "null")
            {
                ORDER = "";
                JOIN = "";
                WHERE = " WHERE b.id BETWEEN " + idFirst + " AND " + idLast ;
                COUNT = idLast - idFirst + 1;
            }
            else
            {                
                TYPE = GetAdvTypeByString(Type);
                WHERE = " WHERE pd.adv_type = " + TYPE;
                ORDER = " ORDER BY b.id";
                COUNT = GetCount(WHERE);
                JOIN = " JOIN Nedvizh.dbo.adv_pdata AS pd ON (b.id = pd.id)";
            }

            var ObjWorkExcel = new Application();
            var ObjWorkBook = ObjWorkExcel.Workbooks.Add();
            ObjWorkBook.SaveAs(fileName);
            var count = COUNT;
            var res1 = new string[count][];
            var res2 = new string[count][];
            var res3 = new string[count][];
            var res4 = new string[count][];
            var res5 = new double[count][];
            using (var sql = new SqlCommand("SELECT b.id, date_ins, time_ins, date_act_str, date_act, source, is_new, adv_status FROM Nedvizh.dbo.main_table as b " + JOIN + " " + WHERE + ORDER +";", Connection))
            {
                var dataReader = sql.ExecuteReader();
                var result = new object[8];
                var j = 0;
                while (dataReader.Read())
                {
                    dataReader.GetValues(result);
                    res1[j] = new string[8];
                    for (int i = 0; i < 8; i++)
                    {
                        if (i == 4) res1[j][i] = result[i].ToString().Remove(11);
                        else
                            res1[j][i] = result[i].ToString();
                    }
                    j++;
                }
                dataReader.Close();
            }
            using (var sql = new SqlCommand("SELECT adv_link, adv_rlink, adv_title, adv_nlink, adv_curlink, adv_num FROM Nedvizh.dbo.adv_links as b" + JOIN + " " + WHERE + ORDER + ";", Connection))
            {
                var dataReader = sql.ExecuteReader();
                var result = new object[6];
                var j = 0;
                while (dataReader.Read())
                {
                    dataReader.GetValues(result);
                    res2[j] = new string[6];
                    for (int i = 0; i < 6; i++)
                    {
                        res2[j][i] = result[i].ToString();
                    }
                    j++;
                }
                dataReader.Close();
            }
            using (var sql = new SqlCommand("SELECT b.adv_cost, b.adv_area FROM Nedvizh.dbo.adv_pdata as b" + JOIN + " " + WHERE + ORDER + ";", Connection))
            {
                var dataReader = sql.ExecuteReader();
                var result = new object[2];
                var j = 0;
                while (dataReader.Read())
                {
                    dataReader.GetValues(result);
                    res3[j] = new string[3];
                    for (int i = 0; i < 2; i++)
                    {
                        res3[j][i] = result[i].ToString();
                    }
                    j++;
                }
                dataReader.Close();
            }
            using (var sql = new SqlCommand(" SELECT \"type\" FROM Nedvizh.dbo.adv_types ty JOIN Nedvizh.dbo.adv_pdata b ON (b.adv_type = ty.id)" + JOIN + " " + WHERE + ORDER + ";", Connection))
            {
                var dataReader = sql.ExecuteReader();
                var j = 0;
                string type;
                while (dataReader.Read())
                {
                    type = dataReader.GetString(0);
                    res3[j][2] = type;
                    j++;
                }
                dataReader.Close();
            }

            using (var sql = new SqlCommand("SELECT addr_coord, addr01, addr02 FROM Nedvizh.dbo.addr_table as b" + JOIN + " " + WHERE + ORDER + ";", Connection))
            {
                var dataReader = sql.ExecuteReader();
                var result = new object[3];
                var j = 0;
                while (dataReader.Read())
                {
                    dataReader.GetValues(result);
                    res4[j] = new string[3];
                    for (var i = 0; i < 3; i++)
                    {                        
                        res4[j][i] = result[i].ToString();                    
                    }
                    j++;
                }
                dataReader.Close();
            }
            for (int j = 0; j < count; j++)
            {
                try { res4[j][2] = GetTownById(Int16.Parse(res4[j][2])); } catch { }
            }
           
            var j1 = 0;

            using (var sql = new SqlCommand("SELECT adv_cost_red, adv_area_red FROM  Nedvizh.dbo.main_table b LEFT JOIN Nedvizh.dbo.adv_sec_data as m ON (b.id = m.id)" + JOIN + " " + WHERE + ORDER + ";", Connection))
            {
                var dataReader = sql.ExecuteReader();
                var result = new object[2];
                while (dataReader.Read())
                {
                    res5[j1] = new double[2];
                    try
                    {
                        dataReader.GetValues(result);
                        for (int i = 0; i < 2; i++)
                        {
                            try { res5[j1][i] = Convert.ToDouble(result[i]); } catch { res5[j1][i] = 0; }
                        }
                    }
                    catch
                    {
                        res5[j1][0] = 0;
                        res5[j1][1] = 0;
                    }                   
                    j1++;
                }
               // System.Windows.Forms.MessageBox.Show(j1.ToString());
                dataReader.Close();
            }
            
            var xlSht = ObjWorkBook.Sheets[1];
            var titles = new[] { "id", "Дата сохранения", "Время сохранения", "Дата актуальности", "Обработанная актуальность", "Лента(архив\\актуаль)", "Новое", "Статус", "Корневая ссылка", "Ссылка объявления", "Название объявления", "Место хранения", "Из ленты", "id на сайте", "Цена", "Площадь", "Тип объявления", "Координаты", "Район", "Город", "Цена Обработанная", "Площадь обработанная" };
            xlSht.Cells[1, 2].Resize[1, titles.GetUpperBound(0) + 1].Value = titles;
            xlSht.Cells[1, 2].Resize[1, titles.GetUpperBound(0) + 1].EntireRow.Font.Bold = true;
            for (var i = 0; i < count; i++)
            {
                xlSht.Cells[i + 2, 2].Resize[1, res1[i].GetUpperBound(0) + 1].Value = res1[i];
                xlSht.Cells[i + 2, 10].Resize[1, res2[i].GetUpperBound(0) + 1].Value = res2[i];
                xlSht.Cells[i + 2, 16].Resize[1, res3[i].GetUpperBound(0) + 1].Value = res3[i];
                xlSht.Cells[i + 2, 19].Resize[1, res4[i].GetUpperBound(0) + 1].Value = res4[i];
                xlSht.Cells[i + 2, 22].Resize[1, res5[i].GetUpperBound(0) + 1].Value = res5[i]; 
            }
            ObjWorkExcel.UserControl = true;
            ObjWorkBook.Save();
            ObjWorkBook.Close(0);
            ObjWorkExcel.Quit();
            ExportToExcelFull(idFirst, idLast, fileName, Type);
        }
        public int GetCount(string Where)
        {
            var sql = new SqlCommand("SELECT count(pd.id) FROM Nedvizh.dbo.adv_pdata pd " + Where, Connection);    
            var dataReader = sql.ExecuteReader();
            dataReader.Read();
            int res = dataReader.GetInt32(0);
            dataReader.Close();
            return res;            
        }

        public bool UpdateUserInfo(int id, string username, string phone, string mail, int advs)
        {
            if (Connection == null)
            {
                return false;
            }
            using (var sql = new SqlCommand(null, Connection)
            {
                CommandText = "UPDATE Nedvizh.dbo.adv_users_info SET \"adv_user_name\" = @adv_user_name, \"adv_user_phone\" = @adv_user_phone, \"adv_user_mail\" = @adv_user_mail, \"adv_user_num_of_advs\" = @adv_user_num_of_advs WHERE \"id\" = @id"
            })
            {
                var nameParam = new SqlParameter("@adv_user_name", SqlDbType.VarChar, 16) { Value = username };
                var phoneParam = new SqlParameter("@adv_user_phone", SqlDbType.VarChar, 128) { Value = phone };
                var mailParam = new SqlParameter("@adv_user_mail", SqlDbType.VarChar, 128) { Value = mail };
                var advsParam = new SqlParameter("@adv_user_num_of_advs", SqlDbType.SmallInt) { Value = advs };
                var idParam = new SqlParameter("@id", SqlDbType.Int) { Value = id };
                sql.Parameters.Add(nameParam);
                sql.Parameters.Add(phoneParam);
                sql.Parameters.Add(mailParam);
                sql.Parameters.Add(advsParam);
                sql.Parameters.Add(idParam);
                sql.Prepare();
                try
                {
                    sql.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    General.WriteLog(ex.ToString());
                }
            }
            return true;
        }

        public bool UpdateAdvData(int id, int is_new, int source, string num, string link, string rlink, string nlink, string curlink, string title, string price, string square, string geo, string town, string distr, int views, int type)
        {
            if (Connection == null)
            {
                return false;
            }
            using (var sql = new SqlCommand("UPDATE Nedvizh.dbo.main_table SET \"is_new\" = @is_new, \"source\" = @source  WHERE \"id\" = @id;", Connection))
            {
                var newParam = new SqlParameter("@is_new", SqlDbType.TinyInt, 255) { Value = is_new };
                var sourceParam = new SqlParameter("@source", SqlDbType.TinyInt, 255) { Value = source };
                var idParam = new SqlParameter("@id", SqlDbType.Int) { Value = id };
                sql.Parameters.Add(newParam);
                sql.Parameters.Add(sourceParam);
                sql.Parameters.Add(idParam);
                sql.Prepare();
                try
                {
                    sql.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    General.WriteLog(ex.Message);
                    return false;
                }
            }
                using (var sql = new SqlCommand("UPDATE Nedvizh.dbo.adv_links SET \"adv_link\" = @adv_link, \"adv_rlink\" = @adv_rlink, \"adv_nlink\" = @adv_nlink, \"adv_curlink\" = @adv_curlink, \"adv_title\" = @adv_title, \"adv_num\" = @adv_num WHERE \"id\" = @id;", Connection))
            {
                var linkParam = new SqlParameter("@adv_link", SqlDbType.VarChar, 255) { Value = link };
                var rlinkParam = new SqlParameter("@adv_rlink", SqlDbType.VarChar, 255) { Value = rlink };
                var nlinkParam = new SqlParameter("@adv_nlink", SqlDbType.VarChar, 255) { Value = nlink };
                var curlinkParam = new SqlParameter("@adv_curlink", SqlDbType.VarChar, 255) { Value = curlink };
                var titleParam = new SqlParameter("@adv_title", SqlDbType.VarChar, 255) { Value = title };
                var numParam = new SqlParameter("@adv_num", SqlDbType.VarChar, 255) { Value = num };
                var idParam = new SqlParameter("@id", SqlDbType.Int) { Value = id };
                sql.Parameters.Add(linkParam);
                sql.Parameters.Add(rlinkParam);
                sql.Parameters.Add(nlinkParam);
                sql.Parameters.Add(curlinkParam);
                sql.Parameters.Add(titleParam);
                sql.Parameters.Add(numParam);
                sql.Parameters.Add(idParam);
                sql.Prepare();
                try
                {
                    sql.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    General.WriteLog(ex.Message);
                    return false;
                }
            }
            using (var sql = new SqlCommand("UPDATE Nedvizh.dbo.adv_pdata SET \"adv_cost\" = @adv_cost, \"adv_area\" = @adv_area, \"adv_type\" = @adv_type WHERE \"id\" = @id;", Connection))
            {
                var costParam = new SqlParameter("@adv_cost", SqlDbType.VarChar, 255) { Value = price };
                var areaParam = new SqlParameter("@adv_area", SqlDbType.VarChar, 255) { Value = square };
                var typeParam = new SqlParameter("@adv_type", SqlDbType.Int, 255) { Value = type };
                var idParam = new SqlParameter("@id", SqlDbType.Int) { Value = id };
                sql.Parameters.Add(costParam);
                sql.Parameters.Add(areaParam);
                sql.Parameters.Add(typeParam);
                sql.Parameters.Add(idParam);
                sql.Prepare();
                try
                {
                    sql.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    General.WriteLog(ex.Message);
                    return false;
                }
            }
            int townid = GetTownByName(town);
            using (var sql = new SqlCommand("UPDATE Nedvizh.dbo.addr_table SET \"addr_coord\" = @addr_coord, \"addr01\" = @addr01, \"addr02\" = " + townid + " WHERE \"id\" = @id;", Connection))
            {
                var coordParam = new SqlParameter("@addr_coord", SqlDbType.VarChar, 255) { Value = geo };
                var distrParam = new SqlParameter("@addr01", SqlDbType.VarChar, 255) { Value = distr };
                var idParam = new SqlParameter("@id", SqlDbType.Int, 255) { Value = id };
                sql.Parameters.Add(coordParam);
                sql.Parameters.Add(distrParam);
                sql.Parameters.Add(idParam);
                sql.Prepare();
                try
                {
                    sql.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    General.WriteLog(ex.Message);
                    return false;
                }
            }
            using (var sql = new SqlCommand("UPDATE Nedvizh.dbo.adv_users_info SET \"adv_user_num_of_views\" = @adv_user_num_of_views WHERE \"id\" = @id;", Connection))
            {
                var idParam = new SqlParameter("@id", SqlDbType.Int) { Value = id };
                var viewsParam = new SqlParameter("@adv_user_num_of_views", SqlDbType.SmallInt) { Value = views };
                sql.Parameters.Add(idParam);
                sql.Parameters.Add(viewsParam);
                sql.Prepare();
                try
                {
                    sql.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    General.WriteLog(ex.Message);
                    return false;
                }
            }
            return true;
        }
        public void InsertSecData(int id)
        {
            float recost = 0, rearea = 0;
            // ЦЕНА
            using (var sql = new SqlCommand("SELECT adv_cost FROM Nedvizh.dbo.adv_pdata WHERE \"id\" = " + id, Connection))
            {
                var dataReader = sql.ExecuteReader();
                if (dataReader.Read())
                {
                    var cost = dataReader.GetString(0);
                    if (cost != "n/a")
                    {

                        cost = cost.Replace(" ", "");
                        cost = cost.Replace(".", ",");
                        if (cost.Contains("?"))
                        {
                            cost = cost.Replace("?", "");
                        }
                        recost = float.Parse(cost);
                    }
                    else recost = 0;
                }
                dataReader.Close();
            }
            // ПЛОЩАДЬ 
            using (var sql = new SqlCommand("SELECT adv_area FROM Nedvizh.dbo.adv_pdata WHERE \"id\" = " + id, Connection))
            {
                var dataReader = sql.ExecuteReader();
                if (dataReader.Read())
                {
                    var area = dataReader.GetString(0);
                    if (area != "n/a")
                    {
                        area = area.Replace(" ", "");
                        var areaArr = area.ToCharArray();
                        for (int i = 0; i < areaArr.Length; i++)
                        {
                            try
                            {
                                float.Parse(areaArr[i].ToString());
                            }
                            catch
                            {
                                area = area.Remove(i);
                                break;
                            }
                        }
                        area = area.Replace(" ", "");
                        area = area.Replace("кв.м.", "");
                    }
                    else area = "0";
                    dataReader.Close();
                    if (float.TryParse(area, out rearea))
                    {
                        using (var sql2 = new SqlCommand("INSERT INTO Nedvizh.dbo.adv_sec_data (id, adv_area_red, adv_cost_red) VALUES (" + id + ", @rearea, @recost)", Connection))
                        {
                            var areaParam = new SqlParameter("@rearea", SqlDbType.Float, 255) { Value = rearea };
                            var costParam = new SqlParameter("@recost", SqlDbType.Float, 255) { Value = recost };
                            sql2.Parameters.Add(areaParam);
                            sql2.Parameters.Add(costParam);
                            sql2.ExecuteNonQuery();
                        }
                    }

                }
                dataReader.Close();
            }

        }
        public void UpdateLinks(int id, string nlink)
        {
            using (var sql = new SqlCommand("UPDATE Nedvizh.dbo.adv_links SET \"adv_nlink\" = @adv_nlink WHERE \"id\" = @id;", Connection))
            {
                var linkParam = new SqlParameter("@adv_nlink", SqlDbType.VarChar, 255) { Value = nlink };
                var idParam = new SqlParameter("@id", SqlDbType.Int) { Value = id };
                sql.Parameters.Add(linkParam);
                sql.Parameters.Add(idParam);
                sql.ExecuteNonQuery();
                
            }
        }
        public void UpdateIsNew(int id, int status)
        {
            using (var sql = new SqlCommand("UPDATE Nedvizh.dbo.main_table SET \"is_new\" = @is_new WHERE \"id\" = @id;", Connection))
            {
                var linkParam = new SqlParameter("@is_new", SqlDbType.TinyInt, 255) { Value = status };
                var idParam = new SqlParameter("@id", SqlDbType.Int) { Value = id };
                sql.Parameters.Add(linkParam);
                sql.Parameters.Add(idParam);
                sql.ExecuteNonQuery();
            }
        }
    }
}
