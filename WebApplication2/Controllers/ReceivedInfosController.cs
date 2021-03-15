using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.IO;
using Newtonsoft.Json.Converters;
using Newtonsoft.Json.Linq;
using MySql.Data.MySqlClient;
using Microsoft.AspNetCore.Mvc.ViewFeatures;
using Org.BouncyCastle.Operators;
using System.Runtime.CompilerServices;

namespace WebApplication2.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ReceivedInfosController : ControllerBase
    {
        [HttpPost("PostData")]
        public string PostData(JArray con)
        {
            var result = InsetDbMaterial(con);
            return result.ToString();
        }
        [HttpPost("GetData")]
        public string GetData(JObject idNumber)
        {
            string data = "";
            using (MySqlConnection msconnection = GetConnectionMaterial())
            {
                msconnection.Open();
                bool flag = false;
                var sql = "SELECT * FROM `Materials` WHERE ID="+ (idNumber.ToObject <idNun>()).id + ";";
                using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                {
                    MySqlDataReader reader = mscommand.ExecuteReader();
                    while (reader.Read())
                    {
                        flag = true;
                        var lengths = reader["subLengths"].ToString().Split('&');
                        List<double> subLength = new List<double>();
                        foreach (var item in lengths)
                        {
                            if (item != "")
                            {
                                subLength.Add(Convert.ToDouble(item));
                            }
                        }
                        BaseMaterialReturn resule = new BaseMaterialReturn() {status= flag, id = (long)reader["ID"], length = (double)reader["Length"], subLengths = subLength, user = reader["UserName"].ToString(), itemCode = reader["itemCode"].ToString(), note = reader["note"].ToString() };
                        data += JsonConvert.SerializeObject(resule);
                        break;
                    }
                    reader.Close();
                    if (!flag)
                    {
                        data += JsonConvert.SerializeObject(new BaseMaterialReturn() { status=flag });
                    }
                }
                msconnection.Close();
            }
            return data;
        }        
        [HttpPost("POSTLOG")]
        public bool Post(JArray con)
        {
            var result = InsetDb(con);
            return result;
        }
        [HttpPost("PostLine")]
        public string PostLine(JObject con)
        {
            var result = "";
            using (MySqlConnection msconnection = GetConnection())
            {
                msconnection.Open();
                LineData data = con.ToObject<LineData>();
                if (data.data == "Get")
                {
                    var sql = "SELECT lastLine FROM lastRead where serverName='" + data.serverName + "';";
                    bool flag = true;
                    using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                    {
                        MySqlDataReader reader = mscommand.ExecuteReader();
                        while (reader.Read())
                        {
                            result = reader["lastLine"].ToString();
                            flag = false;
                            break;
                        }
                        reader.Close();
                    }
                    if (flag)
                    {
                        sql = "INSERT INTO lastRead (ID, lastLine, serverName) VALUES (NULL, 0, '" + data.serverName + "');";
                        using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                        {
                            if (mscommand.ExecuteNonQuery() == -1)
                            {
                                result = "false";
                            }
                        }
                        result = "0";
                    }
                }
                else
                {
                    var sql = "update lastRead set lastLine = " + data.lineNo + " where serverName='" + data.serverName + "';";
                    using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                    {
                        if (mscommand.ExecuteNonQuery() == -1)
                        {
                            result = "false";
                        }
                    }
                }
            }
            return result;
        }
        [HttpPost("GetLicData")]
        public string GetLicData(JObject con)
        {
            var result = "";
            using (MySqlConnection msconnection = GetConnection())
            {
                msconnection.Open();
                LineData data = con.ToObject<LineData>();
                if (data.data == "Get")
                {
                    var sql = "SELECT `Module` FROM `ALL_ServerStatus` where serverName= '" + data.serverName + "';";
                    using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                    {
                        MySqlDataReader reader = mscommand.ExecuteReader();
                        while (reader.Read())
                        {
                            result = reader["Module"].ToString();
                            break;
                        }
                        reader.Close();
                    }
                }
                else
                {
                    var sql = "update `ALL_ServerStatus` set `Assignment` = " + data.lineNo + " where `serverName` = '" + data.serverName + "';";
                    using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                    {
                        if (mscommand.ExecuteNonQuery() == -1)
                        {
                            result = "false";
                        }
                    }
                }
            }
            return result;
        }
        [HttpPost("GetAllMTONos")]
        public string GetAllMTONos(JObject con)
        {
            List<string> result = new List<string>();
            using (MySqlConnection msconnection = GetConnectionMaterial())
            {
                msconnection.Open();
                var sql = "SELECT MTONo FROM `TrackNames` WHERE status=false order by id desc;";
                using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                {
                    MySqlDataReader reader = mscommand.ExecuteReader();
                    while (reader.Read())
                    {
                        result.Add (reader["MTONo"].ToString());
                    }
                    reader.Close();
                }
                msconnection.Close();
            }
            return JsonConvert.SerializeObject(result);
        }
        [HttpPost("GetItemCode")]
        public string GetItemCode(JObject con)
        {
            StringClass code = con.ToObject<StringClass>();
            List<TrackingListContain> result = new List<TrackingListContain>();
            List<string> codes = new List<string>();
            using (MySqlConnection msconnection = GetConnectionMaterial())
            {
                msconnection.Open();
                var sql = "SELECT DISTINCT IdentCode FROM `TrackingList` WHERE `MTONo.`='" + code.str +"' order by id desc;";
                using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                {
                    MySqlDataReader reader = mscommand.ExecuteReader();
                    while (reader.Read())
                    {
                        codes.Add(reader["IdentCode"].ToString());
                    }
                    reader.Close();
                }
                foreach(var item in codes)
                {
                    string MaterialLongDescription = "";
                    string PartMainSize = "";
                    double consumption = 0;
                    double totalLength = 0;
                    List<double> FinalLength = new List<double>();
                    sql = "SELECT consumption.consumption,MaterialLongDescription,PartMainSize from TrackingList INNER JOIN consumption on TrackingList.IdentCode=consumption.IdentCode where `MTONo.`='" + code.str + "' and TrackingList.IdentCode = '" + item + "' limit 1;";
                    using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                    {
                        MySqlDataReader reader = mscommand.ExecuteReader();
                        while (reader.Read())
                        {
                            MaterialLongDescription = reader["MaterialLongDescription"].ToString();
                            PartMainSize = reader["PartMainSize"].ToString();
                            consumption = (double)reader["consumption"];
                            break;
                        }
                        reader.Close();
                    }
                    sql = "select FinalLength from TrackingList where `MTONo.` = '" + code.str + "' and IdentCode = '" + item  + "';";
                    using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                    {
                        MySqlDataReader reader = mscommand.ExecuteReader();
                        while (reader.Read())
                        {
                            var leng = (double)reader["FinalLength"];
                            totalLength += leng;
                            FinalLength.Add(leng);
                        }
                        reader.Close();
                    }
                    result.Add(new TrackingListContain(item, MaterialLongDescription, PartMainSize, consumption, FinalLength));
                }
                msconnection.Close();
            }
            return JsonConvert.SerializeObject(result);
        }
        [HttpGet("{server}")]
        public string Get(string server)
        {
            string returnStr = "";
            int currAll = 0;
            int AssiAll = 0;
            string title1 = "";
            string title2 = "";
            DateTime dt;
            List<SheetItems> items = new List<SheetItems>();
            using (MySqlConnection msconnection = GetConnection())
            {
                msconnection.Open();
                string module = "";
                string serverName = "";
                var sql = "SELECT * FROM `WebTitleStatus` WHERE LinkStr = '" + server + "';";
                using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                {
                    MySqlDataReader reader = mscommand.ExecuteReader();
                    if (reader.Read())
                    {
                        returnStr += "<h4>" + reader["title"] + "</h4>\r\n<table>\r\n<tr>\r\n<td valign=\"top\">\r\n<table border=\"1\">\r\n<tr>\r\n<th>" + reader["title3"] + "</th>\r\n<th>正在使用</th>\r\n<th>分配数量</th>\r\n<th>超出/相差数量</th>\r\n</tr>";
                        sql = "SELECT * FROM `" + reader["dbName"] + "` ORDER BY ID ASC; ";
                        title1 = (string)reader["title1"];
                        title2 = (string)reader["title2"];
                        module = (string)reader["Module"];
                        serverName = (string)reader["serverName"];
                    }
                    else
                    {
                        reader.Close();
                        return "连接不存在！";
                    }
                    reader.Close();
                }
                using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                {
                    MySqlDataReader reader = mscommand.ExecuteReader();
                    while (reader.Read())
                    {
                        items.Add(new SheetItems() { Current = 0 ,Dept = (string)reader["Department"], Assi = (int)reader["Assignment"], users = new List<string>() }); ;
                    }
                    reader.Close();
                }
                var sqlmissing = "SELECT userInfos.alias FROM userStatus INNER JOIN userInfos ON userStatus.userName= userInfos.userName where userStatus.module =  '" + module + "' and userStatus.status = 1 and userStatus.serverName = '" + serverName + "'";
                foreach(var item in items)
                {
                    sql = "SELECT userInfos.alias FROM userStatus INNER JOIN userInfos ON userStatus.userName= userInfos.userName where userStatus.module = '"+module+"' and userStatus.status = 1 and userStatus.department = '"+item.Dept+"' and userStatus.serverName = '"+serverName+"';" ;
                    using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                    {
                        MySqlDataReader reader = mscommand.ExecuteReader();
                        while (reader.Read())
                        {
                            item.Current++;
                            item.users.Add((string)reader["alias"]);
                            sqlmissing += "and userStatus.department <>'" +item.Dept +"'";
                        }
                        reader.Close();
                    }
                    currAll += item.Current;
                    AssiAll += item.Assi;
                    returnStr += "<tr>\r\n";
                    returnStr += "<td align=\"center\">"+item.Dept+ "</td>\r\n";
                    returnStr += "<td align=\"center\">" + item.Current + "</td>\r\n";
                    returnStr += "<td align=\"center\">" + item.Assi + "</td>\r\n";
                    int diff =  item.Current- item.Assi;
                    if (diff>0)
                    {
                        returnStr += "<td bgcolor=\"#ff0000\" align=\"center\">" + diff + "</td>\r\n";
                    }
                    else
                    {
                        returnStr += "<td align=\"center\">" + diff + "</td>\r\n";
                    }
                    returnStr += "</tr>\r\n";
                }
                sqlmissing += ";";
                SheetItems sheetitem = new SheetItems() { Current = 0, Assi = 0, Dept = "其他" ,users=new List<string> () };
                items.Add(sheetitem);
                using (MySqlCommand mscommand = new MySqlCommand(sqlmissing, msconnection))
                {
                    MySqlDataReader reader = mscommand.ExecuteReader();
                    while (reader.Read())
                    {
                        sheetitem.Current ++;
                        sheetitem.users.Add((string)reader["alias"]);
                    }
                    reader.Close();
                }
                currAll += sheetitem.Current;
                sql = "SELECT timeStamp FROM log where serverName = '"+serverName+"'ORDER BY ID DESC LIMIT 1;";
                long unixTimeStamp = 0;
                using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                {
                    MySqlDataReader reader = mscommand.ExecuteReader();
                    if (reader.Read())
                    {
                        unixTimeStamp = (int)reader["timeStamp"];
                    }
                    reader.Close();
                }
                System.DateTime startTime = new System.DateTime(1970, 1, 1,8,0,0); // 当地时区
                dt = startTime.AddSeconds(unixTimeStamp);
                returnStr += "<tr>\r\n";
                returnStr += "<td align=\"center\">" + sheetitem.Dept + "</td>\r\n";
                returnStr += "<td align=\"center\">" + sheetitem.Current + "</td>\r\n";
                returnStr += "<td align=\"center\">" + sheetitem.Assi + "</td>\r\n";
                int diff1 = sheetitem.Current- sheetitem.Assi;
                if (diff1 > 0)
                {
                    returnStr += "<td bgcolor=\"#ff0000\" align=\"center\">" + diff1 + "</td>\r\n";
                }
                else
                {
                    returnStr += "<td align=\"center\">" + diff1 + "</td>\r\n";
                }
                returnStr += "</tr>\r\n";
            }
            returnStr += "</table>\r\n</td>\r\n";
            foreach (var item in items)
            {
                returnStr += "<td valign=\"top\">\r\n";
                returnStr += "<table border=\"1\">\r\n";
                returnStr += "<tr>\r\n<th>"+item.Dept+"</th>\r\n</tr>\r\n";
                foreach (var user in item.users)
                {
                    returnStr += "<tr>\r\n<td>" + user + "</td>\r\n</tr>\r\n";
                }
                returnStr += "</table>\r\n</td>\r\n";
            }            
            returnStr += "</tr>\r\n</table>\r\n";
            returnStr += "<br>\r\n";
            returnStr += "<br>\r\n更新时间：" + dt.ToString("yyyy-MM-dd HH:mm:ss") + "\r\n<br>\r\n";
            returnStr += title2+"共分配 "+ AssiAll + " 个，目前使用 " + currAll +" 个。";
            returnStr += "<br>\r\n";
            returnStr += "<progress value=\""+ currAll+ "\" max=\""+ AssiAll + "\" ></progress>";
            returnStr += "<br>\r\n";
            return returnStr;
        }
        private bool InsetDb(JArray con)
        {
            var result = true;
            using (MySqlConnection msconnection = GetConnection())
            {
                msconnection.Open();
                List<ReceivedInfos> obj2 = con.ToObject<List<ReceivedInfos>>();
                foreach (var item in obj2)
                {
                    var sql = "INSERT INTO `log` (`ID`, `timeStamp`, `module`, `status`, `currentUsage`, `compName`, `userName`, `serverName`) VALUES (NULL, " + item.timeStamp + ", '" + item.module + "', " + item.status + ", " + item.currentUsage + ", '" + item.compName + "', '" + item.userName + "', '" + item.serverName + "');";
                    using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                    {
                        if (mscommand.ExecuteNonQuery() == -1)
                        {
                            result = false;
                        }
                    }
                    var nowStatus = 3;
                    sql = "SELECT usageCount FROM usageNumCount where module = '" + item.module + "' and serverName = '" + item.serverName + "';";
                    bool flag = true;
                    using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                    {
                        MySqlDataReader reader = mscommand.ExecuteReader();
                        while (reader.Read())
                        {
                            sql = "UPDATE usageNumCount set usageCount = " + item.currentUsage + " where module = '" + item.module + "' and serverName = '" + item.serverName + "';";
                            var usageChange = item.currentUsage - (int)reader["usageCount"];
                            if (usageChange > 0)
                            {
                                nowStatus = 1;
                            }
                            else if (usageChange < 0)
                            {
                                nowStatus = 0;
                            }
                            flag = false;
                            break;
                        }
                        if (flag)
                        {
                            sql = "INSERT INTO usageNumCount(ID,module,usageCount,serverName) VALUES (NULL,'" + item.module + "'," + item.currentUsage + ",'" + item.serverName + "');";
                        }
                        reader.Close();
                    }
                    using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                    {
                        if (mscommand.ExecuteNonQuery() == -1)
                        {
                            result = false;
                        }
                    }
                    if (nowStatus != 3)
                    {
                        sql = "SELECT ID FROM userStatus where userName = '" + item.userName + "'and compName = '"+item.compName +"' and  module = '" + item.module + "' and serverName = '" + item.serverName + "';";
                        using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                        {
                            MySqlDataReader reader = mscommand.ExecuteReader();
                            if (reader.Read())
                            {
                                sql = "UPDATE userStatus set status = " + nowStatus.ToString() + " where ID = " + reader["ID"].ToString() + ";";
                            }
                            else
                            {
                                //var depart = item.userName.Substring(item.userName.Length - 3).ToUpper();
                                var depart = "";
                                sql = "INSERT INTO userStatus(ID,userName,compName,department,status,module,serverName) VALUES (NULL,'" + item.userName + "','" + item.compName + "','" + depart + "'," + nowStatus.ToString() + ",'" + item.module + "','" + item.serverName + "');";
                            }
                            reader.Close();
                        }
                        using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                        {
                            if (mscommand.ExecuteNonQuery() == -1)
                            {
                                result = false;
                            }
                        }
                    }
                    sql = "SELECT userName FROM userInfos where userName = '" + item.userName + "';";
                    using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                    {
                        MySqlDataReader reader = mscommand.ExecuteReader();
                        if (!reader.Read())
                        {
                            sql = "INSERT INTO `userInfos` (`ID`, `userName`, `alias`) VALUES (NULL, '" + item.userName + "', '" + item.userName + "');";
                        }
                        reader.Close();
                    }

                    using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                    {
                        if (mscommand.ExecuteNonQuery() == -1)
                        {
                            result = false;
                        }
                    }
                }
                msconnection.Close();
            }
            return result;
        }
        private bool InsetDbMaterial(JArray con)
        {
            var result = true;
            using (MySqlConnection msconnection = GetConnectionMaterial())
            {
                msconnection.Open();
                List<BaseMaterial> obj2 = con.ToObject<List<BaseMaterial>>();
                foreach (var item in obj2)
                {
                    string subLength = "";
                    foreach (var item1 in item.subLengths)
                    {
                        subLength += item1.ToString() + "&";
                    }
                    var sql = "INSERT INTO `Materials` (`ID`, `Length`, `UserName`, `itemCode`, `note`, `subLengths`,`TagNo`) VALUES (" + item.id + ", " + item.length + ", '" + item.user + "', '" + item.itemCode + "', '" + item.note + "', '" + subLength + "','"+item.tagNo +"');";
                    using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                    {
                        if (mscommand.ExecuteNonQuery() == -1)
                        {
                            result = false;
                        }
                    }
                }
                msconnection.Close();
            }
            return result;
        }
        private MySqlConnection GetConnection()
        {
            return new MySqlConnection("server=127.0.0.1;database=E3DLicManagement;user=E3DLic;password=LKdMOThGi39eF4Yo=");
        }
        private MySqlConnection GetConnectionMaterial()
        {
            return new MySqlConnection("server=127.0.0.1;database=MaterialDatabase;user=MaterialAdmin;password=fO_bj5G)._nuyeIg");
        }
    }
    public class SheetItems
    {
        public string Dept { get; set; }
        public int Current { get; set; }
        public int Assi { get; set; }
        public List<string> users { get; set; }
    }
}
