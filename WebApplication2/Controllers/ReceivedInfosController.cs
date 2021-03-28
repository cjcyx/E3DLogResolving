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
using System.Reflection.Metadata;
using Org.BouncyCastle.Bcpg.OpenPgp;
using System.Net.Mail;
using NPOI.HSSF.UserModel;

namespace WebApplication2.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ReceivedInfosController : ControllerBase
    {
        public IdWorker worker = new IdWorker(1);
        [HttpPost("PostData")]
        public string PostData(JArray con)
        {
            var result = InsetDbMaterial(con);
            return result.ToString();
        }
        [HttpPost("GetData")]
        public string GetData(JObject idNumber)
        {
            BaseMaterialReturn returnInfo = new BaseMaterialReturn();
            long id = (idNumber.ToObject<idNun>()).id;
            using (MySqlConnection msconnection = GetConnectionMaterial())
            {
                msconnection.Open();
                if (id == 0)
                {
                    var sql1 = "SELECT ID FROM `Materials` Order BY ID desc limit 1;";
                    using (MySqlCommand mscommand = new MySqlCommand(sql1, msconnection))
                    {
                        MySqlDataReader reader = mscommand.ExecuteReader();
                        {
                            while (reader.Read())
                            {
                                id = (long)reader["ID"];
                                break;
                            }
                        }
                        reader.Close();
                    }
                }
                bool flag = false;
                var sql = "SELECT * FROM `Materials` WHERE ID="+ id + ";";
                using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                {
                    MySqlDataReader reader = mscommand.ExecuteReader();
                    while (reader.Read())
                    {
                        flag = true;
                        BaseMaterialReturn resule = new BaseMaterialReturn() {status= flag, id = (long)reader["ID"], length = (double)reader["Length"],  user = reader["UserName"].ToString(), itemCode = reader["itemCode"].ToString(), note = reader["note"].ToString() };
                        returnInfo = resule;
                        break;
                    }
                    reader.Close();
                    sql = "SELECT `TrackingList`.`FinalLength` FROM `TrackingList` WHERE `BaseMetalId` = '" + id + "';";
                    mscommand.CommandText = sql;
                    List<double> subLengths = new List<double>();
                    reader = mscommand.ExecuteReader();
                    while (reader.Read())
                    {
                        flag = true;
                        subLengths.Add((double)reader["FinalLength"]);
                    }
                    returnInfo.subLengths = subLengths;
                    reader.Close();
                    if (!flag)
                    {
                        returnInfo.status = flag;
                    }
                }
                msconnection.Close();
            }
            return JsonConvert.SerializeObject(returnInfo);
        }
        [HttpPost("Getoddments")]
        public string Getoddments(JObject idNumber)
        {
            string data = "";
            List<Oddments> OddmentsReturn = new List<Oddments>();
            using (MySqlConnection msconnection = GetConnectionMaterial())
            {
                msconnection.Open();
                bool flag = false;
                var sql = "SELECT * FROM `oddments` WHERE IdentCode='" + (idNumber.ToObject<StringClass>()).str + "';";
                using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                {
                    MySqlDataReader reader = mscommand.ExecuteReader();
                    while (reader.Read())
                    {
                        OddmentsReturn.Add(new Oddments() { id=(long)reader["ID"],packageNum = (string)reader["PackageNum"] ,length = (double)reader["Length"] });
                    }
                    reader.Close();
                }
                msconnection.Close();
            }
            if (OddmentsReturn.Count != 0)
            {
                data = JsonConvert.SerializeObject(OddmentsReturn);
            }
            return data;
        }
        [HttpPost("Getpicked")]
        public string Getpicked(JObject idNumber)
        {
            string data = "";
            List<Oddments> OddmentsReturn = new List<Oddments>();
            using (MySqlConnection msconnection = GetConnectionMaterial())
            {
                msconnection.Open();
                bool flag = false;
                var sql = "SELECT * FROM `Materials` WHERE itemCode ='" + (idNumber.ToObject<LengthUpdate>()).IdentCode  + "' and note='" + (idNumber.ToObject<LengthUpdate>()).MtoNO +"'; ";
                using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                {
                    MySqlDataReader reader = mscommand.ExecuteReader();
                    while (reader.Read())
                    {
                        OddmentsReturn.Add(new Oddments() { id = (long)reader["ID"],  length = (double)reader["Length"] });
                    }
                    reader.Close();
                }
                msconnection.Close();
            }
            if (OddmentsReturn.Count != 0)
            {
                data = JsonConvert.SerializeObject(OddmentsReturn);
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
            StringClass code = con.ToObject<StringClass>();
            List<string> result = new List<string>();
            using (MySqlConnection msconnection = GetConnectionMaterial())
            {
                msconnection.Open();
                var sql = "SELECT MTONo FROM `TrackNames` WHERE status=false and User = '"+code.str +"'order by id desc;";
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
            List<ItemCodeinfos> result = new List<ItemCodeinfos>();
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
                foreach (var item in codes)
                {
                    string MaterialLongDescription = "";
                    string PartMainSize = "";
                    sql = "select MaterialLongDescription,	PartMainSize from TrackingList where `MTONo.` = '" + code.str + "' and IdentCode = '" + item + "';";
                    using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                    {
                        MySqlDataReader reader = mscommand.ExecuteReader();
                        while (reader.Read())
                        {
                            MaterialLongDescription = reader["MaterialLongDescription"].ToString();
                            PartMainSize = reader["PartMainSize"].ToString();
                            break;
                        }
                        reader.Close();
                    }
                    sql = "SELECT FinalLength FROM `TrackingList` WHERE `MTONo.`= '" + code.str + "' and `IdentCode`= '" + item  + "' and BaseMetalId = 0;";
                    double finalLength = 0;
                    using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                    {
                        MySqlDataReader reader = mscommand.ExecuteReader();
                        while (reader.Read())
                        {
                            finalLength+=(double)reader["FinalLength"];
                        }
                        reader.Close();
                    }
                    result.Add(new ItemCodeinfos(item, MaterialLongDescription, PartMainSize, finalLength));
                }
                msconnection.Close();
            }
            return JsonConvert.SerializeObject(result);
        }
        [HttpPost("DeleteMaterial")]
        public string DeleteMaterial(JObject con)
        {
            var code = con.ToObject<LengthReturn>();
            double length = 0;
            using (MySqlConnection msconnection = GetConnectionMaterial())
            {
                msconnection.Open();
                var sql = "SELECT FinalLength FROM `TrackingList` WHERE `BaseMetalId`=" + code.ID + ";";
                using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                {
                    MySqlDataReader reader = mscommand.ExecuteReader();
                    while (reader.Read())
                    {
                        length += (double)reader["FinalLength"];
                    }
                    reader.Close();
                }
                //sql = "SELECT Length FROM `oddments` WHERE `BaseMetalId`=" + code.ID + ";";
                //using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                //{
                //    MySqlDataReader reader = mscommand.ExecuteReader();
                //    while (reader.Read())
                //    {
                //        length += (double)reader["Length"];
                //    }
                //    reader.Close();
                //}
                //sql = "SELECT Length FROM `saraplength` WHERE `BaseMetalId`=" + code.ID + ";";
                //using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                //{
                //    MySqlDataReader reader = mscommand.ExecuteReader();
                //    while (reader.Read())
                //    {
                //        length += (double)reader["Length"];
                //    }
                //    reader.Close();
                //}
                sql = "UPDATE `TrackingList` SET `BaseMetalId` = '0' WHERE `TrackingList`.`BaseMetalId` = "+code.ID +";";
                using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                {
                    if (mscommand.ExecuteNonQuery() == -1)
                    {
                        
                    }
                }
                sql = "DELETE FROM `oddments` WHERE `oddments`.`BaseMetalId` = " + code.ID + ";";
                using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                {
                    if (mscommand.ExecuteNonQuery() == -1)
                    {

                    }
                }
                sql = "DELETE FROM `saraplength` WHERE `saraplength`.`BaseMetalId` = " + code.ID + ";";
                using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                {
                    if (mscommand.ExecuteNonQuery() == -1)
                    {

                    }
                }
                sql = "DELETE FROM `Materials` WHERE ID = " + code.ID + ";";
                using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                {
                    if (mscommand.ExecuteNonQuery() == -1)
                    {

                    }
                }
                msconnection.Close();
            }
            code.Length = length;
            return JsonConvert.SerializeObject(code);
        }
        [HttpPost("SubmitPackage")]
        public string SubmitPackage(JObject con)
        {
            var code = con.ToObject<StringClass>();
            List<XlsContain> list = new List<XlsContain>();
            using (MySqlConnection msconnection = GetConnectionMaterial())
            {
                msconnection.Open();
                var datetime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                var sql = "UPDATE `TrackNames` SET `HandleDate` = '" + datetime + "' , `status` = '1' WHERE MTONo = '" + code.str + "';";
                using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                {
                    if (mscommand.ExecuteNonQuery() == -1)
                    {

                    }
                }
                sql = "SELECT * FROM `TrackingList` WHERE `MTONo.` = '" + code.str + "' ORDER BY `TrackingList`.`IdentCode` ASC , `TrackingList`.`BaseMetalId` ASC;";
                using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                {
                    MySqlDataReader reader = mscommand.ExecuteReader();
                    while (reader.Read())
                    {
                        list.Add(new XlsContain() { finalLength = (double)reader["FinalLength"],IdentCode = reader["IdentCode"].ToString (), materialLongDescription = reader["MaterialLongDescription"].ToString() ,partMainSize = reader["PartMainSize"].ToString() ,baseMetalId = (long)reader["BaseMetalId"] });
                    }
                    reader.Close();
                }
                msconnection.Close();
            }
            code.str = GenerateXLS(code.str, list).ToString();
            return JsonConvert.SerializeObject(code);
        }
        [HttpPost("InsertMaterial")]
        public string InsertMaterial(JObject con)
        {
            var code = con.ToObject<LengthUpdate>();
            LengthReturn lengthReturn = new LengthReturn() { ID = worker.nextId()};
            double cutloss = 0;
            List<TrackingListItem> oriItems = new List<TrackingListItem>();
            double saraplength = 100;
            using (MySqlConnection msconnection = GetConnectionMaterial())
            {
                msconnection.Open();
                var sql = "SELECT ID,FinalLength FROM `TrackingList` WHERE `MTONo.`= '" + code.MtoNO  + "' and `IdentCode`= '" + code.IdentCode + "' and BaseMetalId = 0 order by FinalLength DESC;";
                using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                {
                    MySqlDataReader reader = mscommand.ExecuteReader();
                    while (reader.Read())
                    {
                        oriItems.Add(new TrackingListItem() { id = (Int32)reader["ID"],length =(double)reader["FinalLength"] });
                    }
                    reader.Close();
                }
                sql = "SELECT consumption FROM `consumption` WHERE `IdentCode`='" + code.IdentCode  + "';";
                using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                {
                    MySqlDataReader reader = mscommand.ExecuteReader();
                    while (reader.Read())
                    {
                        cutloss =(double)reader["consumption"];
                        break;
                    }
                    reader.Close();
                }
                List<long> returnIds = Typesetting(oriItems,code.Length ,saraplength ,cutloss ,out double remainLength,out double totalLength);
                lengthReturn.Length = totalLength; 
                if (totalLength != 0) 
                { 
                    if (remainLength> saraplength)
                    {
                        sql = "INSERT INTO `oddments` (`ID`, `IdentCode`, `PackageNum`, `Length`, `BaseMetalId`, `Storage`) VALUES (NULL, '" + code.IdentCode +"', '"+code.MtoNO + "', "+ remainLength + ", "+ lengthReturn .ID+ ",'"+code.storage+ "');";
                    }
                    else
                    {
                        sql = "INSERT INTO `saraplength` (`ID`, `IdentCode`, `PackageNum`, `Length`, `BaseMetalId`) VALUES (NULL, '" + code.IdentCode + "', '" + code.MtoNO + "', " + remainLength + ", " + lengthReturn.ID + ");";
                    }
                    using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                    {
                        if (mscommand.ExecuteNonQuery() == -1)
                        {

                        }
                    }
                    sql = "INSERT INTO `Materials` (`ID`, `Length`, `UserName`, `itemCode`, `note`, `TagNo`) VALUES ("+lengthReturn.ID +", '" + code.Length  + "', '"+code.user +"', '" + code.IdentCode + "','"+code.MtoNO +"','');";
                    using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                    {
                        if (mscommand.ExecuteNonQuery() == -1)
                        {

                        }
                    }
                    foreach (var item in returnIds)
                    {
                        sql = "UPDATE `TrackingList` SET `BaseMetalId` = '"+ lengthReturn.ID + "' WHERE `TrackingList`.`ID` = " +item+";";
                        using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                        {
                            if (mscommand.ExecuteNonQuery() == -1)
                            {

                            }
                        }

                    }
                }
                msconnection.Close();
            }            
            return JsonConvert.SerializeObject(lengthReturn);
        }
        [HttpPost("LogInGetPassword")]
        public string LogInGetPassword(JObject con)
        {
            StringClass code = con.ToObject<StringClass>();
            StringClass result = new StringClass() { str = "NotFound"};
            using (MySqlConnection msconnection = GetConnectionMaterial())
            {
                msconnection.Open();
                var sql = "SELECT Password,authority,Alies FROM `UserStorage` WHERE `UserName`='" + code.str + "';";                
                using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                {
                    MySqlDataReader reader = mscommand.ExecuteReader();
                    while (reader.Read())
                    {
                        result.str = reader["Password"].ToString()+"&";
                        result.str += reader["authority"].ToString() + "&";
                        result.str += reader["Alies"].ToString();
                        break;
                    }
                    reader.Close();
                }               
                msconnection.Close();
            }
            return JsonConvert.SerializeObject(result);
        }

        [HttpPost("GetUsersByAuth")]
        public string GetUsersByAuth(JObject con)
        {
            StringClass code = con.ToObject<StringClass>();
            List<string> result = new List<string>();
            using (MySqlConnection msconnection = GetConnectionMaterial())
            {
                msconnection.Open();
                var sql = "SELECT Alies FROM `UserStorage` WHERE `authority` like '%" + code.str + "%';";
                using (MySqlCommand mscommand = new MySqlCommand(sql, msconnection))
                {
                    MySqlDataReader reader = mscommand.ExecuteReader();
                    while (reader.Read())
                    {
                        result.Add( reader["Alies"].ToString());
                    }
                    reader.Close();
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
        private bool sendMail(string MtoNo)
        {
            bool status = false;
            MailMessage msg = new MailMessage();
            msg.To.Add("chaijingchuan@bomesc.com");
            msg.CC.Add("zhaofugui@bomesc.com");
            msg.From = new MailAddress("mero2piping@bomesc.com", "BOMESC领料工具", System.Text.Encoding.UTF8);
            msg.Subject = MtoNo+"领料结果";
            msg.SubjectEncoding = System.Text.Encoding.UTF8;
            msg.Body = "邮件内容";
            msg.BodyEncoding = System.Text.Encoding.UTF8;
            msg.IsBodyHtml = false;
            msg.Attachments.Add(new Attachment(@"tmp/"+MtoNo+".xls"));
            SmtpClient client = new SmtpClient();
            client.Credentials = new System.Net.NetworkCredential("mero2piping@bomesc.com", "gAP8MrenfgGnF7K");
            client.Host = "192.168.0.7";
            object userState = msg;
            try
            {
                client.SendAsync(msg, userState);
                status = true;
            }
            catch (System.Net.Mail.SmtpException ex)
            {

            }
            return status;
        }
        private bool GenerateXLS(string MtoNo,List<XlsContain> list)
        {
            HSSFWorkbook xls = new HSSFWorkbook();
            xls.CreateSheet("sheet1"); 
            HSSFSheet sheet1 = (HSSFSheet)xls.GetSheet("sheet1");
            int k = 0;
            sheet1.CreateRow(k);
            sheet1.GetRow(k).CreateCell(1).SetCellValue("IdentCode");
            sheet1.GetRow(k).CreateCell(2).SetCellValue("MaterialLongDescription");
            sheet1.GetRow(k).CreateCell(3).SetCellValue("PartMainSize");
            sheet1.GetRow(k).CreateCell(4).SetCellValue("FinalLength");
            sheet1.GetRow(k).CreateCell(5).SetCellValue("BaseMetalId");
            k++;
            foreach(var item in list)
            {
                sheet1.CreateRow(k);
                sheet1.GetRow(k).CreateCell(0).SetCellValue(k);
                sheet1.GetRow(k).CreateCell(1).SetCellValue(item.IdentCode );
                sheet1.GetRow(k).CreateCell(2).SetCellValue(item.materialLongDescription);
                sheet1.GetRow(k).CreateCell(3).SetCellValue(item.partMainSize);
                sheet1.GetRow(k).CreateCell(4).SetCellValue(item.finalLength );
                sheet1.GetRow(k).CreateCell(5).SetCellValue(item.baseMetalId.ToString ());
                k++;
            }
            FileStream fileStream = new FileStream(@"tmp/" + MtoNo + ".xls", FileMode.Create);
            xls.Write(fileStream); 
            fileStream.Close();
            return sendMail(MtoNo);
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
        private List<long> Typesetting(List<TrackingListItem> oriItems, double length, double saraplength, double cutloss, out double remainLength,out double totalLength)//saraplength是废料长度，cutloss是切割损耗    
        {
            List<long> NestingID = new List<long>();
            totalLength = 0;

            for (int i = 0; i < oriItems.Count; i++)//去除等于输入长度的。
            {
                if (oriItems[i].length != 0)
                    if (length - oriItems[i].length <= cutloss && 0 <= length - oriItems[i].length)
                    {
                        NestingID.Add(oriItems[i].id);
                        totalLength += oriItems[i].length;
                        oriItems[i].length = 0;
                        length = 0;
                        break;
                    }
            };

            for (int i = 0; i < oriItems.Count; i++)
            {
                if (oriItems[i].length != 0)
                    if (length - oriItems[i].length > cutloss && length - oriItems[i].length <= saraplength)//小于废料长
                    {
                        NestingID.Add(oriItems[i].id);
                        totalLength += oriItems[i].length;
                        length -= oriItems[i].length;
                        length -= cutloss;
                        oriItems[i].length = 0;

                        break;
                    }
                    else if (length - oriItems[i].length > saraplength)//大于废料长
                    {
                        NestingID.Add(oriItems[i].id);
                        totalLength += oriItems[i].length;
                        length -= oriItems[i].length;//减长度
                        length -= cutloss;//减切损
                        oriItems[i].length = 0;

                    }
            };
            remainLength = length;
            return NestingID;
        }
        public class IdWorker
        {
            //机器ID
            private static long workerId;
            private static long twepoch = 687888001020L; //唯一时间，这是一个避免重复的随机量，自行设定不要大于当前时间戳
            private static long sequence = 0L;
            private static int workerIdBits = 4; //机器码字节数。4个字节用来保存机器码(定义为Long类型会出现，最大偏移64位，所以左移64位没有意义)
            public static long maxWorkerId = -1L ^ -1L << workerIdBits; //最大机器ID
            private static int sequenceBits = 10; //计数器字节数，10个字节用来保存计数码
            private static int workerIdShift = sequenceBits; //机器码数据左移位数，就是后面计数器占用的位数
            private static int timestampLeftShift = sequenceBits + workerIdBits; //时间戳左移动位数就是机器码和计数器总字节数
            public static long sequenceMask = -1L ^ -1L << sequenceBits; //一微秒内可以产生计数，如果达到该值则等到下一微妙在进行生成
            private long lastTimestamp = -1L;

            /// <summary>
            /// 机器码
            /// </summary>
            /// <param name="workerId"></param>
            public IdWorker(long workerId)
            {
                if (workerId > maxWorkerId || workerId < 0)
                    throw new Exception(string.Format("worker Id can't be greater than {0} or less than 0 ", workerId));
                IdWorker.workerId = workerId;
            }

            public long nextId()
            {
                lock (this)
                {
                    long timestamp = timeGen();
                    if (this.lastTimestamp == timestamp)
                    { //同一微妙中生成ID
                        IdWorker.sequence = (IdWorker.sequence + 1) & IdWorker.sequenceMask; //用&运算计算该微秒内产生的计数是否已经到达上限
                        if (IdWorker.sequence == 0)
                        {
                            //一微妙内产生的ID计数已达上限，等待下一微妙
                            timestamp = tillNextMillis(this.lastTimestamp);
                        }
                    }
                    else
                    { //不同微秒生成ID
                        IdWorker.sequence = 0; //计数清0
                    }
                    if (timestamp < lastTimestamp)
                    { //如果当前时间戳比上一次生成ID时时间戳还小，抛出异常，因为不能保证现在生成的ID之前没有生成过
                        throw new Exception(string.Format("Clock moved backwards.  Refusing to generate id for {0} milliseconds",
                            this.lastTimestamp - timestamp));
                    }
                    this.lastTimestamp = timestamp; //把当前时间戳保存为最后生成ID的时间戳
                    long nextId = (timestamp - twepoch << timestampLeftShift) | IdWorker.workerId << IdWorker.workerIdShift | IdWorker.sequence;
                    return nextId;
                }
            }

            /// <summary>
            /// 获取下一微秒时间戳
            /// </summary>
            /// <param name="lastTimestamp"></param>
            /// <returns></returns>
            private long tillNextMillis(long lastTimestamp)
            {
                long timestamp = timeGen();
                while (timestamp <= lastTimestamp)
                {
                    timestamp = timeGen();
                }
                return timestamp;
            }

            /// <summary>
            /// 生成当前时间戳
            /// </summary>
            /// <returns></returns>
            private long timeGen()
            {
                return (long)(DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalMilliseconds;
            }
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
