using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Threading;

namespace WebApplication2.Controllers
{
    class BaseMaterial
    {
        public long id { get; set; }
        public double length { get; set; }
        public string user { get; set; }
        public string itemCode { get; set; }
        public string note { get; set; }
        public string tagNo { get; set; }
        public List<double> subLengths { get; set; }
    }
    class BaseMaterialReturn
    {
        public bool status { get; set; }
        public long id { get; set; }
        public double length { get; set; }
        public string user { get; set; }
        public string itemCode { get; set; }
        public string note { get; set; }
        public string tagNo { get; set; }
        public List<double> subLengths { get; set; }
    }
    class idNun
    {
        public long id { get; set; }
    }
    public class TrackingListContain
    {
        public string identCode { get; }
        public string materialDesc { get; }
        public string partSize { get; }
        public double consumption { get; }
        public double totalLength { get; }
        public List<double> FinalLength { get; }
        public string note { get; set; }
        public TrackingListContain(string identCode, string materialDesc, string partSize, double consumption, List<double> FinalLength)
        {
            this.identCode = identCode;
            this.materialDesc = materialDesc;
            this.partSize = partSize;
            this.consumption = consumption;
            this.FinalLength = FinalLength;
            this.totalLength = 0;
            this.note = "";
            foreach (var item in FinalLength)
            {
                totalLength += item;
            }
        }
    }
    public class Oddments
    {
        public long id { get; set; }
        public double length { get; set; }
        public string packageNum { get; set; }
    }
    public class LengthReturn
    {
        public long ID;
        public double Length;
    }
    public class LengthUpdate
    {
        public string IdentCode;
        public string MtoNO;
        public double Length;
        public string user;
    }
    public class ItemCode
    {
        public string code { get; }
        public string MtoNo { get; }
        public ItemCode(string itemCode, string MtoNo)
        {
            this.code = itemCode;
            this.MtoNo = MtoNo;
        }
    }
    public class ItemCodeinfos
    {
        public string IdentCode;
        public string materialLongDescription;
        public string partMainSize;
        public double finalLength;
        public ItemCodeinfos(string IdentCode,string materialLongDescription,string partMainSize,double length)
        {
            this.IdentCode=IdentCode;
            this.materialLongDescription = materialLongDescription;
            this.partMainSize = partMainSize;
            this.finalLength = length;
        }
    }
    public class StringClass
    {
        public string str { get; set; }
    }
    class TrackingListItem
    {
        public long id;
        public double length;

    }
    public class XlsContain
    {
        public string IdentCode;
        public string materialLongDescription;
        public string partMainSize;
        public double finalLength;
        public long baseMetalId;
    }
}