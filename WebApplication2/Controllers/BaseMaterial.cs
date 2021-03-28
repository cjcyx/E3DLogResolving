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
        public List<SubLengthMtoNo> subLengths { get; set; }
    }
    public class SubLengthMtoNo
    {
        public double length;
        public string mtoNo;
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
    //public class TrackingListUpdate
    //{
    //    public string IdentCode;
    //    public string materialLongDescription;
    //    public string partMainSize;
    //    public string packageNo;
    //    public double finalLength;
    //}
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
        public string storage;
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
    public class TrackingListUpdate
    {
        public string workPackageNo;
        public string testPackageNo;
        public string lineNo;
        public string shopISODrawingNo;
        public string pageNo;
        public string totalPage;
        public string shopISODrawingRev;
        public string systemCode;
        public string pipeClass;
        public string spoolNo;
        public string categoryType;
        public string cutPipeNo;
        public string identCode;
        public string partMainSize;
        public string part2ndSize;
        public string pressureClass;
        public string mainThicknessGrade;
        public string thicknessGrade;
        public string mainConnectionType;
        public string connectionType;
        public string materialGrade;
        public string materialLongDescription;
        public string qty;
        public string unit;
        public string finalLength;
        public string materialRequisitionFormNo;
        public string itemNo;
        public string note;
        public TrackingListUpdate(string workPackageNo, string testPackageNo, string lineNo, string shopISODrawingNo, string pageNo, string totalPage, string shopISODrawingRev, string systemCode, string pipeClass, string spoolNo, string categoryType, string cutPipeNo, string identCode, string partMainSize, string part2ndSize, string pressureClass, string mainThicknessGrade, string thicknessGrade, string mainConnectionType, string connectionType, string materialGrade, string materialLongDescription, string qty, string unit, string finalLength, string materialRequisitionFormNo, string itemNo, string note)
        {
            this.workPackageNo = workPackageNo;
            this.testPackageNo = testPackageNo;
            this.lineNo = lineNo;
            this.shopISODrawingNo = shopISODrawingNo;
            this.pageNo = pageNo;
            this.totalPage = totalPage;
            this.shopISODrawingRev = shopISODrawingRev;
            this.systemCode = systemCode;
            this.pipeClass = pipeClass;
            this.spoolNo = spoolNo;
            this.categoryType = categoryType;
            this.cutPipeNo = cutPipeNo;
            this.identCode = identCode;
            this.partMainSize = partMainSize;
            this.part2ndSize = part2ndSize;
            this.pressureClass = pressureClass;
            this.mainThicknessGrade = mainThicknessGrade;
            this.thicknessGrade = thicknessGrade;
            this.mainConnectionType = mainConnectionType;
            this.connectionType = connectionType;
            this.materialGrade = materialGrade;
            this.materialLongDescription = materialLongDescription;
            this.qty = qty;
            this.unit = unit;
            this.finalLength = finalLength;
            this.materialRequisitionFormNo = materialRequisitionFormNo;
            this.itemNo = itemNo;
            this.note = note;
        }
    }
    public class TrackingListInfos
    {
        public string mTONo;
        public string user;
        public string updater;
        public string projectName;
        public string date;
        public string jobNo;
        public string documentNo;
        public string materialPurchasedBy;
        public string discipline;
        public string centralizeWarehouse;
        public string issuedToFabricator;
        public string materialDescription;
        public string issueFor;
        public string workingPackage;
        public string materialsUsedFor;
        public string issuedBy;
        public string checkedBy;
        public string reviewedBy;
    }
    public class ComBinedInfos
    {
        public List<TrackingListUpdate> trackingList;
        public TrackingListInfos infos;
    }
    public class SubmitInfos
    {
        public List<ComBinedInfos> infos;
        public string picker;
        public string submitter;
        public string PickNo;
    }
}