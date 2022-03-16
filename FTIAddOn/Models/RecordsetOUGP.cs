using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace AddOn_AC_AL.Models.OUGP
{
    // using System.Xml.Serialization;
    // XmlSerializer serializer = new XmlSerializer(typeof(BOM));
    // using (StringReader reader = new StringReader(xml))
    // {
    //    var test = (BOM)serializer.Deserialize(reader);
    // }

    [XmlRoot(ElementName = "AdmInfo")]
    public class AdmInfo
    {

        [XmlElement(ElementName = "Object")]
        public int Object { get; set; }
    }

    [XmlRoot(ElementName = "row")]
    public class Row
    {

        [XmlElement(ElementName = "UgpEntry")]
        public int UgpEntry { get; set; }

        [XmlElement(ElementName = "UgpCode")]
        public string UgpCode { get; set; }

        [XmlElement(ElementName = "UgpName")]
        public string UgpName { get; set; }

        [XmlElement(ElementName = "BaseUom")]
        public int BaseUom { get; set; }

        [XmlElement(ElementName = "DataSource")]
        public string DataSource { get; set; }

        [XmlElement(ElementName = "UserSign")]
        public int UserSign { get; set; }

        [XmlElement(ElementName = "LogInstanc")]
        public int LogInstanc { get; set; }

        [XmlElement(ElementName = "UserSign2")]
        public int UserSign2 { get; set; }

        [XmlElement(ElementName = "UpdateDate")]
        public object UpdateDate { get; set; }

        [XmlElement(ElementName = "CreateDate")]
        public object CreateDate { get; set; }

        [XmlElement(ElementName = "Locked")]
        public string Locked { get; set; }
    }

    [XmlRoot(ElementName = "OUGP")]
    public class OUGP
    {

        [XmlElement(ElementName = "row")]
        public List<Row> Row { get; set; }
    }

    [XmlRoot(ElementName = "BO")]
    public class BO
    {

        [XmlElement(ElementName = "AdmInfo")]
        public AdmInfo AdmInfo { get; set; }

        [XmlElement(ElementName = "OUGP")]
        public OUGP OUGP { get; set; }
    }

    [XmlRoot(ElementName = "BOM")]
    public class BOM
    {

        [XmlElement(ElementName = "BO")]
        public BO BO { get; set; }
    }


}
