using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace AddOn_AC_AL.Models.DataTable
{
    // using System.Xml.Serialization;
    // XmlSerializer serializer = new XmlSerializer(typeof(DataTable));
    // using (StringReader reader = new StringReader(xml))
    // {
    //    var test = (DataTable)serializer.Deserialize(reader);
    // }

    [XmlRoot(ElementName = "Column")]
    public class Column
    {

        [XmlAttribute(AttributeName = "Uid")]
        public string Uid { get; set; }

        [XmlAttribute(AttributeName = "Type")]
        public int Type { get; set; }

        [XmlAttribute(AttributeName = "MaxLength")]
        public int MaxLength { get; set; }
    }

    [XmlRoot(ElementName = "Columns")]
    public class Columns
    {

        [XmlElement(ElementName = "Column")]
        public List<Column> Column { get; set; }
    }

    [XmlRoot(ElementName = "Cell")]
    public class Cell
    {

        [XmlElement(ElementName = "ColumnUid")]
        public string ColumnUid { get; set; }

        [XmlElement(ElementName = "Value")]
        public object Value { get; set; }
    }

    [XmlRoot(ElementName = "Cells")]
    public class Cells
    {

        [XmlElement(ElementName = "Cell")]
        public List<Cell> Cell { get; set; }
    }

    [XmlRoot(ElementName = "Row")]
    public class Row
    {

        [XmlElement(ElementName = "Cells")]
        public Cells Cells { get; set; }
    }

    [XmlRoot(ElementName = "Rows")]
    public class Rows
    {

        [XmlElement(ElementName = "Row")]
        public List<Row> Row { get; set; }
    }

    [XmlRoot(ElementName = "DataTable")]
    public class DataTable
    {

        [XmlElement(ElementName = "Columns")]
        public Columns Columns { get; set; }

        [XmlElement(ElementName = "Rows")]
        public Rows Rows { get; set; }

        [XmlAttribute(AttributeName = "Uid")]
        public string Uid { get; set; }

        [XmlText]
        public string Text { get; set; }
    }
}
