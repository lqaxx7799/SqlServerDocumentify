using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SqlServerDocumentify
{
    class TableSchema
    {
        public string TableName { get; set; }
        public string ColumnName { get; set; }
        public string DataType { get; set; }
        public int? CharacterMaximumLength { get; set; }
    }
}
