using System;
using System.Data;
using System.Linq;

namespace SigningDocumentExample
{
    class SignPersonsRepository : RepositoryBase
    {
        private static DataTable _dataTable;

        public SignPersonsRepository(DataTable dataTable) : base(dataTable)
        {
            _dataTable = dataTable;
        }

        public new Object SearchFor(Func<DataRow, bool> predicate, Func<DataRow, object> selector)
        {
            return _dataTable.AsEnumerable().Where(predicate).Select(selector).FirstOrDefault();
        }

        public string GetSignerPositionByName(string signerName)
        {
            return (string)this.SearchFor(p => p.Field<string>("Name") == signerName, row => row.Field<string>("Position"));
        }

        public Guid GetSignerIdByName(string signerName)
        {
            return (Guid)this.SearchFor(p => p.Field<string>("Name") == signerName, row => row.Field<Guid>("Id"));
        }

        public Byte[] GetSignerImageByName(string signerName)
        {
            return (byte[])this.SearchFor(p => p.Field<string>("Name") == signerName, row => row.Field<Byte[]>("Image"));
        }
    }
}
