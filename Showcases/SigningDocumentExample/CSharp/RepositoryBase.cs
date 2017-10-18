using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace SigningDocumentExample
{
    public class RepositoryBase: IRepository
    {
        private static DataTable _dataTable;
        
        public RepositoryBase(DataTable dataTable)
        {
            _dataTable = dataTable;
        }

        public void Insert(DataRow row)
        {
            _dataTable.Rows.Add(row);
        }

        public void Delete(DataRow row)
        {
            _dataTable.Rows.Remove(row);
        }

        public IEnumerable<DataRow> SearchFor(Func<DataRow, bool> predicate)
        {
            return _dataTable.AsEnumerable().Where(predicate);
        }

        public Object SearchFor(Func<DataRow, bool> predicate, Func<DataRow, object> selector)
        {
            return _dataTable.AsEnumerable().Where(predicate).Select(selector).FirstOrDefault();
        }

        public IEnumerable<DataRow> GetAll()
        {
            return _dataTable.AsEnumerable();
        }

        public DataRow GetById(Guid id)
        {
            return _dataTable.AsEnumerable().FirstOrDefault(p => p.Field<Guid>("Id") == id);
        }
    }
}
