using System;
using System.Collections.Generic;
using System.Data;

namespace SigningDocumentExample
{
    public interface IRepository
    {
        void Insert(DataRow row);
        void Delete(DataRow row);
        IEnumerable<DataRow> SearchFor(Func<DataRow, bool> predicate);
        Object SearchFor(Func<DataRow, bool> predicate, Func<DataRow, object> selector);
        IEnumerable<DataRow> GetAll();
        DataRow GetById(Guid id);
    }
}
