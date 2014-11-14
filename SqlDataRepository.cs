using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using System.Reflection;

namespace OkCupidAutoBot
{
    class SqlDataRepository
    {
        public void SaveItems<T>(List<T> girlsList, string tableName) where T:class
        {
            using (SqlConnection conn = new SqlConnection(OkSettings.Default.sqlConnection))
            using (SqlBulkCopy bulkCopy = new SqlBulkCopy(conn))
            {
                //get the type to iterate over their properties
                Type _type = typeof(T);

                DataTable dt = new DataTable();

                //add columns to datatable for each property
                foreach (PropertyInfo pInfo in _type.GetProperties())
                {
                    dt.Columns.Add(new DataColumn(pInfo.Name, pInfo.PropertyType));
                }

                foreach (T item in girlsList)
                {
                    DataRow newRow = dt.NewRow();

                    //copy property to data row
                    foreach (PropertyInfo pInfo in _type.GetProperties())
                    {
                        //form the row
                        newRow[pInfo.Name] = pInfo.GetValue(item, null);

                        ////get the property name, type, value
                        //Console.WriteLine(pInfo.Name);
                        //Console.WriteLine(pInfo.PropertyType);
                        //Console.WriteLine(pInfo.GetValue(person, null));
                    }

                    //add the row to the datatable
                    dt.Rows.Add(newRow);
                }

                //map the column from source to destination
                dt.Columns.Cast<DataColumn>()
                    .ToList()
                    .ForEach(c => bulkCopy.ColumnMappings.Add(c.ColumnName, c.ColumnName));

                bulkCopy.DestinationTableName = tableName.ToString();

                if (conn.State != System.Data.ConnectionState.Open)
                    conn.Open();

                try
                {
                    bulkCopy.WriteToServer(dt);
                }
                catch (SqlException exc)
                {
                    Console.WriteLine(exc.Message);
                }
                finally
                {
                    conn.Close();
                    conn.Dispose();
                }
            }
        }
    }
}
