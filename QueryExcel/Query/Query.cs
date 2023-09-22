using System;
using System.Data;
using System.Windows.Forms;


public class Query
{
    public readonly DataTable result;

    public Query()
    { 
    
    }

    /// <summary>
    /// 使用 SQL 语句过滤数据
    /// </summary>
    /// <param name="dt"></param>
    /// <param name="filter"></param>
    public DataTable Select(DataTable dt, string filter)
    {
        // 使用 SQL 语句过滤数据
        DataRow[] rows = dt.Select(filter);

        // 将过滤后的数据绑定到 DataGridView 控件中
        return rows.CopyToDataTable();
    }
}
