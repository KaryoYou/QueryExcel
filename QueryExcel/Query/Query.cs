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

    /// <summary>
    /// 将DataGridView的数据转换为DataTable。
    /// </summary>
    /// <param name="dgv">要转换的DataGridView。</param>
    /// <returns>包含DataGridView数据的DataTable。</returns>
    public DataTable GetDgvToTable(DataGridView dgv)
    {
        // 创建一个新的DataTable
        DataTable dt = new();

        // 遍历DataGridView的所有列
        for (int i = 0; i < dgv.Columns.Count; i++)
        {
            // 创建一个新的DataColumn并设置其列名
            DataColumn dc = new DataColumn(dgv.Columns[i].Name);
            // 将DataColumn添加到DataTable的列集合中
            dt.Columns.Add(dc);
        }

        // 遍历DataGridView的所有行
        for (int i = 0; i < dgv.Rows.Count; i++)
        {
            // 创建一个新的DataRow
            DataRow dr = dt.NewRow();
            // 遍历当前行的所有单元格
            for (int j = 0; j < dgv.Columns.Count; j++)
            {
                // 将单元格的值转换为字符串并赋值给DataRow
                dr[j] = Convert.ToString(dgv.Rows[i].Cells[j].Value);
            }
            // 将DataRow添加到DataTable的行集合中
            dt.Rows.Add(dr);
        }

        // 返回包含DataGridView数据的DataTable
        return dt;
    }
}
