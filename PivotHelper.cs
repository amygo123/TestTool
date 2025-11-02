#nullable enable
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace StyleWatcherWin
{
    public static class PivotHelper
    {
        // 将明细行转换为 颜色 x 尺码 的矩阵 DataTable，单元格为 "在库/可用" 字符串
        public static DataTable ToMatrix(IEnumerable<InventoryRow> rows, IEnumerable<string>? sizeOrder = null)
        {
            var list = rows.ToList();
            var colors = list.Select(r => r.Color).Distinct().OrderByDescending(c => list.Where(x=>x.Color==c).Sum(x=>x.QtyOut)).ToList();
            var sizes = (sizeOrder?.ToList() ?? list.Select(r => r.Size).Distinct().ToList());
            sizes = SortSizes(sizes);

            var dt = new DataTable();
            dt.Columns.Add("颜色");
            foreach (var s in sizes) dt.Columns.Add(s);

            foreach (var color in colors)
            {
                var row = dt.NewRow();
                row[0] = color;
                var byColor = list.Where(x => x.Color == color).ToList();
                var dict = byColor.ToDictionary(k => k.Size, v => $"{v.QtyIn} / {v.QtyOut}");
                for (int i=0;i<sizes.Count;i++)
                {
                    var key = sizes[i];
                    row[i+1] = dict.TryGetValue(key, out var txt) ? txt : string.Empty;
                }
                dt.Rows.Add(row);
            }
            return dt;
        }

        public static List<string> SortSizes(IEnumerable<string> sizes)
        {
            string[] order = new[]{"XS","S","M","L","XL","2XL","3XL","4XL","5XL","6XL"};
            int Rank(string s)
            {
                for (int i=0;i<order.Length;i++) if (string.Equals(s, order[i], StringComparison.OrdinalIgnoreCase)) return i;
                return 100 + s.Length;
            }
            return sizes.Distinct().OrderBy(Rank).ThenBy(s=>s).ToList();
        }
    }
}
