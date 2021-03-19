using System.Collections.Generic;

using ExcelDna.Integration;
using ExcelDna.Integration.Helpers;

namespace XLToolkit
{
  public static partial class SheetFunctions
  {

    public static object MatchAll(object Value, object[,] Range, bool NoCheckSize, bool Transpose, bool Ensure2d) {

      if (Value is ExcelMissing)
        return Strings.ERR("Invalid value");

      if (Range.Length == 1 && Range[0, 0] is ExcelMissing)
        return Strings.ERR("Invalid lookup range");

      int rRows = Range.GetLength(0);
      int rCols = Range.GetLength(1);

      if (rRows > 1 && rCols > 1)
        return Strings.ERR("Lookup range must be a vector");

      // As it is one-dimensional, Range determines direction
      bool isColumn = rRows > rCols;

      List<int> indices = new List<int>();

      int i = 0; // Excel 1-based indices
      foreach (object obj in Range) {
        ++i;
        if (Value.Equals(obj)) indices.Add(i);
      }

      int nv = indices.Count;

      if (nv == 0)
        return ExcelError.ExcelErrorNA;

      if (Transpose) isColumn = !isColumn;

      if (!NoCheckSize) {
        var caller = new Caller();
        if ((caller.Rows * caller.Columns > 1) && caller.TooSmall(!isColumn, nv, out string msg)) return msg;
      }

      if (Ensure2d && nv == 1) ++nv;

      object[,] ans = new object[
        (isColumn ? nv : 1),
        (isColumn ? 1 : nv)
      ];

      i = -1;
      foreach (int idx in indices) {
        ++i;
        ans[(isColumn ? i : 0), (isColumn ? 0 : i)] = idx;
      }

      if (Ensure2d && i == 0) ans[(isColumn ? 1 : 0), (isColumn ? 0 : 1)] = ExcelError.ExcelErrorNA;

      return ans;

    }

  }
}
