using System;
using System.Collections.Generic;

using ExcelDna.Integration;
using ExcelDna.Integration.Helpers;

namespace XLToolkit
{
  public static partial class SheetFunctions
  {

    public static object FilterRange(object[,] Range, object[,] Condition, bool NoCheckSize, bool Transpose, bool Ensure2d) {

      if (Range.Length == 1 && Range[0, 0] is ExcelMissing)
        return Strings.ERR("Invalid input range");

      if (Condition.Length == 1 && Condition[0, 0] is ExcelMissing)
        return Strings.ERR("Invalid condition range");

      int cRows = Condition.GetLength(0);
      int cCols = Condition.GetLength(1);

      if ((cRows > 1) && (cCols > 1))
        return Strings.ERR("Criteria must be a vector");

      // As it is one-dimensional, Condition determines direction
      bool isRowFilter = cRows > cCols;

      int rRows = Range.GetLength(0);
      int rCols = Range.GetLength(1);

      if (isRowFilter ? (rRows != cRows) : (rCols != cCols)) return Strings.ERR("Incompatible sizes");

      List<int> indices = new List<int>();

      int i = -1;
      foreach (object c in Condition) {
        ++i;

        if (!(c is ExcelError)) {
          try {
            if (Convert.ToBoolean(c)) indices.Add(i);
            continue;
          }
          catch { }
        } else if ((ExcelError)c == ExcelError.ExcelErrorNA)
          continue;

        return Strings.ERR($"Condition: {1 + i}");

      }

      int nv = indices.Count;

      if (nv == 0)
        return ExcelError.ExcelErrorNA;

      if (!NoCheckSize) {

        var caller = new Caller();
        string msg;

        if (caller.TooSmall(isRowFilter == Transpose, nv, out msg)) return msg;
        if (caller.TooSmall(!(isRowFilter == Transpose), isRowFilter ? rCols : rRows, out msg)) return msg;

      }

      if (Ensure2d && nv == 1) ++nv;

      object[,] ans = new object[
        (isRowFilter == Transpose) ? (isRowFilter ? rCols : rRows) : nv,
        (isRowFilter == Transpose) ? nv : (isRowFilter ? rCols : rRows)
      ];

      i = -1;
      foreach (int n in indices) {
        ++i;
        for (int j = 0; j < (isRowFilter ? rCols : rRows); ++j) {
          ans[
            (isRowFilter == Transpose) ? j : i,
            (isRowFilter == Transpose) ? i : j
          ] = Range[
            (isRowFilter ? n : j),
            (isRowFilter ? j : n)
          ];
        }
      }

      if (Ensure2d && i == 0) {
        for (int j = 0; j < (isRowFilter ? rCols : rRows); ++j) {
          ans[
            (isRowFilter == Transpose) ? j : 1,
            (isRowFilter == Transpose) ? 1 : j
          ] = ExcelError.ExcelErrorNA;
        }
      }

      return ans;

    }

  }
}
