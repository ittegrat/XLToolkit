using System;
using System.Collections.Generic;
using System.Linq;

using ExcelDna.Integration;
using ExcelDna.Integration.Helpers;

namespace XLToolkit
{
  public static partial class SheetFunctions
  {

    private enum DistinctOpType { Unique, NotUnique, IsUnique };

    public static object Distinct(object[,] Range, string Sort, bool ErrorsFirst, object IfEmpty, bool NoCheckSize, bool Transpose, bool Ensure2d) {
      return Unique(DistinctOpType.Unique, Range, Sort, ErrorsFirst, IfEmpty, NoCheckSize, Transpose, Ensure2d);
    }
    public static object Duplicates(object[,] Range, string Sort, bool ErrorsFirst, object IfEmpty, bool NoCheckSize, bool Transpose, bool Ensure2d) {
      return Unique(DistinctOpType.NotUnique, Range, Sort, ErrorsFirst, IfEmpty, NoCheckSize, Transpose, Ensure2d);
    }
    public static object IsDistinct(object[,] Range) {
      return Unique(DistinctOpType.IsUnique, Range, String.Empty, false, null, true, false, false);
    }

    private static object Unique(DistinctOpType op, object[,] Range, string Sort, bool ErrorsFirst, object IfEmpty, bool NoCheckSize, bool Transpose, bool Ensure2d) {

      if (Range.Length == 1 && Range[0, 0] is ExcelMissing)
        return Strings.ERR("Invalid input range");

      if (IfEmpty is ExcelError xe && xe != ExcelError.ExcelErrorNA)
        return Strings.ERR("Invalid 'IfEmpty' value");

      HashSet<object> uniques = new HashSet<object>();
      HashSet<object> notUniques = new HashSet<object>();

      // Comparison is case sensitive. As there is no standard way to select
      // which value to output when two values are case insensitive equal, it
      // is better to handle the problem preprocessing input values.
      foreach (object obj in Range)
        if (!uniques.Add(obj)) {
          if (op == DistinctOpType.IsUnique)  // IsDistinct early exit at expenses of Duplicates
            return false;
          else
            notUniques.Add(obj);
        }

      if (op == DistinctOpType.IsUnique)
        return true;

      HashSet<object> values;
      if (op == DistinctOpType.NotUnique) {
        if (notUniques.Count == 0)
          return ExcelError.ExcelErrorNA;
        else
          values = notUniques;
      } else {
        values = uniques;
      }

      int nv = values.Count;

      if (!NoCheckSize) {
        var caller = new Caller();
        // Default layout is columnar, so Transpose really means horizontal.
        if ((caller.Rows * caller.Columns > 1) && caller.TooSmall(Transpose, nv, out string msg)) return msg;
      }

      IEnumerable<object> sorted;
      Sort = Sort.Trim();
      if (Sort.Length > 0) {

        string errPrefix = ErrorsFirst ? "A" : "Z";

        // Objects are compared lexicographically
        // Symbols come first
        string sortkey(object x) {
          if (x is ExcelError) {
            return errPrefix + x.ToString();
          } else if (x is ExcelEmpty) {
            return "B";
          } else if (x is bool) {
            return "C" + x.ToString();
          } else return "D" + x.ToString();
        }

        char scode = Char.ToUpperInvariant(Sort[0]);
        if ('A' == scode) {
          sorted = values.OrderBy(sortkey);
        } else if ('D' == scode) {
          sorted = values.OrderByDescending(sortkey);
        } else {
          return Strings.ERR($"Invalid sort code '{Sort}'");
        }

      } else {
        sorted = values;
      }

      if (Ensure2d && nv == 1) ++nv;

      object[,] ans = new object[
        (Transpose ? 1 : nv),
        (Transpose ? nv : 1)
      ];

      var empty = IfEmpty is ExcelMissing || IfEmpty is ExcelEmpty
        ? Strings.EXCEL_EMPTY
        : IfEmpty
      ;

      int i = -1;
      foreach (object obj in sorted) {
        ++i;
        ans[(Transpose ? 0 : i), (Transpose ? i : 0)] = (obj is ExcelEmpty ? empty : obj);
      }

      if (Ensure2d && i == 0) ans[(Transpose ? 0 : 1), (Transpose ? 1 : 0)] = ExcelError.ExcelErrorNA;

      return ans;

    }

  }
}
