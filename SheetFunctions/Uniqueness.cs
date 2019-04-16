using System;
using System.Collections.Generic;
using System.Linq;

using ExcelDna.Integration;
using ExcelDna.Integration.Helpers;

namespace XLToolkit
{
  public static partial class SheetFunctions
  {

    private enum UniquenessType { Unique, NotUnique };

    public static object Unique(object[,] Range, string Sort, bool ErrorsFirst, bool NoCheckSize, bool Transpose, bool Ensure2d) {
      return Unique(UniquenessType.Unique, Range, Sort, ErrorsFirst, NoCheckSize, Transpose, Ensure2d);
    }
    public static object NotUnique(object[,] Range, string Sort, bool ErrorsFirst, bool NoCheckSize, bool Transpose, bool Ensure2d) {
      return Unique(UniquenessType.NotUnique, Range, Sort, ErrorsFirst, NoCheckSize, Transpose, Ensure2d);
    }

    private static object Unique(UniquenessType uniq, object[,] Range, string Sort, bool ErrorsFirst, bool NoCheckSize, bool Transpose, bool Ensure2d) {

      if (Range.Length == 1 && Range[0, 0] is ExcelMissing)
        return "#ERR{Invalid input range}";

      HashSet<object> uniques = new HashSet<object>();
      HashSet<object> notUniques = new HashSet<object>();

      // Comparison is case sensitive. As there is no standard way to select
      // which value to output when two values are case insensitive equal, it
      // is better to handle the problem preprocessing input values.
      HashSet<object> values;
      if (uniq == UniquenessType.Unique) {
        foreach (object obj in Range)
          uniques.Add(obj);
        values = uniques;
      } else {
        foreach (object obj in Range)
          if (!uniques.Add(obj)) notUniques.Add(obj);
        if (notUniques.Count == 0)
          return ExcelError.ExcelErrorNA;
        values = notUniques;
      }

      int nv = values.Count;

      if (!NoCheckSize) {
        Caller caller = new Caller();
        // Default layout is columnar, so Transpose really means horizontal.
        if (caller.TooSmall(Transpose, nv, out string msg)) return msg;
      }

      IEnumerable<object> sorted;
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

        char scode = Char.ToUpperInvariant(Sort.Trim()[0]);
        if ('A' == scode) {
          sorted = values.OrderBy(sortkey);
        } else if ('D' == scode) {
          sorted = values.OrderByDescending(sortkey);
        } else {
          return $"#INVALID_SORT_CODE{{{Sort}}}";
        }

      } else {
        sorted = values;
      }

      if (Ensure2d && nv == 1) ++nv;

      object[,] ans = new object[
        (Transpose ? 1 : nv),
        (Transpose ? nv : 1)
      ];

      int i = -1;
      foreach (object obj in sorted) {
        ++i;
        ans[(Transpose ? 0 : i), (Transpose ? i : 0)] = (obj == ExcelEmpty.Value ? "<ExcelEmpty>" : obj);
      }

      if (Ensure2d && i == 0) ans[(Transpose ? 0 : 1), (Transpose ? 1 : 0)] = ExcelError.ExcelErrorNA;

      return ans;

    }

  }
}
