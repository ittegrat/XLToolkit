using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using ExcelDna.Integration;
using ExcelDna.Integration.Helpers;

namespace XLToolkit
{

  using Store;

  public static partial class SheetFunctions
  {

    const string STORE = "Store.";

    static Dictionary<string, IStorable> store = new Dictionary<string, IStorable>();

    [ExcelFunction(Prefix = STORE)]
    public static object SetValue(string Name, object[,] Range, bool Shared, object Trigger) {

      if (Trigger is ExcelError)
        return Trigger;

      Name = Name?.Trim().ToUpperInvariant();
      if (String.IsNullOrEmpty(Name))
        return Strings.ERR("Invalid Name");

      if (Range.Length == 1 && Range[0, 0] is ExcelMissing)
        return Strings.ERR("Invalid input range");

      if (ExcelDnaUtil.IsInFunctionWizard())
        return Strings.FUNC_WIZARD;

      for (int i = 0; i < Range.GetLength(0); ++i) {
        for (int j = 0; j < Range.GetLength(1); ++j) {
          if (Range[i, j] is ExcelError xe && xe != ExcelError.ExcelErrorNA) {
            return Strings.ERR($"Invalid value at ({i + 1},{j + 1})");
          } else if (Range[i, j] is ExcelEmpty) {
            Range[i, j] = null;
          }
        }
      }

      var caller = new Caller();
      var source = caller.Range?.Address() ?? "Unknown";

      if (Shared)
        store[SStorable.Prefix + Name] = new SStorable(Range, source);
      else
        store[LStorable.Prefix + Name] = new LStorable(Range, source);

      return DateTime.Now.ToString(Strings.DATETIME_FMT);

    }
    [ExcelFunction(Prefix = STORE)]
    public static object GetValue(
      string Name, int Row, int Column, bool Shared, object IfEmpty, bool NoCheckSize,
      bool Transpose, bool Ensure2d, object Trigger
    ) {

      if (Trigger is ExcelError)
        return Trigger;

      Name = Name?.Trim().ToUpperInvariant();
      if (String.IsNullOrEmpty(Name))
        return Strings.ERR("Invalid Name");

      if (Row < 0)
        return Strings.ERR("Invalid Row");

      if (Column < 0)
        return Strings.ERR("Invalid Column");

      if (IfEmpty is ExcelError xe && xe != ExcelError.ExcelErrorNA)
        return Strings.ERR("Invalid 'IfEmpty' value");

      var pfx = Shared ? SStorable.Prefix : LStorable.Prefix;
      if (!store.TryGetValue(pfx + Name, out var storable)) {
        pfx = Shared ? LStorable.Prefix : SStorable.Prefix;
        if (!store.TryGetValue(pfx + Name, out storable))
          return ExcelError.ExcelErrorNA;
      }

      if (Row > storable.Rows)
        return Strings.ERR("Row out of range");

      if (Column > storable.Columns)
        return Strings.ERR("Column out of range");

      var caller = new Caller();
      var checkSize = !NoCheckSize && (caller.Rows * caller.Columns > 1);
      var empty = IfEmpty is ExcelMissing || IfEmpty is ExcelEmpty
        ? (caller.IsRange ? String.Empty : null)
        : IfEmpty
      ;

      object[,] ans;

      if (Row > 0 && Column > 0) {

        return storable.Value[Row - 1, Column - 1] ?? empty;

      } else if (Row > 0) {

        var cc = storable.Columns;

        if (checkSize) {
          if (caller.TooSmall(!Transpose, cc, out string msg)) return msg;
        }

        if (Ensure2d && cc == 1) {

          var r = Transpose ? 1 : 0;
          var c = Transpose ? 0 : 1;

          ans = new object[r + 1, c + 1];
          ans[0, 0] = storable.Value[Row - 1, 0] ?? empty;
          ans[r, c] = ExcelError.ExcelErrorNA;

        } else {

          ans = new object[
            Transpose ? cc : 1,
            Transpose ? 1 : cc
          ];

          for (var j = 0; j < cc; ++j)
            ans[Transpose ? j : 0, Transpose ? 0 : j] = storable.Value[Row - 1, j] ?? empty;

        }

      } else if (Column > 0) {

        var rr = storable.Rows;

        if (checkSize) {
          if (caller.TooSmall(Transpose, rr, out string msg)) return msg;
        }

        if (Ensure2d && rr == 1) {

          var r = Transpose ? 0 : 1;
          var c = Transpose ? 1 : 0;

          ans = new object[r + 1, c + 1];
          ans[0, 0] = storable.Value[0, Column - 1] ?? empty;
          ans[r, c] = ExcelError.ExcelErrorNA;

        } else {

          ans = new object[
            Transpose ? 1 : rr,
            Transpose ? rr : 1
          ];

          for (var i = 0; i < rr; ++i)
            ans[Transpose ? 0 : i, Transpose ? i : 0] = storable.Value[i, Column - 1] ?? empty;

        }

      } else {

        var rr = storable.Rows;
        var cc = storable.Columns;

        if (checkSize) {
          if (caller.TooSmall(Transpose, rr, out string msg)) return msg;
          if (caller.TooSmall(!Transpose, cc, out msg)) return msg;
        }

        var r = rr + (Ensure2d && rr == 1 ? 1 : 0);
        var c = cc + (Ensure2d && cc == 1 ? 1 : 0);

        ans = new object[
          Transpose ? c : r,
          Transpose ? r : c
        ];

        int i = 0, j = 0;
        for (i = 0; i < rr; ++i)
          for (j = 0; j < cc; ++j)
            ans[Transpose ? j : i, Transpose ? i : j] = storable.Value[i, j] ?? empty;

        if (Ensure2d && j == 1) {
          for (var ii = 0; ii < r; ++ii)
            ans[(Transpose ? cc : ii), (Transpose ? ii : cc)] = ExcelError.ExcelErrorNA;
        }
        if (Ensure2d && i == 1) {
          for (var jj = 0; jj < c; ++jj)
            ans[(Transpose ? jj : rr), (Transpose ? rr : jj)] = ExcelError.ExcelErrorNA;
        }

      }

      return ans;

    }
    [ExcelFunction(Prefix = STORE)]
    public static object Info(string Name, bool Shared, object Trigger) {

      if (Trigger is ExcelError)
        return Trigger;

      Name = Name?.Trim().ToUpperInvariant();
      if (String.IsNullOrEmpty(Name))
        return Strings.ERR("Invalid Name");

      var pfx = Shared ? SStorable.Prefix : LStorable.Prefix;
      if (!store.TryGetValue(pfx + Name, out var storable)) {
        pfx = Shared ? LStorable.Prefix : SStorable.Prefix;
        if (!store.TryGetValue(pfx + Name, out storable))
          return ExcelError.ExcelErrorNA;
      }

      var ans = new object[4];
      ans[0] = storable.Rows;
      ans[1] = storable.Columns;
      ans[2] = storable.IsShared ? "Shared" : "Local";
      ans[3] = storable.Source;

      return ans;

    }
    [ExcelFunction(Prefix = STORE)]
    public static object Clear(string Pattern, short Type, object Trigger) {

      if (Trigger is ExcelError)
        return Trigger;

      Pattern = Pattern?.Trim().ToUpperInvariant();
      if (String.IsNullOrEmpty(Pattern))
        return Strings.ERR("Invalid pattern");

      if (ExcelDnaUtil.IsInFunctionWizard())
        return Strings.FUNC_WIZARD;

      try {

        var keys = List(Pattern, Type).ToArray();

        int cnt = 0;
        foreach (var k in keys) {
          store[k].Dispose();
          cnt += store.Remove(k) ? 1 : 0;
        }

        return cnt;

      }
      catch (Exception ex) {
        return Strings.ERR(ex.Message);
      }

    }
    [ExcelFunction(Prefix = STORE)]
    public static object ClearAll(short Type, object Trigger) {

      return Clear(".+", Type, Trigger);

    }
    [ExcelFunction(Prefix = STORE)]
    public static object List(string Pattern, short Type, object Trigger) {

      if (Trigger is ExcelError)
        return Trigger;

      Pattern = Pattern?.Trim().ToUpperInvariant();
      if (String.IsNullOrEmpty(Pattern))
        Pattern = ".+";

      try {

        var keys = List(Pattern, Type)
          .Select(s => s.Substring(2))
          .ToArray()
        ;

        if (keys.Length > 0)
          return keys;
        else
          return ExcelError.ExcelErrorNA;

      }
      catch (Exception ex) {
        return Strings.ERR(ex.Message);
      }

    }

    static IEnumerable<string> List(string pattern, short type) {

      pattern = pattern?.Trim().ToUpperInvariant();
      if (String.IsNullOrEmpty(pattern))
        pattern = ".+";

      switch (type) {
        case 0:
          pattern = $"^({LStorable.Prefix}|{SStorable.Prefix}){pattern}$";
          break;
        case 1:
          pattern = $"^{LStorable.Prefix}{pattern}$";
          break;
        case 2:
          pattern = $"^{SStorable.Prefix}{pattern}$";
          break;
        default:
          throw new ArgumentOutOfRangeException(null, "Invalid type code");
      }

      return store.Keys.Where(k => Regex.Match(k, pattern).Success);

    }

  }
}
