using System;
using System.Text.RegularExpressions;

using ExcelDna.Integration;
using ExcelDna.Integration.Helpers;

namespace XLToolkit
{
  public static partial class SheetFunctions
  {

    public static object SplitString(string Text, string Delimiters, bool NoCheckSize, bool Transpose, bool Ensure2d) {

      string[] sa = Text.Split(Delimiters.ToCharArray());

      int nv = sa.Length;

      if (!NoCheckSize) {
        Caller caller = new Caller();
        // Default layout is horizontal.
        if (caller.TooSmall(!Transpose, nv, out string msg)) return msg;
      }

      if (Ensure2d && nv == 1) ++nv;

      object[,] ans = new object[
        (Transpose ? nv : 1),
        (Transpose ? 1 : nv)
      ];

      int i; for (i = 0; i < sa.Length; ++i) {
        ans[(Transpose ? i : 0), (Transpose ? 0 : i)] = sa[i];
      }

      if (Ensure2d && i == 1) ans[(Transpose ? 1 : 0), (Transpose ? 0 : 1)] = ExcelError.ExcelErrorNA;

      return ans;

    }

    // RE Quick Reference: https://msdn.microsoft.com/en-us/library/az24scfc.aspx
    public static object RegMatch(string Text, string Pattern, bool IgnoreCase) {

      try {

        Match m = Regex.Match(Text, Pattern, IgnoreCase ? RegexOptions.IgnoreCase : RegexOptions.None);
        if (m.Success)
          return m.Value;
        else
          return ExcelError.ExcelErrorNA;

      }
      catch (Exception e) {
        return $"#ERR{{{e.Message}}}";
      }

    }
    public static string RegReplace(string Text, string Pattern, string Replacement, bool IgnoreCase) {
      try {
        return Regex.Replace(Text, Pattern, Replacement, IgnoreCase ? RegexOptions.IgnoreCase : RegexOptions.None);
      }
      catch (Exception e) {
        return $"#ERR{{{e.Message}}}";
      }
    }

  }
}
