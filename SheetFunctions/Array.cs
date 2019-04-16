using System.Collections.Generic;

using ExcelDna.Integration;

namespace XLToolkit
{
  public static partial class SheetFunctions
  {

    public static object Array(
      object Arg1, [ExcelArgument(Name = "...")] object Arg02,
      object Arg03, object Arg04, object Arg05, object Arg06, object Arg07, object Arg08, object Arg09, object Arg10, object Arg11,
      object Arg12, object Arg13, object Arg14, object Arg15, object Arg16, object Arg17, object Arg18, object Arg19, object Arg20,
      object Arg21, object Arg22, object Arg23, object Arg24, object Arg25, object Arg26, object Arg27, object Arg28, object Arg29
    ) {

      var ans = new List<object>(29);

      if (Arg1 is ExcelMissing) return ExcelError.ExcelErrorNA; else ans.Add(Arg1);

      if (Arg02 is ExcelMissing) return ans.ToArray(); else ans.Add(Arg02);
      if (Arg03 is ExcelMissing) return ans.ToArray(); else ans.Add(Arg03);
      if (Arg04 is ExcelMissing) return ans.ToArray(); else ans.Add(Arg04);
      if (Arg05 is ExcelMissing) return ans.ToArray(); else ans.Add(Arg05);
      if (Arg06 is ExcelMissing) return ans.ToArray(); else ans.Add(Arg06);
      if (Arg07 is ExcelMissing) return ans.ToArray(); else ans.Add(Arg07);
      if (Arg08 is ExcelMissing) return ans.ToArray(); else ans.Add(Arg08);
      if (Arg09 is ExcelMissing) return ans.ToArray(); else ans.Add(Arg09);

      if (Arg10 is ExcelMissing) return ans.ToArray(); else ans.Add(Arg10);
      if (Arg11 is ExcelMissing) return ans.ToArray(); else ans.Add(Arg11);
      if (Arg12 is ExcelMissing) return ans.ToArray(); else ans.Add(Arg12);
      if (Arg13 is ExcelMissing) return ans.ToArray(); else ans.Add(Arg13);
      if (Arg14 is ExcelMissing) return ans.ToArray(); else ans.Add(Arg14);
      if (Arg15 is ExcelMissing) return ans.ToArray(); else ans.Add(Arg15);
      if (Arg16 is ExcelMissing) return ans.ToArray(); else ans.Add(Arg16);
      if (Arg17 is ExcelMissing) return ans.ToArray(); else ans.Add(Arg17);
      if (Arg18 is ExcelMissing) return ans.ToArray(); else ans.Add(Arg18);
      if (Arg19 is ExcelMissing) return ans.ToArray(); else ans.Add(Arg19);

      if (Arg20 is ExcelMissing) return ans.ToArray(); else ans.Add(Arg20);
      if (Arg21 is ExcelMissing) return ans.ToArray(); else ans.Add(Arg21);
      if (Arg22 is ExcelMissing) return ans.ToArray(); else ans.Add(Arg22);
      if (Arg23 is ExcelMissing) return ans.ToArray(); else ans.Add(Arg23);
      if (Arg24 is ExcelMissing) return ans.ToArray(); else ans.Add(Arg24);
      if (Arg25 is ExcelMissing) return ans.ToArray(); else ans.Add(Arg25);
      if (Arg26 is ExcelMissing) return ans.ToArray(); else ans.Add(Arg26);
      if (Arg27 is ExcelMissing) return ans.ToArray(); else ans.Add(Arg27);
      if (Arg28 is ExcelMissing) return ans.ToArray(); else ans.Add(Arg28);
      if (Arg29 is ExcelMissing) return ans.ToArray(); else ans.Add(Arg29);

      return ans.ToArray();

    }

  }
}
