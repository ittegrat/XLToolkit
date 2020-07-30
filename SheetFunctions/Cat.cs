using ExcelDna.Integration;
using ExcelDna.Integration.Helpers;

namespace XLToolkit
{
  public static partial class SheetFunctions
  {

    public static object HCat(object[,] First, object[,] Second, bool NoCheckSize, bool Transpose) {

      if (First.Length == 1 && First[0, 0] is ExcelMissing)
        return Strings.ERR("Invalid First range");

      if (Second.Length == 1 && Second[0, 0] is ExcelMissing)
        return Strings.ERR("Invalid Second range");

      int r = First.GetLength(0);

      if (r != Second.GetLength(0))
        return Strings.ERR("Incompatible row sizes");

      int c1 = First.GetLength(1);
      int c = c1 + Second.GetLength(1);

      if (!NoCheckSize) {

        var caller = new Caller();
        string msg;

        if (caller.TooSmall(Transpose, r, out msg)) return msg;
        if (caller.TooSmall(!Transpose, c, out msg)) return msg;

      }

      object[,] ans = new object[
        (Transpose ? c : r),
        (Transpose ? r : c)
      ];

      for (int i = 0; i < r; ++i) {
        for (int j1 = 0; j1 < c1; ++j1) {
          ans[(Transpose ? j1 : i), (Transpose ? i : j1)] = First[i, j1];
        }
        for (int j2 = c1; j2 < c; ++j2) {
          ans[(Transpose ? j2 : i), (Transpose ? i : j2)] = Second[i, j2 - c1];
        }
      }

      return ans;

    }
    public static object VCat(object[,] First, object[,] Second, bool NoCheckSize, bool Transpose) {

      if (First.Length == 1 && First[0, 0] is ExcelMissing)
        return Strings.ERR("Invalid First range");

      if (Second.Length == 1 && Second[0, 0] is ExcelMissing)
        return Strings.ERR("Invalid Second range");

      int c = First.GetLength(1);

      if (c != Second.GetLength(1))
        return Strings.ERR("Incompatible column sizes");

      int r1 = First.GetLength(0);
      int r = r1 + Second.GetLength(0);

      if (!NoCheckSize) {

        var caller = new Caller();
        string msg;

        if (caller.TooSmall(Transpose, r, out msg)) return msg;
        if (caller.TooSmall(!Transpose, c, out msg)) return msg;

      }

      object[,] ans = new object[
        (Transpose ? c : r),
        (Transpose ? r : c)
      ];

      for (int j = 0; j < c; ++j) {
        for (int i1 = 0; i1 < r1; ++i1) {
          ans[(Transpose ? j : i1), (Transpose ? i1 : j)] = First[i1, j];
        }
        for (int i2 = r1; i2 < r; ++i2) {
          ans[(Transpose ? j : i2), (Transpose ? i2 : j)] = Second[i2 - r1, j];
        }
      }

      return ans;

    }

  }
}
