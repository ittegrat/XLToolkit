using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelDna.Integration.Helpers
{
  public static class CallerExtensions
  {

    public static bool TooSmall(
      this Caller caller, bool checkCols, int desired, out string msg,
      Func<bool, int, int, string> CustomMsg = null
    ) {

      if (caller.Type == Caller.CallerType.Range) {

        int n = checkCols ? caller.Columns : caller.Rows;

        if (n < desired) {

          msg = CustomMsg == null
            ? (checkCols ? "#COLS" : "#ROWS") + "{" + desired + "}"
            : CustomMsg(checkCols, desired, n)
          ;

          return true;

        }

      }

      msg = String.Empty;
      return false;

    }

#if DEBUG
    public static string CheckLengthTest(int desired, bool cols, bool custom) {
      Caller caller = new Caller();
      Func<bool, int, int, string> cmsg = (c, d, n) => String.Format("I need {0} more {1}s", d - n, c ? "row" : "column");
      if (caller.TooSmall(cols, desired, out string msg, CustomMsg: custom ? cmsg : null)) return msg;
      return "OK";
    }
#endif

  }
}
