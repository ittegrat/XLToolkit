using System;
using System.Linq;
using System.Text.RegularExpressions;

using ExcelDna.Integration;

namespace XLToolkit
{

  public static partial class SheetFunctions
  {

    [ExcelFunction(Prefix = "RTD.")]
    public static object NotifyPath(string Path, object Trigger) {

      if (Trigger is ExcelError)
        return Trigger;

      Path = Path.Trim().ToUpperInvariant();
      if (String.IsNullOrEmpty(Path))
        return Strings.ERR("Invalid Path");

      if (ExcelDnaUtil.IsInFunctionWizard())
        return Strings.FUNC_WIZARD;

      var ans = RTD.TopicsMap.Notify(Path);

      return ans.ToString(Strings.DATETIME_FMT);

    }
    [ExcelFunction(Prefix = "RTD.")]
    public static object SubscribePath(string Path) {

      Path = Path.Trim().ToUpperInvariant();
      if (String.IsNullOrEmpty(Path))
        return Strings.ERR("Invalid Path");

      if (ExcelDnaUtil.IsInFunctionWizard())
        return Strings.FUNC_WIZARD;

      object ans = XlCall.RTD(RTD.TriggerServer.ProgId, null, Path);
      return ans;

    }
    [ExcelFunction(Prefix = "RTD.")]
    public static object PathList(string Pattern, object Trigger) {

      if (Trigger is ExcelError)
        return Trigger;

      Pattern = Pattern?.Trim().ToUpperInvariant();
      if (String.IsNullOrEmpty(Pattern))
        Pattern = ".+";

      var ans = RTD.TopicsMap.ListPaths()
        .Where(p => Regex.Match(p, Pattern, RegexOptions.IgnoreCase).Success)
        .ToList()
      ;

      ans.Sort();

      if (ans.Count > 0)
        return ans.ToArray();
      else
        return ExcelError.ExcelErrorNA;

    }

  }
}
