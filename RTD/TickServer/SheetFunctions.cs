using ExcelDna.Integration;

namespace XLToolkit
{
  public static partial class SheetFunctions
  {

    [ExcelFunction(Prefix = "RTD.")]
    public static object TimeTick(double Seconds) {

      var ms = ((uint)(1000 * Seconds));

      if (ms <= 0)
        return Strings.ERR("Invalid period");

      if (ExcelDnaUtil.IsInFunctionWizard())
        return Strings.FUNC_WIZARD;

      var ans = XlCall.RTD(RTD.TickServer.ProgId, null, ms.ToString());
      return ans;

    }

  }
}
