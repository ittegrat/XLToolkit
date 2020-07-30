
namespace XLToolkit
{
  internal static class Strings
  {

    public const string DATETIME_FMT = "yyyy-MM-dd HH:mm:ss.fff";
    public const string EXCEL_EMPTY = "<ExcelEmpty>";
    public const string PROG_ID = "XLToolkit.RTD.";

    public const string FUNC_WIZARD = "#FUNC_WIZARD";
    public const string INVALID_TOPIC = "#INVALID_TOPIC";
    public const string WAIT = "#WAIT!";

    public static string ERR(string s) {
      return $"#ERR{{{s}}}";
    }

  }
}
