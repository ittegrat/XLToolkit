using System;
using System.Reflection;

using ExcelDna.Integration;

namespace XLToolkit
{
  public static partial class SheetFunctions
  {

    [ExcelFunction(Name ="XLToolkit.Version")]
    public static object XLToolkitVersion(object Trigger) {

      if (Trigger is ExcelError)
        return Trigger;

      if (
        Attribute.GetCustomAttribute(typeof(SheetFunctions).Assembly, typeof(AssemblyInformationalVersionAttribute))
        is AssemblyInformationalVersionAttribute va
      ) return va.InformationalVersion;

      return typeof(SheetFunctions).Assembly.GetName().Version.ToString();

    }

  }
}
