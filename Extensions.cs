using System;
using System.Collections.Generic;

using ExcelDna.Integration;
using ExcelDna.Integration.Helpers;

namespace XLToolkit
{
  public static class Extensions
  {

    public static string Address(this ExcelReference reference, int abs = 1, bool A1 = true) {

      var sheet = XlCall.Excel(XlCall.xlSheetNm, reference);

      var address = (string)XlCall.Excel(XlCall.xlfAddress,
        1 + reference.RowFirst,
        1 + reference.ColumnFirst,
        abs,  // abs_num: 1 abs, 2 abs-row, 3 abs-col, 4 rel
        A1,   // A1: true, false RC
        sheet
      );

      //?? Are 3D references supported ?
      if (reference.RowLast != reference.RowFirst || reference.ColumnLast != reference.ColumnFirst) {

        var lr = XlCall.Excel(XlCall.xlfAddress,
          1 + reference.RowLast,
          1 + reference.ColumnLast,
          abs,  // abs_num: 1 abs, 2 abs-row, 3 abs-col, 4 rel
          A1
        );

        address = $"{address}:{lr}";

      }

      return address;

    }

  }
}
