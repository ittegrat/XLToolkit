using System;

namespace XLToolkit.Store
{
  internal interface IStorable : IDisposable
  {

    int Rows { get; }
    int Columns { get; }
    bool IsShared { get; }
    string Source { get; }
    object[,] Value { get; }

  }
}
