
namespace XLToolkit.Store
{
  internal class LStorable : IStorable
  {

    public const string Prefix = "L_";

    public int Rows { get; }
    public int Columns { get; }
    public bool IsShared => false;
    public string Source { get; }
    public object[,] Value { get; }

    public LStorable(object[,] value, string source) {
      Value = value;
      Source = source;
      Rows = value.GetLength(0);
      Columns = value.GetLength(1);
    }

    public void Dispose() { }

  }
}
