using System.IO;
using System.Runtime.Serialization.Formatters.Binary;

namespace XLToolkit.Store
{
  internal class SStorable : IStorable
  {

    public const string Prefix = "S_";

    static BinaryFormatter formatter = new BinaryFormatter();

    byte[] bytes;

    public int Rows { get; }
    public int Columns { get; }
    public bool IsShared => true;
    public string Source { get; }

    public object[,] Value {
      get {
        using (var ms = new MemoryStream(bytes)) {
          return (object[,])formatter.Deserialize(ms);
        }
      }
    }

    public SStorable(object[,] value, string source) {

      Source = source;
      Rows = value.GetLength(0);
      Columns = value.GetLength(1);

      using (var ms = new MemoryStream()) {
        formatter.Serialize(ms, value);
        bytes = ms.ToArray();
      }

    }

    public void Dispose() { }

  }
}
