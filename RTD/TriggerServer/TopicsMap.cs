using System;
using System.Collections.Generic;
using System.Linq;

using ExcelDna.Integration.Rtd;

namespace XLToolkit.RTD
{
  internal static class TopicsMap
  {

    static readonly Path root = new Path();

    public static DateTime Notify(string path) {
      var now = DateTime.Now;
      root.Notify(path, now);
      return now;
    }
    public static void SetTopic(string path, ExcelRtdServer.Topic topic) {
      root.SetTopic(path, topic);
    }
    public static void RemoveTopic(string path) {
      root.RemoveTopic(path);
    }
    public static List<string> ListPaths() {
      return root.ListPaths();
    }

    class Path
    {

      readonly Path parent;
      readonly Dictionary<string, Path> children = new Dictionary<string, Path>();

      ExcelRtdServer.Topic topic;

      public Path() { }

      public void Notify(string path, object value) {
        GetPath(path).Notify(value);
      }
      public void SetTopic(string path, ExcelRtdServer.Topic topic) {
        GetPath(path).topic = topic;
      }
      public void RemoveTopic(string path) {
        GetPath(path).topic = null;
      }
      public List<string> ListPaths() {
        var ans = new List<string>(children.Count);
        foreach (var kv in children) {
          ans.Add(kv.Key);
          if (kv.Value.children.Count > 0)
            ans.AddRange(kv.Value.ListPaths().Select(p => $"{kv.Key}.{p}"));
        }
        return ans;
      }

      Path(Path parent) : this() {
        this.parent = parent;
      }

      Path GetPath(string path) {

        if (path == String.Empty)
          return this;

        (var first, var rest) = Reduce(path);

        if (!children.TryGetValue(first, out var child)) {
          child = new Path(this);
          children[first] = child;
        }

        return child.GetPath(rest);

      }
      void Notify(object value) {

        if (parent == null)
          return;

        if (topic != null)
          topic.UpdateValue(value);

        parent.Notify(value);

      }
      (string first, string rest) Reduce(string path) {
        var i = path.IndexOf('.');
        if (i < 0)
          return (path, String.Empty);
        else
          return (path.Substring(0, i), path.Substring(i + 1));
      }

    }

  }
}
