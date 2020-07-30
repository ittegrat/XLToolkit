using System.Collections.Generic;
using System.Runtime.InteropServices;

using ExcelDna.Integration.Rtd;

namespace XLToolkit.RTD
{
  [ComVisible(true)]
  [ProgId(ProgId)]
  internal partial class TriggerServer : ExcelRtdServer
  {

    public const string ProgId = Strings.PROG_ID + nameof(TriggerServer);

    class TriggerTopic : Topic
    {
      public readonly string Path;
      public TriggerTopic(ExcelRtdServer server, int topicId, string path)
        : base(server, topicId, Strings.WAIT) {
        Path = path;
      }
    }

    protected override Topic CreateTopic(int topicId, IList<string> topicInfo) {
      var path = topicInfo[0];
      var topic = new TriggerTopic(this, topicId, path);
      TopicsMap.SetTopic(path, topic);
      return topic;
    }
    protected override object ConnectData(Topic topic, IList<string> topicInfo, ref bool newValues) {
      newValues = true;
      return topic.Value;
    }
    protected override void DisconnectData(Topic topic) {
      if (topic is TriggerTopic tt)
        TopicsMap.RemoveTopic(tt.Path);
    }

  }
}
