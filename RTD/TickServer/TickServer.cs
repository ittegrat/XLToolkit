using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;

using ExcelDna.Integration.Rtd;

namespace XLToolkit.RTD
{
  [ComVisible(true)]
  [ProgId(ProgId)]
  internal class TickServer : ExcelRtdServer
  {

    public const string ProgId = Strings.PROG_ID + nameof(TickServer);

    class TickTopic : Topic, IDisposable
    {

      Timer timer;

      public TickTopic(ExcelRtdServer server, int topicId, string msecs)
        : base(server, topicId, Strings.WAIT) {
        var period = uint.Parse(msecs);
        timer = new Timer(Notify, null, 0, period);
      }

      public void Dispose() {
        timer.Dispose();
      }

      void Notify(object state) {
        UpdateValue(DateTime.Now.ToString(Strings.DATETIME_FMT));
      }

    }

    protected override Topic CreateTopic(int topicId, IList<string> topicInfo) {
      return new TickTopic(this, topicId, topicInfo[0]);
    }
    protected override object ConnectData(Topic topic, IList<string> topicInfo, ref bool newValues) {
      // On input, newValue = false => Excel has a cached value
      // On output, newValue = false => ignore ConnectData return and use #N/A or cached value
      newValues = true;
      return topic.Value; // return the value already set in RtdTopic constructor
    }
    protected override void DisconnectData(Topic topic) {
      ((TickTopic)topic).Dispose();
    }

  }
}
