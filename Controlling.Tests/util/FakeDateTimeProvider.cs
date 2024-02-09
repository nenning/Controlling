using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Controlling.Tests.util
{
    public class FakeTimeProvider : ITimeProvider
    {
        public void PresetNow(DateTime now)
        {
            Now = now;
            TimeProvider.Instance = this;
        }
        public DateTime Now { get; private set; }
    }
}
