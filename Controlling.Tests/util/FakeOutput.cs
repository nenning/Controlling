using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Controlling.Tests.util
{

    public class FakeOutput : IOutput
    {
        private List<string> expectations = new List<string>();
        private List<string> actual = new List<string>();

        public void SetExpectation(string expected)
        {

        }

        public void Verify()
        {

        }
        public void Write(string value, bool isError = false)
        {
            actual.Add(value);
        }
        public void WriteLine(string value, bool isError = false)
        {
            Write(value, isError);
        }
        public void WriteLine()
        {
        }
    }
}
