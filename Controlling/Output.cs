using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Controlling
{
    public interface IOutput
    {
        void Write(string value, bool isError = false);
        void WriteLine(string value, bool isError = false);
        void WriteLine();
    }

    public class Output : IOutput
    {
        public void Write(string value, bool isError = false)
        {
            if (isError) UseErrorColors();
            Console.Write(value);
            if (isError) UseStandardColors();
        }
        public void WriteLine(string value, bool isError = false)
        {
            if (isError) UseErrorColors();
            Console.Write(value);
            if (isError) UseStandardColors();
            Console.WriteLine();
        }
        public void WriteLine()
        {
            Console.WriteLine();
        }

        private static void UseErrorColors()
        {
            Console.BackgroundColor = ConsoleColor.DarkRed;
            Console.ForegroundColor = ConsoleColor.White;
        }
        private static void UseStandardColors()
        {
            Console.ResetColor();
        }
    }
}
