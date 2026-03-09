using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Xlsx2PO
{
    public sealed class ScopedWorkingDirectories : IDisposable
    {
        private readonly string prevDirectory;
        private bool isDisposed;

        public ScopedWorkingDirectories(string newDirectory)
        {
            prevDirectory = Directory.GetCurrentDirectory();
            if (!String.IsNullOrEmpty(newDirectory))
            {
                Directory.SetCurrentDirectory(newDirectory);
            }
        }

        public void Dispose()
        {
            if (isDisposed)
            {
                return;
            }
            Directory.SetCurrentDirectory(prevDirectory);
            isDisposed = true;
        }
    }
}
