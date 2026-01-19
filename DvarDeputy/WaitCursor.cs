using System.Windows.Input;

namespace Mooseware.DvarDeputy
{
    /// <summary>
    /// Wait cursor utility. Include a new instance of this class in a using{} block containing a long-running operation to show an hourglass cursor
    /// </summary>
    public class WaitCursor : IDisposable
    {
        private readonly System.Windows.Input.Cursor _previousCursor;

        public WaitCursor()
        {
            _previousCursor = Mouse.OverrideCursor;

            Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;
        }

        #region IDisposable Members

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Usage", "CA1816:Dispose methods should call SuppressFinalize")]
        public void Dispose()
        {
            Mouse.OverrideCursor = _previousCursor;
        }

        #endregion
    }
}
