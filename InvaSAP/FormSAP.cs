using System.Diagnostics;
using System.Runtime.InteropServices;

namespace InvaSAP
{
    public partial class FormSAP : Form
    {

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        private IntPtr sapWindowHandle;

        public FormSAP(IntPtr windowHandle)
        {
            InitializeComponent();
            sapWindowHandle = windowHandle;

            Shown += FormSAP_Shown;
        }
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        private void FormSAP_Shown(object sender, EventArgs e)
        {
            // Set the parent of the SAP window to be the handle of the panel control
            //SetParent(sapWindowHandle, panelSAP.Handle);
            SetForegroundWindow(sapWindowHandle);

            Debug.WriteLine("FormSAP_Load() - Handle: " + sapWindowHandle);
        }
    }

}
