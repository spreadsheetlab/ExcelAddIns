using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Windows.Forms;
using NetOffice;
using NetOffice.Tools;
using Excel = NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;
using NetOffice.ExcelApi.Tools;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using NetOffice.OfficeApi.Tools;
using VBIDE = NetOffice.VBIDEApi;
using NetOffice.VBIDEApi.Enums;

namespace Sagitta
{
	[COMAddin("Sagitta", "Sagitta", 3), ProgId("Sagitta.Addin"), Guid("9DE5DAC8-BFE0-4F76-AE17-5AE7F3C675D7")]
	[RegistryLocation(RegistrySaveLocation.CurrentUser), CustomUI("Sagitta.RibbonUI.xml")]
	public class Addin : Excel.Tools.COMAddin
	{
		public Addin()
		{
			this.OnStartupComplete += new OnStartupCompleteEventHandler(Addin_OnStartupComplete);
			this.OnDisconnection += new OnDisconnectionEventHandler(Addin_OnDisconnection);
		}

		internal Office.IRibbonUI RibbonUI { get; private set; }

		private void Addin_OnStartupComplete(ref Array custom)
		{
			Console.WriteLine("Addin started in Excel Version {0}", Application.Version);
		}

		private void Addin_OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
		{
		}

        public void AboutButton_Click(Office.IRibbonControl control)
        {
			MessageBox.Show(String.Format("Sagitta Version {0}", this.GetType().Assembly.GetName().Version),
				"About Sagitta", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

		public void OnLoadRibonUI(Office.IRibbonUI ribbonUI)
        {
			RibbonUI = ribbonUI;
        }

		protected override void OnError(ErrorMethodKind methodKind, System.Exception exception)
		{
			MessageBox.Show("An error occurend in " + methodKind.ToString(), "Sagitta");
		}

		[RegisterErrorHandler]
		public static void RegisterErrorHandler(RegisterErrorMethodKind methodKind, System.Exception exception)
		{
			MessageBox.Show("An error occurend in " + methodKind.ToString(), "Sagitta");
		}
    }
}

