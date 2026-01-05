using Extensibility;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.OneNote;
using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace NoteQuickFormatter
{
    [Guid("358F3B33-5284-47D2-AD0B-7338B991CD03"), ProgId("NoteQuickFormatter.AddIn")]
    public class AddIn : IDTExtensibility2, IRibbonExtensibility
    {
        protected Application OneNoteApplication { get; set; }

        public IStream GetImage(string imageName)
        {
            MemoryStream mem = new MemoryStream();
            //Properties.Resources.test.Save(mem, ImageFormat.Png);
            return new CCOMStreamWrapper(mem);
        }

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            Application = OneNoteApplication;
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            OneNoteApplication = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public void OnAddInsUpdate(ref Array custom) { }

        public void OnStartupComplete(ref Array custom) { }

        public void OnBeginShutdown(ref Array custom)
        {
            if (OneNoteApplication != null)
            { OneNoteApplication = null; }
        }
        // IRibbonExtensibility唯一需要被實作的項目
        // 載入ribbon.xml
        public string GetCustomUI(string RibbonID)
        {
            return Properties.Resources.ribbon;
        }

        // 參考NoteHighlight2016寫法
        public void AddNextMonthSectionButtonClicked(IRibbonControl control)
        {
            Thread t = new Thread(new ThreadStart(() => System.Windows.Forms.Application.Run(new NextMonthToDoTemplateDialog())));
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
        }

        public void QuickAddNextMonthSectionButtonClicked(IRibbonControl control)
        {
            try
            {
                OneNoteService oneNoteService = new OneNoteService();
                oneNoteService.CreateNewSection(true);
            }
            catch (Exception ex)
            {
                /* TODO */
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }
        public void QuickAddCurrentMonthSectionButtonClicked(IRibbonControl control)
        {
            try
            {
                OneNoteService oneNoteService = new OneNoteService();
                oneNoteService.CreateNewSection(false);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }
    }
}
