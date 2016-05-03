using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;


namespace LinkSourceChanger
{
    class ExcelObj
    {
        public ExcelObj(System.Windows.Forms.TextBox textBoxLog)
        {
            this.textBoxLog = textBoxLog;
            openExcel();
        }
        Application excel;
        Workbook activeWorkBook;
        System.Windows.Forms.TextBox textBoxLog;

        private XlWorkBook activeXlWorkBook;
        private string workBookPath;

        private void openExcel()
        {
            excel = new Application();
            excel.DisplayAlerts = false;
            excel.AskToUpdateLinks = false;
        }

        public void openWorkBook(XlWorkBook activeXlWorkBook)
        {
            if (excel != null)
            {
                this.activeXlWorkBook = activeXlWorkBook;
                activeWorkBook = excel.Workbooks.Open(activeXlWorkBook.XlFullPath, 0, true);
                // activeXlWorkBook = new XlWorkBook();
                activeXlWorkBook.XlFileName = activeWorkBook.Name;
                activeXlWorkBook.XlFilePath = activeWorkBook.Path;
                getLinks(activeWorkBook);
                if (excel.ActiveWorkbook != null)
                {
                    excel.ActiveWorkbook.Close(false);
                }
                //return activeXlWorkBook;
            }
            else
            {
                openExcel();
                openWorkBook(activeXlWorkBook);
            }
        }

        private void getLinks(Workbook activeWorkBook)
        {
            try
            {
                if (excel != null)
                {
                    object linkList = activeWorkBook.LinkSources();
                    Array convertedLinkList = (Array)linkList;
                    if (convertedLinkList != null && convertedLinkList.Length > 0)
                    {
                        activeXlWorkBook.XlLinkPaths = null;
                        activeXlWorkBook.XlLinkPaths = new List<string[]>();

                        textBoxLog.Text += activeWorkBook.FullName.ToString() + "\r\n";
                        foreach (object obj in convertedLinkList)
                        {

                            string[] linkTemp = new string[4];
                            linkTemp[0] = (string)obj;
                            activeXlWorkBook.XlLinkPaths.Add(linkTemp);
                            linkTemp = null;
                        }
                    }
                }
                else
                {
                    openExcel();
                    getLinks(activeWorkBook);
                }
            }
            catch (Exception ex)
            {
                textBoxLog.Text += ex.Message + ex.StackTrace + "\r\n";
                textBoxLog.Select();
                textBoxLog.Select(textBoxLog.Text.Length, 0);
                textBoxLog.ScrollToCaret();
            }
        }

        public void updateLinkPath(XlWorkBook workBook)
        {
            try
            {
                activeWorkBook = excel.Workbooks.Open(workBook.XlFullPath, false, false);
                if (!activeWorkBook.ReadOnly || !activeWorkBook.HasPassword)
                {
                    foreach (string[] targetLink in workBook.XlLinkPaths)
                    {
                        if (!String.IsNullOrWhiteSpace(targetLink[1]))
                        {
                            activeWorkBook.ChangeLink(targetLink[0], targetLink[1], XlLinkType.xlLinkTypeExcelLinks);
                    
                        }
                    }
                    activeWorkBook.Save();
                    activeWorkBook.Close();
                }
                else
                {
                    textBoxLog.Text += "Workbook: " + workBook.XlFullPath + " is read only.";
                }
            }
            catch (Exception ex)
            {
                textBoxLog.Text += ex.Message + ex.StackTrace + "\r\n";
            }
        }

        public void closeExcel()
        {
            excel.Quit();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}