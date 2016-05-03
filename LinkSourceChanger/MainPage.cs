using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;


namespace LinkSourceChanger
{
    public partial class MainPage : Form
    {
        private static ExcelObj excelObj;
        private List<XlWorkBook> xlWorkBooks;
        //private List<string> targetExcelDocumentPaths;
        private string selectedPath;
        private int numberOfIndividualLinks = 0;
        private int totalNumberOfLinks = 0;
        private XlWorkBook activeWorkBook;
        private int numberOfTargetetLinks = 0;
        private string[] pathReplaceItems;

        public MainPage()
        {
            InitializeComponent();
            excelObj = new ExcelObj(textBoxLog);
            pathReplaceItems = new string[2] { "", "" };
            textBoxLog.Text = "";
            labelProcessStatus.Text = "";
            labelProgLabel.Text = "";
            progBar.Visible = false;
            buttonGetLinks.Visible = false;
            SizeColumns(listLinks);
        }

        private void findExcelFiles()
        {
            xlWorkBooks = null;
            xlWorkBooks = new List<XlWorkBook>();

            numberOfIndividualLinks = 0;
            numberOfTargetetLinks = 0;
            totalNumberOfLinks = 0;
            progBar.Visible = true;

            textBoxLog.Text += "Searching for excel files, please wait..." + "\r\n";
            try
            {
                var files = Directory.EnumerateFiles(selectedPath, "*", SearchOption.AllDirectories)
                    .Where(s => s.EndsWith(".xlsx") || s.EndsWith(".xls") || s.EndsWith(".xlsm"))
                    .Where(n => !n.Contains("~$"));
                 
                
                foreach (object file in files)
                {
                    FileInfo fi = new FileInfo(file.ToString());
                    if (fi.Length > 0)
                    {
                        XlWorkBook workBook = new XlWorkBook();
                        workBook.XlFullPath = file.ToString();
                        xlWorkBooks.Add(workBook);
                    }
                    else
                    {
                        textBoxLog.Text += "File size is zero, will not process: " + file.ToString()+"\r\n";
                    }
                }
                progBar.Maximum = xlWorkBooks.Count;
                textBoxLog.Text += "Found " + xlWorkBooks.Count.ToString() + " excel documents.\r\n";
                textBoxLog.Text += "Obtaining link sources, please wait...\r\n";
                labelProgLabel.Text = "Obtaining link sources, please wait...";

                foreach (XlWorkBook workBook in xlWorkBooks)
                {
                    progBar.Value++;
                    labelProcessStatus.Text = "(" + progBar.Value + " / " + progBar.Maximum + ")";
                    excelObj.openWorkBook(workBook);
                    listIndividualLinks(workBook.XlLinkPaths);

                    //if (workBook.XlLinkPaths.Count == 0)
                    //{
                    //    workBook.XlLinkPaths = null;
                    //    // xlWorkBooks.Remove(workBook);
                    //}
                    //else
                    //{
                    //    foreach (string[] link in workBook.XlLinkPaths)
                    //    {
                    //        totalNumberOfLinks++;
                    //    }
                    //}
                }

                xlWorkBooks = xlWorkBooks.FindAll(
                    delegate (XlWorkBook wB)
                    {
                        if (wB.XlLinkPaths != null && wB.XlLinkPaths.Count > 0)
                        {
                            return true;
                        }
                        else
                        {
                            return false;
                        }
                    }
                    );

                foreach (XlWorkBook workBook in xlWorkBooks)
                {

                    totalNumberOfLinks += workBook.XlLinkPaths.Count;
                }

                textBoxLog.Text += xlWorkBooks.Count + " Excel documents contain link sources.\r\n";
                textBoxLog.Text += "Number of individual links: " + numberOfIndividualLinks.ToString() + " out of total: " + totalNumberOfLinks.ToString() + " links\r\n";
                labelProgLabel.Text = "Completed.";
                labelProcessStatus.Text = "";
                progBar.Value = 0;
                progBar.Visible = false;


            }
            catch (Exception ex)
            {
                textBoxLog.Text += ex.Message + " in: " + ex.StackTrace +  "\r\n";
                textBoxLog.Select();
                textBoxLog.Select(textBoxLog.Text.Length, 0);
                textBoxLog.ScrollToCaret();
            }
        }

        private void validateLinkChange()
        {
            progBar.Visible = true;
            progBar.Value = 0;
            numberOfTargetetLinks = 0;
            labelProgLabel.Text = "Validation link source changes, please wait...";

            foreach (XlWorkBook workBook in xlWorkBooks)
            {
                foreach (string[] linkPath in workBook.XlLinkPaths)
                {
                    if (!String.IsNullOrWhiteSpace(linkPath[3]))
                    {

                        numberOfTargetetLinks++;
                    }
                }
            }


            progBar.Maximum = numberOfTargetetLinks;
            //}
            foreach (XlWorkBook workBook in xlWorkBooks)
            {
                foreach (string[] linkPath in workBook.XlLinkPaths)
                {
                    if (!String.IsNullOrWhiteSpace(linkPath[3]))
                    {
                        progBar.Value++;
                        labelProcessStatus.Text = "(" + progBar.Value + " / " + progBar.Maximum + ")";
                    }
                }
             
               
                excelObj.openWorkBook(workBook);
                listIndividualLinks(workBook.XlLinkPaths);
            }

                   textBoxLog.Text += "Validation complete.";
            labelProgLabel.Text = "Completed validation.";
            labelProcessStatus.Text = "";
            progBar.Value = 0;
            progBar.Visible = false;
        }

        private void listIndividualLinks(List<string[]> linkList)
        {
            if (linkList != null)
            {
                foreach (string[] path in linkList)
                {
                    if (path[0] != "")
                    {
                        char backSlash = '\\';
                        char forwardSlash = '/';
                        char[] slashArray = new char[] { backSlash, forwardSlash };
                        int countSlash = path[0].Count(x => x == forwardSlash || x == backSlash);
                        if (countSlash > 0)
                        {
                            string[] checkPath = new string[] { path[0], path[1], path[2], path[3] };
                            checkPath[0] = checkPath[0].Substring(0, checkPath[0].LastIndexOfAny(slashArray));
                            ListViewItem newListItem = new ListViewItem(checkPath[0]);
                            if (!listLinks.Items.ContainsKey(checkPath[0]))
                            {
                                newListItem.Name = checkPath[0];
                                newListItem.SubItems.Add(checkPath[1]);
                                newListItem.SubItems.Add(checkPath[2]);
                                newListItem.SubItems.Add(checkPath[3]);
                                listLinks.Items.Add(newListItem);
                                numberOfIndividualLinks++;
                            }
                        }
                    }
                }
            }
        }

        private void findWorkBooksWithTarget()
        {
            if (xlWorkBooks != null)
            {
                progBar.Value = 0;
                numberOfTargetetLinks = 0;
                progBar.Visible = true;
                labelProgLabel.Text += "Updating objects...";
                progBar.Maximum = xlWorkBooks.Count;

                foreach (XlWorkBook workBook in xlWorkBooks)
                {
                    progBar.Value++;
                    labelProcessStatus.Text = "( " + progBar.Value.ToString() + " / " + progBar.Maximum + " )";

                    foreach (string[] link in workBook.XlLinkPaths)
                    {
                        foreach (ListViewItem listItem in listLinks.Items)
                        {
                            if (listItem.SubItems[2].Text != "" && link[0].Contains(listItem.SubItems[2].Text))
                            {
                                link[1] = link[0].Replace(listItem.SubItems[2].Text, listItem.SubItems[3].Text);
                                link[2] = listItem.SubItems[2].Text;
                                link[3] = listItem.SubItems[3].Text;
                                workBook.XlSetToUpdateLinks = true;
                                break;
                            }
                        }
                        if (link[2] != null && link[0].Contains(link[2]))
                        {
                            numberOfTargetetLinks++;
                        }
                    }
                }
                progBar.Visible = false;
                labelProcessStatus.Text = "";
                textBoxLog.Text += "Target updates completed. ( " + numberOfTargetetLinks.ToString() + " / " + totalNumberOfLinks.ToString() + " links set to be replaced )\r\n";
                labelProgLabel.Text = "";
            }
            else
            {
                MessageBox.Show("Please set a target folder, and get the links...");
            }
        }

        private void SizeColumns(ListView listLinks)
        {
            listLinks.Columns[listLinks.Columns.Count - 1].Width = -2;
            listLinks.Columns[0].Width = listLinks.Width / 4;
            listLinks.Columns[1].Width = listLinks.Columns[0].Width;
            listLinks.Columns[2].Width = listLinks.Columns[0].Width;
            listLinks.Columns[3].Width = -2;
        }

        private void buttonSetNewDesitnation_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem listItem in listLinks.Items)
            {
                if (listItem.SubItems[0].Text.Contains(pathReplaceItems[0]) && pathReplaceItems[0] != "")
                {
                    listItem.SubItems[1].Text = listItem.SubItems[0].Text.Replace(pathReplaceItems[0], pathReplaceItems[1]);
                    listItem.SubItems[2].Text = (pathReplaceItems[0]);
                    listItem.SubItems[3].Text = (pathReplaceItems[1]);
                }
            }
            textBoxLog.Text += "Replacing: " + pathReplaceItems[0] + " ---> " + pathReplaceItems[1] + "\r\n";
            pathReplaceItems[0] = "";
            pathReplaceItems[1] = "";
            listLinks.Refresh();
            findWorkBooksWithTarget();

        }

        private void buttonReplaceAllLinkSources_Click(object sender, EventArgs e)
        {
            progBar.Visible = true;
            progBar.Value = 0;
            numberOfTargetetLinks = 0;
            listLinks.Items.Clear();
            labelProgLabel.Text = "Updating link sources...";
            textBoxLog.Text += "Updating link sources..., please wait.\r\n";
            foreach (XlWorkBook workBook in xlWorkBooks)
            {
                foreach (string[] linkPath in workBook.XlLinkPaths)
                {
                    if (linkPath[1] != "")
                    {
                        numberOfTargetetLinks++;

                    }
                }
            }

            progBar.Maximum = numberOfTargetetLinks;

            foreach (XlWorkBook workBook in xlWorkBooks)
            {
                progBar.Value++;
                if (workBook.XlSetToUpdateLinks)
                {
                    foreach (string[] linkPath in workBook.XlLinkPaths)
                    {
                        labelProcessStatus.Text = "( " + progBar.Value + " / " + progBar.Maximum + " )";
                    }
                    excelObj.updateLinkPath(workBook);
                }
            }
            progBar.Visible = false;
            progBar.Value = 0;
            validateLinkChange();
        }

        private void listLinks_Click(object sender, EventArgs e)
        {
            if (listLinks.SelectedItems.Count == 1)
            {
                textReplaceSource.Text = listLinks.SelectedItems[0].Name;
                textReplaceDestination.Text = listLinks.SelectedItems[0].Name;
            }
        }

        private void buttonTargetFolder_Click(object sender, EventArgs e)
        {
            folderBrowser.ShowDialog();
            labelTargetFolder.Text = "Target Folder: \r\n" + folderBrowser.SelectedPath;
            if (folderBrowser.SelectedPath != "")
            {
                buttonGetLinks.Visible = true;
                buttonGetLinks.Focus();
            }
            else
            {
                buttonGetLinks.Visible = false;
            }
        }

        private void buttonGetLinks_Click(object sender, EventArgs e)
        {
            if (folderBrowser.SelectedPath != "")
            {
                textReplaceSource.Text = "";
                textReplaceDestination.Text = "";
                numberOfTargetetLinks = 0;
                numberOfIndividualLinks = 0;
                totalNumberOfLinks = 0;
                xlWorkBooks = null;
                //targetExcelDocumentPaths = null;
                listLinks.Items.Clear();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                selectedPath = folderBrowser.SelectedPath;
                findExcelFiles();
            }
            else
            {
                buttonGetLinks.Visible = false;
                MessageBox.Show("Please select a target folder...");
            }
        }

        private void listLinks_Resize(object sender, EventArgs e)
        {
            SizeColumns((ListView)sender);
            listLinks.Refresh();
        }

        private void MainPage_FormClosing(object sender, FormClosingEventArgs e)
        {
            excelObj.closeExcel();
            GC.Collect();
        }

        private void textReplaceSource_TextChanged(object sender, EventArgs e)
        {
            pathReplaceItems[0] = textReplaceSource.Text;
            pathReplaceItems[1] = textReplaceDestination.Text;
        }

        private void textReplaceDestination_TextChanged(object sender, EventArgs e)
        {
            pathReplaceItems[0] = textReplaceSource.Text;
            pathReplaceItems[1] = textReplaceDestination.Text;
        }
    }
}


