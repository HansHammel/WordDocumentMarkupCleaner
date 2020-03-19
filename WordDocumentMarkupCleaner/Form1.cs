using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace WordDocumentMarkupCleaner
{

    public partial class Form1 : System.Windows.Forms.Form
    {
        private System.Collections.Specialized.StringCollection folderCol;

        private System.Windows.Forms.ImageList ilLarge;
        private System.Windows.Forms.ImageList ilSmall;
        private System.Windows.Forms.ListView listView1;
        private ListViewItemComparer listViewItemComparer;

        public Form1()
        {
            InitializeComponent();

            // for Transparency
            ilLarge.ColorDepth = ColorDepth.Depth32Bit;
            ilSmall.ColorDepth = ColorDepth.Depth32Bit;

            // Init ListView and folder collection
            // folderCol = new System.Collections.Specialized.StringCollection();
            CreateHeadersAndFillListView();

            //folderBrowserDialog1.ShowDialog();
            //string[] files = Directory.GetFiles(folderBrowserDialog1.SelectedPath);
            //PaintListView(folderBrowserDialog1.SelectedPath); // @"C:\");
            // folderCol.Add(@"C:\");

            this.listView1.ItemActivate += ListView1_ItemActivate;
            // this.Show();
            // this.BringToFront();
            //this.Activate();
            //this.Focus();

        }

        private void ListView1_ItemActivate(object sender, EventArgs e)
        {
            System.Windows.Forms.ListView lw = (System.Windows.Forms.ListView)sender;
            string filename = (lw.SelectedItems[0].Tag as FileInfo).FullName;

            if (lw.SelectedItems[0].ImageIndex != 0) //known Filetypes only
            {
                try
                {
                    System.Diagnostics.Process.Start(filename);
                }
                catch
                {
                    return;
                }
            }
        }

        private void CreateHeadersAndFillListView()
        {
            // Set the view to show details.
            listView1.View = View.Details;

            // Allow the user to edit item text.
            listView1.LabelEdit = true;

            // Allow the user to rearrange columns.
            listView1.AllowColumnReorder = true;

            // Select the item and subitems when selection is made.
            listView1.FullRowSelect = true;

            // Display grid lines.
            listView1.GridLines = true;

            // Sort the items in the list in ascending order.
            listView1.Sorting = SortOrder.Ascending;

            listView1.AllowDrop = true;

            listView1.ShowItemToolTips = true;

            ColumnHeader colHead;

            colHead = new ColumnHeader();
            colHead.Text = "Filename";
            this.listView1.Columns.Add(colHead);

            colHead = new ColumnHeader();
            colHead.Text = "Size";
            this.listView1.Columns.Add(colHead);

            colHead = new ColumnHeader();
            colHead.Text = "Last accessed";
            this.listView1.Columns.Add(colHead);

            // The ListViewItemSorter property allows you to specify the
            // object that performs the sorting of items in the ListView.
            // You can use the ListViewItemSorter property in combination
            // with the Sort method to perform custom sorting.
            listViewItemComparer = new ListViewItemComparer();
            listView1.ListViewItemSorter = listViewItemComparer;

        }

        private void listView1_DragDrop(object sender, DragEventArgs e)
        {
            List<string> files = ((string[])e.Data.GetData(DataFormats.FileDrop, false)).ToList();
            try
            {
                // Perform drag-and-drop, depending upon the effect.
                if (e.Effect != DragDropEffects.Copy)
                {
                    this.listView1.Items.Clear();
                }


                this.listView1.BeginUpdate();

                //remove duplicates bevore add
                foreach (ListViewItem lvi in this.listView1.Items)
                {
                    if (files.Contains(lvi.Text))
                        files.Remove(lvi.Text);
                }

                foreach (var filepath in files)
                {
                    var file = new FileInfo(filepath);
                    if (IsValidTarget(file))
                    {
                        // Set a default icon for the file.
                        Icon iconForFile = SystemIcons.WinLogo;

                        ListViewItem lvi = new ListViewItem();
                        //lvi.Text = file.Name;
                        lvi.Text = file.FullName;
                        lvi.ImageIndex = 1;
                        lvi.Tag = file;

                        ListViewItem.ListViewSubItem lvsi1 = new ListViewItem.ListViewSubItem();
                        lvsi1.Text = file.Length.ToString();
                        lvi.SubItems.Add(lvsi1);

                        ListViewItem.ListViewSubItem lvsi2 = new ListViewItem.ListViewSubItem();
                        lvsi2.Text = file.LastAccessTime.ToString();
                        lvi.SubItems.Add(lvsi2);

                        // Check to see if the image collection contains an image
                        // for this extension, using the extension as a key.
                        if (!ilSmall.Images.ContainsKey(file.Extension))
                        {
                            // If not, add the image to the image list.
                            iconForFile = System.Drawing.Icon.ExtractAssociatedIcon(file.FullName);
                            ilSmall.Images.Add(file.Extension, iconForFile);
                        }

                        lvi.ImageKey = file.Extension;
                        lvi.ToolTipText = "Unsaved";
                        this.listView1.Items.Add(lvi);
                    }
                }

                listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);

                this.listView1.EndUpdate();
            }
            catch (System.Exception err)
            {
                MessageBox.Show("Error: " + err.Message);
            }

            this.listView1.View = View.Details;
        }

        private bool IsValidTarget(FileInfo file)
        {
            if (!file.Attributes.HasFlag(FileAttributes.Directory) && !file.Attributes.HasFlag(FileAttributes.ReadOnly) && file.Exists)
            {
                if ((file.Extension == ".docx") || (file.Extension == ".docx"))
                {
                    return true;
                }
            }
            return false;
        }

        private void listView1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
            else
                e.Effect = DragDropEffects.None;
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void saveAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem lvi in this.listView1.Items)
            {
                CleanAndSaveItem(lvi);
            }
        }

        private void CleanAndSaveItem(ListViewItem lvi)
        {
            //listView1.EnsureVisible(lvi.Index);
            var file = lvi.Tag as FileInfo;
            if (lvi.ToolTipText == "Unsaved" //prevent double saving
                && IsValidTarget(file))
            {
                try
                {
                    using (WordprocessingDocument doc = WordprocessingDocument.Open(file.FullName, true))
                    {
                        SimplifyMarkupSettings settings = new SimplifyMarkupSettings
                        {
                            AcceptRevisions = true,
                            //setting this to false reduces translation quality, but if true some documents have XML format errors when opening
                            NormalizeXml = true,        // Merges Run's in a paragraph with similar formatting 
                            RemoveBookmarks = true,
                            RemoveComments = true,
                            RemoveContentControls = true,
                            RemoveEndAndFootNotes = true,
                            RemoveFieldCodes = false, //true,
                            RemoveGoBackBookmark = true,
                            RemoveHyperlinks = false,
                            RemoveLastRenderedPageBreak = true,
                            RemoveMarkupForDocumentComparison = true,
                            RemovePermissions = false,
                            RemoveProof = true,
                            RemoveRsidInfo = true,
                            RemoveSmartTags = true,
                            RemoveSoftHyphens = true,
                            RemoveWebHidden = true,
                            ReplaceTabsWithSpaces = false
                        };
                        MarkupSimplifier.SimplifyMarkup(doc, settings);
                        // OpenXmlPowerTools.WmlComparer.Compare
                        doc.Save();
                        lvi.BackColor = Color.Green;
                        lvi.ToolTipText = "Saved";
                    }
                }
                catch (Exception ex)
                {
                    lvi.BackColor = Color.Red;
                    lvi.ToolTipText = ex.Message;
                    //Console.WriteLine("Error in File: " + file.FullName + ". " + ex.Message);
                }
            }
        }

        private void listView1_ColumnClick(object sender, ColumnClickEventArgs e)
        {

            // Determine if clicked column is already the column that is being sorted.
            if (e.Column == listViewItemComparer.SortColumn)
            {
                // Reverse the current sort direction for this column.
                if (listViewItemComparer.Order == SortOrder.Ascending)
                {
                    listViewItemComparer.Order = SortOrder.Descending;
                }
                else
                {
                    listViewItemComparer.Order = SortOrder.Ascending;
                }
            }
            else
            {
                // Set the column number that is to be sorted; default to ascending.
                listViewItemComparer.SortColumn = e.Column;
                listViewItemComparer.Order = SortOrder.Ascending;
            }

            // Perform the sort with these new sort options.
            this.listView1.Sort();
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Added to prevent errors when nothing was selected
            if (listView1.SelectedItems.Count > 0)
            {
                CleanAndSaveItem(listView1.SelectedItems[0]);
            }
        }

        private void clearToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.listView1.Items.Clear();
        }

        private void saveAllBackupToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem lvi in this.listView1.Items)
            {
                if (CreateBackup(lvi))
                    CleanAndSaveItem(lvi);
            }
        }

        private static bool CreateBackup(ListViewItem lvi)
        {
            var filename = (lvi.Tag as FileInfo).FullName;
            var backupname = filename + ".bak";
            var backupfile = new FileInfo(backupname);
            if (!backupfile.Exists)
            {
                try
                {
                    File.Copy(filename, backupname);
                    return true;
                } catch (Exception ex)
                {
                    lvi.BackColor = Color.Red;
                    lvi.ToolTipText = ex.Message;
                    return false;
                }
            } else
            {
                lvi.BackColor = Color.Yellow;
                lvi.ToolTipText = "Backup present. File will be ingored.";
                return false;
            }
        }

        private void saveBackupToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var lvi = listView1.SelectedItems[0];
            // Added to prevent errors when nothing was selected
            if (listView1.SelectedItems.Count > 0)
            {
                if (CreateBackup(lvi))
                    CleanAndSaveItem(lvi);
            }
        }
    }
}


