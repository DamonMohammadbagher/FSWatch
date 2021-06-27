using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Drawing.Drawing2D;
using System.Reflection;
using System.Threading;
// using C1.C1Excel;
using System.Globalization;

namespace progress
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        #region Assembly Attribute Accessors

        public string AssemblyTitle
        {
            get
            {
                // Get all Title attributes on this assembly
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyTitleAttribute), false);
                // If there is at least one Title attribute
                if (attributes.Length > 0)
                {
                    // Select the first one
                    AssemblyTitleAttribute titleAttribute = (AssemblyTitleAttribute)attributes[0];
                    // If it is not an empty string, return it
                    if (titleAttribute.Title != "")
                        return titleAttribute.Title;
                }
                // If there was no Title attribute, or if the Title attribute was the empty string, return the .exe name
                return System.IO.Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().CodeBase);
            }
        }

        public string AssemblyVersion
        {
            get
            {
                return Assembly.GetExecutingAssembly().GetName().Version.ToString();
            }
        }

        public string AssemblyDescription
        {
            get
            {
                // Get all Description attributes on this assembly
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false);
                // If there aren't any Description attributes, return an empty string
                if (attributes.Length == 0)
                    return "";
                // If there is a Description attribute, return its value
                return ((AssemblyDescriptionAttribute)attributes[0]).Description;
            }
        }

        public string AssemblyProduct
        {
            get
            {
                // Get all Product attributes on this assembly
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyProductAttribute), false);
                // If there aren't any Product attributes, return an empty string
                if (attributes.Length == 0)
                    return "";
                // If there is a Product attribute, return its value
                return ((AssemblyProductAttribute)attributes[0]).Product;
            }
        }

        public string AssemblyCopyright
        {
            get
            {
                // Get all Copyright attributes on this assembly
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
                // If there aren't any Copyright attributes, return an empty string
                if (attributes.Length == 0)
                    return "";
                // If there is a Copyright attribute, return its value
                return ((AssemblyCopyrightAttribute)attributes[0]).Copyright;
            }
        }

        public string AssemblyCompany
        {
            get
            {
                // Get all Company attributes on this assembly
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCompanyAttribute), false);
                // If there aren't any Company attributes, return an empty string
                if (attributes.Length == 0)
                    return "";
                // If there is a Company attribute, return its value
                return ((AssemblyCompanyAttribute)attributes[0]).Company;
            }
        }
        #endregion
        ListBox myabouttext = new ListBox();
        Button ok = new Button();
        Form m = new Form();        
        TableLayoutPanel tableLayoutPanel = new System.Windows.Forms.TableLayoutPanel();
        PictureBox logoPictureBox = new System.Windows.Forms.PictureBox();
        Label ProductName = new System.Windows.Forms.Label();
        Label Version = new System.Windows.Forms.Label();
        Label labelCopyright = new System.Windows.Forms.Label();
        Label labelCompanyName = new System.Windows.Forms.Label();
        TextBox textBoxDescription = new System.Windows.Forms.TextBox();
        Button okButton = new System.Windows.Forms.Button();
        RichTextBox Des = new RichTextBox();
        private delegate void Change_Delegate(System.IO.FileSystemEventArgs e);
        private string drv = @"c:\";
        private string st_file, st_fol;
        private static ThreadStart threadDelegate1;
        private static Thread newThread1;
        private static Thread bingo = null;
        private double opacityIncrease = 0.05;

        //public C1.C1Excel.C1XLBook c1XLBook1;

        //private void AutoSizeColumns(XLSheet sheet , C1.C1Excel.C1XLBook xls)
        //{
        //    using (Graphics g = Graphics.FromHwnd(IntPtr.Zero))
        //    {
        //        for (int c = 0; c < sheet.Columns.Count; c++)
        //        {
        //            int colWidth = -1;
        //            for (int r = 0; r < sheet.Rows.Count; r++)
        //            {
        //                object value = sheet[r, c].Value;
        //                if (value != null)
        //                {
        //                    // get value (unformatted at this point)
        //                    string text = value.ToString();

        //                    // format value if cell has a style with format set
        //                    C1.C1Excel.XLStyle s = sheet[r, c].Style;
        //                    if (s != null && s.Format.Length > 0 && value is IFormattable)
        //                    {
        //                        string fmt = XLStyle.FormatXLToDotNet(s.Format);
        //                        text = ((IFormattable)value).ToString(fmt, CultureInfo.CurrentCulture);
        //                    }

        //                    // get font (default or style)
        //                    Font font = xls.DefaultFont;
        //                    if (s != null && s.Font != null)
        //                    {
        //                        font = s.Font;
        //                    }

        //                    // measure string (add a little tolerance)
        //                    Size sz = Size.Ceiling(g.MeasureString(text + "x", font));

        //                    // keep widest so far
        //                    if (sz.Width > colWidth)
        //                        colWidth = sz.Width;
        //                }
        //            }

        //            // done measuring, set column width
        //            if (colWidth > -1)
        //                sheet.Columns[c].Width = C1XLBook.PixelsToTwips(colWidth);
        //        }
        //    }
        //}
        public class __Log_Class
        {
            public static void Log(String logMessage, TextWriter w)
            {
                w.WriteLine("  {0}", logMessage);
                w.Flush();
            }
            public static void DumpLog(StreamReader r)
            {
                String line;
                while ((line = r.ReadLine()) != null)
                {
                    //  Console.WriteLine(line);
                }
                r.Close();
            }
        }

        private void ok_Click(object sender, EventArgs e)
        {
            m.Close();                        
        }
        private void tableLayoutPanel_Paint(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            Rectangle rect = this.ClientRectangle;
            rect.Inflate(0, 0);
            LinearGradientBrush filler = new LinearGradientBrush(rect, Color.FromArgb(236, 241, 253), Color.FromArgb(225, 240, 255), 180);
            g.FillRectangle(filler, rect);
        }
        private void m_Paint(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            Rectangle rect = this.ClientRectangle;
            rect.Inflate(0, 0);
            LinearGradientBrush filler = new LinearGradientBrush(rect, Color.FromArgb(236, 241, 253), Color.FromArgb(225,240,255), 180);
            g.FillRectangle(filler, rect);
        }
       
        private string return_Folderandfiles(string add)
        {
            st_fol = "";
            st_file = "";
            
            string [] currentItems_container_folder = Directory.GetDirectories(add);
		    string [] currentItems_container_files = Directory.GetFiles(add);
        
        int FOLDERS_COUNT = 0;
               
                for ( FOLDERS_COUNT = 0; FOLDERS_COUNT < currentItems_container_folder.Length; FOLDERS_COUNT++)
                {
                
                }


                int Files_COUNT = 0;

            for (Files_COUNT = 0; Files_COUNT < currentItems_container_files.Length; Files_COUNT++)
            {
              
            }
            int rtn = (Files_COUNT + FOLDERS_COUNT);
            return rtn.ToString();
        }

        /// <summary>
        ///  Class for log file (Realtime)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 
      
        private void _Changes(System.IO.FileSystemEventArgs e)
        {
//            System.IO.FileSystemEventArgs e = new FileSystemEventArgs(WatcherChangeTypes.Changed, toolStripComboBox1.Text, e.FullPath);
            string myevent = e.FullPath;
            DirectoryInfo g = new DirectoryInfo(myevent);
            FileInfo f = new FileInfo(myevent);
            string a2 = e.ChangeType.ToString();
            string myeventName = "Folder : " + e.FullPath + " is " + a2 + " : " + f.Name;
            string a3 = " --> ";
            string a4 = f.Name + a3 + a2;
            string a5 = myeventName;
           
            try
            {
                if (Directory.Exists(f.FullName))
                {

                    ListViewItem folder_items = new ListViewItem();
                    folder_items.SubItems.Add(DateTime.Now.ToString());
                    folder_items.SubItems.Add(e.ChangeType.ToString());
                    folder_items.SubItems.Add(e.FullPath);
                    folder_items.SubItems.Add(f.Name);
                    folder_items.SubItems.Add("Folder");
                    folder_items.SubItems.Add("-");
                    folder_items.SubItems.Add(return_Folderandfiles(e.FullPath));
                    listView1.Items.AddRange(new ListViewItem[] { folder_items });
                    listView1.CheckBoxes = false;
                    listBox5.Items.Add(listBox5.Items.Count + " <" + DateTime.Now.ToString() + ">" + "  <Event: " + e.ChangeType.ToString() + ">  <Name: " + f.Name + ">  <Path: " + e.FullPath);
                    listBox5.SelectedIndex = listBox5.Items.Count - 1;
                }
                else
                {
                    if (File.Exists(f.FullName))
                    {
                        if (f.Name.ToUpper() != "NTUSER.DAT.LOG")
                        {
                            FileInfo myinfo = new FileInfo(e.FullPath);
                            System.IO.FileInfo k = new FileInfo(e.FullPath);
                            ListViewItem folder_items = new ListViewItem();
                            folder_items.SubItems.Add(DateTime.Now.ToString());
                            folder_items.SubItems.Add(e.ChangeType.ToString());
                            folder_items.SubItems.Add(e.FullPath);
                            folder_items.SubItems.Add(f.Name);
                            folder_items.SubItems.Add("File");

                            if (f.Name.ToUpper() != "NTUSER.DAT.LOG")
                            {

                                if (e.Name.ToUpper().StartsWith("WINDOWS") || toolStripComboBox1.Text.ToUpper().Substring(3).StartsWith("WINDOWS")) { listBox3.Items.Add(listBox3.Items.Count + " <" + DateTime.Now.ToString() + ">" + "  <Event: " + e.ChangeType.ToString() + ">  <Name: " + f.Name + ">  <Path: " + e.FullPath + ">  <Size: " + (k.Length).ToString() + " Bytes>"); listBox3.SelectedIndex = listBox3.Items.Count - 1; }
                                if (e.Name.ToUpper().StartsWith("PROGRAM FILES") || toolStripComboBox1.Text.ToUpper().Substring(3).StartsWith("PROGRAM FILES")) { listBox1.Items.Add(listBox1.Items.Count + " <" + DateTime.Now.ToString() + ">" + "  <Event: " + e.ChangeType.ToString() + ">  <Name: " + f.Name + ">  <Path: " + e.FullPath + ">  <Size: " + (k.Length).ToString() + " Bytes>"); listBox1.SelectedIndex = listBox1.Items.Count - 1; }
                                if (e.Name.ToUpper().StartsWith("DOCUMENTS AND SETTINGS") || toolStripComboBox1.Text.ToUpper().Substring(3).StartsWith("DOCUMENTS AND SETTINGS")) { listBox2.Items.Add(listBox2.Items.Count + " <" + DateTime.Now.ToString() + ">" + "  <Event: " + e.ChangeType.ToString() + ">  <Name: " + f.Name + ">  <Path: " + e.FullPath + ">  <Size: " + (k.Length).ToString() + " Bytes>"); listBox2.SelectedIndex = listBox2.Items.Count - 1; }
                                toolStripTextBox1.Text = f.Name + "   " + e.ChangeType.ToString();
                            }

                            if (k.Length >= 1024)
                            {
                                long kk = k.Length / 1024;
                                folder_items.SubItems.Add(kk.ToString() + " Kb");
                                folder_items.SubItems.Add("-");
                                folder_items.SubItems.Add(_GetFileInfo(e.FullPath));
                                listBox5.Items.Add(listBox5.Items.Count + " <" + DateTime.Now.ToString() + ">" + "  <Event: " + e.ChangeType.ToString() + ">  <Name: " + f.Name + ">  <Path: " + e.FullPath + ">  <Size: " + (kk.ToString() + " Kb>"));
                            }
                            else
                            {
                                folder_items.SubItems.Add((k.Length).ToString() + " Bytes");
                                folder_items.SubItems.Add("-");
                                folder_items.SubItems.Add(_GetFileInfo(e.FullPath));
                                listBox5.Items.Add(listBox5.Items.Count + " <" + DateTime.Now.ToString() + ">" + "  <Event: " + e.ChangeType.ToString() + ">  <Name: " + f.Name + ">  <Path: " + e.FullPath + ">  <Size: " + (k.Length).ToString() + " Bytes>");
                            }
                            listView1.Items.AddRange(new ListViewItem[] { folder_items });
                            listView1.CheckBoxes = false;

                            listBox5.SelectedIndex = listBox5.Items.Count - 1;
                        }
                    }

                }

                toolStripStatusLabel1.Text = "Monitor Status: On  " + " | Path: " + fileSystemWatcher1.Path;
                // this.Text = "File System Watch - Events: " + (listBox5.Items.Count + 1);
                string times = DateTime.Now.ToString();
                listView1.EnsureVisible(listView1.Items.Count - 1);
               
            }
            catch (Exception Change_Err) { }
        }
        private void _Create(System.IO.FileSystemEventArgs e)
        {
            string myevent = e.FullPath;
            DirectoryInfo g = new DirectoryInfo(myevent);
            FileInfo f = new FileInfo(myevent);

            string a2 = e.ChangeType.ToString();
            string myeventName = "Folder : " + e.FullPath + " is " + a2 + " : " + f.Name;

            string a3 = " --> ";
            string a4 = f.Name + a3 + a2;
            string a5 = myeventName;
            try
            {

                if (Directory.Exists(f.FullName))
                {
                    ListViewItem folder_items = new ListViewItem();
                    folder_items.SubItems.Add(DateTime.Now.ToString());
                    folder_items.SubItems.Add(e.ChangeType.ToString());
                    folder_items.SubItems.Add(e.FullPath);
                    folder_items.SubItems.Add(f.Name);
                    folder_items.SubItems.Add("Folder");
                    folder_items.SubItems.Add("-");
                    folder_items.SubItems.Add(return_Folderandfiles(e.FullPath));
                    listView1.Items.AddRange(new ListViewItem[] { folder_items });
                    listView1.CheckBoxes = false;
                    listBox5.Items.Add(listBox5.Items.Count + " <" + DateTime.Now.ToString() + ">" + "  <Event: " + e.ChangeType.ToString() + ">  <Name: " + f.Name + ">  <Path: " + e.FullPath);
                    listBox6.Items.Add(listBox6.Items.Count + " <" + DateTime.Now.ToString() + ">" + "  <Event: " + e.ChangeType.ToString() + ">  <Name: " + f.Name + ">  <Path: " + e.FullPath + ">  <Size: " + " >"); listBox6.SelectedIndex = listBox6.Items.Count - 1;
                    listBox5.SelectedIndex = listBox5.Items.Count - 1;
                }
                else
                {
                    if (File.Exists(f.FullName))
                    {
                        FileInfo myinfo = new FileInfo(e.FullPath);
                        System.IO.FileInfo k = new FileInfo(e.FullPath);
                        ListViewItem folder_items = new ListViewItem();
                        folder_items.SubItems.Add(DateTime.Now.ToString());
                        folder_items.SubItems.Add(e.ChangeType.ToString());
                        folder_items.SubItems.Add(e.FullPath);
                        folder_items.SubItems.Add(f.Name);
                        folder_items.SubItems.Add("File");
                        listBox6.Items.Add(listBox6.Items.Count + " <" + DateTime.Now.ToString() + ">" + "  <Event: " + e.ChangeType.ToString() + ">  <Name: " + f.Name + ">  <Path: " + e.FullPath + ">  <Size: " + (k.Length).ToString() + " Bytes>"); listBox6.SelectedIndex = listBox6.Items.Count - 1;
                        if (f.Name != "ntuser.dat.LOG")
                        {

                            if (e.Name.ToUpper().StartsWith("WINDOWS") || toolStripComboBox1.Text.ToUpper().Substring(3).StartsWith("WINDOWS")) { listBox3.Items.Add(listBox3.Items.Count + " <" + DateTime.Now.ToString() + ">" + "  <Event: " + e.ChangeType.ToString() + ">  <Name: " + f.Name + ">  <Path: " + e.FullPath + ">  <Size: " + (k.Length).ToString() + " Bytes>"); listBox3.SelectedIndex = listBox3.Items.Count - 1; }
                            if (e.Name.ToUpper().StartsWith("PROGRAM FILES") || toolStripComboBox1.Text.ToUpper().Substring(3).StartsWith("PROGRAM FILES")) { listBox1.Items.Add(listBox1.Items.Count + " <" + DateTime.Now.ToString() + ">" + "  <Event: " + e.ChangeType.ToString() + ">  <Name: " + f.Name + ">  <Path: " + e.FullPath + ">  <Size: " + (k.Length).ToString() + " Bytes>"); listBox1.SelectedIndex = listBox1.Items.Count - 1; }
                            if (e.Name.ToUpper().StartsWith("DOCUMENTS AND SETTINGS") || toolStripComboBox1.Text.ToUpper().Substring(3).StartsWith("DOCUMENTS AND SETTINGS")) { listBox2.Items.Add(listBox2.Items.Count + " <" + DateTime.Now.ToString() + ">" + "  <Event: " + e.ChangeType.ToString() + ">  <Name: " + f.Name + ">  <Path: " + e.FullPath + ">  <Size: " + (k.Length).ToString() + " Bytes>"); listBox2.SelectedIndex = listBox2.Items.Count - 1; }
                            toolStripTextBox1.Text = f.Name + "   " + e.ChangeType.ToString();
                        }

                        if (k.Length >= 1024)
                        {
                            long kk = k.Length / 1024;
                            folder_items.SubItems.Add(kk.ToString() + " Kb");
                            folder_items.SubItems.Add("-");
                            folder_items.SubItems.Add(_GetFileInfo(e.FullPath));
                            listBox5.Items.Add(listBox5.Items.Count + " <" + DateTime.Now.ToString() + ">" + "  <Event: " + e.ChangeType.ToString() + ">  <Name: " + f.Name + ">  <Path: " + e.FullPath + ">  <Size: " + (kk.ToString() + " Kb>"));
                        }
                        else
                        {
                            folder_items.SubItems.Add((k.Length).ToString() + " Bytes");
                            folder_items.SubItems.Add("-");
                            folder_items.SubItems.Add(_GetFileInfo(e.FullPath));
                            listBox5.Items.Add(listBox5.Items.Count + " <" + DateTime.Now.ToString() + ">" + "  <Event: " + e.ChangeType.ToString() + ">  <Name: " + f.Name + ">  <Path: " + e.FullPath + ">  <Size: " + (k.Length).ToString() + " Bytes>");
                        }

                        listView1.Items.AddRange(new ListViewItem[] { folder_items });
                        listView1.CheckBoxes = false;
                        listBox5.SelectedIndex = listBox5.Items.Count - 1;
                    }
                }
                string times = DateTime.Now.ToString();
                toolStripStatusLabel1.Text = "Monitor Status: On  " + " | Path: " + fileSystemWatcher1.Path;
                //this.Text = "File System Watch - Events: " + (listBox5.Items.Count + 1);
                //listBox6.Items.Add(listBox6.Items.Count + " <" + DateTime.Now.ToString() + ">" + "  <Event: " + e.ChangeType.ToString() + ">  <Name: " + f.Name + ">  <Path: " + e.FullPath + ">  <Size: " + (k.Length).ToString() + " Bytes>"); listBox6.SelectedIndex = listBox6.Items.Count - 1;




                toolStripTextBox1.Text = f.Name + "   " + e.ChangeType.ToString();
                listView1.EnsureVisible(listView1.Items.Count - 1);


            }
            catch (Exception Create_Err) { }
        }
        private void _Delete(System.IO.FileSystemEventArgs e)
        {
            string myevent = e.FullPath;
            DirectoryInfo g = new DirectoryInfo(myevent);
            FileInfo f = new FileInfo(myevent);


            string a2 = e.ChangeType.ToString();
            string myeventName = "Folder : " + e.FullPath + " is " + a2 + " : " + f.Name;

            string a3 = " --> ";
            string a4 = f.Name + a3 + a2;
            string a5 = myeventName;

            try
            {

                ListViewItem folder_items = new ListViewItem();
                folder_items.SubItems.Add(DateTime.Now.ToString());
                folder_items.SubItems.Add(e.ChangeType.ToString());
                folder_items.SubItems.Add(e.FullPath);
                folder_items.SubItems.Add(e.Name);
                folder_items.SubItems.Add("-");
                folder_items.SubItems.Add("-");
                folder_items.SubItems.Add("-");
                listView1.Items.AddRange(new ListViewItem[] { folder_items });
                listView1.CheckBoxes = false;
                string times = DateTime.Now.ToString();
                toolStripStatusLabel1.Text = "Monitor Status: On  " + " | Path: " + fileSystemWatcher1.Path;
                listBox4.Items.Add(listBox4.Items.Count + " <" + DateTime.Now.ToString() + ">" + "  <Event: " + e.ChangeType.ToString() + ">  <Name: " + f.Name + ">  <Path: " + e.FullPath + ">  <Size: " + " >"); listBox4.SelectedIndex = listBox4.Items.Count - 1;
                listBox5.Items.Add(listBox5.Items.Count + " <" + DateTime.Now.ToString() + ">" + "  <Event: " + e.ChangeType.ToString() + ">  <Name: " + f.Name + ">  <Path: " + e.FullPath);
                listBox5.SelectedIndex = listBox5.Items.Count - 1;


                // this.Text = "File System Watch - Events: " + (listBox5.Items.Count + 1);
                if (f.Name != "ntuser.dat.LOG")
                {

                    if (e.Name.ToUpper().StartsWith("WINDOWS") || toolStripComboBox1.Text.ToUpper().Substring(3).StartsWith("WINDOWS")) { listBox3.Items.Add(listBox3.Items.Count + " <" + DateTime.Now.ToString() + ">" + "  <Event: " + e.ChangeType.ToString() + ">  <Name: " + f.Name + ">  <Path: " + e.FullPath + ">  <Size: " + " >"); listBox3.SelectedIndex = listBox3.Items.Count - 1; }
                    if (e.Name.ToUpper().StartsWith("PROGRAM FILES") || toolStripComboBox1.Text.ToUpper().Substring(3).StartsWith("PROGRAM FILES")) { listBox1.Items.Add(listBox1.Items.Count + " <" + DateTime.Now.ToString() + ">" + "  <Event: " + e.ChangeType.ToString() + ">  <Name: " + f.Name + ">  <Path: " + e.FullPath + ">  <Size: " + " >"); listBox1.SelectedIndex = listBox1.Items.Count - 1; }
                    if (e.Name.ToUpper().StartsWith("DOCUMENTS AND SETTINGS") || toolStripComboBox1.Text.ToUpper().Substring(3).StartsWith("DOCUMENTS AND SETTINGS")) { listBox2.Items.Add(listBox2.Items.Count + " <" + DateTime.Now.ToString() + ">" + "  <Event: " + e.ChangeType.ToString() + ">  <Name: " + f.Name + ">  <Path: " + e.FullPath + ">  <Size: " + " >"); listBox2.SelectedIndex = listBox2.Items.Count - 1; }
                    toolStripTextBox1.Text = f.Name + "   " + e.ChangeType.ToString();
                }
                toolStripTextBox1.Text = f.Name + "   " + e.ChangeType.ToString();
                listView1.EnsureVisible(listView1.Items.Count - 1);
            }
            catch (Exception Delete_Err) { }
        }
        private void _Rename(string OldName, System.IO.FileSystemEventArgs e)
        {
            string myevent = e.FullPath;            
            DirectoryInfo g = new DirectoryInfo(myevent);
            FileInfo f = new FileInfo(myevent);

            string a2 = e.ChangeType.ToString();
            string myeventName = "Folder : " + g.FullName + " is " + a2 + " : " + f.Name;

            string a3 = " --> ";
            string a4 = f.Name + a3 + a2;
            string a5 = myeventName;
            listBox7.Items.Add(listBox7.Items.Count + " <" + DateTime.Now.ToString() + ">" + "  <Event: " + e.ChangeType.ToString() + ">  <Old Name: " + OldName + ">  <New Name: " + f.Name + ">  <Path: " + e.FullPath + ">");
            listBox7.SelectedIndex = listBox7.Items.Count - 1;
            try
            {
                if (Directory.Exists(f.FullName))
                {
                    ListViewItem folder_items = new ListViewItem();
                    folder_items.SubItems.Add(DateTime.Now.ToString());
                    folder_items.SubItems.Add(e.ChangeType.ToString());
                    folder_items.SubItems.Add(e.FullPath);
                    folder_items.SubItems.Add(f.Name);
                    folder_items.SubItems.Add("Folder");
                    folder_items.SubItems.Add("-");
                    folder_items.SubItems.Add(return_Folderandfiles(e.FullPath));
                    listView1.Items.AddRange(new ListViewItem[] { folder_items });
                    listView1.CheckBoxes = false;
                    listBox5.Items.Add(listBox5.Items.Count + " <" + DateTime.Now.ToString() + ">" + "  <Event: " + e.ChangeType.ToString() + ">  <Name: " + f.Name + ">  <Path: " + e.FullPath+">");
                    listBox5.SelectedIndex = listBox5.Items.Count - 1;

                }
                else
                {
                    if (File.Exists(f.FullName))
                    {
                        FileInfo myinfo = new FileInfo(e.FullPath);
                        System.IO.FileInfo k = new FileInfo(e.FullPath);
                        ListViewItem folder_items = new ListViewItem();
                        folder_items.SubItems.Add(DateTime.Now.ToString());
                        folder_items.SubItems.Add(e.ChangeType.ToString());
                        folder_items.SubItems.Add(e.FullPath);
                        folder_items.SubItems.Add(f.Name);
                        folder_items.SubItems.Add("File");
                        if (f.Name != "ntuser.dat.LOG")
                        {

                            if (e.Name.ToUpper().StartsWith("WINDOWS") || toolStripComboBox1.Text.ToUpper().Substring(3).StartsWith("WINDOWS")) { listBox3.Items.Add(listBox3.Items.Count + " <" + DateTime.Now.ToString() + ">" + "  <Event: " + e.ChangeType.ToString() + ">  <Name: " + f.Name + ">  <Path: " + e.FullPath + ">  <Size: " + (k.Length).ToString() + " Bytes>"); listBox3.SelectedIndex = listBox3.Items.Count - 1; }
                            if (e.Name.ToUpper().StartsWith("PROGRAM FILES") || toolStripComboBox1.Text.ToUpper().Substring(3).StartsWith("PROGRAM FILES")) { listBox1.Items.Add(listBox1.Items.Count + " <" + DateTime.Now.ToString() + ">" + "  <Event: " + e.ChangeType.ToString() + ">  <Name: " + f.Name + ">  <Path: " + e.FullPath + ">  <Size: " + (k.Length).ToString() + " Bytes>"); listBox1.SelectedIndex = listBox1.Items.Count - 1; }
                            if (e.Name.ToUpper().StartsWith("DOCUMENTS AND SETTINGS") || toolStripComboBox1.Text.ToUpper().Substring(3).StartsWith("DOCUMENTS AND SETTINGS")) { listBox2.Items.Add(listBox2.Items.Count + " <" + DateTime.Now.ToString() + ">" + "  <Event: " + e.ChangeType.ToString() + ">  <Name: " + f.Name + ">  <Path: " + e.FullPath + ">  <Size: " + (k.Length).ToString() + " Bytes>"); listBox2.SelectedIndex = listBox2.Items.Count - 1; }
                            toolStripTextBox1.Text = f.Name + "   " + e.ChangeType.ToString();
                        }

                        if (k.Length >= 1024)
                        {
                            long kk = k.Length / 1024;
                            folder_items.SubItems.Add(kk.ToString() + " Kb");
                            folder_items.SubItems.Add("-");
                            folder_items.SubItems.Add(_GetFileInfo(e.FullPath));
                            listBox5.Items.Add(listBox5.Items.Count + " <" + DateTime.Now.ToString() + ">" + "  <Event: " + e.ChangeType.ToString() + ">  <Name: " + f.Name + ">  <Path: " + e.FullPath + ">  <Size: " + (kk.ToString() + " Kb>"));
                        }
                        else
                        {
                            folder_items.SubItems.Add((k.Length).ToString() + " Bytes");
                            folder_items.SubItems.Add("-");
                            folder_items.SubItems.Add(_GetFileInfo(e.FullPath));
                            listBox5.Items.Add(listBox5.Items.Count + " <" + DateTime.Now.ToString() + ">" + "  <Event: " + e.ChangeType.ToString() + ">  <Name: " + f.Name + ">  <Path: " + e.FullPath + ">  <Size: " + (k.Length).ToString() + " Bytes>");
                        }

                        listView1.Items.AddRange(new ListViewItem[] { folder_items });
                        listView1.CheckBoxes = false;
                        listBox5.SelectedIndex = listBox5.Items.Count - 1;
                    }
                }
                string times = DateTime.Now.ToString();
                toolStripStatusLabel1.Text = "Monitor Status: On  " + " | Path: " + fileSystemWatcher1.Path;
                // this.Text = "File System Watch - Events: " + (listBox5.Items.Count + 1);

                listView1.EnsureVisible(listView1.Items.Count - 1);

            }
            catch (Exception Rename_Err) { }
        }

        private void fileSystemWatcher1_Changed(object sender, System.IO.FileSystemEventArgs e)
        {

            Change_Delegate a = new Change_Delegate(_Changes);
            BeginInvoke(a, e);
             //_Changes(e);
             //Thread.Sleep(1);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (tabPage1.Focus())
                {
                    fileSystemWatcher1.Path = toolStripComboBox1.Text;

                }
                fileSystemWatcher1.Path = toolStripComboBox1.Text;
            }
            catch (Exception f) { MessageBox.Show(f.Message); }
                
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            fileSystemWatcher1.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.DirectoryName | NotifyFilters.FileName | NotifyFilters.LastWrite;
            fileSystemWatcher1.Error += new ErrorEventHandler(fileSystemWatcher1_Error);
            timer1.Start();

            fileSystemWatcher1.EnableRaisingEvents = false;
            string [] mydrive = Directory.GetLogicalDrives();
            foreach (string mydrv in mydrive) { toolStripComboBox1.Items.Add(mydrv);            
            toolStripComboBox1.SelectedIndex =0;
                //------
            listView1.Items.Clear();
            listView1.View = View.Details;
            // Allow the user to rearrange columns.
            listView1.AllowColumnReorder = true;
            // Display check boxes.
            listView1.CheckBoxes = true;
            // Select the item and subitems when selection is made.
            listView1.FullRowSelect = true;
            // Display grid lines.
            listView1.GridLines = true;
           

            listView1.Columns.Clear();
            listView1.Columns.Add("NA", 10, HorizontalAlignment.Left);
            listView1.Columns.Add("Last Access Time", 120, HorizontalAlignment.Left);
            listView1.Columns.Add("Event", 60, HorizontalAlignment.Left);
            listView1.Columns.Add("Full Path", 250, HorizontalAlignment.Left);
            listView1.Columns.Add("Name", 100, HorizontalAlignment.Left);
            listView1.Columns.Add("Type", 48, HorizontalAlignment.Left);
            listView1.Columns.Add("File Size", 60, HorizontalAlignment.Left);
            listView1.Columns.Add("Total : file and folder", 80, HorizontalAlignment.Left);
            listView1.Columns.Add("Other Information", 500, HorizontalAlignment.Left);            
          
                //------
        }
        }
        
        private void fileSystemWatcher1_Created(object sender, System.IO.FileSystemEventArgs e)
        {

            _Create(e);
         
        }

        private void fileSystemWatcher1_Deleted(object sender, System.IO.FileSystemEventArgs e)
        {


            _Delete(e);
        }

        private void fileSystemWatcher1_Renamed(object sender, System.IO.RenamedEventArgs e)
        {
            FileInfo RenameOldName = new FileInfo(e.OldName);

            _Rename(RenameOldName.Name,e);
        }

        private void start()
        {
            fileSystemWatcher1.EnableRaisingEvents = true;
            try
            {
                if (fileSystemWatcher1.EnableRaisingEvents == true)
                {
                    fileSystemWatcher1.EnableRaisingEvents = true;
                    fileSystemWatcher1.IncludeSubdirectories = true;
                    //fileSystemWatcher1.Filter = "*.*";
                }
            }
            catch (Exception f) { fileSystemWatcher1.EnableRaisingEvents = false; MessageBox.Show(f.Message); }

            //if (fileSystemWatcher1.EnableRaisingEvents == false) { toolStripStatusLabel1.Text = "Monitor Status: Off  "; }
        }
        private void toolStripButton1_Click(object sender, EventArgs e)
        {

            drv = toolStripComboBox1.Text;
            bingo = new Thread(start);
            bingo.Start();

            //  fileSystemWatcher1.EnableRaisingEvents = true;
           // toolStripStatusLabel1.Text = "Monitor Status: On  ";

            //try
            //{

            //    if (fileSystemWatcher1.EnableRaisingEvents == true)
            //    {
            //        toolStripStatusLabel1.Text = "Monitor Status: On  ";
            //       // fileSystemWatcher1.EnableRaisingEvents = true;
            //       // fileSystemWatcher1.IncludeSubdirectories = true;
            //       // fileSystemWatcher1.Path = toolStripComboBox1.Text;
            //        toolStripStatusLabel1.Text = "Monitor Status: On  " + " | Path: " + fileSystemWatcher1.Path;
            //    }


            //    fileSystemWatcher1.Path = toolStripComboBox1.Text;
            //}
            //catch (Exception f) { fileSystemWatcher1.EnableRaisingEvents = false; MessageBox.Show(f.Message); }
            //if (fileSystemWatcher1.EnableRaisingEvents == false) { toolStripStatusLabel1.Text = "Monitor Status: Off  "; }
        }

        void fileSystemWatcher1_Error(object sender, ErrorEventArgs e)
        {
            
            
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            listView1.Items.Clear();
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            Form.CheckForIllegalCrossThreadCalls = false;
            if (listBox5.Items.Count > 0) { button1.Enabled = true; } else { button1.Enabled = false; }
            if (listBox5.Items.Count == 0) { button1.Enabled = false; }
            try
            {

                listBox1.Items.Clear();
                listBox2.Items.Clear();
                listBox3.Items.Clear();
                listBox5.Items.Clear();
                listBox4.Items.Clear();
                listBox6.Items.Clear();
                listView1.Items.Clear();
           
        
            tabPage6.Text = "File and Folder (Create Events) " + listBox6.Items.Count; 
            tabPage5.Text = "File and Folder (Delete Events) " + listBox4.Items.Count; 
            tabPage4.Text = "All " + listBox5.Items.Count; 
            tabPage3.Text = "Program Files " + listBox1.Items.Count; 
            tabPage2.Text = "Windows Directory " + listBox3.Items.Count; 
            tabPage1.Text = "Documents and Settings " + listBox2.Items.Count; 
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            if (fileSystemWatcher1.EnableRaisingEvents == true)
            {
                toolStripStatusLabel1.Text = "Monitor Status: Off";
                fileSystemWatcher1.EnableRaisingEvents = false;
            }
        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {
           
            tableLayoutPanel.ColumnCount = 1;
            tableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33F));
            tableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 67F));
            tableLayoutPanel.Controls.Add(ProductName, 1, 0);
            tableLayoutPanel.Controls.Add(Version, 1, 1);
            tableLayoutPanel.Controls.Add(labelCopyright, 1, 2);
            tableLayoutPanel.Controls.Add(labelCompanyName, 1, 3);
            tableLayoutPanel.Controls.Add(textBoxDescription, 1, 4);
            tableLayoutPanel.Controls.Add(okButton, 1, 5);
            tableLayoutPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            tableLayoutPanel.Location = new System.Drawing.Point(9, 9);
            tableLayoutPanel.Name = "tableLayoutPanel";
            tableLayoutPanel.RowCount = 6;
            tableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10F));
            tableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10F));
            tableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10F));
            tableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10F));
            tableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            tableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10F));
            tableLayoutPanel.Size = new System.Drawing.Size(417, 265);
            tableLayoutPanel.TabIndex = 0;
            // 
            // ProductName
            // 
            ProductName.Dock = System.Windows.Forms.DockStyle.Fill;
            ProductName.Location = new System.Drawing.Point(143, 0);
            ProductName.Margin = new System.Windows.Forms.Padding(6, 0, 3, 0);
            ProductName.MaximumSize = new System.Drawing.Size(0, 17);
            ProductName.Name = "ProductName";
            ProductName.Size = new System.Drawing.Size(271, 17);
            ProductName.TabIndex = 19;
            ProductName.Text = "Product Name: " + AssemblyProduct;
            ProductName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // Version
            // 
            Version.Dock = System.Windows.Forms.DockStyle.Fill;
            Version.Location = new System.Drawing.Point(143, 26);
            Version.Margin = new System.Windows.Forms.Padding(6, 0, 3, 0);
            Version.MaximumSize = new System.Drawing.Size(0, 17);
            Version.Name = "Version";
            Version.Size = new System.Drawing.Size(271, 17);
            Version.TabIndex = 0;
            Version.Text = "Version: " + AssemblyVersion;
            Version.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // labelCopyright
            // 
            labelCopyright.Dock = System.Windows.Forms.DockStyle.Fill;
            labelCopyright.Location = new System.Drawing.Point(143, 52);
            labelCopyright.Margin = new System.Windows.Forms.Padding(6, 0, 3, 0);
            labelCopyright.MaximumSize = new System.Drawing.Size(0, 17);
            labelCopyright.Name = "labelCopyright";
            labelCopyright.Size = new System.Drawing.Size(271, 17);
            labelCopyright.TabIndex = 21;
            labelCopyright.Text = AssemblyCopyright;
            labelCopyright.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // labelCompanyName
            // 
            labelCompanyName.Dock = System.Windows.Forms.DockStyle.Fill;
            labelCompanyName.Location = new System.Drawing.Point(143, 78);
            labelCompanyName.Margin = new System.Windows.Forms.Padding(6, 0, 3, 0);
            labelCompanyName.MaximumSize = new System.Drawing.Size(0, 17);
            labelCompanyName.Name = "labelCompanyName";
            labelCompanyName.Size = new System.Drawing.Size(271, 17);
            labelCompanyName.TabIndex = 22;
            labelCompanyName.Text = "Company Name: " + AssemblyCompany;
            labelCompanyName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // textBoxDescription
            // 
            textBoxDescription.Dock = System.Windows.Forms.DockStyle.Fill;
            textBoxDescription.Location = new System.Drawing.Point(143, 107);
            textBoxDescription.Margin = new System.Windows.Forms.Padding(6, 3, 3, 3);
            textBoxDescription.Multiline = true;
            textBoxDescription.Name = "textBoxDescription";
            textBoxDescription.ReadOnly = true;
            textBoxDescription.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            textBoxDescription.Size = new System.Drawing.Size(271, 126);
            textBoxDescription.TabIndex = 23;
            textBoxDescription.TabStop = false;
            textBoxDescription.Text = "Description: " + AssemblyDescription + "   ( .Net Framework 2.0 )";
            // 
            // okButton
            // 
            okButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            okButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            okButton.Location = new System.Drawing.Point(339, 239);
            okButton.Name = "okButton";
            okButton.Size = new System.Drawing.Size(75, 23);
            okButton.TabIndex = 24;
            okButton.Text = "&OK";

            //
            m.Bounds = new Rectangle(new Point(250, 0), new Size(250, 140));
            m.MaximizeBox = false;
            m.MinimizeBox = false;
            m.AcceptButton = okButton;
            m.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            m.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            m.ClientSize = new System.Drawing.Size(435, 283);
            m.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            m.MaximizeBox = false;
            m.MinimizeBox = false;
            m.Padding = new System.Windows.Forms.Padding(9);
            m.ShowIcon = false;
            m.BackColor = Color.FromArgb(225, 240, 255);
            m.ShowInTaskbar = false;
            m.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            m.Text = "About " + AssemblyProduct;
            tableLayoutPanel.Paint += new PaintEventHandler(tableLayoutPanel_Paint);
            m.Controls.Add(tableLayoutPanel);
            tableLayoutPanel.ResumeLayout(false);
            tableLayoutPanel.PerformLayout();
            m.ResumeLayout(false);
            m.Paint += new PaintEventHandler(m_Paint);
            tableLayoutPanel.Show();
            m.ShowDialog();    
                
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {

            fileSystemWatcher1.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.DirectoryName | NotifyFilters.FileName | NotifyFilters.LastWrite;
            
            
            if (fileSystemWatcher1.EnableRaisingEvents == true)
            {
                fileSystemWatcher1.EnableRaisingEvents = false;                
            }
            try
            {
                if (newThread1.ThreadState == ThreadState.Running)
                {
                    newThread1.Abort();
                    newThread1 = null;


                }
            }
            catch (Exception thexe) { }
            
            
        }

        private void listBox5_MouseClick(object sender, MouseEventArgs e)
        {
            toolTip1.SetToolTip(listBox5, " ");
            toolTip1.ToolTipIcon = ToolTipIcon.Info;
            toolTip1.ToolTipTitle = "All Events";
            toolTip1.SetToolTip(listBox5,listBox5.Text);
            
        }

        private void listBox4_MouseClick(object sender, MouseEventArgs e)
        {
            toolTip1.SetToolTip(listBox4, " ");
            toolTip1.ToolTipIcon = ToolTipIcon.Error;
            toolTip1.ToolTipTitle = "Delete Events";
            toolTip1.SetToolTip(listBox4, listBox4.Text);
        }

        private void listBox6_MouseClick(object sender, MouseEventArgs e)
        {
            toolTip1.SetToolTip(listBox6, " ");
            toolTip1.ToolTipIcon = ToolTipIcon.Info;
            toolTip1.ToolTipTitle = "Create Events";
            toolTip1.SetToolTip(listBox6, listBox6.Text);
        }

        private void listBox1_MouseClick(object sender, MouseEventArgs e)
        {
            toolTip1.SetToolTip(listBox1, " ");
            toolTip1.ToolTipIcon = ToolTipIcon.Info;
            toolTip1.ToolTipTitle = "Program Files";
            toolTip1.SetToolTip(listBox1, listBox1.Text);
        }

        private void listBox3_MouseClick(object sender, MouseEventArgs e)
        {
            toolTip1.SetToolTip(listBox3, " ");
            toolTip1.ToolTipIcon = ToolTipIcon.Info;
            toolTip1.ToolTipTitle = "Windows Directory";
            toolTip1.SetToolTip(listBox3, listBox3.Text);
        }

        private void listBox2_MouseClick(object sender, MouseEventArgs e)
        {
            toolTip1.SetToolTip(listBox2, " ");
            toolTip1.ToolTipIcon = ToolTipIcon.Info;
            toolTip1.ToolTipTitle = "Documents and Settings";
            toolTip1.SetToolTip(listBox2, listBox2.Text);
        }

        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.ShowDialog();
            string myaddress = folderBrowserDialog1.SelectedPath;
            toolStripComboBox1.Text = myaddress ;
        }

        private void listBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            tabPage6.Text = "Create Events " + listBox6.Items.Count; 
        }

        private void listBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            tabPage5.Text = "Delete Events " + listBox4.Items.Count; 
        }

        private void listBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            Form.CheckForIllegalCrossThreadCalls = false;
            tabPage4.Text = "All " + listBox5.Items.Count;
            Form.CheckForIllegalCrossThreadCalls = false;
            if (listBox5.Items.Count > 0) { button1.Enabled = true; } else { button1.Enabled = false; }
            
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            tabPage3.Text = "Program Files " + listBox1.Items.Count; 
        }

        private void listBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            tabPage2.Text = "Windows Directory " + listBox3.Items.Count; 
        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            tabPage1.Text = "Documents and Settings " + listBox2.Items.Count; 
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            if (listBox5.Items.Count > 0)
            {
                threadDelegate1 = new ThreadStart(Thread1);
                newThread1 = new Thread(threadDelegate1);
                newThread1.Start();
            }

        }
        
        public  void Thread1()
        {
            ///-------------------xls codes--------------------
            /// dll removed ;)
            ///

           // C1.C1Excel.C1XLBook c1XLBook1 = new C1XLBook();

            //c1XLBook1.Clear();

            //// add some styles                   
            //XLStyle s1 = new XLStyle(c1XLBook1); //id
            //XLStyle s2 = new XLStyle(c1XLBook1); //time
            //XLStyle s3 = new XLStyle(c1XLBook1); //event type
            //XLStyle s4 = new XLStyle(c1XLBook1); //file_Name
            //XLStyle s5 = new XLStyle(c1XLBook1); //file_fullpath
            //XLStyle s6 = new XLStyle(c1XLBook1); //size


            //s5.Font = new Font("Tahoma", 10, FontStyle.Bold);

            //s1.Format = "general";
            //s2.Format = "dd/MM/yyyy hh:mm";
            //s3.Format = "general";
            //s4.Format = "general";
            //s5.Format = "general";
            //s6.Format = "general";

            //s3.ForeColor = Color.LightGreen;

            //s7.Format = "general";
            //s7.Font = new Font("Courier New", 14);
            //s8.Format = "general";
            //s9.Format = "general";
            //s10.ForeColor = Color.LightYellow;
            //s10.BackColor = Color.Black;
            //s10.Font = new Font("Tahoma", 8, FontStyle.Bold);
            //s11.ForeColor = Color.Red;
            //s11.BackColor = Color.Black;
            //s11.Font = new Font("Tahoma", 8, FontStyle.Bold);
            //s12.Format = "general";
            // populate sheet with some random values

            //C1.C1Excel.XLSheet sheet = c1XLBook1.Sheets[0];

           
            //sheet[0, 1].Value = "Index";
            //sheet[0, 2].Value = "TimeWritten";
            //sheet[0, 3].Value = "Event_Type";
            //sheet[0, 4].Value = "FileName";
            //sheet[0, 5].Value = "File_FullPath";
            //sheet[0, 6].Value = "Size";

            //sheet[0, 1].Style = (cu % 13 == 0) ? s5 : s5;
            //sheet[0, 2].Style = (cu % 13 == 0) ? s5 : s5;
            //sheet[0, 3].Style = (cu % 13 == 0) ? s5 : s5;
            //sheet[0, 4].Style = (cu % 13 == 0) ? s5 : s5;
            //sheet[0, 5].Style = (cu % 13 == 0) ? s5 : s5;
            //sheet[0, 6].Style = (cu % 13 == 0) ? s5 : s5;

            //-------------------xls codes--------------------


            Thread.Sleep(1000);            
            string mycache = this.Text;
            this.Text = "Please Wait.....";
            Form.CheckForIllegalCrossThreadCalls = false;                            
            this.tabControl1.Visible = false;            
            Form.CheckForIllegalCrossThreadCalls = false;                            
                try
                {
                    Form.CheckForIllegalCrossThreadCalls = false;
                if (checkBox1.Checked)
                {
                    string Filename = "Documents and Settings_log.txt";
                    StreamWriter w = File.AppendText(Filename);
                    for (int plus = 0; plus <= listBox2.Items.Count - 1; plus++)
                    {
                        listBox2.SelectedIndex = plus;
                        Form.CheckForIllegalCrossThreadCalls = false;
                        __Log_Class.Log(" " + listBox2.SelectedItem.ToString(), w);
                       
                    }
                    MessageBox.Show(Filename.ToString(), "Create Report: <Documents and Settings>", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    w.Close();
                }
                Form.CheckForIllegalCrossThreadCalls = false;
                    if (checkBox2.Checked)
                    {
                        string Filename = "Windows Directory_log.txt";
                        StreamWriter w = File.AppendText(Filename);
                        for (int plus = 0; plus <= listBox3.Items.Count - 1; plus++)
                        {
                            listBox3.SelectedIndex = plus;
                            Form.CheckForIllegalCrossThreadCalls = false;
                            __Log_Class.Log(" " + listBox3.SelectedItem.ToString(), w);
                        }
                        MessageBox.Show(Filename.ToString(), "Create Report: <Windows Directory>", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        w.Close();
                    }
                    Form.CheckForIllegalCrossThreadCalls = false;
                    if (checkBox3.Checked)
                    {
                        string Filename = "Program Files_log.txt";
                        StreamWriter w = File.AppendText(Filename);
                        for (int plus = 0; plus <= listBox1.Items.Count - 1; plus++)
                        {
                            listBox1.SelectedIndex = plus;
                            Form.CheckForIllegalCrossThreadCalls = false;
                            __Log_Class.Log(" " + listBox1.SelectedItem.ToString(), w);
                        }

                        MessageBox.Show(Filename.ToString(), "Create Report: <Program Files>", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        w.Close();
                    }
                    Form.CheckForIllegalCrossThreadCalls = false;
                    if (checkBox4.Checked)
                    {
                        string Filename = "All_log.txt";
                        StreamWriter w = File.AppendText(Filename);
                        for (int plus = 0; plus <= listBox5.Items.Count - 1; plus++)
                        {
                            listBox5.SelectedIndex = plus;
                            Form.CheckForIllegalCrossThreadCalls = false;
                            __Log_Class.Log(" " + listBox5.SelectedItem.ToString(), w);
                        }

                        MessageBox.Show(Filename.ToString(), "Create Report: <All>", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        w.Close();
                    }
                    Form.CheckForIllegalCrossThreadCalls = false;
                    if (checkBox5.Checked)
                    {
                        string Filename = "Delete Events_log.txt";
                        StreamWriter w = File.AppendText(Filename);
                        for (int plus = 0; plus <= listBox4.Items.Count - 1; plus++)
                        {
                            listBox4.SelectedIndex = plus;
                            Form.CheckForIllegalCrossThreadCalls = false;
                            __Log_Class.Log(" " + listBox4.SelectedItem.ToString(), w);
                        }

                        MessageBox.Show(Filename.ToString(), "Create Report: <Delete Events>", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        w.Close();
                    }
                    Form.CheckForIllegalCrossThreadCalls = false;
                    if (checkBox6.Checked)
                    {
                        string Filename = "Create Events_log.txt";
                        StreamWriter w = File.AppendText(Filename);
                        for (int plus = 0; plus <= listBox6.Items.Count - 1; plus++)
                        {
                            listBox6.SelectedIndex = plus;
                            Form.CheckForIllegalCrossThreadCalls = false;
                            __Log_Class.Log(" " + listBox6.SelectedItem.ToString(), w);
                        }
                        MessageBox.Show(Filename.ToString(), "Create Report: <Create Events>", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        w.Close();
                    }
                    Form.CheckForIllegalCrossThreadCalls = false;
                    if (checkBox15.Checked)
                    {
                        string Filename = "Rename Events_log.txt";
                        StreamWriter w = File.AppendText(Filename);
                        for (int plus = 0; plus <= listBox7.Items.Count - 1; plus++)
                        {
                            listBox7.SelectedIndex = plus;
                            Form.CheckForIllegalCrossThreadCalls = false;
                            __Log_Class.Log(" " + listBox7.SelectedItem.ToString(), w);
                        }
                        MessageBox.Show(Filename.ToString(), "Create Report: <Rename Events>", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        w.Close();
                    }
                                      
                    this.Text = mycache;
                }
                catch (Exception exc)
                {
                    Form.CheckForIllegalCrossThreadCalls = false;
                    this.tabControl1.Visible = true;            
                    newThread1.Abort();
                    newThread1 = null;
                    MessageBox.Show(exc.Message); 
                }
                Form.CheckForIllegalCrossThreadCalls = false;                                
                this.tabControl1.Visible = true;       
        }
        private void timer1_Tick(object sender, EventArgs e)
        {            
            if (this.Opacity == 1) { timer1.Stop(); }
            this.Opacity += this.opacityIncrease;
            this.timer1.Interval = 30;
            this.ResumeLayout();
        }        

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
           
        }
        private string _GetFileInfo(string File_FullPath) 
        {
            FileInfo myfile = new FileInfo(File_FullPath);
            System.Security.AccessControl.FileSystemSecurity MYACL;            
            return ("{CreationTime: "+ myfile.CreationTime + "} {LastAccessTime: "+myfile.LastAccessTime.ToString() + "} {LastWriteTime: "+myfile.LastWriteTime.ToString()+"} {Attributes: "+myfile.Attributes+"}");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (checkBox7.Checked && checkBox8.Checked == false && checkBox13.Checked == false && checkBox14.Checked == false) { fileSystemWatcher1.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.DirectoryName | NotifyFilters.FileName | NotifyFilters.LastWrite | NotifyFilters.Attributes; }
                else
                    if (checkBox8.Checked && checkBox7.Checked == false && checkBox13.Checked == false && checkBox14.Checked == false) { fileSystemWatcher1.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.DirectoryName | NotifyFilters.FileName | NotifyFilters.LastWrite | NotifyFilters.CreationTime; }
                    else
                        if (checkBox14.Checked && checkBox7.Checked == false && checkBox13.Checked == false && checkBox8.Checked == false) { fileSystemWatcher1.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.DirectoryName | NotifyFilters.FileName | NotifyFilters.LastWrite | NotifyFilters.Size; }
                        else
                            if (checkBox13.Checked && checkBox8.Checked == false && checkBox7.Checked == false && checkBox14.Checked == false) { fileSystemWatcher1.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.DirectoryName | NotifyFilters.FileName | NotifyFilters.LastWrite | NotifyFilters.Security; }
                            else
                                if (checkBox13.Checked == false && checkBox8.Checked == false && checkBox7.Checked == false && checkBox14.Checked == false) { fileSystemWatcher1.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.DirectoryName | NotifyFilters.FileName | NotifyFilters.LastWrite; }
            }
            catch (Exception _exee) { MessageBox.Show(_exee.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void checkBox7_CheckStateChanged(object sender, EventArgs e)
        {
            if (checkBox7.Checked) { checkBox8.Checked = false; checkBox13.Checked = false; checkBox14.Checked = false; }
        }

        private void checkBox8_CheckStateChanged(object sender, EventArgs e)
        {
            if (checkBox8.Checked) { checkBox7.Checked = false; checkBox13.Checked = false; checkBox14.Checked = false; }
        }

        private void checkBox13_CheckStateChanged(object sender, EventArgs e)
        {
            if (checkBox13.Checked) { checkBox8.Checked = false; checkBox7.Checked = false; checkBox14.Checked = false; }
        }

        private void checkBox14_CheckStateChanged(object sender, EventArgs e)
        {
            if (checkBox14.Checked) { checkBox8.Checked = false; checkBox13.Checked = false; checkBox7.Checked = false; }
        }

        private void listBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
            tabPage9.Text = "Rename Events " + listBox7.Items.Count; 
        }
      
    }
}


