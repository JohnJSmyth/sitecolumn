// ***************************************
// ** Voltar Sharepoint Admin Tools
// ** Site Colums Copy
// ** Version 1.0
// ** www.voltar.ch
// ***************************************

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.SharePoint;
using System.IO;
using System.Xml;

namespace SiteColumnCopy
{
    public partial class Form1 : Form
    {
        public bool SiteFound; // is true when a sharepoint site exists
        
        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        { // This is out [...] button.
            saveFileDialog1.ShowDialog();
            textBox2.Text = saveFileDialog1.FileName;

        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            SPSite mySourceSite;
            // We check if a sharepoint site exists at all...
            try
            {
                mySourceSite = new SPSite(textBox1.Text);
                SPWeb spWeb = mySourceSite.OpenWeb();
                spWeb.Dispose();
                mySourceSite.Dispose();
                SiteFound = true; // ... the site exists.
                ErrorLabel.Visible = false;

                //fill the listbox with Site Colums group names
                SPFieldCollection mygroups = spWeb.Fields;
                listBox1.Items.Clear();
                listBox1.Items.Add("all");
                listBox1.SelectedIndex = listBox1.TopIndex; // Select the top item (all)

                foreach (SPField field in mygroups)
                {

                    if (listBox1.Items.Contains(field.Group))
                    { }
                    else
                    {
                        listBox1.Items.Add(field.Group);
                    }
                }
            }
            catch // ... the site does not exist (or you do not have permission)
            {
                ErrorLabel.Text = "This site could not be found";
                SiteFound = false;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        { // this is our [export] button
            try
            {
                if (SiteFound)
                {
                    string url = textBox1.Text;
                   
                    string groupName = listBox1.SelectedItem.ToString();
                   
                    bool useGroup;
                    if (groupName.Equals("all"))
                        useGroup = false;
                    else
                        useGroup = true;

                    string OutputFile = textBox2.Text;
                   // we need an XML String Builder in order to write out xml code.
                    StringBuilder MyStringBuilder = new StringBuilder();
                    XmlTextWriter MyxmlWriter = new XmlTextWriter(new StringWriter(MyStringBuilder));
                    MyxmlWriter.Formatting = Formatting.Indented;
                    int counter = 0; // counts how many site columns we have exported.
                    try
                    {
                        MyxmlWriter.WriteStartElement("Elements");
                       
                       // Here we get out site columns...
                        using (SPSite site = new SPSite(url))
                        {
                            using (SPWeb web = site.AllWebs[GetWebURL(url)])
                            {
                                SPFieldCollection fields = web.Fields;
                                foreach (SPField field in fields)
                                {
                                    if (!useGroup ||
                                    (useGroup && groupName == field.Group))
                                    {
                                        MyxmlWriter.WriteString("\r\n");
                                        MyxmlWriter.WriteRaw(field.SchemaXml);
                                        counter++;  // counts how many site columns we have exported.
                                    }
                                }
                                label9.Text = counter.ToString() + " Fields exported.";

                            }
                        }

                        MyxmlWriter.WriteString("\r\n");
                        MyxmlWriter.WriteEndElement();
                      
                    }
                    finally
                    {
                        // close the writer
                        MyxmlWriter.Flush();
                        MyxmlWriter.Close();
                    }
                    // write the output file.
                    File.WriteAllText(OutputFile, MyStringBuilder.ToString());
                    label9.Text += " Export succeded.";
                }
            }
            catch (Exception exp)
            {
                label9.Text = "Export failed. " + exp.Message;
            }


        }
        public string GetWebURL(string url)
            // This method comes from CodePlex.com customstsadmtemplate
        {
            int index = url.IndexOf("//");
            if ((index < 0) || (index == (url.Length - 2)))
            {
                throw new ArgumentException();
            }

            int startIndex = url.IndexOf('/', index + 2);
            if (startIndex < 0)
            {
                return "/";
            }

            string str = url.Substring(startIndex);
            if (str.IndexOf("?") >= 0)
                str = str.Substring(0, str.IndexOf("?"));

            if (str.IndexOf(".aspx") > 0)
                str = str.Substring(0, str.LastIndexOf("/"));

            if ((str.Length > 1) && (str[str.Length - 1] == '/'))
            {
                return str.Substring(0, str.Length - 1);
            }
            return str;
        }

        private void button3_Click(object sender, EventArgs e)
        { // this is our [import] button
            int counter = 0; // counts the site columns imported.
            try
            {
                if (SiteFound)
                {

                    string url = textBox5.Text;
                    string InputFile = textBox4.Text;
                    XmlDocument MyXMLFile = new XmlDocument();
                    MyXMLFile.Load(InputFile);
                    // this is our traget web 
                    using (SPSite site = new SPSite(url))
                    using (SPWeb web = site.AllWebs[GetWebURL(url)])
                    {   // now we read the xml file
                        foreach (XmlElement fieldNode in MyXMLFile.SelectNodes("//Field"))
                        {
                            web.Fields.AddFieldAsXml(fieldNode.OuterXml);
                            counter++;// counts the site columns imported.
                        }
                    }
                    label8.Text = counter.ToString() + " Fields imported";
                }
            }
            catch (Exception ex)
            {
                
                label8.Text = "Import failed." + ex.Message;
            }

        }

        private void button4_Click(object sender, EventArgs e)
        { // this is our [...] button
            openFileDialog1.ShowDialog();
            textBox4.Text = openFileDialog1.FileName;
        }

        private void textBox5_Leave(object sender, EventArgs e)
        {   // check if a sharepoint site exists
            SPSite mySourceSite;
            try
            {
                mySourceSite = new SPSite(textBox5.Text);
                SPWeb spWeb = mySourceSite.OpenWeb();
                spWeb.Dispose();
                mySourceSite.Dispose();
                SiteFound = true; // the site exists
                ErrorLabel2.Visible = false;
            }
            catch // the site does not exist.
            {
                ErrorLabel2.Text = "This site could not be found";
                ErrorLabel2.Visible = true;
                SiteFound = false;
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {   // visit the Voltar Home Page
            System.Diagnostics.Process.Start("http://www.voltar.ch");
        }

     
      
    }
   
}
