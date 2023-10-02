﻿using CsvHelper;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static RichemontMailMerge.Form1;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace RichemontMailMerge
{
    public partial class Form1 : Form
    {

        private List<EeidNote> eeidNotes = new List<EeidNote>();
        private List<EmployeeData> employeeDataList = new List<EmployeeData>();

        public Form1()
        {
            InitializeComponent();
            // Set initial helper text
            textBox1.Text = "Enter employee information";

            // Clear helper text when TextBox receives focus
            textBox1.GotFocus += (sender, e) =>
            {
                if (textBox1.Text == "Enter employee information")
                {
                    textBox1.Text = "";
                }
            };

            // Restore helper text when TextBox loses focus and is empty
            textBox1.LostFocus += (sender, e) =>
            {
                if (string.IsNullOrEmpty(textBox1.Text))
                {
                    textBox1.Text = "Enter employee information";
                }
            };

            // Set initial helper text for textBox2
            textBox2.Text = "Enter EEID";

            // Clear helper text when textBox2 receives focus
            textBox2.GotFocus += (sender, e) =>
            {
                if (textBox2.Text == "Enter EEID")
                {
                    textBox2.Text = "";
                }
            };

            // Restore helper text when textBox2 loses focus and is empty
            textBox2.LostFocus += (sender, e) =>
            {
                if (string.IsNullOrEmpty(textBox2.Text))
                {
                    textBox2.Text = "Enter EEID";
                }
            };

            // Set initial helper text
            textBox4.Text = "Enter employee information";

            // Clear helper text when TextBox receives focus
            textBox4.GotFocus += (sender, e) =>
            {
                if (textBox4.Text == "Enter employee information")
                {
                    textBox4.Text = "";
                }
            };

            // Restore helper text when TextBox loses focus and is empty
            textBox3.LostFocus += (sender, e) =>
            {
                if (string.IsNullOrEmpty(textBox4.Text))
                {
                    textBox3.Text = "Enter employee information";
                }
            };

            // Set initial helper text for textBox3
            textBox3.Text = "Enter EEID";

            // Clear helper text when textBox3 receives focus
            textBox3.GotFocus += (sender, e) =>
            {
                if (textBox3.Text == "Enter EEID")
                {
                    textBox3.Text = "";
                }
            };

            // Restore helper text when textBox3 loses focus and is empty
            textBox3.LostFocus += (sender, e) =>
            {
                if (string.IsNullOrEmpty(textBox3.Text))
                {
                    textBox3.Text = "Enter EEID";
                }
            };

            // Set initial helper text
            textBox5.Text = "Enter EEID";

            // Clear helper text when TextBox receives focus
            textBox5.GotFocus += (sender, e) =>
            {
                if (textBox5.Text == "Enter EEID")
                {
                    textBox5.Text = "";
                }
            };

            // Restore helper text when TextBox loses focus and is empty
            textBox5.LostFocus += (sender, e) =>
            {
                if (string.IsNullOrEmpty(textBox5.Text))
                {
                    textBox5.Text = "Enter EEID";
                }
            };

            // Set initial helper text
            textBox6.Text = "Enter EE name";

            // Clear helper text when TextBox receives focus
            textBox6.GotFocus += (sender, e) =>
            {
                if (textBox6.Text == "Enter EE name")
                {
                    textBox6.Text = "";
                }
            };

            // Restore helper text when TextBox loses focus and is empty
            textBox6.LostFocus += (sender, e) =>
            {
                if (string.IsNullOrEmpty(textBox6.Text))
                {
                    textBox6.Text = "Enter EE Name";
                }
            };

            // Set initial helper text
            textBox7.Text = "Enter EE email address";

            // Clear helper text when TextBox receives focus
            textBox7.GotFocus += (sender, e) =>
            {
                if (textBox7.Text == "Enter EE email address")
                {
                    textBox7.Text = "";
                }
            };

            // Restore helper text when TextBox loses focus and is empty
            textBox7.LostFocus += (sender, e) =>
            {
                if (string.IsNullOrEmpty(textBox7.Text))
                {
                    textBox7.Text = "Enter EE email address";
                }
            };


            panel4.Visible = true;
            panel5.Visible = false;
            panel6.Visible = false;
            panel7.Visible = false;
            panel8.Visible = false;
        }

        public class Record
        {
            public string CaseID { get; set; }
            public string EEID { get; set; }
            public string NoteText { get; set; }
            public string Private { get; set; }
        }

        public class EmployeeData
        {
            public string EEID { get; set; }
            public string Name { get; set; }
            public string Email { get; set; }
            public DateTime Date { get; set; }
            public string Group1Selection { get; set; }
            public string Group2Selection { get; set; }
        }

        public class EeidNote
        {
            public string Eeid { get; set; }
            public string Note { get; set; }
        }

        private void CreateNoteload1(string eeid, string name, string note)
        {
            // Create a parent node for the letter
            TreeNode letterNode = new TreeNode($"Letter for {name} (EEID: {eeid})");

            // Create child nodes for the EEID, Name, and Note
            TreeNode eeidNode = new TreeNode($"EEID: {eeid}");
            TreeNode nameNode = new TreeNode($"Name: {name}");
            TreeNode noteNode = new TreeNode($"Note: {note}");

            // Add the child nodes to the parent node
            letterNode.Nodes.Add(eeidNode);
            letterNode.Nodes.Add(nameNode);
            letterNode.Nodes.Add(noteNode);

            // Add the parent node to the TreeView
            treeView1.Nodes.Add(letterNode);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "CSV Files|*.csv", // Filter files by extension
                Title = "Select a CSV File"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {

                DateTime currentDate = DateTime.Now;

                string dateString1 = currentDate.ToString("yyyyMMdd");
                string dateString2 = currentDate.ToString("MM/dd/yyyy");

                string mailMergeExcelFileName = $"NewHires format for mailmerge {dateString1}";
                string mailMergeWordfilename = $"Richemont_NH Welcome Letter_Merged {dateString1}";

                string inputFilePath = openFileDialog.FileName;

                string outputExcelFilePath = Path.Combine(@"G:\Account Support Team\Richemont\Communications\Dennis Automated Communications\output\new_hire_mail_merges", mailMergeExcelFileName + ".xlsx");
                string outputWordFilePath = Path.Combine(@"G:\Account Support Team\Richemont\Communications\Dennis Automated Communications\output\nh_letters_generated", mailMergeWordfilename + ".docx");
                string csvFilePath = Path.Combine(@"G:\Account Support Team\Richemont\Communications\Dennis Automated Communications\output\noteload", $"noteload {dateString1}.csv");
                //string csvOutputFilePath = Path.Combine(@"\\WFMLOCAL\shared\Employees\dmcmorrow\Desktop\RiMaiMerg\output\noteload", $"Cleaned noteload {dateString1}.csv");

                Excel.Application excelApp = new Excel.Application();
                excelApp.Visible = true;

                Excel.Workbook inputWorkbook = excelApp.Workbooks.Open(inputFilePath);
                Excel.Worksheet inputWorksheet = inputWorkbook.Sheets[1];

                Excel.Range headerRange = inputWorksheet.UsedRange.Rows[1];
                object[,] headerValues = headerRange.Value;

                Excel.Workbook outputWorkbook = excelApp.Workbooks.Add();
                Excel.Worksheet outputWorksheet = outputWorkbook.Sheets[1];

                Excel.Range outputHeaderRange = outputWorksheet.Range[outputWorksheet.Cells[1, 1], outputWorksheet.Cells[1, headerValues.Length]];
                outputHeaderRange.Value = headerValues;

                outputWorksheet.Cells[1, 1] = "Date of Letter";

                Excel.Range inputDataRange = inputWorksheet.UsedRange.Offset[1, 1];

                inputDataRange.Copy(outputWorksheet.Range["B2"]);

                // Get the last row number in the output worksheet
                int lastRow = outputWorksheet.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

                // Fill the first column with the current date
                for (int i = 2; i <= lastRow; i++)
                {
                    outputWorksheet.Cells[i, 1] = dateString2;
                }

                outputWorkbook.SaveAs(outputExcelFilePath);

                // Define the data for the CSV file
                var data = new List<Record>();

                // Loop over the rows in the output Excel file
                for (int i = 2; i <= lastRow; i++)
                {
                    // Get the EEID from the current row
                    string eeid = outputWorksheet.Cells[i, 2].Value.ToString(); // Assuming EEID is in the second column

                    // Create a record and add it to the list
                    data.Add(new Record { CaseID = "93", EEID = string.Format("=\"{0}\"", eeid.PadLeft(8, '0')), NoteText = "**NH Welcome Letter Mailed**", Private = "" });
                }

                // Create the CSV file noteload
                using (var writer = new StreamWriter(csvFilePath))
                using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
                {
                    // Write the header
                    csv.WriteHeader<Record>();
                    csv.NextRecord();

                    // Write the records
                    csv.WriteRecords(data);
                }

                // Code trims EEID and converts to number for noteload
                /*using (StreamReader sr = new StreamReader(csvFilePath))
                using (StreamWriter sw = new StreamWriter(csvOutputFilePath))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        string[] fields = line.Split(',');

                        // Assuming the EEID is the second field in the CSV
                        if (fields.Length > 1 && fields[1].StartsWith("=\"") && fields[1].EndsWith("\""))
                        {
                            fields[1] = fields[1].Substring(2, fields[1].Length - 3);
                        }

                        sw.WriteLine(string.Join(",", fields));
                    }
                }

                Console.WriteLine("CSV file cleaned successfully.");*/

                Word.Application wordApp = new Word.Application();

                // Open the mail merge template
                Word.Document document = wordApp.Documents.Open(@"G:\Account Support Team\Richemont\Communications\Dennis Automated Communications\input\new_hire_templates\Richemont New Hire letter template.docx");

                // Set up the mail merge
                document.MailMerge.MainDocumentType = Word.WdMailMergeMainDocType.wdFormLetters;
                document.MailMerge.OpenDataSource(outputExcelFilePath);

                // Execute the mail merge
                document.MailMerge.DataSource.ActiveRecord = Word.WdMailMergeActiveRecord.wdFirstRecord;
                document.MailMerge.Execute(Pause: true);

                /*// Generate the filename
                string filename = $"Richemont_NH Welcome Letter_Merged {dateString2}_{document.MailMerge.DataSource.ActiveRecord}";*/

                wordApp.ActiveDocument.SaveAs2(outputWordFilePath);
                wordApp.ActiveDocument.Close();

                inputWorkbook.Close(false);
                outputWorkbook.Close(false);
                document.Close(false);
                excelApp.Quit();
                wordApp.Quit();
                System.Windows.Forms.MessageBox.Show("New hire letters and noteload have generated sucessfully.");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // Split the text from the TextBox into lines
            string[] lines = textBox1.Text.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            string note = "";
            string eeid = textBox2.Text;

            int i = 0;

            // Extract the data from the lines
            string[] nameParts = lines[i].Split(' ');
            string firstName = nameParts[0];
            string lastName = nameParts[1];
            string address = lines[i + 1];
            string address2 = "";
            string[] cityStateZipParts;
            string city;
            string[] stateZipParts;
            string state;
            string zip;
            try
            {
                cityStateZipParts = lines[i + 2].Split(',');
                city = cityStateZipParts[0].Trim();
                stateZipParts = cityStateZipParts[1].Trim().Split(' ');
                state = stateZipParts[0];
                zip = stateZipParts[1];
            }
            catch (Exception ex) // if EE has an address2 aka apartment number
            {
                address2 = lines[i + 2];
                cityStateZipParts = lines[i + 3].Split(',');
                city = cityStateZipParts[0].Trim();
                stateZipParts = cityStateZipParts[1].Trim().Split(' ');
                state = stateZipParts[0];
                zip = stateZipParts[1];
            }

            // Define the path for the output Word document
            string outputFilePath = "";

            string templateFilePath = "";
            if (panel5.Visible == true)
            {
                note = "**CT letter sent " + DateTime.Now.ToString("MM/dd/yyyy") + "**";
                templateFilePath = @"G:\Account Support Team\Richemont\Communications\Dennis Automated Communications\input\new_hire_templates\Richemont CT Letter Template.docx";
            }
            else if (panel6.Visible == true)
            {
                note = "**Insufficient letter sent " + DateTime.Now.ToString("MM/dd/yyyy") + "**";
                templateFilePath = @"G:\Account Support Team\Richemont\Communications\Dennis Automated Communications\input\new_hire_templates\Richemont Insufficient Doc template.docx";
            }

            eeidNotes.Add(new EeidNote { Eeid = eeid, Note = note });
            CreateNoteload1(eeid, firstName + " " + lastName, note);

            // Create Word application object
            Word.Application wordApp = new Word.Application();

            // Open the template
            Word.Document document = wordApp.Documents.Open(templateFilePath);

            // Fill in the fields
            document.FormFields["Date_of_letter"].Result = DateTime.Now.ToString("MM/dd/yyyy");
            document.FormFields["FirstName1"].Result = firstName;
            document.FormFields["LastName1"].Result = lastName;
            document.FormFields["Address"].Result = address;
            document.FormFields["Address2"].Result = address2;
            document.FormFields["City"].Result = city;
            document.FormFields["State"].Result = state;
            document.FormFields["Zip"].Result = zip;
            document.FormFields["FirstName2"].Result = firstName;
            document.FormFields["LastName2"].Result = lastName;
            if (panel5.Visible == true)
            {
                document.FormFields["Date_of_letter_15_1"].Result = DateTime.Now.AddDays(15).ToString("MM/dd/yyyy");
                document.FormFields["Date_of_letter_15_2"].Result = DateTime.Now.AddDays(15).ToString("MM/dd/yyyy");
                outputFilePath = Path.Combine(@"G:\Account Support Team\Richemont\Communications\Dennis Automated Communications\output\ct_letters_generated", $"CT Doc {firstName} {lastName} {DateTime.Now:yyyyMMdd}.docx");
            }
            else if (panel6.Visible == true)
            {
                outputFilePath = Path.Combine(@"G:\Account Support Team\Richemont\Communications\Dennis Automated Communications\output\insufficient_docs_letters_generated", $"{firstName} {lastName} Insufficient Docs {DateTime.Now:yyyyMMdd}.docx");
            }

            // Save the document
            document.SaveAs2(outputFilePath);

            // Close the document
            document.Close();

            // Quit Word application
            wordApp.Quit();
            System.Windows.Forms.MessageBox.Show($"CT letter for {firstName} generated sucessfully and is ready for noteload.");

            // Process each line
            /*for (int i = 0; i < lines.Length; i += 3)
            {
                // Extract the data from the lines
                string[] nameParts = lines[i].Split(' ');
                string firstName = nameParts[0];
                string lastName = nameParts[1];
                string address = lines[i + 1];
                string address2 = "";
                string[] cityStateZipParts;
                string city;
                string[] stateZipParts;
                string state;
                string zip;
                try
                {
                    cityStateZipParts = lines[i + 2].Split(',');
                    city = cityStateZipParts[0].Trim();
                    stateZipParts = cityStateZipParts[1].Trim().Split(' ');
                    state = stateZipParts[0];
                    zip = stateZipParts[1];
                }
                catch (Exception ex) // if EE has an address2 aka apartment number
                {
                    address2 = lines[i + 2];
                    cityStateZipParts = lines[i + 3].Split(',');
                    city = cityStateZipParts[0].Trim();
                    stateZipParts = cityStateZipParts[1].Trim().Split(' ');
                    state = stateZipParts[0];
                    zip = stateZipParts[1];
                }

                // Define the path for the output Word document
                string outputFilePath;

                string templateFilePath;
                if (radioButton1.Checked)
                {
                    note = "CT letter sent " + DateTime.Now.ToString("MM/dd/yyyy");
                    templateFilePath = @"\\WFMLOCAL\shared\Employees\dmcmorrow\Desktop\RiMaiMerg\input\new_hire_templates\Richemont CT Letter Template.docx";
                }
                else if (radioButton2.Checked)
                {
                    note = "Insufficient letter sent " + DateTime.Now.ToString("MM/dd/yyyy");
                    templateFilePath = @"\\WFMLOCAL\shared\Employees\dmcmorrow\Desktop\RiMaiMerg\input\new_hire_templates\Richemont Insufficient Doc template.docx";
                }
                else
                {
                    continue;
                }

                eeidNotes.Add(new EeidNote { Eeid = eeid, Note = note });

                // Create Word application object
                Word.Application wordApp = new Word.Application();

                // Open the template
                Word.Document document = wordApp.Documents.Open(templateFilePath);

                // Fill in the fields
                document.FormFields["Date_of_letter"].Result = DateTime.Now.ToString("MM/dd/yyyy");
                document.FormFields["FirstName1"].Result = firstName;
                document.FormFields["LastName1"].Result = lastName;
                document.FormFields["Address"].Result = address;
                document.FormFields["Address2"].Result = address2;
                document.FormFields["City"].Result = city;
                document.FormFields["State"].Result = state;
                document.FormFields["Zip"].Result = zip;
                document.FormFields["FirstName2"].Result = firstName;
                document.FormFields["LastName2"].Result = lastName;
                if (radioButton1.Checked)
                {
                    document.FormFields["Date_of_letter_15_1"].Result = DateTime.Now.AddDays(15).ToString("MM/dd/yyyy");
                    document.FormFields["Date_of_letter_15_2"].Result = DateTime.Now.AddDays(15).ToString("MM/dd/yyyy");
                    outputFilePath = Path.Combine(@"\\WFMLOCAL\shared\Employees\dmcmorrow\Desktop\RiMaiMerg\output\ct_letters_generated", $"CT Doc {firstName} {lastName} {DateTime.Now:yyyyMMdd}.docx");
                }
                else if (radioButton2.Checked)
                {
                    outputFilePath = Path.Combine(@"\\WFMLOCAL\shared\Employees\dmcmorrow\Desktop\RiMaiMerg\output\insufficient_docs_letters_generated", $"{firstName} {lastName} Insufficient Docs {DateTime.Now:yyyyMMdd}.docx");
                }
                else
                {
                    // No radio button is selected, so skip this iteration
                    continue;
                }

                // Save the document
                document.SaveAs2(outputFilePath);

                // Close the document
                document.Close();

                // Quit Word application
                wordApp.Quit();
            }*/
        }

        private void button4_Click(object sender, EventArgs e)
        {
            panel4.Visible = true;
            panel5.Visible = false;
            panel6.Visible = false;
            panel7.Visible = false;
            panel8.Visible = false;
            panel2.Height = button4.Height;
            panel2.Top = button4.Top;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            panel8.Visible = true;
            panel4.Visible = false;
            panel5.Visible = true;
            panel6.Visible = false;
            panel7.Visible = false;
            panel8.Visible = true;
            panel2.Height = button5.Height;
            panel2.Top = button5.Top;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            panel4.Visible = false;
            panel5.Visible = false;
            panel6.Visible = true;
            panel7.Visible = false;
            panel8.Visible = true;
            panel2.Height = button6.Height;
            panel2.Top = button6.Top;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            panel4.Visible = false;
            panel5.Visible = false;
            panel6.Visible = false;
            panel7.Visible = true;
            panel8.Visible = true;
            panel2.Height = button9.Height;
            panel2.Top = button9.Top;
        }

        private void button8_Click_1(object sender, EventArgs e)
        {
            // Split the text from the TextBox into lines
            string[] lines = textBox4.Text.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            string note = "";
            string eeid = textBox3.Text;

            int i = 0;

            // Extract the data from the lines
            string[] nameParts = lines[i].Split(' ');
            string firstName = nameParts[0];
            string lastName = nameParts[1];
            string address = lines[i + 1];
            string address2 = "";
            string[] cityStateZipParts;
            string city;
            string[] stateZipParts;
            string state;
            string zip;
            try
            {
                cityStateZipParts = lines[i + 2].Split(',');
                city = cityStateZipParts[0].Trim();
                stateZipParts = cityStateZipParts[1].Trim().Split(' ');
                state = stateZipParts[0];
                zip = stateZipParts[1];
            }
            catch (Exception ex) // if EE has an address2 aka apartment number
            {
                address2 = lines[i + 2];
                cityStateZipParts = lines[i + 3].Split(',');
                city = cityStateZipParts[0].Trim();
                stateZipParts = cityStateZipParts[1].Trim().Split(' ');
                state = stateZipParts[0];
                zip = stateZipParts[1];
            }

            // Define the path for the output Word document
            string outputFilePath = "";

            string templateFilePath = "";
            if (panel5.Visible == true)
            {
                note = "**CT letter sent " + DateTime.Now.ToString("MM/dd/yyyy") + "**";
                templateFilePath = @"G:\Account Support Team\Richemont\Communications\Dennis Automated Communications\input\new_hire_templates\Richemont CT Letter Template.docx";
            }
            else if (panel6.Visible == true)
            {
                note = "**Insufficient letter sent " + DateTime.Now.ToString("MM/dd/yyyy") + "**";
                templateFilePath = @"G:\Account Support Team\Richemont\Communications\Dennis Automated Communications\input\new_hire_templates\Richemont Insufficient Doc template.docx";
            }

            eeidNotes.Add(new EeidNote { Eeid = eeid, Note = note });
            CreateNoteload1(eeid, firstName + " " + lastName, note);

            // Create Word application object
            Word.Application wordApp = new Word.Application();

            // Open the template
            Word.Document document = wordApp.Documents.Open(templateFilePath);

            // Fill in the fields
            document.FormFields["Date_of_letter"].Result = DateTime.Now.ToString("MM/dd/yyyy");
            document.FormFields["FirstName1"].Result = firstName;
            document.FormFields["LastName1"].Result = lastName;
            document.FormFields["Address"].Result = address;
            document.FormFields["Address2"].Result = address2;
            document.FormFields["City"].Result = city;
            document.FormFields["State"].Result = state;
            document.FormFields["Zip"].Result = zip;
            document.FormFields["FirstName2"].Result = firstName;
            document.FormFields["LastName2"].Result = lastName;
            if (panel5.Visible == true)
            {
                document.FormFields["Date_of_letter_15_1"].Result = DateTime.Now.AddDays(15).ToString("MM/dd/yyyy");
                document.FormFields["Date_of_letter_15_2"].Result = DateTime.Now.AddDays(15).ToString("MM/dd/yyyy");
                outputFilePath = Path.Combine(@"G:\Account Support Team\Richemont\Communications\Dennis Automated Communications\output\ct_letters_generated", $"CT Doc {firstName} {lastName} {DateTime.Now:yyyyMMdd}.docx");
            }
            else if (panel6.Visible == true)
            {
                outputFilePath = Path.Combine(@"G:\Account Support Team\Richemont\Communications\Dennis Automated Communications\output\insufficient_docs_letters_generated", $"{firstName} {lastName} Insufficient Docs {DateTime.Now:yyyyMMdd}.docx");
            }

            // Save the document
            document.SaveAs2(outputFilePath);

            // Close the document
            document.Close();

            // Quit Word application
            wordApp.Quit();
            System.Windows.Forms.MessageBox.Show($"Insufficent document letter for {firstName} generated sucessfully and is ready for noteload.");

            // Process each line
            /*for (int i = 0; i < lines.Length; i += 3)
            {
                // Extract the data from the lines
                string[] nameParts = lines[i].Split(' ');
                string firstName = nameParts[0];
                string lastName = nameParts[1];
                string address = lines[i + 1];
                string address2 = "";
                string[] cityStateZipParts;
                string city;
                string[] stateZipParts;
                string state;
                string zip;
                try
                {
                    cityStateZipParts = lines[i + 2].Split(',');
                    city = cityStateZipParts[0].Trim();
                    stateZipParts = cityStateZipParts[1].Trim().Split(' ');
                    state = stateZipParts[0];
                    zip = stateZipParts[1];
                }
                catch (Exception ex) // if EE has an address2 aka apartment number
                {
                    address2 = lines[i + 2];
                    cityStateZipParts = lines[i + 3].Split(',');
                    city = cityStateZipParts[0].Trim();
                    stateZipParts = cityStateZipParts[1].Trim().Split(' ');
                    state = stateZipParts[0];
                    zip = stateZipParts[1];
                }

                // Define the path for the output Word document
                string outputFilePath;

                string templateFilePath;
                if (radioButton1.Checked)
                {
                    note = "CT letter sent " + DateTime.Now.ToString("MM/dd/yyyy");
                    templateFilePath = @"\\WFMLOCAL\shared\Employees\dmcmorrow\Desktop\RiMaiMerg\input\new_hire_templates\Richemont CT Letter Template.docx";
                }
                else if (radioButton2.Checked)
                {
                    note = "Insufficient letter sent " + DateTime.Now.ToString("MM/dd/yyyy");
                    templateFilePath = @"\\WFMLOCAL\shared\Employees\dmcmorrow\Desktop\RiMaiMerg\input\new_hire_templates\Richemont Insufficient Doc template.docx";
                }
                else
                {
                    continue;
                }

                eeidNotes.Add(new EeidNote { Eeid = eeid, Note = note });

                // Create Word application object
                Word.Application wordApp = new Word.Application();

                // Open the template
                Word.Document document = wordApp.Documents.Open(templateFilePath);

                // Fill in the fields
                document.FormFields["Date_of_letter"].Result = DateTime.Now.ToString("MM/dd/yyyy");
                document.FormFields["FirstName1"].Result = firstName;
                document.FormFields["LastName1"].Result = lastName;
                document.FormFields["Address"].Result = address;
                document.FormFields["Address2"].Result = address2;
                document.FormFields["City"].Result = city;
                document.FormFields["State"].Result = state;
                document.FormFields["Zip"].Result = zip;
                document.FormFields["FirstName2"].Result = firstName;
                document.FormFields["LastName2"].Result = lastName;
                if (radioButton1.Checked)
                {
                    document.FormFields["Date_of_letter_15_1"].Result = DateTime.Now.AddDays(15).ToString("MM/dd/yyyy");
                    document.FormFields["Date_of_letter_15_2"].Result = DateTime.Now.AddDays(15).ToString("MM/dd/yyyy");
                    outputFilePath = Path.Combine(@"\\WFMLOCAL\shared\Employees\dmcmorrow\Desktop\RiMaiMerg\output\ct_letters_generated", $"CT Doc {firstName} {lastName} {DateTime.Now:yyyyMMdd}.docx");
                }
                else if (radioButton2.Checked)
                {
                    outputFilePath = Path.Combine(@"\\WFMLOCAL\shared\Employees\dmcmorrow\Desktop\RiMaiMerg\output\insufficient_docs_letters_generated", $"{firstName} {lastName} Insufficient Docs {DateTime.Now:yyyyMMdd}.docx");
                }
                else
                {
                    // No radio button is selected, so skip this iteration
                    continue;
                }

                // Save the document
                document.SaveAs2(outputFilePath);

                // Close the document
                document.Close();

                // Quit Word application
                wordApp.Quit();
            }*/
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            DateTime currentDate = DateTime.Now;
            string dateString1 = currentDate.ToString("yyyyMMdd");
            string noteLoadFilePath = Path.Combine(@"G:\Account Support Team\Richemont\Communications\Dennis Automated Communications\output\noteload\", $"CT Insufficient noteload {dateString1}.csv");

            using (var writer = new StreamWriter(noteLoadFilePath))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                csv.WriteHeader<Record>();
                csv.NextRecord();

                // Iterate through the top-level nodes in the TreeView
                foreach (TreeNode letterNode in treeView1.Nodes)
                {
                    // Access the text of the parent node (the letter node)
                    string letterText = letterNode.Text;
                    var record = new Record();
                    record.CaseID = "93";
                    int counter = 0;
                    // Iterate through the child nodes of the letter node
                    foreach (TreeNode childNode in letterNode.Nodes)
                    {
                        // Access the text of the child node
                        string childText = childNode.Text;
                        string value = childText.Split(':')[1].Trim();
                        if (counter == 0)
                        {
                            value = string.Format("=\"{0}\"", value.PadLeft(8, '0'));
                            record.EEID = value;
                        }
                        else if (counter == 1)
                        {

                        }
                        else if (counter == 2)
                        {
                            record.NoteText = value;
                        }

                        counter++;

                        // Do something with the value...
                    }
                    record.Private = "";
                    csv.WriteRecord(record);
                    csv.NextRecord();
                }
            }
            treeView1.Nodes.Clear();
            System.Windows.Forms.MessageBox.Show("CT letter, insufficient document letter, canceled transaction note load has generated successfully.");
        }

        private void button10_Click(object sender, EventArgs e)
        {
            string currentDate = DateTime.Now.ToString("yyyyMMdd");
            // Get data from the controls
            string eeid = textBox5.Text;
            string name = textBox6.Text;
            string email = textBox7.Text;
            DateTime date = dateTimePicker1.Value;

            // Get selected values from the radio buttons
            string group1Selection = radioButton1.Checked ? "LOC" : radioButton2.Checked ? "GOC" : radioButton3.Checked ? "Marriage" : radioButton4.Checked ? "Birth/Adoption" : 
                                     radioButton5.Checked ? "Divorce" : radioButton6.Checked ? "Death of Spouse" : radioButton7.Checked ? "Death of Child" : radioButton8.Checked ? "Medicare Change" : "";
            string group2Selection = radioButton9.Checked ? "Insufficient documents" : radioButton10.Checked ? "no docs received" : "";

            string note = $"**Canceled Transaction email sent {currentDate}**";

            // Create an EmployeeData object and add it to the list
            employeeDataList.Add(new EmployeeData
            {
                EEID = eeid,
                Name = name,
                Email = email,
                Date = date,
                Group1Selection = group1Selection,
                Group2Selection = group2Selection
            });
            CreateNoteload1(eeid, name, note);
            System.Windows.Forms.MessageBox.Show("Employee has been added to the cancel transaction list");
        }

        private void button11_Click(object sender, EventArgs e)
        {
            string dateString1 = DateTime.Now.ToString("yyyyMMdd");
            string canceledTransactionsFileName = $"Richemont Cancel Life Event {dateString1}";
            string canceledTransactionsFilePath = Path.Combine(@"G:\Account Support Team\Richemont\Communications\Dennis Automated Communications\output\canceled_transactions_generated", canceledTransactionsFileName + ".xlsx");

            // Create a new Excel application
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.Sheets[1];

            // Add headers
            worksheet.Cells[1, 1].Value = "EE ID";
            worksheet.Cells[1, 2].Value = "EE Name";
            worksheet.Cells[1, 3].Value = "LE Type";
            worksheet.Cells[1, 4].Value = "Date of Completion";
            worksheet.Cells[1, 5].Value = "Drop Letter Date";
            worksheet.Cells[1, 6].Value = "Docs sent - sufficient or no";
            worksheet.Cells[1, 7].Value = "Email";

            // Add data from the list to the worksheet
            for (int i = 0; i < employeeDataList.Count; i++)
            {
                worksheet.Cells[i + 2, 1].Value = string.Format("=\"{0}\"", employeeDataList[i].EEID.PadLeft(8, '0'));
                worksheet.Cells[i + 2, 2].Value = employeeDataList[i].Name;
                worksheet.Cells[i + 2, 3].Value = employeeDataList[i].Group1Selection;
                worksheet.Cells[i + 2, 4].Value = employeeDataList[i].Date.ToShortDateString();
                worksheet.Cells[i + 2, 5].Value = employeeDataList[i].Date.AddDays(15).ToShortDateString();
                worksheet.Cells[i + 2, 6].Value = employeeDataList[i].Group2Selection;
                worksheet.Cells[i + 2, 7].Value = employeeDataList[i].Email;
            }

            // Save the workbook to a file
            workbook.SaveAs(canceledTransactionsFilePath);

            // Cleanup
            workbook.Close();
            excelApp.Quit();
            System.Windows.Forms.MessageBox.Show("Cancel transaction spreadsheet has been created successfully");
        }

        private void webBrowser2_DocumentCompleted_1(object sender, WebBrowserDocumentCompletedEventArgs e)
        {

        }
    }
}

