using CsvHelper;
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
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ToolBar;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace RichemontMailMerge
{
    public partial class Form1 : Form
    {
        private List<EeidNote> eeidNotes = new List<EeidNote>();
        private List<EmployeeData> employeeDataList = new List<EmployeeData>();
        private List<System.Windows.Forms.Button> managedButtons = new List<System.Windows.Forms.Button>();
        private List<string> clients = new List<string> { "Please select a client", "Richemont", "Sunland", "Primetals", "Caromont" };

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

        public Form1()
        {
            InitializeComponent();

            InitializeManagedButtons();
            InitializeComboBox();
            InitializePlaceholderTexts();
            InitializePanelVisibility();
            InitializeButtonVisibility();

        }

        private void InitializeManagedButtons()
        {
            managedButtons.Add(button4);
            managedButtons.Add(button5);
            managedButtons.Add(button6);
            managedButtons.Add(button7);
            managedButtons.Add(button9);
        }

        private void InitializeComboBox()
        {
            comboBox1.DataSource = clients;
        }

        private void InitializePlaceholderTexts()
        {
            SetupPlaceholderText(textBox1, "Enter employee information");
            SetupPlaceholderText(textBox2, "Enter EEID");
            SetupPlaceholderText(textBox4, "Enter employee information");
            SetupPlaceholderText(textBox3, "Enter EEID");
            SetupPlaceholderText(textBox5, "Enter EEID");
            SetupPlaceholderText(textBox6, "Enter EE name");
            SetupPlaceholderText(textBox7, "Enter EE email address");
        }

        private void InitializePanelVisibility()
        {
            Panel[] panels = { panel4, panel5, panel6, panel7, panel8, panel12, panel13 };
            foreach (var panel in panels)
            {
                panel.Visible = false;
            }
        }

        private void InitializeButtonVisibility()
        {
            button12.Visible = false;
            button7.Visible = false;

            button4.Visible = false;
            button5.Visible = false;
            button6.Visible = false;
            button9.Visible = false;
        }

        private void SetupForRichemont()
        {
            panel11.Visible = false;
            panel4.Visible = true;
            panel12.Visible = false;
            panel5.Visible = false;
            panel6.Visible = false;
            panel7.Visible = false;
            panel8.Visible = false;
            panel2.Height = button4.Height;
            panel2.Top = button4.Top;

            button12.Visible = false;
            button7.Visible = false;

            button4.Visible = true;
            button5.Visible = true;
            button6.Visible = true;
            button9.Visible = true;
        }

        private void SetupForSunland()
        {
            panel11.Visible = false;
            panel12.Visible = true;
            panel4.Visible = false;
            panel5.Visible = false;
            panel6.Visible = false;
            panel7.Visible = false;
            panel8.Visible = false;
            button4.Visible = false;
            button5.Visible = false;
            button6.Visible = false;
            button9.Visible = false;

            button12.Visible = true;
            button7.Visible = true;

            button4.Visible = false;
            button5.Visible = false;
            button6.Visible = false;
            button9.Visible = false;
        }

        private void SetupForPrimetals()
        {

        }

        private void SetupForCaromont()
        {

        }

        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e) // Method that handles when user selects a client
        {
            if (comboBox1.SelectedIndex == 0)
            {
                return;
            }
            string selectedClient = comboBox1.SelectedItem.ToString();
            
            MessageBox.Show($"You selected: {selectedClient}");

            if (selectedClient == "Richemont")
            {
                SetupForRichemont();
            }
            if (selectedClient == "Sunland")
            {
                SetupForSunland();
            }
            if (selectedClient == "Primetals")
            {
                SetupForPrimetals();
            }
            if (selectedClient == "Caromont")
            {
                SetupForCaromont();
            }
        }

        private void SetupPlaceholderText(System.Windows.Forms.TextBox textBox, string placeholderText) // Method for setting placeholder text
        {
            textBox.Text = placeholderText;

            textBox.GotFocus += (sender, e) =>
            {
                if (textBox.Text == placeholderText)
                {
                    textBox.Text = "";
                }
            };

            textBox.LostFocus += (sender, e) =>
            {
                if (string.IsNullOrEmpty(textBox.Text))
                {
                    textBox.Text = placeholderText;
                }
            };
        }

        private void CreateNoteload1(string eeid, string name, string note) // Method used to create Treenode display for generating a noteload
        {
            TreeNode letterNode = new TreeNode($"Letter for {name} (EEID: {eeid})");

            TreeNode eeidNode = new TreeNode($"EEID: {eeid}");
            TreeNode nameNode = new TreeNode($"Name: {name}");
            TreeNode noteNode = new TreeNode($"Note: {note}");

            letterNode.Nodes.Add(eeidNode);
            letterNode.Nodes.Add(nameNode);
            letterNode.Nodes.Add(noteNode);

            treeView1.Nodes.Add(letterNode);
        }

        private void button1_Click(object sender, EventArgs e) // Button for Richemont New Hire Letters
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "CSV Files|*.csv",
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

                string basePath = @"G:\Account Support Team\Richemont\Communications\Dennis Automated Communications";
                string outputExcelFilePath = Path.Combine(basePath, "output", "new_hire_mail_merges", mailMergeExcelFileName + ".xlsx");
                string outputWordFilePath = Path.Combine(basePath, "output", "nh_letters_generated", mailMergeWordfilename + ".docx");
                string csvFilePath = Path.Combine(basePath, "output", "noteload", $"noteload {dateString1}.csv");

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

                int lastRow = outputWorksheet.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

                for (int i = 2; i <= lastRow; i++)
                {
                    outputWorksheet.Cells[i, 1] = dateString2;
                }

                outputWorkbook.SaveAs(outputExcelFilePath);

                var data = new List<Record>();
                for (int i = 2; i <= lastRow; i++)
                {
                    string eeid = outputWorksheet.Cells[i, 2].Value.ToString();
                    data.Add(new Record { CaseID = "93", EEID = string.Format("=\"{0}\"", eeid.PadLeft(8, '0')), NoteText = "**NH Welcome Letter Mailed**", Private = "" });
                }

                using (var writer = new StreamWriter(csvFilePath))
                using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
                {
                    csv.WriteHeader<Record>();
                    csv.NextRecord();
                    csv.WriteRecords(data);
                }

                Word.Application wordApp = new Word.Application();
                Word.Document document = wordApp.Documents.Open(Path.Combine(basePath, "input", "new_hire_templates", "Richemont New Hire letter template.docx"));
                document.MailMerge.MainDocumentType = Word.WdMailMergeMainDocType.wdFormLetters;
                document.MailMerge.OpenDataSource(outputExcelFilePath);
                document.MailMerge.DataSource.ActiveRecord = Word.WdMailMergeActiveRecord.wdFirstRecord;
                document.MailMerge.Execute(Pause: true);
                wordApp.ActiveDocument.SaveAs2(outputWordFilePath);
                wordApp.ActiveDocument.Close();

                inputWorkbook.Close(false);
                outputWorkbook.Close(false);
                document.Close(false);
                excelApp.Quit();
                wordApp.Quit();

                System.Windows.Forms.MessageBox.Show("New hire letters and noteload have generated successfully.");
            }
        }

        private void button2_Click(object sender, EventArgs e) // Button for Richemont Completed Transactions
        {
            string[] lines = textBox1.Text.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            string note = "";
            string eeid = textBox2.Text;

            int i = 0;

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
            catch (Exception ex)
            {
                address2 = lines[i + 2];
                cityStateZipParts = lines[i + 3].Split(',');
                city = cityStateZipParts[0].Trim();
                stateZipParts = cityStateZipParts[1].Trim().Split(' ');
                state = stateZipParts[0];
                zip = stateZipParts[1];
            }

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

            Word.Application wordApp = new Word.Application();
            Word.Document document = wordApp.Documents.Open(templateFilePath);

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

            document.SaveAs2(outputFilePath);
            document.Close();
            wordApp.Quit();

            System.Windows.Forms.MessageBox.Show($"CT letter for {firstName} generated successfully and is ready for noteload.");
        }

        private void button3_Click_1(object sender, EventArgs e) // Button for Richemont noteload Completed Transactions and Insufficient Documents generation
        {
            DateTime currentDate = DateTime.Now;
            string dateString1 = currentDate.ToString("yyyyMMdd");
            string noteLoadFilePath = Path.Combine(@"G:\Account Support Team\Richemont\Communications\Dennis Automated Communications\output\noteload\", $"CT Insufficient noteload {dateString1}.csv");

            using (var writer = new StreamWriter(noteLoadFilePath))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                csv.WriteHeader<Record>();
                csv.NextRecord();

                foreach (TreeNode letterNode in treeView1.Nodes)
                {
                    string letterText = letterNode.Text;
                    var record = new Record
                    {
                        CaseID = "93"
                    };
                    int counter = 0;

                    foreach (TreeNode childNode in letterNode.Nodes)
                    {
                        string childText = childNode.Text;
                        string value = childText.Split(':')[1].Trim();
                        if (counter == 0)
                        {
                            value = string.Format("=\"{0}\"", value.PadLeft(8, '0'));
                            record.EEID = value;
                        }
                        else if (counter == 2)
                        {
                            record.NoteText = value;
                        }

                        counter++;
                    }
                    record.Private = "";
                    csv.WriteRecord(record);
                    csv.NextRecord();
                }
            }
            treeView1.Nodes.Clear();
            System.Windows.Forms.MessageBox.Show("CT letter, insufficient document letter, canceled transaction note load has generated successfully.");
        }

        private void button4_Click(object sender, EventArgs e) // Button to display Panel4, used for Richemont New Hire Letters
        {
            panel11.Visible = false;
            panel4.Visible = true;
            panel5.Visible = false;
            panel6.Visible = false;
            panel7.Visible = false;
            panel8.Visible = false;
            panel2.Height = button4.Height;
            panel2.Top = button4.Top;
        }

        private void button5_Click(object sender, EventArgs e) // Button to display Panel5, used for Richemont Completed Transactions
        {
            panel11.Visible = false;
            panel8.Visible = true;
            panel4.Visible = false;
            panel5.Visible = true;
            panel6.Visible = false;
            panel7.Visible = false;
            panel8.Visible = true;
            panel2.Height = button5.Height;
            panel2.Top = button5.Top;
        }

        private void button6_Click(object sender, EventArgs e) // Button to display Panel6, used for Richemont Insufficient Documents
        {
            panel11.Visible = false;
            panel4.Visible = false;
            panel5.Visible = false;
            panel6.Visible = true;
            panel7.Visible = false;
            panel8.Visible = true;
            panel2.Height = button6.Height;
            panel2.Top = button6.Top;
        }

        private void button9_Click(object sender, EventArgs e) // Button to display Panel7, used for Richemont Canceled Life Events
        {
            panel11.Visible = false;
            panel4.Visible = false;
            panel5.Visible = false;
            panel6.Visible = false;
            panel7.Visible = true;
            panel8.Visible = true;
            panel2.Height = button9.Height;
            panel2.Top = button9.Top;
        }

        private void button7_Click(object sender, EventArgs e) // Button to display Panel12, used Sunland New Hire emails
        {
            panel12.Visible = true;
            panel13.Visible = false;
            panel2.Height = button7.Height;
            panel2.Top = button7.Top;
        }

        private void button12_Click(object sender, EventArgs e) // Button to display Panel13, used Sunland New Hire Final emails
        {
            panel12.Visible = false;
            panel13.Visible = true;
            panel2.Height = button12.Height;
            panel2.Top = button12.Top;
        }

        private void button8_Click_1(object sender, EventArgs e) // Button for Richemont Insufficient Documents
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

        private void button10_Click(object sender, EventArgs e) // Button for adding Richemont EE to Canceled Life events 
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

        private void button11_Click(object sender, EventArgs e) // Button for creating Richemont Canceled Life events spreadsheet
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

        private void pictureBox1_Click(object sender, EventArgs e) // Button to go to Client Communication Manager home screen
        {
            panel11.Visible = true;
            panel12.Visible = false;
            panel13.Visible = false;
            panel8.Visible = false;
            panel4.Visible = false;
            panel5.Visible = false;
            panel6.Visible = false;
            panel7.Visible = false;
            panel8.Visible = false;

            button12.Visible = false;
            button7.Visible = false;

            button4.Visible = false;
            button5.Visible = false;
            button6.Visible = false;
            button9.Visible = false;

            panel2.Height = button4.Height;
            panel2.Top = button4.Top;
        }

        private void button13_Click(object sender, EventArgs e) // Button for creating Sunland New Hire emails
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "CSV Files|*.csv",
                Title = "Select a CSV File"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                DateTime currentDate = DateTime.Now;
                string dateString1 = currentDate.ToString("yyyyMMdd");

                string mailMergeExcelFileName = $"NHWE.{dateString1}.sunland";
                string inputFilePath = openFileDialog.FileName;

                string basePath = @"G:\Account Support Team\Sunland Asphalt\Communications\Dennis Automated Communications";
                string outputExcelFilePath = Path.Combine(basePath, "output", "new_hire_emails_generated", mailMergeExcelFileName + ".xlsx");
                string csvFilePath = Path.Combine(basePath, "output", "noteload", $"new hire emails noteload {dateString1}.csv");

                Excel.Application excelApp = new Excel.Application();
                excelApp.Visible = true;

                Excel.Workbook inputWorkbook = excelApp.Workbooks.Open(inputFilePath);
                Excel.Worksheet inputWorksheet = inputWorkbook.Sheets[1];
                Excel.Range usedRange = inputWorksheet.UsedRange;

                List<int> rowsToCopy = new List<int>();
                List<string> columnsToKeep = new List<string> { "EmployeeId", "FirstName", "LastName", "WorkEmail", "PersonalEmail" };
                List<int> columnIndicesToKeep = new List<int>();
                int enrollmentStatusColumnIndex = 0;

                for (int col = 1; col <= usedRange.Columns.Count; col++)
                {
                    string header = ((Excel.Range)usedRange.Cells[1, col]).Value.ToString();
                    if (columnsToKeep.Contains(header))
                    {
                        columnIndicesToKeep.Add(col);
                    }
                    if (header == "EnrollmentStatus")
                    {
                        enrollmentStatusColumnIndex = col;
                    }
                }

                for (int row = 2; row <= usedRange.Rows.Count; row++)
                {
                    if (int.TryParse(((Excel.Range)usedRange.Cells[row, 1]).Value.ToString(), out int daysRemaining) && daysRemaining > 14)
                    {
                        string enrollmentStatus = ((Excel.Range)usedRange.Cells[row, enrollmentStatusColumnIndex]).Value.ToString();
                        if (enrollmentStatus != "Canceled")
                        {
                            rowsToCopy.Add(row);
                        }
                    }
                }

                Excel.Workbook outputWorkbook = excelApp.Workbooks.Add();
                Excel.Worksheet outputWorksheet = outputWorkbook.Sheets[1];

                // Copy headers
                for (int i = 0; i < columnIndicesToKeep.Count; i++)
                {
                    outputWorksheet.Cells[1, i + 1].Value = ((Excel.Range)usedRange.Cells[1, columnIndicesToKeep[i]]).Value;
                }

                // Copy rows
                int outputRow = 2;
                foreach (int row in rowsToCopy)
                {
                    for (int i = 0; i < columnIndicesToKeep.Count; i++)
                    {
                        outputWorksheet.Cells[outputRow, i + 1].Value = ((Excel.Range)usedRange.Cells[row, columnIndicesToKeep[i]]).Value;
                    }
                    outputRow++;
                }

                outputWorkbook.SaveAs(outputExcelFilePath);

                var data = new List<Record>();
                for (int i = 2; i < outputRow; i++)
                {
                    string eeid = outputWorksheet.Cells[i, 1].Value.ToString();
                    data.Add(new Record { CaseID = "194", EEID = eeid, NoteText = "**NH Emails Sent**", Private = "" });
                }

                using (var writer = new StreamWriter(csvFilePath))
                using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
                {
                    csv.WriteHeader<Record>();
                    csv.NextRecord();
                    csv.WriteRecords(data);
                }

                inputWorkbook.Close(false);
                outputWorkbook.Close(false);
                excelApp.Quit();

                System.Windows.Forms.MessageBox.Show("New hire emails and noteload have generated successfully.");
            }
        }

        private void button14_Click(object sender, EventArgs e) // Button for creating Sunland New Hire Final emails
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "CSV Files|*.csv",
                Title = "Select a CSV File"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                DateTime currentDate = DateTime.Now;
                string dateString1 = currentDate.ToString("yyyyMMdd");

                string mailMergeExcelFileName = $"NHFRE.{dateString1}.sunland";
                string inputFilePath = openFileDialog.FileName;

                string basePath = @"G:\Account Support Team\Sunland Asphalt\Communications\Dennis Automated Communications";
                string outputExcelFilePath = Path.Combine(basePath, "output", "new_hire_final_emails_generated", mailMergeExcelFileName + ".xlsx");
                string csvFilePath = Path.Combine(basePath, "output", "noteload", $"new hire final emails noteload {dateString1}.csv");

                Excel.Application excelApp = new Excel.Application();
                excelApp.Visible = true;

                Excel.Workbook inputWorkbook = excelApp.Workbooks.Open(inputFilePath);
                Excel.Worksheet inputWorksheet = inputWorkbook.Sheets[1];
                Excel.Range usedRange = inputWorksheet.UsedRange;

                List<int> rowsToCopy = new List<int>();
                List<string> columnsToKeep = new List<string> { "EmployeeId", "FirstName", "LastName", "WorkEmail", "PersonalEmail" };
                List<int> columnIndicesToKeep = new List<int>();
                int enrollmentStatusColumnIndex = 0;

                for (int col = 1; col <= usedRange.Columns.Count; col++)
                {
                    string header = ((Excel.Range)usedRange.Cells[1, col]).Value.ToString();
                    if (columnsToKeep.Contains(header))
                    {
                        columnIndicesToKeep.Add(col);
                    }
                    if (header == "EnrollmentStatus")
                    {
                        enrollmentStatusColumnIndex = col;
                    }
                }

                for (int row = 2; row <= usedRange.Rows.Count; row++)
                {
                    if (int.TryParse(((Excel.Range)usedRange.Cells[row, 1]).Value.ToString(), out int daysRemaining) && daysRemaining <= 14)
                    {
                        string enrollmentStatus = ((Excel.Range)usedRange.Cells[row, enrollmentStatusColumnIndex]).Value.ToString();
                        if (enrollmentStatus != "Canceled")
                        {
                            rowsToCopy.Add(row);
                        }
                    }
                }

                Excel.Workbook outputWorkbook = excelApp.Workbooks.Add();
                Excel.Worksheet outputWorksheet = outputWorkbook.Sheets[1];

                // Copy headers
                for (int i = 0; i < columnIndicesToKeep.Count; i++)
                {
                    outputWorksheet.Cells[1, i + 1].Value = ((Excel.Range)usedRange.Cells[1, columnIndicesToKeep[i]]).Value;
                }

                // Copy rows
                int outputRow = 2;
                foreach (int row in rowsToCopy)
                {
                    for (int i = 0; i < columnIndicesToKeep.Count; i++)
                    {
                        outputWorksheet.Cells[outputRow, i + 1].Value = ((Excel.Range)usedRange.Cells[row, columnIndicesToKeep[i]]).Value;
                    }
                    outputRow++;
                }

                outputWorkbook.SaveAs(outputExcelFilePath);

                var data = new List<Record>();
                for (int i = 2; i < outputRow; i++)
                {
                    string eeid = outputWorksheet.Cells[i, 1].Value.ToString();
                    data.Add(new Record { CaseID = "194", EEID = eeid, NoteText = "**NH Final Emails Sent**", Private = "" });
                }

                using (var writer = new StreamWriter(csvFilePath))
                using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
                {
                    csv.WriteHeader<Record>();
                    csv.NextRecord();
                    csv.WriteRecords(data);
                }

                inputWorkbook.Close(false);
                outputWorkbook.Close(false);
                excelApp.Quit();

                System.Windows.Forms.MessageBox.Show("New hire finals emails and noteload have generated successfully.");
            }
        }

        private void webBrowser2_DocumentCompleted_1(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            // Your code here
        }

    }
}

