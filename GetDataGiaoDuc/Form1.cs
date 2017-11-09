using GetDataGiaoDuc.APISMAS;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.ServiceModel.Security;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GetDataGiaoDuc
{
    public partial class Form1 : Form
    {
        public SchoolProfile[] arrSchool;
        public PupilProfile[] arrStudent;
        public Employee[] arrEmployee;
        public List<Item> arrDistinctName;
        public List<Item> arrSchoolName;
        public String fileDanhSachTruong = @"\DanhSachTruong.xls";

        public Form1()
        {
            InitializeComponent();
            //Khởi tạo các dữ liệu cho UI
            UIInit();



        }
       
        private void Form1_Load(object sender, EventArgs e)
        {
            arrDistinctName = new List<Item>();
            arrSchoolName = new List<Item>();

            
        }

        private void UIInit()
        {
            //Lấy dữ liệu trường từ file excel
            getDataFromFile(Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName+fileDanhSachTruong);


        }

        public void getDataFromFile(String filePath)
        {
            try { 
                // create the Application object we can use in the member functions.
                Microsoft.Office.Interop.Excel.Application _excelApp = new Microsoft.Office.Interop.Excel.Application();
                _excelApp.Visible = true;
                //open the workbook
                Workbook workbook = _excelApp.Workbooks.Open(filePath,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);

                //select the first sheet        
                Worksheet worksheet = (Worksheet)workbook.Worksheets[1];

                //find the used range in worksheet
                Range excelRange = worksheet.UsedRange;

                //get an object array of all of the cells in the worksheet (their values)
                object[,] valueArray = (object[,])excelRange.get_Value(
                            XlRangeValueDataType.xlRangeValueDefault);

                //access the cells
                for (int row = 2; row <= worksheet.UsedRange.Rows.Count; ++row)
                {
                    for (int col = 2; col <= worksheet.UsedRange.Columns.Count; ++col)
                    {
                        //access each cell
                    
                        Debug.Print(valueArray[row, col].ToString());
                    }
                }

                //clean up stuffs
                workbook.Close(false, Type.Missing, Type.Missing);
                _excelApp.Quit();
                }
            catch(Exception es)
            {
                MessageBox.Show(es.ToString(), "Lỗi khi đọc dữ liệu trường từ file", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnGetSchool_Click(object sender, EventArgs e)
        {
            GiaoDucClient client = new GiaoDucClient();
            client.ClientCredentials.UserName.UserName = "test";
            client.ClientCredentials.UserName.Password = "test123";
            client.ClientCredentials.ServiceCertificate.Authentication.CertificateValidationMode = X509CertificateValidationMode.None;

            BindingList<SchoolProfile> sPList = new BindingList<SchoolProfile>();

            try
            {
                arrSchool = client.GetSchoolProfile();
                client.Close();
                for (int i = 0; i < arrSchool.Length; i++)
                {
                    sPList.Add(arrSchool[i]);
                }

            labelNumberSchool.Text = "Tổng số trường: " + arrSchool.Length;
            schoolProfileBindingSource.DataSource = sPList;
                
            }
            catch (Exception es)
            {
                MessageBox.Show(es.ToString(), "Lỗi kết nối đến máy chủ", MessageBoxButtons.OK,MessageBoxIcon.Error);
            }

            
        }
        public class Item
        {
            public String name;
            public int value;
            public int parent;
            public Item(String name, int value)
            {
                this.name = name;
                this.value = value;
            }
            public Item(String name, int value, int parent)
            {
                this.name = name;
                this.value = value;
                this.parent = parent;
            }
            public override string ToString()
            {
                // Generates the text shown in the combo box
                return name;
            }
        }

        private void btnExportShoolExcel_Click(object sender, EventArgs e)
        {
            if(arrSchool != null)
            {

                    using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel workbook|*.xls" })
                    {
                        if (sfd.ShowDialog() == DialogResult.OK)
                        {
                            String fileName = sfd.FileName;
                            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                            Workbook wb = excel.Workbooks.Add(XlSheetType.xlWorksheet);
                            Worksheet ws = excel.ActiveSheet;
                            ws.EnableSelection = Microsoft.Office.Interop.Excel.XlEnableSelection.xlNoSelection;
                            excel.Visible = false;
                            //Đặt tên cột trong file Excel
                            ws.Cells[1, 1] = "Số thứ tự";
                            ws.Cells[1, 2] = "Mã trường";
                            ws.Cells[1, 3] = "Cấp học";
                            ws.Cells[1, 4] = "Mã huyện";
                            ws.Cells[1, 5] = "Tên trường";
                            ws.Cells[1, 6] = "User name";

                            int index = 1;
                            foreach (SchoolProfile sf in arrSchool)
                                {
                                    index++;
                                    ws.Cells[index, 1] = index - 1;
                                    ws.Cells[index, 2] = sf.SchoolProfileID;
                                    ws.Cells[index, 3] = sf.EducationGrade.ToString();
                                    ws.Cells[index, 4] = sf.District;
                                    ws.Cells[index, 5] = sf.SchoolName;
                                    ws.Cells[index, 6] = sf.UserName;
                                }
                            //Lưu file
                            ws.SaveAs(fileName, XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, true, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing);
                            excel.Quit();
                            }
                        }
                
                    }
            else MessageBox.Show("Danh sách trường hiện tại đang trống, vui lòng lấy danh sách từ cổng thông tin", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void btnGetDataStudent_Click(object sender, EventArgs e)
        {
            GiaoDucClient client = new GiaoDucClient();
            client.ClientCredentials.UserName.UserName = "test";
            client.ClientCredentials.UserName.Password = "test123";
            client.ClientCredentials.ServiceCertificate.Authentication.CertificateValidationMode = X509CertificateValidationMode.None;


            BindingList<PupilProfile> studentList = new BindingList<PupilProfile>();

            try
            {
                arrStudent= client.GetPupilProfile(int.Parse(comboBoxSchoolStudent.Text), 2017,2000,1);
                client.Close();
                for (int i = 0; i < arrStudent.Length; i++)
                {

                    Console.WriteLine("\n"+arrStudent[i].PupilCode);
                    studentList.Add(arrStudent[i]);
                }

                //labelNumberSchool.Text = "Tổng số trường: " + arrSchool.Length;
                pupilProfileBindingSource.DataSource = studentList;
                MessageBox.Show(arrStudent[2].FullName, "Kết quả", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception es)
            {
                MessageBox.Show(es.ToString(), "Lỗi kết nối đến máy chủ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnGetDataEmployee_Click(object sender, EventArgs e)
        {
            GiaoDucClient client = new GiaoDucClient();
            client.ClientCredentials.UserName.UserName = "test";
            client.ClientCredentials.UserName.Password = "test123";
            client.ClientCredentials.ServiceCertificate.Authentication.CertificateValidationMode = X509CertificateValidationMode.None;

            BindingList<Employee> eList = new BindingList<Employee>();

            try
            {
                arrEmployee = client.GetEmployee(int.Parse(comboBoxSchoolEmployee.Text));
                client.Close();
                for (int i = 0; i < arrEmployee.Length; i++)
                {
                    eList.Add(arrEmployee[i]);
                }

                labelNumberSchool.Text = "Tổng số trường: " + arrSchool.Length;
                employeeBindingSource.DataSource = eList;

            }
            catch (Exception es)
            {
                MessageBox.Show(es.ToString(), "Lỗi kết nối đến máy chủ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridViewStudent_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void btnExportStudent_Click(object sender, EventArgs e)
        {
            if (arrStudent != null)
            {

                using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel workbook|*.xls" })
                {
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        String fileName = sfd.FileName;
                        Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                        Workbook wb = excel.Workbooks.Add(XlSheetType.xlWorksheet);
                        Worksheet ws = excel.ActiveSheet;
                        ws.EnableSelection = Microsoft.Office.Interop.Excel.XlEnableSelection.xlNoSelection;
                        excel.Visible = false;
                        //Đặt tên cột trong file Excel
                        ws.Cells[1, 1] = "Số thứ tự";
                        ws.Cells[1, 2] = "Mã học sinh";
                        ws.Cells[1, 3] = "Họ và tên";
                        ws.Cells[1, 4] = "Ngày sinh";
                        ws.Cells[1, 5] = "Nơi sinh";
                        ws.Cells[1, 6] = "Dân tộc";
                        ws.Cells[1, 7] = "Class ID";

                        int index = 1;
                        foreach (PupilProfile pf in arrStudent)
                        {
                            index++;
                            ws.Cells[index, 1] = index - 1;
                            ws.Cells[index, 2] = pf.PupilCode;
                            ws.Cells[index, 3] = pf.FullName;
                            ws.Cells[index, 4] = pf.BirthDate;
                            ws.Cells[index, 5] = pf.BirthPlace;
                            ws.Cells[index, 6] = pf.Ethnic;
                            ws.Cells[index, 7] = pf.ClassID;
                        }
                        //Lưu file
                        ws.SaveAs(fileName, XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, true, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing);
                        excel.Quit();
                    }
                }

            }
            else MessageBox.Show("Danh sách trường hiện tại đang trống, vui lòng lấy danh sách từ cổng thông tin", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void comboBoxDistinctEmployee_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
