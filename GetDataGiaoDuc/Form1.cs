using GetDataGiaoDuc.APISMAS;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.ServiceModel.Security;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GetDataGiaoDuc
{
    public partial class Form1 : Form
    {
        SchoolProfile[] arrSchool;
        
        public Form1()
        {
            InitializeComponent();

            //Khởi tạo dữ liệu cho Combobox huyện
            comboBoxDistinctStudent.Items.Add(new Item("A Lưới", 1));
            comboBoxDistinctStudent.Items.Add(new Item("Nam Đông", 2));
            comboBoxDistinctStudent.Items.Add(new Item("Phú Lộc", 3));
            comboBoxDistinctStudent.Items.Add(new Item("Phú Vang", 1));
            comboBoxDistinctStudent.Items.Add(new Item("Hương Thủy", 1));
            comboBoxDistinctStudent.Items.Add(new Item("Thành phố Huế", 1));
            comboBoxDistinctStudent.Items.Add(new Item("Quảng Điền", 1));
            comboBoxDistinctStudent.Items.Add(new Item("Hương Trà", 1));
            comboBoxDistinctStudent.Items.Add(new Item("Phong Điền", 1));
            comboBoxDistinctStudent.Items.Add(new Item("Toàn tỉnh", 1));
            //Khởi tạo dữ liệu cho Combobox huyện
            comboBoxDistinctEmployee.Items.Add(new Item("A Lưới", 1));
            comboBoxDistinctEmployee.Items.Add(new Item("Nam Đông", 1));
            comboBoxDistinctEmployee.Items.Add(new Item("Phú Lộc", 1));
            comboBoxDistinctEmployee.Items.Add(new Item("Phú Vang", 1));
            comboBoxDistinctEmployee.Items.Add(new Item("Hương Thủy", 1));
            comboBoxDistinctEmployee.Items.Add(new Item("Thành phố Huế", 1));
            comboBoxDistinctEmployee.Items.Add(new Item("Quảng Điền", 1));
            comboBoxDistinctEmployee.Items.Add(new Item("Hương Trà", 1));
            comboBoxDistinctEmployee.Items.Add(new Item("Phong Điền", 1));
            comboBoxDistinctEmployee.Items.Add(new Item("Toàn tỉnh", 1));


            
        }
       
        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void btnGetSchool_Click(object sender, EventArgs e)
        {

            GiaoDucClient client = new GiaoDucClient();
            client.ClientCredentials.UserName.UserName = "test";
            client.ClientCredentials.UserName.Password = "test123";
            client.ClientCredentials.ServiceCertificate.Authentication.CertificateValidationMode  = X509CertificateValidationMode.None;
            BindingList<SchoolProfile> sPList = new BindingList<SchoolProfile>();

            try
            {
                arrSchool = client.GetSchoolProfile();
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
        private class Item
        {
            public String name;
            public int value;
            public Item(String name, int value)
            {
                this.name = name;
                this.value = value;
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
                                    ws.Cells[index, 2] = sf.SchoolCode;
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



    }
}
