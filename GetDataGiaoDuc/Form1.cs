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
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GetDataGiaoDuc
{
    public partial class Form1 : Form
    {
        public SchoolProfile[] arrSchool;
        //public PupilProfile[] arrStudent;
        public List<PupilProfile> listStudent;
        public Employee[] arrEmployee;
        public List<Item> arrDistinctName = new List<Item>();
        public List<Item> arrSchoolName = new List<Item>();
        public String fileDanhSachTruong = @"\DanhSachTruong.xls";
        private String username = "test";
        private String password = "test123";

        public Form1()
        {
            InitializeComponent();
            //Khởi tạo các dữ liệu cho UI
            UIInit();



        }

        private void Form1_Load(object sender, EventArgs e)
        {



        }

        private void UIInit()
        {
            listStudent = new List<PupilProfile>();

            //khởi tạo dữ liệu cho comboBox Huyện
            //arrDistinctName = new List<Item>();
            arrDistinctName.Add(new Item("A Lưới", 1));
            arrDistinctName.Add(new Item("Hương Thủy", 2));
            arrDistinctName.Add(new Item("Hương Trà", 3));
            arrDistinctName.Add(new Item("Thành phố Huế", 4));
            arrDistinctName.Add(new Item("Nam Đông", 5));
            arrDistinctName.Add(new Item("Phú Lộc", 6));
            arrDistinctName.Add(new Item("Phú Vang", 7));
            arrDistinctName.Add(new Item("Phong Điền", 8));
            arrDistinctName.Add(new Item("Quảng Điền", 9));
            arrDistinctName.Add(new Item("Toàn tỉnh", 10));

            //đẩy dữ liệu vào combobox Huyện.
            foreach (Item i in arrDistinctName)
            {
                comboBoxDistinctEmployee.Items.Add(i);
                comboBoxDistinctStudent.Items.Add(i);

            }
            //thiết lập giá trị mặc định cho combobox Huyện
            comboBoxDistinctEmployee.SelectedIndex = 9;
            comboBoxDistinctStudent.SelectedIndex = 9;

            //thiết lập giá trị cho combobox trường
            //arrSchoolName = new List<Item>();
            //Lấy dữ liệu trường từ file excel
            getDataFromFile(Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName + fileDanhSachTruong);
            foreach (Item i in arrSchoolName)
            {
                comboBoxSchoolEmployee.Items.Add(i);
                comboBoxSchoolStudent.Items.Add(i);
            }


        }

        public void getDataFromFile(String filePath)
        {
            try {
                // create the Application object we can use in the member functions.
                Microsoft.Office.Interop.Excel.Application _excelApp = new Microsoft.Office.Interop.Excel.Application();
                _excelApp.Visible = false;
                //open the workbook
                Workbook workbook = _excelApp.Workbooks.Open(filePath, ReadOnly: true);


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


                    Item i = new Item(valueArray[row, 5].ToString(), int.Parse(valueArray[row, 2].ToString()), int.Parse(valueArray[row, 4].ToString()));
                    arrSchoolName.Add(i);
                    //Debug.Print(""+i.value+"-----"+i.name+"-----"+i.parent);
                    /*
                    for (int col = 2; col <= worksheet.UsedRange.Columns.Count; ++col)
                    {
                        //access each cell
                    
                        Debug.Print(valueArray[row, col].ToString());
                    */
                }

                //clean up stuffs
                workbook.Close(false, Type.Missing, Type.Missing);
                _excelApp.Quit();
            }
            catch (Exception es)
            {
                MessageBox.Show(es.ToString(), "Lỗi khi đọc dữ liệu trường từ file", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnGetSchool_Click(object sender, EventArgs e)
        {
            GiaoDucClient client = new GiaoDucClient();
            client.ClientCredentials.UserName.UserName = this.username;
            client.ClientCredentials.UserName.Password = this.password;
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
                MessageBox.Show(es.ToString(), "Lỗi kết nối đến máy chủ", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            if (arrSchool != null)
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
            //listStudent.Clear();
            //dataGridViewStudent.Rows.Clear();
            //dataGridViewStudent.Refresh();

            GiaoDucClient client = new GiaoDucClient();
            client.ClientCredentials.UserName.UserName = this.username;
            client.ClientCredentials.UserName.Password = this.password;
            client.ClientCredentials.ServiceCertificate.Authentication.CertificateValidationMode = X509CertificateValidationMode.None;


            BindingList<PupilProfile> listStudentBiding = new BindingList<PupilProfile>();

            try
            {
                Item itemSelect = comboBoxSchoolStudent.SelectedItem as Item;
                PupilProfile[] arrStudent = client.GetPupilProfile(itemSelect.value, 2017, 2000, 1);
                client.Close();

                labelSumStudent.Text = "Tổng số học sinh: " + arrStudent.Length;
                for (int i = 0; i < arrStudent.Length; i++)
                {

                    Console.WriteLine("\n" + arrStudent[i].PupilCode);
                    listStudent.Add(arrStudent[i]);
                }


                listStudent.Sort(new NameComparer());
                //listStudentBiding = listStudent;
                //labelNumberSchool.Text = "Tổng số trường: " + arrSchool.Length;
                pupilProfileBindingSource.DataSource = listStudent;
                MessageBox.Show("Dkm thành công !!!!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            catch (NullReferenceException)
            {
                MessageBox.Show("Chọn trường trong danh sách ", "Lỗi con mẹ nó rồi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception es)
            {
                MessageBox.Show(es.ToString(), "Lỗi kết nối đến máy chủ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnGetDataEmployee_Click(object sender, EventArgs e)
        {
            GiaoDucClient client = new GiaoDucClient();
            client.ClientCredentials.UserName.UserName = this.username;
            client.ClientCredentials.UserName.Password = this.password;
            client.ClientCredentials.ServiceCertificate.Authentication.CertificateValidationMode = X509CertificateValidationMode.None;

            BindingList<Employee> eList = new BindingList<Employee>();

            try
            {
                //lấy truongID từ combobox
                Item item = (Item)comboBoxSchoolEmployee.SelectedItem;
                //Debug.Print(item.value.ToString());
                arrEmployee = client.GetEmployee(item.value);


                client.Close();
                for (int i = 0; i < arrEmployee.Length; i++)
                {
                    eList.Add(arrEmployee[i]);
                }
                //eList.ToList<Employee>().Sort(new Nam);
                labelNumberSchool.Text = "Tổng số trường: " + arrSchool.Length;
                employeeBindingSource.DataSource = eList;

            }
            catch (NullReferenceException)
            {
                MessageBox.Show("Chọn trường trong danh sách ", "Lỗi con mẹ nó rồi", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

            GiaoDucClient client = new GiaoDucClient();
            client.ClientCredentials.UserName.UserName = this.username;
            client.ClientCredentials.UserName.Password = this.password;
            client.ClientCredentials.ServiceCertificate.Authentication.CertificateValidationMode = X509CertificateValidationMode.None;

            ClassProfile[] cProfile;

            try
            {
                Item i = comboBoxSchoolStudent.SelectedItem as Item;
                cProfile = client.GetClassProfile(i.value, 2017);
                //for(int j=0; j<cProfile.Length; j++)
                //Debug.Print(cProfile[j].ClassProfileId+"-------------"+cProfile[j].ClassName);

                //if (!listStudent.Any())
              //  {

                    using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel workbook|*.xls" })
                    {
                        sfd.FileName = comboBoxSchoolStudent.Text;
                        if (sfd.ShowDialog() == DialogResult.OK)
                        {
                            String fileName = sfd.FileName;
                            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                            Workbook wb = excel.Workbooks.Add(XlSheetType.xlWorksheet);
                            Worksheet ws = excel.ActiveSheet;
                            //ws.EnableSelection = Microsoft.Office.Interop.Excel.XlEnableSelection.xlNoSelection;
                            ws.UsedRange.NumberFormat = "@";
                            
                            excel.Visible = false;
                            //Đặt tên cột trong file Excel
                            ws.Cells[1, 1] = "Số thứ tự";
                            ws.Cells[1, 2] = "Mã học sinh";
                            ws.Cells[1, 3] = "Họ và tên";
                            ws.Cells[1, 4] = "Giới tính";
                            ws.Cells[1, 5] = "Lớp";
                            ws.Cells[1, 6] = "Ngày sinh";
                            ws.Cells[1, 7] = "Nơi sinh";
                            ws.Cells[1, 8] = "Tỉnh thành";
                            ws.Cells[1, 9] = "Địa chỉ";
                            ws.Cells[1, 10] = "Dân tộc";

                            int index = 1;
                            foreach (PupilProfile pf in listStudent)
                            {
                                index++;

                                ws.Cells[index, 1] = index - 1;
                                ws.Cells[index, 2] = pf.PupilCode;
                                ws.Cells[index, 3] = pf.FullName;
                                ws.Cells[index, 4] = pf.Genre;
                                String classname = "";                        
                                for (int k = 0; k < cProfile.Length; k++)

                                    if (pf.ClassID == cProfile[k].ClassProfileId)
                                    {
                                        classname = cProfile[k].ClassName;
                                        //Debug.Print("Tìm được tên lớp: " + classname);
                                        break;
                                        //else ws.Cells[index, 7] = "Looix";
                                    }
                                ws.Cells[index, 5].value = "'"+classname;
                                ws.Cells[index, 6] = pf.BirthPlace;
                                ws.Cells[index, 7] = pf.BirthDate;
                                ws.Cells[index, 8] = pf.Province;
                                ws.Cells[index, 9] = pf.Area;
                                ws.Cells[index, 10] = pf.Ethnic;
                                

                            }
                            
                            //Lưu file
                           // Thread th = new Thread( ()=>{
                                ws.SaveAs(fileName, XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, true, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing);
                                excel.Quit();
                          //  });
                           // th.Start();
                          //  if(th.IsAlive == false) { MessageBox.Show("Xuất danh sách học sinh thành công", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                            
                        }
                    }

              //  }
              //  else MessageBox.Show("Danh sách trường hiện tại đang trống, vui lòng lấy danh sách từ cổng thông tin", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);

             }
            catch(Exception es)
            {
                MessageBox.Show("Không lấy được danh sách lớp, vui lòng kiểm tra lại\n"+es.ToString(), "Lỗi CMNR!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
} 

        private void comboBoxDistinctEmployee_SelectedIndexChanged(object sender, EventArgs e)
        {

            //Xóa hết dữ liệu trường trong combobox
            comboBoxSchoolEmployee.Text = "";
            comboBoxSchoolEmployee.Items.Clear();
            ComboBox cb = sender as ComboBox;

            Item itemSelected = (Item)cb.SelectedItem;
            //Debug.Print(itemSelected.name);

            //Kiểm tra có phải là dữ liệu toàn tỉnh hay không?
            if (itemSelected.value == 10)
            {
                foreach (Item i in arrSchoolName) comboBoxSchoolEmployee.Items.Add(i);
            }

            else
                foreach (Item i in arrSchoolName)
                {
                    if (i.parent == itemSelected.value) comboBoxSchoolEmployee.Items.Add(i);
                }
           
        }

        private void comboBoxDistinctStudent_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Xóa hết dữ liệu trường trong combobox
            comboBoxSchoolStudent.Text = "";
            comboBoxSchoolStudent.Items.Clear();

            ComboBox cb = sender as ComboBox;

            Item itemSelected = (Item)cb.SelectedItem;
            //Debug.Print(itemSelected.name);
            List<Item> listSchoolItem = new List<Item>();
            //Kiểm tra có phải là dữ liệu toàn tỉnh hay không?
            if (itemSelected.value == 10)
            {
                foreach (Item i in arrSchoolName) listSchoolItem.Add(i);
            }

            else
                foreach (Item i in arrSchoolName)
                {
                    if (i.parent == itemSelected.value) listSchoolItem.Add(i);
                }
            AutoCompleteStringCollection acsc = new AutoCompleteStringCollection();
         
            comboBoxSchoolStudent.Items.AddRange(listSchoolItem.ToArray());


        }

        private void comboBoxSchoolStudent_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
