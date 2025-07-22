using ExcelDataReader;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace TESTWEBBOOK
{
    internal class ExcelDataProvider
    {
        private static DataTable _excelDataTable;
        private static string _filePath = @"C:\Users\84395\Desktop\ExcelData.xlsx";
        private static string _sheetName = "DEMO"; 
        //private static int rowStart = 3; // Dòng bắt đầu (F3)
        //private static int rowEnd = 8;   // Dòng kết thúc (F8)

        public static DataTable ReadExcel(string filePath)
        {
            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();
                    return result.Tables[0]; // Trả về sheet đầu tiên
                }
            }
        }
        // gộp ô
        private static string ReadMergedCell(string filePath, int row, int column)
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[_sheetName];
                if (worksheet == null)
                    throw new Exception($"Không tìm thấy sheet '{_sheetName}' trong file Excel!");

                var cell = worksheet.Cells[row, column];

                if (cell.Merge)
                {
                    var mergedRange = worksheet.MergedCells[row, column];
                    return worksheet.Cells[mergedRange].Text;
                }
                return cell.Text;
            }
        }
        // ghi kết quả thực tế và trạng thái
        public static void WriteResultToExcel(int row, string actualResult, string status)
        {
            using (var package = new ExcelPackage(new FileInfo(_filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[_sheetName];
                if (worksheet == null)
                    throw new Exception($"Không tìm thấy sheet '{_sheetName}' trong file Excel!");

                worksheet.Cells[row, 8].Value = actualResult;
                worksheet.Cells[row, 9].Value = status;

                package.Save();
            }
        }
        static string GetContentAfterColon(string input)
        {
            if (string.IsNullOrEmpty(input)) return string.Empty;

            // Tìm vị trí dấu `:`
            int index = input.IndexOf(':');
            if (index != -1)
            {
                // Cắt lấy phần sau dấu `:`
                return input.Substring(index + 1).Trim();
            }

            return input.Trim(); // Nếu không có dấu `:`, giữ nguyên chuỗi
        }
        // Thêm sách
        public static IEnumerable<TestCaseData> GetDataFromEx_TestAddBook()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance); // Đăng ký mã hóa
            string filePath = @"C:\Users\84395\Desktop\ExcelData.xlsx"; // Đường dẫn file Excel
            var testData = new List<TestCaseData>();
            DataTable excelDataTable = ReadExcel(filePath);

            if (excelDataTable.Rows.Count < 9)
            {
                throw new Exception("Dữ liệu trong file Excel không đủ dòng!");
            }

            var tenSach = GetContentAfterColon(excelDataTable.Rows[1][4]?.ToString()); // Hàng 2, Cột E
            string giaBan = GetContentAfterColon(excelDataTable.Rows[2][4]?.ToString())?.Trim();
            var moTa = GetContentAfterColon(excelDataTable.Rows[3][4]?.ToString()); // Hàng 4, Cột E
            var anh = GetContentAfterColon(excelDataTable.Rows[4][4]?.ToString()); // Hàng 5, Cột E
            var ngayCapNhat = GetContentAfterColon(excelDataTable.Rows[5][4]?.ToString()); // Hàng 6, Cột E
            int soLuongInt;

            if (!int.TryParse(GetContentAfterColon(excelDataTable.Rows[6][4]?.ToString()), out soLuongInt))
            {
                throw new Exception("Lỗi: Số lượng không hợp lệ! Hãy kiểm tra file Excel.");
            }

            var theLoai = GetContentAfterColon(excelDataTable.Rows[7][4]?.ToString()); // Hàng 8, Cột E
            var nhaXuatBan = GetContentAfterColon(excelDataTable.Rows[8][4]?.ToString()); // Hàng 9, Cột E

            string expectedResult = excelDataTable.Rows[1][6]?.ToString().Trim();

            // Chỉ thêm test case 1 lần, giữ nguyên kiểu int cho số lượng
            if (!string.IsNullOrEmpty(tenSach))
            {
                testData.Add(new TestCaseData(
                    tenSach,
                    giaBan ?? "",  // Tránh null
                    moTa ?? "",
                    anh ?? "",
                    ngayCapNhat ?? "",
                    soLuongInt,  // ❌ Không cần ToString()
                    theLoai ?? "",
                    nhaXuatBan ?? "",
                    expectedResult ?? ""
                ));
            }

            return testData;
        }

        // thêm loại
        public static IEnumerable<TestCaseData> GetDataFromExcel_TestLoai()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            string filePath = @"C:\Users\84395\Desktop\ExcelData.xlsx";

            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("Không tìm thấy file Excel tại đường dẫn: " + filePath);
            }

            var testData = new List<TestCaseData>();
            DataTable excelDataTable = ExcelDataProvider.ReadExcel(filePath);

            if (excelDataTable.Rows.Count < 10) 
            {
                throw new Exception("Dữ liệu trong file Excel không đủ dòng!");
            }

            int rowIndex = 9;
            string tenLoaiSach = GetContentAfterColon(excelDataTable.Rows[rowIndex][4]?.ToString());
            string expectedResult = excelDataTable.Rows[rowIndex][6]?.ToString().Trim();

            if (!string.IsNullOrEmpty(tenLoaiSach) && !string.IsNullOrEmpty(expectedResult))
            {
                testData.Add(new TestCaseData(tenLoaiSach, expectedResult, rowIndex));
            }
            return testData;
        }
        public static IEnumerable<TestCaseData> GetDataFromExcel_TestAddNXB()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            string filePath = @"C:\Users\84395\Desktop\ExcelData.xlsx";

            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("Không tìm thấy file Excel tại đường dẫn: " + filePath);
            }

            var testData = new List<TestCaseData>();
            DataTable excelDataTable = ExcelDataProvider.ReadExcel(filePath);

            if (excelDataTable.Rows.Count < 18)
            {
                throw new Exception("Dữ liệu trong file Excel không đủ dòng!");
            }

            int rowIndex = 14; // Chỉ số dòng bắt đầu của AddNXB
            string tenNXB = GetContentAfterColon(excelDataTable.Rows[rowIndex][4]?.ToString());
            string diaChi = GetContentAfterColon(excelDataTable.Rows[rowIndex + 1][4]?.ToString()); // Đọc từ E16
            string soDienThoai = GetContentAfterColon(excelDataTable.Rows[rowIndex + 2][4]?.ToString()); // Đọc từ E17

            string expectedResult = excelDataTable.Rows[rowIndex][6]?.ToString().Trim();

            if (!string.IsNullOrEmpty(tenNXB) && !string.IsNullOrEmpty(diaChi) && !string.IsNullOrEmpty(soDienThoai) && !string.IsNullOrEmpty(expectedResult))
            {
                testData.Add(new TestCaseData(tenNXB, diaChi, soDienThoai, expectedResult, rowIndex));
            }

            return testData;
        }
        public static IEnumerable<TestCaseData> GetDataFromExcel_TestSearchBook()
        {
            string filePath = @"C:\Users\84395\Desktop\ExcelData.xlsx";
            DataTable excelData = ReadExcel(filePath);
            var testData = new List<TestCaseData>();

            for (int i = 18; i < excelData.Rows.Count && testData.Count < 1; i++) // Bắt đầu từ hàng 19 (index 18)
            {
                string tuKhoaTimKiem = GetContentAfterColon(excelData.Rows[i][4]?.ToString().Trim()); // Cột E (index 4)
                string expectedResult = GetContentAfterColon(excelData.Rows[i][6]?.ToString().Trim()); // Cột G (index 6)

                if (!string.IsNullOrEmpty(tuKhoaTimKiem) && !string.IsNullOrEmpty(expectedResult))
                {
                    testData.Add(new TestCaseData(tuKhoaTimKiem, expectedResult, i)); // Thêm dữ liệu test
                }
            }
            return testData;
        }

        // đăng ký
        public static IEnumerable<TestCaseData> GetDataFromExcel_TestDangKy()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            string filePath = @"C:\Users\84395\Desktop\ExcelData.xlsx";
            DataTable excelData = ReadExcel(filePath);
            var testData = new List<TestCaseData>();

            
                string hoTen = GetContentAfterColon(excelData.Rows[24][4]?.ToString().Trim());
                string tenDN = GetContentAfterColon(excelData.Rows[25][4]?.ToString().Trim());
                string matKhau = GetContentAfterColon(excelData.Rows[26][4]?.ToString().Trim());
                string matKhauNhapLai = GetContentAfterColon(excelData.Rows[27][4]?.ToString().Trim());
                string email = GetContentAfterColon(excelData.Rows[28][4]?.ToString().Trim());
                string diaChi = GetContentAfterColon(excelData.Rows[29][4]?.ToString().Trim());
                string soDienThoai = GetContentAfterColon(excelData.Rows[30][4]?.ToString().Trim());
                string ngaySinh = GetContentAfterColon(excelData.Rows[31][4]?.ToString().Trim());
                string expectedResult = excelData.Rows[24][6]?.ToString().Trim();

                if (!string.IsNullOrEmpty(tenDN) && !string.IsNullOrEmpty(expectedResult))
                {
                    testData.Add(new TestCaseData(hoTen, tenDN, matKhau, matKhauNhapLai, email, diaChi, soDienThoai, ngaySinh, expectedResult));
                }
            
            return testData;
        }
        // đăng ký sai định dạng mail
        public static IEnumerable<TestCaseData> GetDataFromExcel_TestDangKySaiDinhDangMail()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            string filePath = @"C:\Users\84395\Desktop\ExcelData.xlsx";
            DataTable excelData = ReadExcel(filePath);
            var testData = new List<TestCaseData>();


            string hoTen = GetContentAfterColon(excelData.Rows[32][4]?.ToString().Trim());
            string tenDN = GetContentAfterColon(excelData.Rows[33][4]?.ToString().Trim());
            string matKhau = GetContentAfterColon(excelData.Rows[34][4]?.ToString().Trim());
            string matKhauNhapLai = GetContentAfterColon(excelData.Rows[35][4]?.ToString().Trim());
            string email = GetContentAfterColon(excelData.Rows[36][4]?.ToString().Trim());
            string diaChi = GetContentAfterColon(excelData.Rows[37][4]?.ToString().Trim());
            string soDienThoai = GetContentAfterColon(excelData.Rows[38][4]?.ToString().Trim());
            string ngaySinh = GetContentAfterColon(excelData.Rows[39][4]?.ToString().Trim());
            string expectedResult = excelData.Rows[32][6]?.ToString().Trim();

            if (!string.IsNullOrEmpty(tenDN) && !string.IsNullOrEmpty(expectedResult))
            {
                testData.Add(new TestCaseData(hoTen, tenDN, matKhau, matKhauNhapLai, email, diaChi, soDienThoai, ngaySinh, expectedResult));
            }

            return testData;
        }
        // đăng ký bỏ trống trường mật khẩu
        public static IEnumerable<TestCaseData> GetDataFromExcel_TestBoTrongMK()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            string filePath = @"C:\Users\84395\Desktop\ExcelData.xlsx";
            DataTable excelData = ReadExcel(filePath);
            var testData = new List<TestCaseData>();


            string hoTen = GetContentAfterColon(excelData.Rows[40][4]?.ToString().Trim());
            string tenDN = GetContentAfterColon(excelData.Rows[41][4]?.ToString().Trim());
            string matKhau = GetContentAfterColon(excelData.Rows[42][4]?.ToString().Trim());
            string matKhauNhapLai = GetContentAfterColon(excelData.Rows[43][4]?.ToString().Trim());
            string email = GetContentAfterColon(excelData.Rows[44][4]?.ToString().Trim());
            string diaChi = GetContentAfterColon(excelData.Rows[45][4]?.ToString().Trim());
            string soDienThoai = GetContentAfterColon(excelData.Rows[46][4]?.ToString().Trim());
            string ngaySinh = GetContentAfterColon(excelData.Rows[47][4]?.ToString().Trim());
            string expectedResult = excelData.Rows[40][6]?.ToString().Trim();

            if (!string.IsNullOrEmpty(tenDN) && !string.IsNullOrEmpty(expectedResult))
            {
                testData.Add(new TestCaseData(hoTen, tenDN, matKhau, matKhauNhapLai, email, diaChi, soDienThoai, ngaySinh, expectedResult));
            }

            return testData;
        }
        // đăng nhập thành công
        public static IEnumerable<TestCaseData> GetDataFromExcel_TestDangNhap()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            string filePath = @"C:\Users\84395\Desktop\ExcelData.xlsx";
            DataTable excelData = ReadExcel(filePath);
            var testData = new List<TestCaseData>();


            string tenDN = GetContentAfterColon(excelData.Rows[48][4]?.ToString().Trim());
            string matKhau = GetContentAfterColon(excelData.Rows[49][4]?.ToString().Trim());
            string expectedResult = excelData.Rows[48][6]?.ToString().Trim();

            if (!string.IsNullOrEmpty(tenDN) && !string.IsNullOrEmpty(expectedResult))
            {
                testData.Add(new TestCaseData(tenDN, matKhau, expectedResult));
            }
            return testData;
        }
        public static IEnumerable<TestCaseData> GetDataFromExcel_TestDangNhapInvalid()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            string filePath = @"C:\Users\84395\Desktop\ExcelData.xlsx";
            var testData = new List<TestCaseData>();

            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("Không tìm thấy file Excel tại: " + filePath);
            }

            try
            {
                DataTable excelData = ReadExcel(filePath);
                if (excelData == null || excelData.Rows.Count < 54)
                {
                    throw new Exception("File Excel không đủ dữ liệu (ít hơn 54 dòng).");
                }

                string tenDN = GetContentAfterColon(excelData.Rows[52][4]?.ToString()?.Trim()) ?? "";
                string matKhau = GetContentAfterColon(excelData.Rows[53][4]?.ToString()?.Trim()) ?? "";
                string expectedResult = excelData.Rows[52][6]?.ToString()?.Trim();

                Console.WriteLine($"tenDN: '{tenDN}', matKhau: '{matKhau}', expectedResult: '{expectedResult}'");

                if (!string.IsNullOrEmpty(expectedResult))
                {
                    testData.Add(new TestCaseData(tenDN, matKhau, expectedResult));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Lỗi khi đọc file Excel: " + ex.Message);
                throw;
            }

            Console.WriteLine($"Số test case: {testData.Count}");
            return testData;
        }
        // xem chi tiết sp (GUI)
        public static IEnumerable<TestCaseData> GetDataFromExcel_GUIBookTitle()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            string filePath = @"C:\Users\84395\Desktop\ExcelData.xlsx";
            DataTable excelData = ReadExcel(filePath);
            var testData = new List<TestCaseData>();

            // Đọc dữ liệu từ Excel
            string Tensach = GetContentAfterColon(excelData.Rows[57][4]?.ToString()?.Trim());  // Không cần GetContentAfterColon nếu không có dấu ':'
            string expectedResult = excelData.Rows[56][6]?.ToString()?.Trim(); // Chỉnh lại dòng từ 56 -> 58 để khớp với Tensach

            // Kiểm tra nếu dữ liệu hợp lệ thì thêm vào danh sách test case
            if (!string.IsNullOrEmpty(Tensach) && !string.IsNullOrEmpty(expectedResult))
            {
                testData.Add(new TestCaseData(Tensach, expectedResult));
            }

            return testData;
        }
        // verify nxb
        public static IEnumerable<TestCaseData> GetDataFromExcel_TenNXBNull()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            string filePath = @"C:\Users\84395\Desktop\ExcelData.xlsx";
            var testData = new List<TestCaseData>();

            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("Không tìm thấy file Excel tại: " + filePath);
            }

            try
            {
                DataTable excelDataTable = ExcelDataProvider.ReadExcel(filePath);
                string tenNXB = GetContentAfterColon(excelDataTable.Rows[59][4]?.ToString()) ?? ""; // Cho phép trống
                string diaChi = GetContentAfterColon(excelDataTable.Rows[60][4]?.ToString());
                string soDienThoai = GetContentAfterColon(excelDataTable.Rows[61][4]?.ToString());
                string expectedResult = excelDataTable.Rows[59][6]?.ToString().Trim();

                if (!string.IsNullOrEmpty(diaChi) && !string.IsNullOrEmpty(soDienThoai) && !string.IsNullOrEmpty(expectedResult))
                {
                    testData.Add(new TestCaseData(tenNXB, diaChi, soDienThoai, expectedResult));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Lỗi khi đọc file Excel: " + ex.Message);
                throw;
            }

            return testData;
        }
    }
}
