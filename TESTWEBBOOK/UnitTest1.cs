using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using WebDriverManager;
using WebDriverManager.DriverConfigs.Impl;
using System;
using System.Threading;
using ExcelDataReader;
using System.Globalization;

namespace TESTWEBBOOK
{
    public class WebBook
    {
        IWebDriver driver;
        WebDriverWait wait;
        string baseUrl = "https://localhost:44326";

        [SetUp]
        public void Setup()
        {
            new DriverManager().SetUpDriver(new ChromeConfig());
            driver = new ChromeDriver();
            driver.Manage().Window.Maximize();
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
        }
// đăng ký thành công
        public string Test_DangKy(string HotenKH, string TenDN, string Matkhau, string Matkhaunhaplai, string Email, string Diachi, string Dienthoai, string Ngaysinh)
        {
            driver.Navigate().GoToUrl(baseUrl + "/Nguoidung/Dangky");
            Thread.Sleep(2000);

            driver.FindElement(By.Id("HotenKH")).SendKeys(HotenKH);
            driver.FindElement(By.Id("TenDN")).SendKeys(TenDN);
            driver.FindElement(By.Id("Matkhau")).SendKeys(Matkhau);
            driver.FindElement(By.Id("Matkhaunhaplai")).SendKeys(Matkhaunhaplai);
            driver.FindElement(By.Id("Email")).SendKeys(Email);
            driver.FindElement(By.Id("Diachi")).SendKeys(Diachi);
            driver.FindElement(By.Id("Dienthoai")).SendKeys(Dienthoai);
            if (DateTime.TryParseExact(Ngaysinh, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDate))
            {
                Ngaysinh = parsedDate.ToString("yyyy-MM-dd");
            }
            IWebElement ngaySinhInput = driver.FindElement(By.Name("Ngaysinh"));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].value = arguments[1];", ngaySinhInput, Ngaysinh);
            driver.FindElement(By.XPath("//input[@type='submit']")).Click();
            Thread.Sleep(3000);
            if (driver.Url.Contains(baseUrl + "/Nguoidung/Dangnhap"))
            {
                return "Đăng ký thành công";
            }
            else
            {
                return "Đăng ký thất bại";
            }
        }
        [Test, TestCaseSource(typeof(ExcelDataProvider), nameof(ExcelDataProvider.GetDataFromExcel_TestDangKy))]
        public void Test_DangKy(string HotenKH, string TenDN, string Matkhau, string Matkhaunhaplai, string Email, string Diachi, string Dienthoai, string Ngaysinh, string expected)
        {
            string expectedResult = "Đăng ký thành công";

            string actualResult = Test_DangKy(HotenKH, TenDN, Matkhau, Matkhaunhaplai, Email, Diachi, Dienthoai, Ngaysinh);

            ExcelDataProvider.WriteResultToExcel(28, actualResult, actualResult == expectedResult ? "Pass" : "Fail");
            Assert.AreEqual(expectedResult, actualResult);
        }
 // đăng ký sai định dạng mail
        public string Test_DangKySaiDinhDang(string HotenKH, string TenDN, string Matkhau, string Matkhaunhaplai, string Email, string Diachi, string Dienthoai, string Ngaysinh)
        {
            driver.Navigate().GoToUrl(baseUrl + "/Nguoidung/Dangky");
            Thread.Sleep(2000);
            driver.FindElement(By.Id("HotenKH")).SendKeys(HotenKH);
            driver.FindElement(By.Id("TenDN")).SendKeys(TenDN);
            driver.FindElement(By.Id("Matkhau")).SendKeys(Matkhau);
            driver.FindElement(By.Id("Matkhaunhaplai")).SendKeys(Matkhaunhaplai);
            driver.FindElement(By.Id("Email")).SendKeys(Email);
            driver.FindElement(By.Id("Diachi")).SendKeys(Diachi);
            driver.FindElement(By.Id("Dienthoai")).SendKeys(Dienthoai);

            if (DateTime.TryParseExact(Ngaysinh, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDate))
            {
                Ngaysinh = parsedDate.ToString("yyyy-MM-dd");
            }
            IWebElement ngaySinhInput = driver.FindElement(By.Name("Ngaysinh"));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].value = arguments[1];", ngaySinhInput, Ngaysinh);
            driver.FindElement(By.XPath("//input[@type='submit']")).Click();
            Thread.Sleep(3000);
            if (driver.Url.Contains(baseUrl + "/Nguoidung/Dangnhap"))
            {
                return "Đăng ký thành công"; 
            }
            try
            {
                IWebElement errorElement = driver.FindElement(By.CssSelector("td#p1"));
                string errorText = errorElement.Text;
                return errorText;
            }
            catch (NoSuchElementException)
            {
                return "Không tìm thấy thông báo lỗi";
            }
        }

        [Test, TestCaseSource(typeof(ExcelDataProvider), nameof(ExcelDataProvider.GetDataFromExcel_TestDangKySaiDinhDangMail))]
        public void Test_DangKySaiDinhDang(string HotenKH, string TenDN, string Matkhau, string Matkhaunhaplai, string Email, string Diachi, string Dienthoai, string Ngaysinh, string expected)
        {
            string actualResult = Test_DangKySaiDinhDang(HotenKH, TenDN, Matkhau, Matkhaunhaplai, Email, Diachi, Dienthoai, Ngaysinh);
            string expectedResult = "Lỗi! Email không đúng định dạng";

            ExcelDataProvider.WriteResultToExcel(36, actualResult, actualResult == expectedResult ? "Pass" : "Fail");
            Assert.AreEqual(expectedResult, actualResult, "Kết quả không đúng!");
        }
 // Đăng ký bỏ trống trường mật khẩu
        public string Test_BoTrongMatKhau(string HotenKH, string TenDN, string Matkhau, string Matkhaunhaplai, string Email, string Diachi, string Dienthoai, string Ngaysinh)
        {
            driver.Navigate().GoToUrl(baseUrl + "/Nguoidung/Dangky");
            Thread.Sleep(2000);

            driver.FindElement(By.Id("HotenKH")).SendKeys(HotenKH);
            driver.FindElement(By.Id("TenDN")).SendKeys(TenDN);
            driver.FindElement(By.Id("Matkhau")).SendKeys(Matkhau); 
            driver.FindElement(By.Id("Matkhaunhaplai")).SendKeys(Matkhaunhaplai);
            driver.FindElement(By.Id("Email")).SendKeys(Email);
            driver.FindElement(By.Id("Diachi")).SendKeys(Diachi);
            driver.FindElement(By.Id("Dienthoai")).SendKeys(Dienthoai);

            if (DateTime.TryParseExact(Ngaysinh, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDate))
            {
                Ngaysinh = parsedDate.ToString("yyyy-MM-dd");
            }
            IWebElement ngaySinhInput = driver.FindElement(By.Name("Ngaysinh"));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].value = arguments[1];", ngaySinhInput, Ngaysinh);

            driver.FindElement(By.XPath("//input[@type='submit']")).Click();
            Thread.Sleep(3000);

            if (driver.Url.Contains(baseUrl + "/Nguoidung/Dangnhap"))
            {
                return "Đăng ký thành công";
            }

            try
            {
                IWebElement errorMessage = driver.FindElement(By.XPath("//td[input[@id='Matkhau']]/following-sibling::td"));
                return errorMessage.Text;
            }
            catch (NoSuchElementException)
            {
                return "Không tìm thấy thông báo lỗi";
            }
            finally
            {
                driver.Quit(); 
            }
        }
        [Test, TestCaseSource(typeof(ExcelDataProvider), nameof(ExcelDataProvider.GetDataFromExcel_TestBoTrongMK))]
        public void Test_BoTrongMatKhau(string HotenKH, string TenDN, string Matkhau, string Matkhaunhaplai, string Email, string Diachi, string Dienthoai, string Ngaysinh, string expected)
        {
            string actualResult = Test_BoTrongMatKhau(HotenKH, TenDN, Matkhau, Matkhaunhaplai, Email, Diachi, Dienthoai, Ngaysinh);
            string expectedResult = "Mật khẩu không thể để trống!";

            ExcelDataProvider.WriteResultToExcel(44, actualResult, actualResult == expectedResult ? "Pass" : "Fail");
            Assert.AreEqual(expectedResult, actualResult, "Kết quả không đúng!");
        }

        // đăng ký thành công
        public string Test_DangNhap(string TenDN, string Matkhau)
        {
            driver.Navigate().GoToUrl(baseUrl + "/Nguoidung/Dangnhap");
            Thread.Sleep(2000);
            driver.FindElement(By.Id("TenDN")).SendKeys(TenDN);
            driver.FindElement(By.Id("Matkhau")).SendKeys(Matkhau);
            driver.FindElement(By.XPath("//input[@type='submit']")).Click();
            Thread.Sleep(3000);
            if (driver.Url.Contains(baseUrl))
            {
                return "Đăng nhập thành công";
            }
            else
            {
                return "Đăng nhập thất bại";
            }
        }
        [Test, TestCaseSource(typeof(ExcelDataProvider), nameof(ExcelDataProvider.GetDataFromExcel_TestDangNhap))]
        public void Test_DangNhap(string TenDN, string Matkhau, string expected)
        {
            string expectedResult = "Đăng nhập thành công";

            string actualResult = Test_DangNhap(TenDN, Matkhau);

            ExcelDataProvider.WriteResultToExcel(49, actualResult, actualResult == expectedResult ? "Pass" : "Fail");
            Assert.AreEqual(expectedResult, actualResult);
        }

        //  đăng nhap không thành công (chưa chạy được) 
        public string DangNhapInvalid(string TenDN, string Matkhau)
        {
            driver.Navigate().GoToUrl(baseUrl + "/Nguoidung/Dangnhap");
            Thread.Sleep(2000);
            driver.FindElement(By.Id("TenDN")).SendKeys(TenDN);
            driver.FindElement(By.Id("Matkhau")).SendKeys(Matkhau);
            driver.FindElement(By.XPath("//input[@type='submit']")).Click();
            Thread.Sleep(3000);
            wait.Until(d => d.FindElements(By.XPath("//input[@id='TenDN']/following-sibling::font")).Count > 0 ||
                            !d.Url.Contains("/Nguoidung/Dangnhap"));

            var errorElements = driver.FindElements(By.XPath("//input[@id='TenDN']/following-sibling::font"));
            if (errorElements.Count > 0)
            {
                return errorElements[0].Text; 
            }
            if (!driver.Url.Contains("/Nguoidung/Dangnhap"))
            {
                return "Đăng nhập thành công";
            }
            return "Không tìm thấy thông báo lỗi hoặc chuyển hướng";

        }
        [Test, TestCaseSource(typeof(ExcelDataProvider), nameof(ExcelDataProvider.GetDataFromExcel_TestDangNhapInvalid))]
        public void Test_DangNhapInvalid(string TenDN, string Matkhau, string expected)
        {
            string actualResult = DangNhapInvalid(TenDN, Matkhau);
            ExcelDataProvider.WriteResultToExcel(53, actualResult, actualResult == expected ? "Pass" : "Fail");
            Assert.AreEqual(expected, actualResult);
        }
        //GUI   
        public string Test_GUIBook(string Tensach)
        {
            driver.Navigate().GoToUrl(baseUrl);
            Thread.Sleep(2000);

            var bookLinks = driver.FindElements(By.XPath("//div[@class='product-content text-center']/h3/a"));
            foreach (var book in bookLinks)
            {
                if (book.Text.Trim() == Tensach)
                {
                    book.Click();
                    Thread.Sleep(2000);
                    break;
                }
            }
            IWebElement bookTitleElement = driver.FindElement(By.TagName("h5"));
            string bookTitleDetail = bookTitleElement.Text.Trim();

            return bookTitleDetail;
        }

        [Test, TestCaseSource(typeof(ExcelDataProvider), nameof(ExcelDataProvider.GetDataFromExcel_GUIBookTitle))]
        public void Test_GUIBook(string Tensach, string expected)
        {
            string actualResult = Test_GUIBook(Tensach);
            ExcelDataProvider.WriteResultToExcel(58, actualResult, actualResult == expected ? "Pass" : "Fail");
            Assert.AreEqual(expected, actualResult);
        }

        [Test]
        public void Test_ThemSach_Va_KiemTraGioHangDatHang()
        {
            driver.Navigate().GoToUrl(baseUrl + "/Nguoidung/Dangnhap");
            Thread.Sleep(2000);

            string tenDangNhap = "testuser_1";
            string matKhau = "testpass1";

            driver.FindElement(By.Id("TenDN")).SendKeys(tenDangNhap);
            driver.FindElement(By.Id("Matkhau")).SendKeys(matKhau);
            driver.FindElement(By.XPath("//input[@type='submit']")).Click();
            Thread.Sleep(3000);


            string bookTitle = "Trên Đỉnh Phố Wall";  
            IWebElement bookLink = driver.FindElement(By.XPath($"//a[contains(text(), '{bookTitle}')]"));
            bookLink.Click();
            Thread.Sleep(5000);

            wait.Until(driver => driver.Url.Contains("BookStore/Details"));

            IWebElement buyButton = driver.FindElement(By.XPath("//a[contains(@href, '/GioHang/ThemGiohang')]"));
            buyButton.Click();
            Thread.Sleep(4000);

            IWebElement cartIcon = driver.FindElement(By.CssSelector("i.ti-bag"));
            cartIcon.Click();
            Thread.Sleep(2000);

            try
            {
                IWebElement productInCart = driver.FindElement(By.XPath($"//*[contains(text(), '{bookTitle}')]"));
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", productInCart);
                Thread.Sleep(4000);
            }
            catch (NoSuchElementException)
            {
                Assert.Fail("Lỗi: Không tìm thấy sản phẩm trong giỏ hàng!");
            }

            bool isBookInCart = driver.FindElements(By.XPath($"//*[contains(text(), '{bookTitle}')]")).Count > 0;
            Assert.That(isBookInCart, "Lỗi: Sách không có trong giỏ hàng sau khi đặt mua!");
            IWebElement orderButton = wait.Until(drv => drv.FindElement(By.XPath("//td[@colspan='8' and contains(@class, 'text-white')]/a[contains(@href, '/Giohang/DatHang')]")));
            orderButton.Click();
            IWebElement confirmButton = driver.FindElement(By.XPath("//input[@type='submit']"));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", confirmButton);
            bool isOrderSuccess = wait.Until(drv => drv.Url.Contains("/Giohang/Xacnhandonhang"));

            Assert.That(isOrderSuccess, "Lỗi: Không chuyển về trang xác nhận đơn hàng!");
        }


        [Test]
        public void Test_DangNhap_Admin()
        {
            driver.Navigate().GoToUrl(baseUrl + "/Admin/Login");
            wait.Until(d => d.FindElement(By.Name("username"))); 

            driver.FindElement(By.Name("username")).SendKeys("Admin");
            driver.FindElement(By.Name("password")).SendKeys("Admin");

            driver.FindElement(By.XPath("//button[@type='submit']")).Click();

            wait.Until(d => d.Url.Contains("/Admin"));
            Thread.Sleep(3000);
            string currentUrl = driver.Url.TrimEnd('/');
            string expectedUrl = (baseUrl + "/Admin").TrimEnd('/');

            Assert.That(currentUrl, Is.EqualTo(expectedUrl), "Lỗi: Chưa chuyển hướng đến trang Admin!");
        }


        [Test]
        public void Test_DangNhap_Va_VaoQuanLySach()
        {
            driver.Navigate().GoToUrl(baseUrl + "/Admin/Login");
            wait.Until(d => d.FindElement(By.Name("username"))); 

            driver.FindElement(By.Name("username")).SendKeys("Admin");
            driver.FindElement(By.Name("password")).SendKeys("Admin");

            driver.FindElement(By.XPath("//button[@type='submit']")).Click();

            wait.Until(d => d.Url.Contains("/Admin"));
            Assert.That(driver.Url.Contains("/Admin"), "Lỗi: Chưa chuyển hướng đến trang Admin!");

            IWebElement quanLySachLink = wait.Until(d => d.FindElement(By.XPath("//a[@href='/Admin/Sach']")));
            quanLySachLink.Click();
            Thread.Sleep(3000);

            wait.Until(d => d.Url.Contains("/Admin/Sach"));
            Assert.That(driver.Url.Contains("/Admin/Sach"), "Lỗi: Không vào được trang Quản Lý Sách!");
        }

       
        [Test]
        public void Test_DangNhap_Admin_ThanhCong()
        {
            driver.Navigate().GoToUrl(baseUrl + "/Admin/Login");

            string adminUser = "Admin";
            string adminPass = "Admin";
            var usernameField = driver.FindElement(By.Name("username"));
            var passwordField = driver.FindElement(By.Name("password"));
            var loginButton = driver.FindElement(By.XPath("//button[text()='Đăng Nhập']"));

            usernameField.SendKeys(adminUser);
            passwordField.SendKeys(adminPass);
            loginButton.Click();
            string currentUrl = driver.Url.TrimEnd('/');
            Assert.That(currentUrl, Does.Contain("/Admin"), "Lỗi: Không chuyển hướng đúng sau khi đăng nhập!");
        }

            public string AddBookToSystem(string tenSach, string giaBan, string moTa, string anh, string ngayCapNhat, int soLuong, string nhaXuatBan, string theLoai)
            {
                try
                {
                    driver.Navigate().GoToUrl(baseUrl + "/Admin");
                    var quanLySachLink = driver.FindElement(By.XPath("//a[@href='/Admin/Sach']"));
                    quanLySachLink.Click();
                    Thread.Sleep(2000);

                var themMoiLink = driver.FindElement(By.XPath("//a[@href='/Admin/ThemmoiSach']"));
                    themMoiLink.Click();
                    Thread.Sleep(2000); 
                    driver.FindElement(By.Id("Tensach")).SendKeys(tenSach);
                    driver.FindElement(By.Id("Giaban")).SendKeys(giaBan);
                    driver.FindElement(By.Id("Mota")).SendKeys(moTa);
                    if (File.Exists(anh))
                    {
                        driver.FindElement(By.Name("fileupload")).SendKeys(Path.GetFullPath(anh));
                    }
                    else
                    {
                        throw new Exception("Lỗi: Không tìm thấy ảnh tại đường dẫn " + anh);
                    }
                    if (DateTime.TryParseExact(ngayCapNhat, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDate))
                        {
                            ngayCapNhat = parsedDate.ToString("yyyy-MM-dd"); 
                        }

                    IWebElement ngayCapNhatInput = driver.FindElement(By.Name("Ngaycapnhat"));
                    ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].value = arguments[1];", ngayCapNhatInput, ngayCapNhat);
                    driver.FindElement(By.Id("Soluongton")).SendKeys(soLuong.ToString());
                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
                    IWebElement theLoaiDropdown = wait.Until(d => d.FindElement(By.ClassName("nice-select")));
                    theLoaiDropdown.Click();

                    IWebElement theLoaiOption = wait.Until(d => d.FindElement(By.XPath("//li[contains(@class, 'option') and normalize-space(text())='" + theLoai + "']")));
                    IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                    js.ExecuteScript("arguments[0].click();", theLoaiOption);


                    IWebElement nhaXuatBanDropdown = driver.FindElement(By.ClassName("nice-select"));
                    nhaXuatBanDropdown.Click();
                    Thread.Sleep(2000); 
                    IWebElement nhaXuatBanOption = driver.FindElement(By.XPath("//li[contains(text(),'" + nhaXuatBan + "')]"));
                    nhaXuatBanOption.Click();

                    var btnTao = driver.FindElement(By.XPath("//input[@type='submit' and @value='Tạo']"));
                    btnTao.Click();

                driver.Navigate().GoToUrl(baseUrl + "/Admin/Sach?page=9");
                try
                {
                    var SachMoi = wait.Until(d => d.FindElement(By.XPath($"//td[contains(text(), '{tenSach}')]")));
                    return $"Sách '{tenSach}' đã được thêm thành công!";
                }
                catch (NoSuchElementException)
                {
                    return $"Sách '{tenSach}' không xuất hiện trong danh sách sau khi thêm!";
                }
            }
            catch (Exception e)
            {
                return $"Thêm sách thất bại: {e.Message}";
            }
        }

            [Test, TestCaseSource(typeof(ExcelDataProvider), nameof(ExcelDataProvider.GetDataFromEx_TestAddBook))]
            public void Test_AddBook(string tenSach, string giaBan, string moTa, string anh, string ngayCapNhat, int soLuong, string theLoai, string nhaXuatBan, string expectedResult)
            {
                Test_DangNhap_Admin_ThanhCong(); 

                string actualResult = AddBookToSystem(tenSach, giaBan, moTa, anh, ngayCapNhat, soLuong, theLoai, nhaXuatBan);
                ExcelDataProvider.WriteResultToExcel(3, actualResult, actualResult == expectedResult ? "Pass" : "Fail");
                Assert.AreEqual(expectedResult, actualResult);
        }
        // Thêm loại sách
        public string AddLoaiSach(string tenLoai)
        {
            try
            {
                driver.Navigate().GoToUrl(baseUrl + "/Admin"); 
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

                var quanLyLoaiLink = wait.Until(d => d.FindElement(By.XPath("//a[contains(@href, '/Admin/Loai')]")));
                quanLyLoaiLink.Click();

                var themMoiLink = wait.Until(d => d.FindElement(By.XPath("//a[contains(@href, 'ThemmoiLoai')]")));
                themMoiLink.Click();

                var inputTenLoai = wait.Until(d => d.FindElement(By.Id("TenLoai")));
                inputTenLoai.SendKeys(tenLoai);

                var btnThem = driver.FindElement(By.XPath("//input[@type='submit' and @value='Thêm Mới']"));
                btnThem.Click();

                driver.Navigate().GoToUrl(baseUrl + "/Admin/Loai");
                try
                {
                    var loaiSachMoi = wait.Until(d => d.FindElement(By.XPath($"//td[contains(text(), '{tenLoai}')]")));
                    return $"Loại sách '{tenLoai}' đã được thêm thành công!";
                }
                catch (NoSuchElementException)
                {
                    return $"Loại sách '{tenLoai}' không xuất hiện trong danh sách sau khi thêm!";
                }
            }
            catch (Exception e)
            {
                return $"Thêm loại thất bại: {e.Message}";
            }
        }
        // Test case sử dụng dữ liệu từ Excel
        [Test, TestCaseSource(typeof(ExcelDataProvider), nameof(ExcelDataProvider.GetDataFromExcel_TestLoai))]
        public void AddLoaiSach(string tenSach, string expectedResult, int rowIndex)
        {
            Test_DangNhap_Admin_ThanhCong(); 
            string actualResult = AddLoaiSach(tenSach);
            ExcelDataProvider.WriteResultToExcel(10, actualResult, actualResult == expectedResult ? "Pass" : "Fail");
            Assert.AreEqual(expectedResult, actualResult);
        }

        public string AddNXB(string tenNXB, string diaChi, string soDienThoai)
        {
            try
            {
                driver.Navigate().GoToUrl(baseUrl + "/Admin"); 
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

                var quanLyNXBLink = wait.Until(d => d.FindElement(By.XPath("//a[contains(@href, '/Admin/NXB')]")));
                quanLyNXBLink.Click();

                var themMoiLink = wait.Until(d => d.FindElement(By.XPath("//a[contains(@href, 'ThemNXB')]")));
                themMoiLink.Click();

                var inputTenNXB = wait.Until(d => d.FindElement(By.Id("TenNXB")));
                inputTenNXB.SendKeys(tenNXB);

                var inputDiaChi = wait.Until(d => d.FindElement(By.Id("Diachi")));
                inputDiaChi.SendKeys(diaChi);

                var inputSoDienThoai = wait.Until(d => d.FindElement(By.Id("DienThoai")));
                inputSoDienThoai.SendKeys(soDienThoai);

                var btnThem = wait.Until(d => d.FindElement(By.XPath("//input[@type='submit' and contains(@value, 'Thêm Mới')]")));
                btnThem.Click();

                driver.Navigate().GoToUrl(baseUrl + "/Admin/NXB");
                try
                {
                    var nxbMoi = wait.Until(d => d.FindElement(By.XPath($"//td[contains(text(), '{tenNXB}')]")));
                    return $"Nhà Xuất Bản '{tenNXB}' đã được thêm thành công!";
                }
                catch (NoSuchElementException)
                {
                    return $"Nhà Xuất Bản '{tenNXB}' không xuất hiện trong danh sách sau khi thêm!";
                }
            }
            catch (Exception e)
            {
                return $"Thêm NXB thất bại: {e.Message}";
            }
        }

        [Test, TestCaseSource(typeof(ExcelDataProvider), nameof(ExcelDataProvider.GetDataFromExcel_TestAddNXB))]
        public void AddNXB(string tenNXB, string diaChi, string soDienThoai, string expectedResult, int rowIndex)
        {
            Test_DangNhap_Admin_ThanhCong();
            string actualResult = AddNXB(tenNXB, diaChi, soDienThoai);
            ExcelDataProvider.WriteResultToExcel(15, actualResult, actualResult == expectedResult ? "Pass" : "Fail");
            Assert.AreEqual(expectedResult, actualResult);
        }
        // thêm nhà xuất bản để trống trường tên NXB
        public string AddNXB_TrongTenNXB(string tenNXB, string diaChi, string soDienThoai)
        {
            try
            {
                driver.Navigate().GoToUrl(baseUrl + "/Admin"); 
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                var quanLyNXBLink = wait.Until(d => d.FindElement(By.XPath("//a[contains(@href, '/Admin/NXB')]")));
                quanLyNXBLink.Click();
                var themMoiLink = wait.Until(d => d.FindElement(By.XPath("//a[contains(@href, 'ThemNXB')]")));
                themMoiLink.Click();
                var inputTenNXB = wait.Until(d => d.FindElement(By.Id("TenNXB")));
                inputTenNXB.SendKeys(tenNXB);
                var inputDiaChi = wait.Until(d => d.FindElement(By.Id("Diachi")));
                inputDiaChi.SendKeys(diaChi);
                var inputSoDienThoai = wait.Until(d => d.FindElement(By.Id("DienThoai")));
                inputSoDienThoai.SendKeys(soDienThoai);
                var btnThem = wait.Until(d => d.FindElement(By.XPath("//input[@type='submit' and contains(@value, 'Thêm Mới')]")));
                btnThem.Click();
                try
                {
                    var errorMessage = wait.Until(d => d.FindElement(By.CssSelector(".validation-summary-errors li")));
                    return errorMessage.Text;
                }
                catch (NoSuchElementException)
                {
                    return "Không tìm thấy thông báo lỗi khi bỏ trống tên Nhà Xuất Bản!";
                }
            }
            catch (Exception e)
            {
                return $"Thêm NXB thất bại: {e.Message}";
            }
        }

        [Test, TestCaseSource(typeof(ExcelDataProvider), nameof(ExcelDataProvider.GetDataFromExcel_TenNXBNull))]
        public void AddNXB_TrongTenNXBNull(string tenNXB, string diaChi, string soDienThoai, string expectedResult)
        {
            Test_DangNhap_Admin_ThanhCong();
            string actualResult = AddNXB_TrongTenNXB(tenNXB, diaChi, soDienThoai);
            ExcelDataProvider.WriteResultToExcel(61, actualResult, actualResult == expectedResult ? "Pass" : "Fail");
            Assert.AreEqual(expectedResult, actualResult);
        }

        public string TimKiemSach(string tuKhoaTimKiem)
        {
            try
            {
                driver.Navigate().GoToUrl(baseUrl);
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                var searchBox = wait.Until(d => d.FindElement(By.Name("sTuKhoa"))); 
                searchBox.Clear();
                searchBox.SendKeys(tuKhoaTimKiem);
                searchBox.SendKeys(Keys.Enter); 
                Thread.Sleep(3000);
                var ketQuaSach = driver.FindElements(By.CssSelector(".product-content h3 a"));
                if (ketQuaSach.Count == 0)
                {
                    return "Không có sách nào được tìm thấy!";
                }
                string tenSachKetQua = ketQuaSach[0].Text.Trim();
                return tenSachKetQua; 
            }
            catch (Exception e)
            {
                return $"Lỗi khi tìm kiếm: {e.Message}";
            }

        }
        [Test, TestCaseSource(typeof(ExcelDataProvider), nameof(ExcelDataProvider.GetDataFromExcel_TestSearchBook))]
        public void TimKiemSach(string tuKhoaTimKiem, string expectedResult, int rowIndex)
        {
            //Test_DangNhap();
            string actualResult = TimKiemSach(tuKhoaTimKiem);
            ExcelDataProvider.WriteResultToExcel(19, actualResult, actualResult == expectedResult ? "Pass" : "Fail");
            Assert.AreEqual(expectedResult, actualResult);
        }




        [TearDown]
        public void TearDown()
        {
            if (driver != null)
            {
                driver.Quit();
                driver.Dispose(); 
            }
        }

    }
}
