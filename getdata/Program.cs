using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.IO;
using HtmlAgilityPack;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace getdata
{
    class Program
    {

        static void Main(string[] args)
        {
            
            importData();

            //Process("http://tinbds.com/tk/ho-chi-minh/quan-8/p-");

            Console.ReadKey();
        }

        static String ReadData(string add)
        {
            WebClient wc = new WebClient();
            wc.Encoding = System.Text.Encoding.UTF8;
            string webData = wc.DownloadString(add);
            return webData;
        }
        static List<Person> Process(string add)
        {

            var doc = new HtmlDocument();
            doc.LoadHtml(ReadData(add));

            var centercontent = doc.DocumentNode.SelectNodes("//section[@id='center-content']/div").FirstOrDefault();
            var article = centercontent.SelectNodes("article");
            var items = new List<Person>();

            foreach (var a in article)
            {
                Person person = new Person();
                //Extract các giá trị từ các tag con của tag a
                var linkNode = a.SelectSingleNode(".//a[contains(@class,'title')]");

                person.Name = linkNode.InnerText;

                //lay cac the p
                var pNode = a.SelectNodes(".//figure/figcaption/p").ToList();


                //xóa node strong
                var strongNode = a.SelectNodes(".//figure/figcaption/p/strong").ToList();
                foreach (var c in strongNode)
                {
                    c.Remove();
                }

                // Lấy nội dung file p
                person.Dc = pNode[0].InnerHtml;
                person.Sdt = pNode[1].InnerHtml;
                person.Email = pNode[2].InnerHtml;

                items.Add(person);

            }
            
            return items;
            
            // write file
                    //String filepath = "D:\\test.txt";
                    //FileStream fs = new FileStream(filepath, FileMode.Create);//Tạo file mới tên là test.txt  ;
                    //StreamWriter sWriter = new StreamWriter(fs, Encoding.UTF8);

                    //foreach (var a in items)
                    //{
                    //    Console.WriteLine("ten:" + a.Name + " dc: " + a.Dc + " sdt: " + a.Sdt + " email: " + a.Email);
                    //    sWriter.WriteLine("ten:" + a.Name + " dc: " + a.Dc + " sdt: " + a.Sdt + " email: " + a.Email); ;
                    //}

                    //sWriter.Flush();
                    //fs.Close();
            // end

           

        }
        static void importData()
        {
            string filePath = "F:\\Desktop\\writeData\\working.xlsx";
        

            // nếu đường dẫn null hoặc rỗng thì báo không hợp lệ và return hàm
            if (string.IsNullOrEmpty(filePath))
            {
                Console.WriteLine("Đường dẫn lỗi, kiểm tra lại ở trên ");
                return;
            }

            try
            {
                using (ExcelPackage p = new ExcelPackage())
                {
                    // đặt tên người tạo file
                    p.Workbook.Properties.Author = "DongCutena";

                    // đặt tiêu đề cho file
                    p.Workbook.Properties.Title = "Dong cutena pro vo doi";

                    //Tạo một sheet để làm việc trên đó
                    p.Workbook.Worksheets.Add("working");

                    // lấy sheet vừa add ra để thao tác
                    ExcelWorksheet ws = p.Workbook.Worksheets[1];

                    // đặt tên cho sheet
                    ws.Name = "Working sheet";
                    // fontsize mặc định cho cả sheet
                    ws.Cells.Style.Font.Size = 10;
                    // font family mặc định cho cả sheet
                    ws.Cells.Style.Font.Name = "Arial";

                    //Tạo danh sách các column header
                            string[] arrColumnHeader = {
                                                        "Tên",
                                                        "Điện thoại",
                                                        "Email",
                                                        "Địa danh"
                };
                //#region NotImportant 
                //// đoạn này ko quan trọng lắm, màu vẽ chơi


                //// lấy ra số lượng cột cần dùng dựa vào số lượng header
                var countColHeader = arrColumnHeader.Count();

                //// merge các column lại từ column 1 đến số column header
                //// gán giá trị cho cell vừa merge là Thống kê thông tni User Kteam
                //ws.Cells[1, 1].Value = "Thống kê khác hàng gì gì đó của Dương lợn";
                //ws.Cells[1, 1, 1, countColHeader].Merge = true;
                //// in đậm
                //ws.Cells[1, 1, 1, countColHeader].Style.Font.Bold = true;
                //// căn giữa
                //ws.Cells[1, 1, 1, countColHeader].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                int colIndex = 1;
                int rowIndex = 2;

                    ////tạo các header từ column header đã tạo từ bên trên
                    //foreach (var item in arrColumnHeader)
                    //{
                    //    var cell = ws.Cells[rowIndex, colIndex];

                    //    //set màu thành gray
                    //    var fill = cell.Style.Fill;
                    //    fill.PatternType = ExcelFillStyle.Solid;
                    //    fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);

                    //    //căn chỉnh các border
                    //    var border = cell.Style.Border;
                    //    border.Bottom.Style =
                    //        border.Top.Style =
                    //        border.Left.Style =
                    //        border.Right.Style = ExcelBorderStyle.Thin;

                    //    //gán giá trị
                    //    cell.Value = item;

                    //    colIndex++;
                    //}
                    //#endregion

                    // lấy ra danh sách UserInfo từ ItemSource của DataGrid
                    //List<Person> personlist = dtgExcel.ItemsSource.Cast<UserInfo>().ToList();

                    // với mỗi item trong danh sách sẽ ghi trên 1 dòng
                    //1-103
                    for (int temp = 1; temp < 104; temp++)
                    {
                        List<Person> lstPerson = Process("http://tinbds.com/tk/ho-chi-minh/quan-8/p-" + temp);
                        foreach (var per in lstPerson)
                        {
                            // bắt đầu ghi từ cột 1. Excel bắt đầu từ 1 không phải từ 0
                            colIndex = 1;

                            // rowIndex tương ứng từng dòng dữ liệu
                            rowIndex++;

                            //gán giá trị cho từng cell                      
                            ws.Cells[rowIndex, colIndex++].Value = per.Name;

                            // lưu ý phải .ToShortDateString để dữ liệu khi in ra Excel là ngày như ta vẫn thấy.Nếu không sẽ ra tổng số :v
                            ws.Cells[rowIndex, colIndex++].Value = per.Sdt;


                            ws.Cells[rowIndex, colIndex++].Value = per.Email;

                            ws.Cells[rowIndex, colIndex++].Value = per.Dc;
                        }
                    }

                    //Lưu file lại
                    Byte[] bin = p.GetAsByteArray();
                    File.WriteAllBytes(filePath, bin);
                }
                Console.WriteLine("Thành công, tiếp đi! ");

            }
            catch (Exception EE)
            {
                Console.WriteLine("Lỗi : " + EE);
            }
        }
    }
}
