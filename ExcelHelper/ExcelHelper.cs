using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;


namespace Nlog
{
    public class ExcelHelper
    {
        public static void GenerateReport(string path)
        {
            var userList = GetUserFromTxt();

            IWorkbook workBooks;
            ISheet sheet;
            using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.ReadWrite))
            {
                workBooks = new HSSFWorkbook(fs);
                sheet = workBooks.GetSheetAt(0);
            }

            if (sheet != null)
            {
                for (int i = 0; i < userList.Count; i++)
                {
                    var row = sheet.GetRow(i + 2);
                    var user = userList[i];
                    try
                    {
                        row.Cells[1].SetCellValue(user.Name);
                        row.Cells[2].SetCellValue(user.CarlType);
                        row.Cells[3].SetCellValue(user.Id);
                        row.Cells[4].SetCellValue(user.BirthDate);
                        row.Cells[5].SetCellValue(user.Female);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"try to set value error at line {i}: {user.Name}");
                        Console.WriteLine(ex.Message);
                    }
                }
                Console.WriteLine($"Total insert {userList.Count} records!");
            }
            //save 
            SaveExcel(sheet);
        }

        public static void SaveExcel(ISheet sheet)
        {
            using (FileStream fs2 = File.Create($"sheet.SheetName_{DateTime.Now.ToString("yyyyMMdd")}_final.xls"))
            {
                sheet.Workbook.Write(fs2);
                fs2.Close();
            }
        }

        public static List<UserEntity> GetUserFromTxt()
        {

            var users = new List<UserEntity>();
            var userStrList = new List<string>();
            using (StreamReader reader = new StreamReader("name.txt"))
            {
                //userStr = reader.ReadToEnd();
                var i = 1;
                string userStr;
                while ((userStr = reader.ReadLine()) != null)
                {
    
                    if (userStr == null) continue;
                    try
                    {
                       
                        userStr = userStr.Replace("\n", "");
                        userStr = userStr.Replace("\r", "");
                        
                        if (!string.IsNullOrEmpty(userStr))
                        {
                            if (userStrList.Contains(userStr))
                            {
                                Console.WriteLine($"{userStr} already exist, won't insert again.");
                            }
                            else {
                                userStrList.Add(userStr.Trim());
                                users.Add(GetUserDetail(userStr.Trim()));
                            }
                            i++;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"line {i}  row( {userStr} )  get error as:");
                        Console.WriteLine(ex.Message);
                    }
                }
             
            }

            return users;
        }

        public static UserEntity GetUserDetail(string user)
        {
            UserEntity entity = new UserEntity();
            if (!string.IsNullOrEmpty(user))
            {
                //regex to get name

                //entity.Name = Regex.Replace(user, @"\d{17}([0-9]|X|x)", "");
                //entity.Id = user.Replace(entity.Name, "");

                entity.Name = Regex.Replace(user, @"[a-zA-Z\d]+", "");
                entity.Id = user.Replace(entity.Name, "");
                if (entity.Id == null || (entity.Id.Length != 18 && entity.Id.Length != 15))
                    throw new Exception($"user {user} ID lengh  is not correct.");

                var date = entity.Id.Substring(6, 8);
                entity.BirthDate = date.Insert(4, "-").Insert(7, "-");
                entity.Female = Convert.ToInt16(entity.Id.Substring(16, 1)) % 2 == 0 ? "1" : "0";
            }
            return entity;
        }

        public class UserEntity
        {
            public String Name { get; set; }
            public String Id { get; set; }
            public String Female { get; set; }
            public String BirthDate { get; set; }
            public String CarlType { get; set; }
            public UserEntity()
            {
                this.CarlType = "0";
            }

        }
    }
}
