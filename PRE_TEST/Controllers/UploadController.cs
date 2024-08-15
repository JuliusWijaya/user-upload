using PRE_TEST.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity.Validation;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using LinqToExcel;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Threading.Tasks;

namespace PRE_TEST.Controllers
{
    public class UploadController : Controller
    {
        private readonly SIAKADEntities db = new SIAKADEntities();

        // GET: Upload
        public ActionResult UploadUser()
        {
            var userTemp = db.user_temp.AsQueryable();
            return View(userTemp);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult ImportUser(HttpPostedFileBase theFile)
        {
            List<string> data = new List<string>();

            if (theFile != null && theFile.ContentLength > 0)
            {
                if (theFile.ContentType == "application/vnd.ms-excel" || theFile.ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                {
                    string filename = theFile.FileName;
                    string targetpath = Server.MapPath("~/Doc/");

                    if (!Directory.Exists(targetpath))
                    {
                        Directory.CreateDirectory(targetpath);
                    }

                    string pathToExcelFile = Path.Combine(targetpath, filename);
                    theFile.SaveAs(pathToExcelFile);

                    var connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\";", pathToExcelFile);

                    var adapter = new OleDbDataAdapter("SELECT * FROM [Sheet1$]", connectionString);
                    var ds = new DataSet();
                    adapter.Fill(ds, "ExcelTable");
                    var excelFile = new ExcelQueryFactory(pathToExcelFile);

                    try
                    {
                        var users = from a in excelFile.Worksheet<user_temp>("Sheet1") select a;

                        foreach (var user in users)
                        {
                            if (!string.IsNullOrEmpty(user.email))
                            {
                                InsertUserToDatabase(user);
                            }
                            else
                            {
                                TempData["Error"] = "Email user is required";
                                return RedirectToAction("UploadUser");
                            }
                        }

                        if (!data.Any())
                        {
                            db.SaveChanges();
                        }
                    }
                    catch (DbEntityValidationException ex)
                    {
                        foreach (var entityValidationErrors in ex.EntityValidationErrors)
                        {
                            foreach (var validationError in entityValidationErrors.ValidationErrors)
                            {
                                string errorMessage = $"Property: {validationError.PropertyName}, Error: {validationError.ErrorMessage}";
                                data.Add(errorMessage);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        TempData["Error"] = "Error: " + ex.Message;
                        return RedirectToAction("UploadUser");
                    }

                    // Delete the excel file from the folder
                    if (System.IO.File.Exists(pathToExcelFile))
                    {
                        System.IO.File.Delete(pathToExcelFile);
                    }

                    if (data.Any())
                    {
                        TempData["Error"] = "Validation Error";
                        TempData["ValidationErrors"] = data;
                    }
                    else
                    {
                        TempData["success"] = "Successfully import file excel";
                    }

                    return RedirectToAction("UploadUser");
                }
                else
                {
                    TempData["Error"] = "Only Excel file format is allowed";
                    return RedirectToAction("UploadUser");
                }
            }
            else
            {
                TempData["Error"] = "Please choose an Excel file";
                return RedirectToAction("UploadUser");
            }
        }

        private void InsertUserToDatabase(user_temp user_Temp)
        {
            string strSQL = "INSERT INTO user_temp " +
                            "(name, jk, email, no_telp, address)" +
                            "VALUES(@name, @jk, @email, @no_telp, @address)";


            db.Database.ExecuteSqlCommand(strSQL,
                new SqlParameter("@name", user_Temp.name),
                new SqlParameter("@jk", user_Temp.jk),
                new SqlParameter("@email", user_Temp.email),
                new SqlParameter("@no_telp", user_Temp.no_telp),
                new SqlParameter("@address", user_Temp.address)
            );
        }

        public ActionResult AdditionalUser(string command)
        {
            if (command.Equals("save"))
            {
                string cmd = "INSERT INTO [user] ([name], [jk], [email], [no_telp], [address]) SELECT [name], [jk], [email], [no_telp], [address] FROM [user_temp] ut " +
                             "WHERE NOT EXISTS(SELECT * FROM [user] u WHERE (ut.[name] = u.[name]) and (ut.[email] = u.[email]))";
                db.Database.ExecuteSqlCommand(cmd);
                db.Database.ExecuteSqlCommand("DELETE FROM user_temp");
                db.SaveChanges();

                TempData["success"] = "Successfully import data file excel";
                return Redirect("UploadUser");
            }
            else
            {
                TempData["Error"] = "Failed import data file excel!";
                return Redirect("UploadUser");
            }
        }

        public async Task<JsonResult> DeleteUserTemp()
        {
            try
            {
                db.Database.ExecuteSqlCommand("DELETE FROM user_temp");
                db.SaveChanges();

                var result = new
                {
                    code = 200,
                    success = true,
                    message = "Successfully delete user"
                };

                return Json(result, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                var result = new
                {
                    code = 500,
                    success = false,
                    message = ex.InnerException.Message
                };

                return Json(result, JsonRequestBehavior.AllowGet);
            }
        }
    }
}