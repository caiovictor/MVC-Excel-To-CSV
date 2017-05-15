using Excel.Para.CSV.Extensions;
using System;
using System.IO;
using System.Web;
using System.Web.Mvc;

namespace Excel.Para.CSV.Controllers
{
    public class UploadController : Controller
    {
        [HttpGet]
        public ActionResult UploadFile()
        {
            return View();
        }

        [HttpPost]
        public ActionResult UploadFile(HttpPostedFileBase file, FormCollection form)
        {
            try
            {
                if (file.ContentLength > 0)
                {
                    string _FileName = Path.GetFileName(file.FileName);
                    string _path = Path.Combine(Server.MapPath("~/Content/Upload"), _FileName);
                    file.SaveAs(_path);

                    var excel = ExtentionsExcel.LerArquivoExcel(_path);

                    if (form["planilhas"].Equals("todas"))
                        ExtentionsExcel.ConverterCSV(excel, true, _path + ".csv");
                    else
                        ExtentionsExcel.ConverterCSV(excel, false, _path + ".csv");

                    ViewBag.Status = "success";
                    ViewBag.Message = "Arquivo convertido com sucesso! Clique no link acima para fazer o download.";
                    ViewBag.File = _FileName + ".csv";
                    return View();
                }
                else
                {
                    ViewBag.Status = "danger";
                    ViewBag.Message = "Falha ao converter o arquivo! ";
                    return View();
                }

            }
            catch (Exception ex)
            {
                ViewBag.Status = "danger";
                ViewBag.Message = "Falha ao converter o arquivo! " + ex.Message;
                return View();
            }
        }

        public FileResult Download(string file)
        {
            byte[] arquivo = System.IO.File.ReadAllBytes(Path.Combine(Server.MapPath("~/Content/Upload"), file));

            return File(arquivo, "application/vnd.ms-excel", "conversaoExcel.csv");
        }
    }
}