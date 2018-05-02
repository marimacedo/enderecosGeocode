using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Web;
using System.Web.Configuration;
using System.Web.Mvc;
using System.Web.UI;
using System.Web.UI.WebControls;
using WebApplication1.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace WebApplication1.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View(ViewBag.Mensagem);
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async System.Threading.Tasks.Task<ActionResult> Upload(HttpPostedFileBase upload)
        {
            try
            {
                if (ModelState.IsValid)
                {
                    List<string> listaEnderecos = new List<string>();

                    if (upload != null && upload.ContentLength > 0)
                    {
                        if (upload.FileName.EndsWith(".txt") || upload.FileName.EndsWith(".csv"))
                        {
                            StreamReader stream = new StreamReader(upload.InputStream);

                            string linha = null;
                            while ((linha = stream.ReadLine()) != null)
                            {
                                if (!string.IsNullOrEmpty(linha))
                                    listaEnderecos.Add(linha);
                            }

                            stream.Close();

                            var chaveMaps = WebConfigurationManager.AppSettings["ChaveGoogleMaps"];
                            HttpClient client = new HttpClient();
                            string url = "https://maps.googleapis.com/maps/api/geocode/json?address={0}&sensor=false&key=" + chaveMaps;

                            List<Planilha> planilha = new List<Planilha>();

                            foreach (var endereco in listaEnderecos)
                            {
                                try
                                {
                                    var response = await client.GetAsync(string.Format(url, endereco));
                                    string resultString = await response.Content.ReadAsStringAsync();
                                    ResultadoGoogle lista = JsonConvert.DeserializeObject<ResultadoGoogle>(resultString);

                                    if (lista.results != null && lista.results.Any())
                                    {
                                        var item = lista.results.FirstOrDefault();
                                        planilha.Add(new Planilha() { Endereco = endereco, Latitude = item.geometry.location.lat, Longitude = item.geometry.location.lng });
                                    }
                                }
                                catch { }
                            }

                            var nomeArquivo = TesteExcel(planilha);
                            ViewBag.Mensagem = $"O arquivo {nomeArquivo} foi gerado com sucesso";
                        }
                        else
                        {
                            ViewBag.Mensagem = $"Selecione um arquivo .txt ou .csv.";
                        }
                    }
                    else
                    {
                        ViewBag.Mensagem = "Selecione o arquivo!";
                    }
                }
                else
                {
                    ViewBag.Mensagem = "Selecione o arquivo!";
                }
             
            }catch(Exception e)
            {
                ViewBag.Mensagem = $"Erro ao ler o arquivo {e.Message}";
            }

            return View("Index");
        }

        private string TesteExcel(List<Planilha> planilha)
        {
            Microsoft.Office.Interop.Excel.Application App; // Aplicação Excel
            Microsoft.Office.Interop.Excel.Workbook WorkBook; // Pasta
            Microsoft.Office.Interop.Excel.Worksheet WorkSheet; // Planilha
            object misValue = System.Reflection.Missing.Value;

            App = new Microsoft.Office.Interop.Excel.Application();
            WorkBook = App.Workbooks.Add(misValue);
            WorkSheet = (Excel.Worksheet)WorkBook.Worksheets.get_Item(1);

            List<string> colunas = new List<string>() { "Endereço", "Latitude", "Longitude" };
            int linha = 1;
            int coluna = 1;

            //Response.ClearContent();
            //Response.Buffer = true;
            //Response.AddHeader("content-disposition", string.Format("attachment; filename={0}", "confidential feedback.xls"));
            //Response.ContentType = "application/ms-excel";
            //Response.Write("");
            //Response.Write("<html xmlns:x=\"urn:schemas-microsoft-com:office:excel\">");
            //Response.Write("<body>");
            //Response.Write("<table>");
            //planilha.ForEach(a =>
            //{
            //    Response.Write("<tr>");
            //    Response.Write($"<td>{a.Endereco}</td><td>{string.Concat("'",a.Latitude, "'")}</td><td>{a.Longitude.ToString()}</td>");
            //    Response.Write("<tr>");
            //});

            //Response.Write("</table>");
            //Response.Write("</body>");
            //Response.Write("</html>");

            colunas.ForEach(a =>
            {
                WorkSheet.Cells[linha, coluna] = a;
                coluna++;
            });

            coluna = 1;
            linha = 2;

            planilha.ForEach(item =>
            {
                WorkSheet.Cells[linha, coluna] = item.Endereco;
                coluna++;
                WorkSheet.Cells[linha, coluna] = item.Latitude;
                coluna++;
                WorkSheet.Cells[linha, coluna] = item.Longitude;

                linha++;
                coluna = 1;
            });
            
            // salva o arquivo
            var nomeArquivo = "Enderecos-"+Guid.NewGuid().ToString();
            WorkBook.SaveAs(@"c:\enderecos\"+nomeArquivo, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,

            Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            WorkBook.Close(true, misValue, misValue);
            App.Quit(); // encerra o excel

            return nomeArquivo;
        }
    }
}