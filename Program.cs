using RestSharp;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;

namespace ConsoleApplication5
{
    class Program
    {
        static void Main(string[] args)
        {

            var log = new StreamWriter(ConfigurationManager.AppSettings["output_file"]) { AutoFlush = true };
            var re = new System.Text.RegularExpressions.Regex(@"\d{7,8}-[\dkK]");

            Console.Write("Descomprimiendo el archivo {0}.7z... ", ConfigurationManager.AppSettings["path_file_contribuyentes"]);
            var ms = new MemoryStream();
            var sevenZip = new SevenZip.SevenZipExtractor(ConfigurationManager.AppSettings["path_file_contribuyentes"] + ".7z");
            sevenZip.ExtractFile(ConfigurationManager.AppSettings["path_file_contribuyentes"] + ".csv", ms);
            Console.Write("Ok!\n\r");

            Console.Write("Leyendo el archivo {0}.csv... ", ConfigurationManager.AppSettings["path_file_contribuyentes"]);
            var binaryContent = ms.GetBuffer();
            var stringContent = System.Text.Encoding.GetEncoding("UTF-8").GetString(binaryContent);
            Console.Write("Ok!\n\r");

            Console.Write("Procesando registros...\n\r");
            var contribuyentes = stringContent.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
            foreach(var contribuyente in contribuyentes)
            {

                var arr = contribuyente.Split(';');
                if (re.IsMatch(arr[0]))
                {
                    var rut = arr[0].Split('-')[0];
                    var dv = arr[0].Split('-')[1];

                    var resultado = GetDocumentosAutorizados(rut, dv);

                    Console.WriteLine(resultado);
                    log.WriteLine(resultado);

                }
                
            }

            Console.Write("\n\rPresione cualquier teclar para continuar...\n\r");
            Console.ReadKey();

        }

        private static string GetDocumentosAutorizados(string rut, string dv)
        {
            var ret = string.Empty;
            var tiposAutorizados = new Dictionary<string, bool>()
            {
                { "33", false },
                { "34", false },
                { "39", false },
                { "41", false },
                { "43", false },
                { "46", false },
                { "52", false },
                { "56", false },
                { "61", false },
                { "110", false },
                { "111", false },
                { "112", false }
            };

            var response = ExecuteRequest(rut, dv);

            if (response.StatusCode == System.Net.HttpStatusCode.OK)
            {

                var doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(response.Content);

                // Del <body> obtengo el segundo <center> y de ahí la cuarta <table> y de ahí los <tr> cuya posición sea mayor a 1 y de ese tr el primer <td> y luego el tag <font> cuya clase sea 'texto'.
                var tds = doc.DocumentNode.SelectNodes(ConfigurationManager.AppSettings["xpath_tipo_documento"]);
                if (tds != null)
                {
                    foreach (var tipo in tds)
                    {

                        var tipoAutorizado = tipo.InnerText.Trim();
                        if (!string.IsNullOrEmpty(tipoAutorizado)) tiposAutorizados[tipoAutorizado] = true;

                    }
                }
            }

            ret = string.Format("{0}-{1}{2}", rut, dv, ImprimeAutorizados(tiposAutorizados));

            return ret;
        }

        private static IRestResponse ExecuteRequest(string rut, string dv)
        {

            var client = new RestClient(ConfigurationManager.AppSettings["url_sii"]);
            var request = new RestRequest(Method.POST);
            request.AddHeader("content-type", "application/x-www-form-urlencoded");
            request.AddParameter("application/x-www-form-urlencoded", string.Format("RUT_EMP={0}&DV_EMP={1}", rut, dv), ParameterType.RequestBody);
            IRestResponse response = client.Execute(request);

            return response;
            
        }

        private static string ImprimeAutorizados(Dictionary<string, bool> tipos)
        {
            var str = string.Empty;

            foreach (var tipo in tipos)
            {
                str += string.Format(";{0}", (tipo.Value ? tipo.Key : string.Empty));
            }

            return str;
        }
    }
}
