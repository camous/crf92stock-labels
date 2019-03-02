using Newtonsoft.Json.Linq;
using QRCoder;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Words.NET;

namespace crf92stocklabels
{
    class Program
    {
        static void Main(string[] args)
        {
            // use rebrandly prefix
            var baseuri = "rebrand.ly/XXXXXXXX?k=";
            var articles = "https://raw.githubusercontent.com/camous/crf92stock/master/products.json";

            var rawarticles = new System.Net.WebClient().DownloadString(articles);
            var products = JObject.Parse(rawarticles);

            QRCodeGenerator qrGenerator = new QRCodeGenerator(); 

            using (var doc = DocX.Create("qrcodes.docx")){

                var tableadd = doc.AddTable((products.Count +1) / 2, 4);
   
                var table = doc.InsertTable(tableadd);

                var line = 0;
                var column = 0;
                foreach (var product in products)
                {
                    QRCodeData qrCodeData = qrGenerator.CreateQrCode(baseuri + product.Key, QRCodeGenerator.ECCLevel.Q);

                    QRCode qrCode = new QRCode(qrCodeData);
                    var qrCodeAsBitmap = qrCode.GetGraphic(2);
                    qrCodeAsBitmap.Save(product.Key + ".png", System.Drawing.Imaging.ImageFormat.Png);

                    System.IO.Stream stream = new MemoryStream();
                    qrCodeAsBitmap.Save(stream, System.Drawing.Imaging.ImageFormat.Jpeg);
                    stream.Position = 0;
                    var qrcode = doc.AddImage(stream);

                    var qrcell = table.Rows[line].Cells[column].InsertParagraph();
                    
                    qrcell.AppendPicture(qrcode.CreatePicture());

                    table.Rows[line].Cells[column+1].InsertParagraph(product.Value.Value<string>());

                    if (column != 0 && column % 2 == 0)
                    {
                        column = 0;
                        line++;
                    }
                    else
                    {
                        column += 2;
                    }
                }

                doc.Save();
            }
        }
    }
}
