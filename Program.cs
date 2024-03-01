using iText.Kernel.Colors;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using System.Data;
using System.Data.SqlClient;
using System.Net;
using System.Xml;
using System.Xml.Serialization;
using System.Linq;
using System.Globalization;
using iText.Layout.Borders;
using System.Net.Mail;

[XmlRoot("products")]
public class Products
{

    [XmlElement("product")]
    public Product[] ProductList { get; set; }
}

public class Product
{
    [XmlElement("stock")]
    public int Stock { get; set; }

    [XmlElement("id_product")]
    public string IdProduct { get; set; }

    [XmlElement("designation")]
    public string Designation { get; set; }

    [XmlElement("product_category")]
    public string ProductCategory { get; set; }

    [XmlElement("brand")]
    public string Brand { get; set; }

    [XmlElement("reference")]
    public string Reference { get; set; }

    [XmlElement("ean")]
    public string Ean { get; set; }

    [XmlElement("description")]
    public string Description { get; set; }

    [XmlElement("product_url")]
    public string ProductUrl { get; set; }

    [XmlElement("image_url")]
    public string ImageUrl { get; set; }

    [XmlElement("sale_price")]
    public string SalePrice { get; set; }

    [XmlElement("shipping_cost_iberia")]
    public double ShippingCostIberia { get; set; }

    [XmlElement("shipping_cost_ilhas_PT")]
    public double ShippingCostIlhasPT { get; set; }

    [XmlElement("best_seller")]
    public bool BestSeller { get; set; }
}

namespace stockSend
{
    class Program
    {

        static void Main(string[] args)
        {
            Console.WriteLine("A ligar ao servidor...\n\n");

            var datasource = @"HP-LOJA";//servidor
            var database = "database"; //nome da base de dados
            var username = "sa"; //utilizador
            var password = "Sa"; //password
            //variaveis para familias
            string familias = "", sqlQueryFamilias = "";
            SqlCommand commandFamilias;
            SqlDataReader dataReaderFamilias;
            var familiaId = "";
            var familiaNome = "";
            //variaveis para produtos
            string produtos = "", sqlQueryProdutos = "";
            string pvpXmlString = "";
            string pvpFinalString = "";
            float pvpFinal=0.0f;
            float margemValorFloat=0.0f;
            float margemPerc=0.0f;
            

            SqlCommand commandProdutos;
            SqlDataReader dataReaderProdutos;

            //connection string
            string connString = @"Data Source=" + datasource + ";Initial Catalog="
                        + database + ";Persist Security Info=True;User ID=" + username + ";Password=" + password;

            //cria uma nova ligação
            SqlConnection conn = new SqlConnection(connString);

            //tenta ligar
            try
            {
                Console.WriteLine("A ligar a BD...\n\n");

                //abre ligação
                conn.Open();

                Console.WriteLine("Ligacao bem sucedida!\n\n");
            }
            catch (Exception e)
            {
                Console.WriteLine("Erro: " + e.Message);
            }

            //query de familias
            sqlQueryFamilias = "SELECT FamilyID, Description FROM dbo.Family ORDER BY Description";
            commandFamilias = new SqlCommand(sqlQueryFamilias, conn);
            dataReaderFamilias = commandFamilias.ExecuteReader();
            //tabela de familias
            DataTable tabelaFamilias = new DataTable("tabelaFamilias");
            DataColumn idFamilia = new DataColumn("idFamilia");
            DataColumn familia = new DataColumn("familia");
            tabelaFamilias.Columns.Add(idFamilia);
            tabelaFamilias.Columns.Add(familia);

            int linhasFamilias = 0;

            ////////// LER O XML PARA OBTER O CNP E PVP PARA CALCULAR MARGENS //////////////////////////////////
            //ler um xml

              Console.WriteLine("A ler dados do ficheiro XML...\n\n");

               String urlXml = "https://www.zonpharma.com/extend/catalog_20.xml";

                // Create a web request to the URL
                WebRequest request = WebRequest.Create(urlXml);

                // Get the response from the web request
                WebResponse response = request.GetResponse();

                // Read the XML content from the response
                Stream xmlStream = response.GetResponseStream();

                // Cria uma instância do XmlSerializer para a classe Products
                XmlSerializer serializer = new XmlSerializer(typeof(Products));

                // Desserializa o XML e obtém uma instância da classe Products
                Products products;
                using (StreamReader reader = new StreamReader(xmlStream))
                {
                    products = (Products)serializer.Deserialize(reader);
                }

                // Fecha o fluxo e a resposta
                xmlStream.Close();
                response.Close();

            ////////////////////////////////////////////////////////////////////////////////////////////////////

            while (dataReaderFamilias.Read())
            {
                DataRow linhaNova = tabelaFamilias.NewRow();
                linhaNova["idFamilia"] = dataReaderFamilias.GetValue(0);
                linhaNova["familia"] = dataReaderFamilias.GetValue(1);

                tabelaFamilias.Rows.Add(linhaNova);

                linhasFamilias++;
            }

            dataReaderFamilias.Close();

            // Data atual
            DateTime justDate = DateTime.Today.Date;
            //DateTime justDate = justDate.Date;
            DateTime vendasPassadas = justDate.AddDays(-30);
            // Display da data-debug
            //Console.WriteLine(vendasPassadas.ToString("MM-dd-yyyy"));

            //query de produtos por familia
            sqlQueryProdutos = "SELECT o.ItemID, o.FamilyID, i.Description, k.AvailableQty, x.FamilyID, x.Description, z.Quantity, sp.* FROM dbo.Item o  INNER JOIN dbo.ItemNames i  on o.ItemID = i.ItemID INNER JOIN dbo.Family x  on o.FamilyID = x.FamilyID INNER JOIN dbo.Stock k on i.ItemID = k.ItemID INNER JOIN dbo.SaleTransactionDetails z on i.ItemID = z.ItemID INNER JOIN dbo.ItemSellingPrices sp ON o.ItemID = sp.ItemID WHERE k.WarehouseID = 1 AND k.AvailableQty<10 AND z.TransDocument = 'FR' AND z.CreateDate > " + "'" + vendasPassadas.ToString("MM-dd-yyyy") + "'" + "AND NOT o.ItemID = 'PENVIO' AND o.ItemID = sp.ItemID AND sp.PriceLineID = 0 ORDER BY x.Description ASC  ";
            commandProdutos = new SqlCommand(sqlQueryProdutos, conn);
            dataReaderProdutos = commandProdutos.ExecuteReader();


            //tabela de produtos por familia
            DataTable tabelaProdutos = new DataTable("tabelaProdutos");
            DataColumn cnp = new DataColumn("cnp");
            DataColumn produtoDesc = new DataColumn("descricao");
            DataColumn stock = new DataColumn("stock atual");
            DataColumn vendas = new DataColumn("vendas 30 dias");
            DataColumn idFamilia2 = new DataColumn("id familia");
            DataColumn familia2 = new DataColumn("familia");
            DataColumn pvpAtual = new DataColumn("PVP");
            DataColumn margemPercentagem = new DataColumn("margem percentagem");
            DataColumn margemValor = new DataColumn("margem euros");
            DataColumn pcuSage = new DataColumn("pcu sage");

            tabelaProdutos.Columns.Add(cnp);
            tabelaProdutos.Columns.Add(produtoDesc);
            tabelaProdutos.Columns.Add(stock);
            tabelaProdutos.Columns.Add(vendas);
            tabelaProdutos.Columns.Add(idFamilia2);
            tabelaProdutos.Columns.Add(familia2);

            tabelaProdutos.Columns.Add(pvpAtual);
            tabelaProdutos.Columns.Add(margemPercentagem);
            tabelaProdutos.Columns.Add(margemValor);

            tabelaProdutos.Columns.Add(pcuSage);

            int linhasProdutos = 0;

            while (dataReaderProdutos.Read())
            {
                DataRow linhaNova = tabelaProdutos.NewRow();
                linhaNova["cnp"] = dataReaderProdutos.GetValue(0);

                string cnpBusca = dataReaderProdutos.GetValue(0).ToString();
                
                linhaNova["descricao"] = dataReaderProdutos.GetValue(2);
                linhaNova["stock atual"] = dataReaderProdutos.GetValue(3);
                linhaNova["id familia"] = dataReaderProdutos.GetValue(1);
                linhaNova["familia"] = dataReaderProdutos.GetValue(5);

                //debug
                //Console.WriteLine("CNP:"+ dataReaderProdutos.GetValue(0) + " PREÇO:"+ dataReaderProdutos.GetValue(12));

                string pcuSageString = dataReaderProdutos.GetValue(12).ToString();
                float pcuFloat = float.Parse(pcuSageString);
                
                //debug
                //Console.WriteLine("CNP:"+ cnpBusca+" - PCU:" + pcuFloat);

                // Loop através de todos os produtos da lista xml para obter o pvp atual do produto
                foreach (Product product in products.ProductList)
                {
                    //Console.WriteLine("Produto: " + product.IdProduct);
                    //Console.WriteLine("Produto: " + product.Designation);
                    //Console.WriteLine("Stock: " + product.Stock);
                    //Console.WriteLine("Preço: " + product.SalePrice);
                    //Console.WriteLine();

                    if (product.IdProduct == cnpBusca)
                    {
                        pvpXmlString = product.SalePrice;

                        pvpFinalString = pvpXmlString.Replace(" EUR", "").Trim(); // Remover " EUR" e espaços em branco
                        
                        pvpFinal =(float) double.Parse(pvpFinalString, CultureInfo.InvariantCulture);
                        break;

                    }
                    else
                    {
                        pvpFinal = 0;
                    }
                }


                //calcula a margem de lucro em %
               // margemPerc = (float) ((pvpFinal / 1.23) - pcuFloat) / ((pvpFinal / ((23 / 100) + 1))) * 100;

                margemPerc = (float) (((pvpFinal / 1.23) - pcuFloat) / (pvpFinal / 1.23) * 100);

                //calcula a margem em €
                margemValorFloat =(float) (pvpFinal - (pcuFloat + (pvpFinal - (pvpFinal / 1.23))));
                   
                //debug
                Console.WriteLine("PCU:" + pcuFloat + " PVP" + pvpFinal + " Margem percentagem:" + margemPerc + " Margem euros:" + margemValorFloat);

                linhaNova["PVP"] = pvpFinalString;
                linhaNova["margem percentagem"] = margemPerc.ToString("F2");
                linhaNova["margem euros"] = margemValorFloat.ToString("F2");
                
                tabelaProdutos.Rows.Add(linhaNova);

                linhasProdutos++;
            }

            dataReaderProdutos.Close();
            conn.Close();

            //conta os produtos vendidos nos ultimos 30 dias
            DataTable tabelaFinal = tabelaProdutos.AsEnumerable()
             .GroupBy(r => new { familia = r["familia"], cnp = r["cnp"], descricao = r["descricao"], stockAtual = r["stock atual"], margemAplicadaPerc = r["margem percentagem"], margemAplicadaValor = r["margem euros"] })
                .Select(g =>
                {
                    var row = tabelaProdutos.NewRow();

                    row["familia"] = g.Key.familia;
                    row["CNP"] = g.Key.cnp;
                    row["Descricao"] = g.Key.descricao;
                    row["Stock atual"] = g.Key.stockAtual;
                    row["margem percentagem"] = g.Key.margemAplicadaPerc;
                    row["margem euros"] = g.Key.margemAplicadaValor;
                    row["Vendas 30 dias"] = g.Count();

                    return row;

                }).CopyToDataTable();

            DataTable tabelaFinalOrdenada = tabelaFinal.AsEnumerable()
            .OrderBy(r => r["familia"]) // Ordenar por ordem alfabética da coluna "familia"
            .ThenByDescending(r => float.Parse(r["margem percentagem"].ToString())) // Ordenar dentro de cada família pela coluna "margemAplicadaPerc" de forma decrescente
            .CopyToDataTable();
             

            ////////// CRIAR PDF //////////////////////////////////////////////////////////////////////////////////////////////////////////////

            // Must have write permissions to the path folder
            PdfWriter writer = new PdfWriter("C:\\sendReports\\listagem-stock.pdf");
            PdfDocument pdf = new PdfDocument(writer);
            pdf.SetDefaultPageSize(PageSize.A4.Rotate());

            Document document = new Document(pdf);
            Paragraph header = new Paragraph("ZONPHARMA")
               .SetTextAlignment(TextAlignment.CENTER).SetFontColor(ColorConstants.BLACK)
               .SetFontSize(20);
            Paragraph header2 = new Paragraph("Gestão de encomendas a fornecedores")
               .SetTextAlignment(TextAlignment.CENTER)
               .SetFontSize(14);
            Paragraph header3 = new Paragraph("Vendas dos últimos 30 dias e stock disponível inferior a 10 unidades.")
               .SetTextAlignment(TextAlignment.CENTER)
               .SetFontSize(8);

            document.Add(header);
            document.Add(header2);
            document.Add(header3);

            Table tabelaPdf = new Table(7, false);
      

                //cabeçalhos
                Cell familiaCell = new Cell(1, 1)
               .SetBackgroundColor(ColorConstants.ORANGE)
               .SetTextAlignment(TextAlignment.CENTER).SetFontSize(7)
               .Add(new Paragraph("Família"));
            Cell cnpCell = new Cell(1, 1)
               .SetBackgroundColor(ColorConstants.ORANGE).SetFontSize(7)
               .SetTextAlignment(TextAlignment.CENTER)
               .Add(new Paragraph("CNP"));
            Cell descricaoCell = new Cell(1, 1)
               .SetBackgroundColor(ColorConstants.ORANGE).SetFontSize(7)
               .SetTextAlignment(TextAlignment.CENTER)
               .Add(new Paragraph("Descrição"));
            Cell stockCell = new Cell(1, 1)
               .SetBackgroundColor(ColorConstants.ORANGE).SetFontSize(7)
               .SetTextAlignment(TextAlignment.CENTER)
               .Add(new Paragraph("Stock atual"));
            Cell vendasCell = new Cell(1, 1)
               .SetBackgroundColor(ColorConstants.ORANGE).SetFontSize(7)
               .SetTextAlignment(TextAlignment.CENTER)
               .Add(new Paragraph("Unidades vendidas"));
            Cell MargemPercCell = new Cell(1, 1)
               .SetBackgroundColor(ColorConstants.ORANGE).SetFontSize(7)
               .SetTextAlignment(TextAlignment.CENTER)
               .Add(new Paragraph("Margem %"));
            Cell MargemValorCell = new Cell(1, 1)
               .SetBackgroundColor(ColorConstants.ORANGE).SetFontSize(7)
               .SetTextAlignment(TextAlignment.CENTER)
               .Add(new Paragraph("Margem €"));

            tabelaPdf.AddCell(familiaCell);
            tabelaPdf.AddCell(cnpCell);
            tabelaPdf.AddCell(descricaoCell);
            tabelaPdf.AddCell(stockCell);
            tabelaPdf.AddCell(vendasCell);
            tabelaPdf.AddCell(MargemPercCell);
            tabelaPdf.AddCell(MargemValorCell);

            string previousFamily = null;
            bool isGrayBackground = true;

            /*
            foreach (DataRow row in tabelaFinalOrdenada.Rows)
            {
                string currentFamily = row["familia"].ToString();

                if (previousFamily != null && currentFamily != previousFamily)
                {
                    isGrayBackground = !isGrayBackground; // Alternar a cor de fundo ao mudar de família
                }

                Color backgroundColor = isGrayBackground ? ColorConstants.LIGHT_GRAY : ColorConstants.WHITE;

                // Adicionar as células à tabela
                for (int i = 0; i < tabelaFinalOrdenada.Columns.Count; i++)
                {
                    Cell cell = new Cell(1,1).SetFontSize(7)
               .SetTextAlignment(TextAlignment.CENTER).SetBackgroundColor(backgroundColor).Add(new Paragraph(row[i].ToString()));
                                          
                    tabelaPdf.AddCell(cell);
                }

                previousFamily = currentFamily;
            }
            */

            
            foreach (DataRow row in tabelaFinalOrdenada.Rows)
            {
                string currentFamily = row["familia"].ToString();

                if (previousFamily != null && currentFamily != previousFamily)
                {
                    isGrayBackground = !isGrayBackground; // Alternar a cor de fundo ao mudar de família
                }

                Color backgroundColor = isGrayBackground ? ColorConstants.WHITE : ColorConstants.LIGHT_GRAY;


                Cell cell1 = new Cell(1, 1)
                .SetTextAlignment(TextAlignment.LEFT).SetFontSize(7).SetBackgroundColor(backgroundColor)
                .Add(new Paragraph(row["familia"].ToString()));

                Cell cell2 = new Cell(1, 1)
                .SetTextAlignment(TextAlignment.LEFT).SetFontSize(7).SetBackgroundColor(backgroundColor)
                .Add(new Paragraph(row["CNP"].ToString()));

                Cell cell3 = new Cell(1, 1).SetTextAlignment(TextAlignment.LEFT).SetFontSize(7).SetBackgroundColor(backgroundColor)
                .Add(new Paragraph(row["Descricao"].ToString()));

                Cell cell4 = new Cell(1, 1).SetTextAlignment(TextAlignment.LEFT).SetFontSize(7).SetBackgroundColor(backgroundColor)
                .Add(new Paragraph(row["Stock atual"].ToString()));

                Cell cell5 = new Cell(1, 1).SetTextAlignment(TextAlignment.LEFT).SetFontSize(7).SetBackgroundColor(backgroundColor)
                .Add(new Paragraph(row["Vendas 30 dias"].ToString()));

                Cell cell6 = new Cell(1, 1).SetTextAlignment(TextAlignment.LEFT).SetFontSize(7).SetBackgroundColor(backgroundColor)
                .Add(new Paragraph(row["Margem percentagem"].ToString()+"%"));

                Cell cell7 = new Cell(1, 1).SetTextAlignment(TextAlignment.LEFT).SetFontSize(7).SetBackgroundColor(backgroundColor)
                .Add(new Paragraph(row["Margem euros"].ToString()+"€"));

                tabelaPdf.AddCell(cell1);
                tabelaPdf.AddCell(cell2);
                tabelaPdf.AddCell(cell3);
                tabelaPdf.AddCell(cell4);
                tabelaPdf.AddCell(cell5);
                tabelaPdf.AddCell(cell6);
                tabelaPdf.AddCell(cell7);

                previousFamily = currentFamily;
                //debug
                //Console.WriteLine("Família:" + row["familia"] + " - " + "CNP:" + row["CNP"] + " - " + "Descrição:" + row["Descricao"] + " - " + "Stock atual:" + row["Stock atual"] + " - " + "Vendas 30 dias:" + row["Vendas 30 dias"] + " - " + "Margem %:" + row["Margem percentagem"] + " - " + "Margem Euros:" + row["Margem euros"]);

            }

            document.Add(tabelaPdf);
            document.Close();

            Thread.Sleep(3000);


            //envia email
            
            Console.WriteLine("A enviar email com relatório...");
            try
            {
                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient("mail.zonpharma.com");
                mail.From = new MailAddress("ajuda@zonpharma.com");
                mail.To.Add("Central.amarante@hotmail.com"); //production
               // mail.To.Add("samuel.parente@gmail.com"); //debug
                mail.CC.Add("zonkarma5@gmail.com");
                mail.Subject = "Listagem - stocks e vendas - "+ justDate.ToString("dd-MM-yyyy");
                mail.Body = "Gestão de encomendas a fornecedores\r\nVendas dos últimos 30 dias e stock disponível inferior a 10 unidades.\r\n ESTE EMAIL É GERADO AUTOMATICAMENTE.\r\nFrequência de envio: Diário.\r\n\nSoftware desenvolvido por Samuel Parente\r\n\nTodos os direitos reservados @Zonpharma 2023";

                System.Net.Mail.Attachment attachment;
                attachment = new System.Net.Mail.Attachment("C:\\sendReports\\listagem-stock.pdf");
                mail.Attachments.Add(attachment);

                SmtpServer.Port = 25;
                SmtpServer.Credentials = new System.Net.NetworkCredential("ajuda@zonpharma.com", "#pharma4ALL#");
                SmtpServer.EnableSsl = false;

                SmtpServer.Send(mail);
                Console.WriteLine("Email enviado.");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                Console.WriteLine("Email nao enviado.");
            }

            

            Thread.Sleep(3000);

        }
    }
}
