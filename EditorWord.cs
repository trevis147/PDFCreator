using System;
using Microsoft.Office.Interop.Word;

namespace CreatePDF
{
    public class EditorWord
    {
        public void PropostaVendaServico(DadosDeImpressaoViewModal dados)
        {
            string word = CreateTable("C:\\Dev\\CreatePDF\\CreatePDF\\ArquivoTeste.docx");
            var oWord = new Application();
            object oMissing = System.Reflection.Missing.Value;
            var oWordDoc = new Document();
            oWord.Visible = false;

            object oTemplatePath = @"" + word;
            oWordDoc = oWord.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);

            try
            {
                //cabeçalho
                oWordDoc.Bookmarks.get_Item("NomeEmpresa").Range.Text = dados.NomeEmpresa;
                oWordDoc.Bookmarks.get_Item("TelEmpresa").Range.Text = dados.TelefoneEmpresa;
                oWordDoc.Bookmarks.get_Item("EmailEmpresa").Range.Text = dados.EmailEmpresa;
                oWordDoc.Bookmarks.get_Item("EnderecoEmpresa").Range.Text = dados.EnderecoEmpresa;



                var table = oWordDoc.Bookmarks.get_Item("table").Range;

                int j = 0;
                var rnd = new Random();
                for (int i = 1; i < 10; i++)
                {
                    table.Rows[i + 1].Cells[1].Range.Text = ((int)j++).ToString();
                    table.Rows[i + 1].Cells[2].Range.Text = "grupo :" + rnd.Next(0, 2000).ToString();
                    table.Rows[i + 1].Cells[3].Range.Text = "serviço :" + rnd.Next(0, 2000).ToString();
                    table.Rows[i + 1].Cells[4].Range.Text = "descricao :" + rnd.Next(0, 2000).ToString();
                    table.Rows[i + 1].Cells[5].Range.Text = "subtotal :" + rnd.Next(0, 2000).ToString();
                }

                oWordDoc.SaveAs2(@"C:\Dev\CreatePDF\CreatePDF\PropostaComercialpdf" + DateTime.Now.Millisecond + ".pdf", 17);
                oWordDoc.Close(null, null, null);
                oWord.Quit();
                System.IO.File.Delete(@"" + word);
            }
            catch (Exception e)
            {
                oWord.Quit();
                Console.WriteLine("Deu problema!!!!!!");
                Console.WriteLine(e.ToString());

                Console.ReadKey();
            }
        }
        public void PropostaVendaProduto()
        {
            string word = CreateTable("C:\\Dev\\CreatePDF\\CreatePDF\\PropostaComercialProdutos.docx");
            var oWord = new Application();
            object oMissing = System.Reflection.Missing.Value;
            var oWordDoc = new Document();
            oWord.Visible = false;

            object oTemplatePath = @""+ word;
            oWordDoc = oWord.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);

            try
            {
                //cabeçalho
                oWordDoc.Bookmarks.get_Item("Nome").Range.Text = "Nome sobrenome";
                oWordDoc.Bookmarks.get_Item("Tel").Range.Text = "(11)-5555-5555";
                oWordDoc.Bookmarks.get_Item("email").Range.Text = "email@empresa.com";
                oWordDoc.Bookmarks.get_Item("prop").Range.Text = "17663";

                //pagina-01
                oWordDoc.Bookmarks.get_Item("Empresa").Range.Text = "empresa cli";
                oWordDoc.Bookmarks.get_Item("Tipo").Range.Text = "Remanejamento ";
                oWordDoc.Bookmarks.get_Item("data").Range.Text = "10/10/2018";

                //pagina-02  
                oWordDoc.Bookmarks.get_Item("DadosClienteFaturamento").Range.Text = "Dados de faturamento";
                oWordDoc.Bookmarks.get_Item("DadosEntrega").Range.Text = "dados de entrega";
                oWordDoc.Bookmarks.get_Item("DadosClienteFaturamentoRS").Range.Text = "Dados de faturamento";
                oWordDoc.Bookmarks.get_Item("DadosEntregaRS").Range.Text = "Dados de entrega";
                oWordDoc.Bookmarks.get_Item("TotalUS").Range.Text = "US$ 1000,00";
                oWordDoc.Bookmarks.get_Item("TotalRS").Range.Text = "R$ 5000,00";
                oWordDoc.Bookmarks.get_Item("TotalGarantia").Range.Text = "R$ 900,00";


                //pagina-03
                oWordDoc.Bookmarks.get_Item("validade").Range.Text = "7 dias";
                oWordDoc.Bookmarks.get_Item("condicoes").Range.Text = "60 dd";
                oWordDoc.Bookmarks.get_Item("frete").Range.Text = "60 dd";
                oWordDoc.Bookmarks.get_Item("PTAX").Range.Text = "3,81";
                oWordDoc.Bookmarks.get_Item("obs").Range.Text = "";

                //pagina-04
                oWordDoc.Bookmarks.get_Item("Descricaodosservicos").Range.Text = "descricao";

                //rodapé
                oWordDoc.Bookmarks.get_Item("NomeEmpresa").Range.Text = "Nome Sobrenome";
                oWordDoc.Bookmarks.get_Item("EmailEmpresa").Range.Text = "email@Empresa.com";
                oWordDoc.Bookmarks.get_Item("TelEmpresa").Range.Text = "(11)-5555-55555";
                oWordDoc.Bookmarks.get_Item("CelEmpresa").Range.Text = "(11)-95555-5555";

                var tableUS = oWordDoc.Bookmarks.get_Item("table01").Range;


                int j = 0;
                var rnd = new Random();
                for (int i = 1; i < 10; i++)
                {
                    tableUS.Rows[i + 1].Cells[1].Range.Text = ((int)j++).ToString();
                    tableUS.Rows[i + 1].Cells[2].Range.Text = "grupo :" + rnd.Next(0, 2000).ToString();
                    tableUS.Rows[i + 1].Cells[3].Range.Text = "serviço :" + rnd.Next(0, 2000).ToString();
                    tableUS.Rows[i + 1].Cells[4].Range.Text = "descricao :" + rnd.Next(0, 2000).ToString();
                    tableUS.Rows[i + 1].Cells[5].Range.Text = "subtotal :" + rnd.Next(0, 2000).ToString();
                    tableUS.Rows[i + 1].Cells[6].Range.Text = "subtotal :" + rnd.Next(0, 2000).ToString();
                    tableUS.Rows[i + 1].Cells[7].Range.Text = "subtotal :" + rnd.Next(0, 2000).ToString();
                    tableUS.Rows[i + 1].Cells[8].Range.Text = "subtotal :" + rnd.Next(0, 2000).ToString();
                    tableUS.Rows[i + 1].Cells[9].Range.Text = "subtotal :" + rnd.Next(0, 2000).ToString();
                    tableUS.Rows[i + 1].Cells[10].Range.Text = "subtotal :" + rnd.Next(0, 2000).ToString();

                }

                var tableRS = oWordDoc.Bookmarks.get_Item("table02").Range;


                j = 0;
                for (int i = 1; i < 10; i++)
                {
                    tableRS.Rows[i + 1].Cells[1].Range.Text = ((int)j++).ToString();
                    tableRS.Rows[i + 1].Cells[2].Range.Text = "grupo :" + rnd.Next(0, 2000).ToString();
                    tableRS.Rows[i + 1].Cells[3].Range.Text = "serviço :" + rnd.Next(0, 2000).ToString();
                    tableRS.Rows[i + 1].Cells[4].Range.Text = "descricao :" + rnd.Next(0, 2000).ToString();
                    tableRS.Rows[i + 1].Cells[5].Range.Text = "subtotal :" + rnd.Next(0, 2000).ToString();
                    tableRS.Rows[i + 1].Cells[6].Range.Text = "subtotal :" + rnd.Next(0, 2000).ToString();
                    tableRS.Rows[i + 1].Cells[7].Range.Text = "subtotal :" + rnd.Next(0, 2000).ToString();
                    tableRS.Rows[i + 1].Cells[8].Range.Text = "subtotal :" + rnd.Next(0, 2000).ToString();
                    tableRS.Rows[i + 1].Cells[9].Range.Text = "subtotal :" + rnd.Next(0, 2000).ToString();
                    tableRS.Rows[i + 1].Cells[10].Range.Text = "subtotal :" + rnd.Next(0, 2000).ToString();
                }

                oWordDoc.SaveAs2(@"C:\Dev\CreatePDF\CreatePDF\PropostaComercialpdf" + DateTime.Now.Millisecond + ".pdf", 17);
                oWordDoc.Close(null,null,null);
                oWord.Quit();
                System.IO.File.Delete(@"" + word);
            }
            catch (Exception e)
            {
                oWord.Quit();
                Console.WriteLine("Deu problema!!!!!!");
                Console.WriteLine(e.ToString());

                Console.ReadKey();
            }
        }
        public string CreateTable(string word)
        {
            var oWord = new Application();
            object oMissing = System.Reflection.Missing.Value;
            var oWordDoc = new Document();
            oWord.Visible = false;

            object oTemplatePath = @"" + word;
            var rnd = new Random();
            int j = 0;
            oWordDoc = oWord.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);
            if (oWordDoc.Bookmarks.Exists("table"))
            {
                var table = oWordDoc.Bookmarks.get_Item("table").Range;
                for (int i = 0; i < 10; i++)
                    table.Rows.Add();
            }
            else
            {
                var tableRS = oWordDoc.Bookmarks.get_Item("table02").Range;
                var tableUS = oWordDoc.Bookmarks.get_Item("table01").Range;
                for (int i = 0; i < 10; i++)
                {
                    tableRS.Rows.Add();
                    tableUS.Rows.Add();
                }
            }
            string save = rnd.Next(0, 10000000).ToString();
            oWordDoc.SaveAs2(@"C:\\Dev\\CreatePDF\\CreatePDF\\" + save + ".docx");
            oWordDoc.Close(null, null, null);
            oWord.Quit();


            return "C:\\Dev\\CreatePDF\\CreatePDF\\" + save + ".docx";
        }
    }
}
