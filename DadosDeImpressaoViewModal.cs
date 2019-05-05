

using System.Collections.Generic;

namespace CreatePDF
{
    public class DadosDeImpressaoViewModal
    {
        public string NomeEmpresa { get; set; }
        public string TelefoneEmpresa { get; set; }
        public string EmailEmpresa { get; set; }
        public string EnderecoEmpresa { get; set; }
        public List<Contatos> Contatos { get; set; }
    }
    public class Contatos
    {
        public string NomeContato { get; set; }
        public string TelefoneContato { get; set; }
        public string EnderecoContato { get; set; }
        public string EmailContato { get; set; }
    }
}
