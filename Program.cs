using System;
using Microsoft.Office.Interop.Word;

namespace CreatePDF
{
    class Program
    {
        static void Main(string[] args)
        {
            new EditorWord().PropostaVendaProduto();
        }
    }
}