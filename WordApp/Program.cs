using System;

namespace WordApp
{
    class Program
    {
        static void Main(string[] args)
        {
            WordDocument word = new WordDocument();
            word.SuplyerName = "GBK Unioin";
            word.SuplyerINN = "23456789";
            word.SuplyerKPP = "2345678908765432";
            word.CompanyName = "gazStroyServis";
            word.CompanyINN = "123456543";
            word.CompanyKPP = "234532";
            word.SnabName = "Igor";
            word.SnabPhone = "89600655315";
            word.Email = "mr.shaykheraziev@mail.ru";
            word.Create();
         
        }
    }
}
