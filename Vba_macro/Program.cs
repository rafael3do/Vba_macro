using System;
using System.IO;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//Software criadado para a adequação em linhas de uma planilha ja pronta.
namespace Vba_macro
{
    class Program
    {
        static void Main(string[] args)
        {//abaixo inicialização das variaveis
            string Pasta = @"C:\Temp";
            int Linhas = 0;
            int Val_teste = 1;
            int Sub_ciclo = 0;
            string Num_linhas = "";
            int Valida_escrita = 0;
            int escrevi = 1;
            if (!File.Exists(Pasta))//valida se tem a pasta temp caso contrario ele cria
            {
                System.IO.Directory.CreateDirectory(Pasta);
            }

            FileStream Arquivo = new FileStream(@"C:\Temp\Codigo_vba_macro.txt", FileMode.Create);
            StreamWriter Insere = new StreamWriter(Arquivo);//IO de escrita
                       
            while (Linhas==0)//valida as linhas digitadas
            {               
                Console.WriteLine("Digite a quantidade de linhas na planilha.");
                Num_linhas = Console.ReadLine();
                bool Valida_numero = int.TryParse(Num_linhas, out Linhas);
                if (Valida_numero==false) {
                    Console.WriteLine("Numero invalido.");
                }
                else
                {
                    Insere.WriteLine("Sub Controle_vba");
                    Insere.WriteLine(" ");
                }
            }

            for (int i = 0; i <= Linhas; i += 200)//Crias as sub consultas do vba
            {
                Insere.WriteLine("Call Ciclo" + Val_teste);
                Insere.WriteLine(" ");
                Val_teste += 1;
            }

            Insere.WriteLine("End Sub");
            Insere.WriteLine(" ");
            
            for (Sub_ciclo = 1; Sub_ciclo < Linhas; Sub_ciclo++)//Digita a linha dentro das sub consultas do vba
            {
                Contador();
                Insere.WriteLine("Range(\"E" + (Sub_ciclo + 1) + ":H" + (Sub_ciclo + 1) + "\").Select");
                Insere.WriteLine("Application.CutCopyMode = False");
                Insere.WriteLine("Selection.Cut");
                Insere.WriteLine("Rows(\"" + (1 + (2 * Sub_ciclo)) + ":" + (1 + (2 * Sub_ciclo) + "\").Select"));
                Insere.WriteLine("Selection.Insert Shift:=xlDown");
                Insere.WriteLine(" ");
                Valida_escrita++;
            }
            
            void Contador()//Funçao dentro da classe que inicia o espaço para digitar as linhas organizando nas funções
            {
                if (Valida_escrita==0) {
                    for (int i = 0; i < 1; i++)
                    {
                        Insere.WriteLine(" ");
                        Insere.WriteLine("Sub Ciclo"+ escrevi);
                        Insere.WriteLine("Dim i As Integer");
                        Insere.WriteLine("For i = 1 To 1");
                        Insere.WriteLine(" ");
                    }
                    escrevi += 1;
                    Valida_escrita = 1;
                }
                if (Valida_escrita==200)
                {
                    Insere.WriteLine("Next i");
                    Insere.WriteLine("End Sub");
                    Insere.WriteLine(" ");
                    Insere.WriteLine("Sub Ciclo" + escrevi);
                    Insere.WriteLine("Dim i As Integer");
                    Insere.WriteLine("For i = 1 To 1");
                    Insere.WriteLine(" ");
                    escrevi += 1;
                    Valida_escrita = 1;
                }

            }
            Insere.WriteLine("Next i");
            Insere.WriteLine(" ");
            Insere.WriteLine("End Sub");
            Insere.Close();
            Arquivo.Close();
            Console.WriteLine("Gerado o codigo da macro em C:\\Temp\\Codigo_vba_macro.txt");
            System.Diagnostics.Process.Start("C:\\Temp\\Codigo_vba_macro.txt");
        }

    }
}