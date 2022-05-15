using System;
using System.IO;
// Software criadado para a adequação em linhas de uma planilha ja pronta.
// Software designed to fit rows into a ready-made spreadsheet.
namespace Vba_macro
{
    class Program
    {
        static void Main(string[] args)
        {//abaixo inicialização das variaveis.
            //below variable initialization.
            string pasta = @"C:\Temp";
            int linhas = 0;
            int val_teste = 1;
            int sub_ciclo = 0;
            string num_linhas = "";
            int valida_escrita = 0;
            int escrevi = 1;

            if (!File.Exists(pasta))//valida se tem a pasta temp caso contrario ele cria.
                                    // validate if it has temp folder otherwise it creates.
            {
                System.IO.Directory.CreateDirectory(pasta);
            }
            FileStream Arquivo = new FileStream(@"C:\Temp\Codigo_vba_macro.txt", FileMode.Create);
            StreamWriter Insere = new StreamWriter(Arquivo);//IO de escrita.
                                                            // IO write.
            while (linhas == 0)//valida as linhas digitadas.
                               // validate the typed lines.
            {
                Console.WriteLine("Digite a quantidade de linhas na planilha.");
                num_linhas = Console.ReadLine();
                bool Valida_numero = int.TryParse(num_linhas, out linhas);
                if (Valida_numero == false)
                {
                    Console.WriteLine("Numero invalido.");
                }
                else
                {
                    Insere.WriteLine("Sub Controle_vba");
                    Insere.WriteLine(" ");
                }
            }

            for (int i = 0; i <= linhas; i += 200)//Crias as sub consultas do vba.
                                                  // Create vba sub queries.
            {
                Insere.WriteLine("Call Ciclo" + val_teste);
                Insere.WriteLine(" ");
                val_teste += 1;
            }

            Insere.WriteLine("End Sub");
            Insere.WriteLine(" ");
            for (sub_ciclo = 1; sub_ciclo < linhas; sub_ciclo++)//Digita a linha dentro das sub consultas do vba.
                                                                // Type the line inside vba sub queries.
            {
                Contador();
                Insere.WriteLine("Range(\"E" + (sub_ciclo + 1) + ":H" + (sub_ciclo + 1) + "\").Select");
                Insere.WriteLine("Application.CutCopyMode = False");
                Insere.WriteLine("Selection.Cut");
                Insere.WriteLine("Rows(\"" + (1 + (2 * sub_ciclo)) + ":" + (1 + (2 * sub_ciclo) + "\").Select"));
                Insere.WriteLine("Selection.Insert Shift:=xlDown");
                Insere.WriteLine(" ");
                valida_escrita++;
            }

            void Contador()//Funçao dentro da classe que inicia o espaço para digitar as linhas organizando nas funções.
                           // Function within class that starts the space to type lines by arranging in functions.
            {
                if (valida_escrita == 0)
                {
                    for (int i = 0; i < 1; i++)
                    {
                        Insere.WriteLine(" ");
                        Insere.WriteLine("Sub Ciclo" + escrevi);
                        Insere.WriteLine("Dim i As Integer");
                        Insere.WriteLine("For i = 1 To 1");
                        Insere.WriteLine(" ");
                    }
                    escrevi += 1;
                    valida_escrita = 1;
                }

                if (valida_escrita == 200)
                {
                    Insere.WriteLine("Next i");
                    Insere.WriteLine("End Sub");
                    Insere.WriteLine(" ");
                    Insere.WriteLine("Sub Ciclo" + escrevi);
                    Insere.WriteLine("Dim i As Integer");
                    Insere.WriteLine("For i = 1 To 1");
                    Insere.WriteLine(" ");
                    escrevi += 1;
                    valida_escrita = 1;
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
}//Finish progam.