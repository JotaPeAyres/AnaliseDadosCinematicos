using AnalisarDados;
using System;
using System.ComponentModel.DataAnnotations;
using System.IO;

namespace ValidadorExcel;

class Program
{
    static void Main()
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        Console.WriteLine("===========================================");
        Console.WriteLine("       Validador de Excel - ClosedXML");
        Console.WriteLine("===========================================\n");

        Console.Write("Informe o caminho da pasta raiz: ");
        string pastaRaiz = Console.ReadLine()!;

        if (!Directory.Exists(pastaRaiz))
        {
            Console.WriteLine("Pasta não encontrada!");
            return;
        }

        var processor = new Analisador();
        var resultados = processor.ProcessarTodosExcels(pastaRaiz);

        var medias = processor.CalcularMedia(resultados);
        string[] atividades = { "SLDL", "SLHFD", "SLLV" };

        foreach (var atividade in atividades)
        {
            var resultadosAtividade = resultados
               .Where(r => r.NomeArquivo.StartsWith(atividade))
               .ToList();

            var mediasAtividade = medias
                .Where(m => m.NomeArquivo.StartsWith(atividade))
                .ToList();

            string arquivoSaida = Path.Combine(pastaRaiz, $"{atividade}_Resultados.xlsx");
            processor.ExportarResultadosExcel(resultadosAtividade, mediasAtividade, arquivoSaida);
        }


        Console.WriteLine($"\n✅ Validação concluída!");
    }
}
