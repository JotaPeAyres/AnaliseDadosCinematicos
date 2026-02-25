using AnalisarDados;
using System;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics;
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

        var pastasID = Directory.GetDirectories(pastaRaiz, "ID_*")
                                .OrderBy(p => p)
                                .ToList();

        var analisador = new Analisador();
        string[] atividades = Constantes.Atividades.Todas;

        foreach (string pastaRaizId in pastasID)
        {
            string voluntarioId = Path.GetFileName(pastaRaizId);

            Console.WriteLine($"\n--- Processando: {voluntarioId} ---");

            try
            {
                var resultados = analisador.ProcessarTodosExcels(pastaRaizId);
                var medias = analisador.CalcularMedia(resultados);

                foreach (var atividade in atividades)
                {
                    var resultadosAtividade = resultados
                       .Where(r => r.NomeArquivo.StartsWith(atividade))
                       .ToList();

                    var mediasAtividade = medias
                        .Where(m => m.NomeArquivo.StartsWith(atividade))
                        .ToList();

                    if (resultadosAtividade.Any())
                    {
                        string arquivoSaida = Path.Combine(pastaRaizId, $"{atividade}_Resultados_{voluntarioId}.xlsx");
                        analisador.ExportarResultadosExcel(resultadosAtividade, mediasAtividade, arquivoSaida);
                        Console.WriteLine($"[OK] {atividade} exportado.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ERRO] Falha ao processar {voluntarioId}: {ex.Message}");
            }
        }

        Console.WriteLine($"\n✅ Processamento global concluído!");
    }
}