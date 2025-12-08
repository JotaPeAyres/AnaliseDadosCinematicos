using ClosedXML.Excel;
using DocumentFormat.OpenXml.InkML;
using OfficeOpenXml;
using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using ValidadorExcel;

namespace AnalisarDados;

public class Analisador
{
    // Processa todos os arquivos .xlsx em subpastas
    /// <summary>
    /// Processa todos os arquivos Excel nas subpastas da pasta raiz fornecida.
    /// </summary>
    /// <param name="pastaRaiz">Caminho da pasta raiz contendo as subpastas dos dias</param>
    /// <returns>Lista com os resultados da análise</returns>
    public List<ResultadoAnalise> ProcessarTodosExcels(string pastaRaiz)
    {
        if (string.IsNullOrWhiteSpace(pastaRaiz))
        {
            throw new ArgumentException("O caminho da pasta raiz não pode ser vazio ou nulo.", nameof(pastaRaiz));
        }

        var resultados = new List<ResultadoAnalise>();
        var cronometroGeral = Stopwatch.StartNew();

        Console.WriteLine(string.Format(Constantes.Mensagens.INICIO_PROCESSAMENTO, pastaRaiz));

        try
        {
            // Processar os dias definidos nas constantes
            for (int dia = Constantes.Processamento.DIA_INICIAL; 
                 dia < Constantes.Processamento.DIA_INICIAL + Constantes.Processamento.DIAS_PROCESSAMENTO; 
                 dia++)
            {
                Console.WriteLine(string.Format(Constantes.Mensagens.PROCESSANDO_DIA, dia));
                var resultadosDia = ProcessarPorDia(pastaRaiz, dia);
                resultados.AddRange(resultadosDia);
                
                Console.WriteLine(string.Format(Constantes.Mensagens.DIA_CONCLUIDO, dia));
                Console.WriteLine($"   - {resultadosDia.Count} registros processados no total");
            }

            Console.WriteLine(Constantes.Mensagens.PROCESSAMENTO_CONCLUIDO);
            return resultados;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Erro durante o processamento: {ex.Message}");
            throw; // Propaga a exceção para que o chamador possa lidar com ela
        }
        finally
        {
            cronometroGeral.Stop();
            Console.WriteLine(string.Format(Constantes.Mensagens.TEMPO_EXECUCAO, cronometroGeral.Elapsed));
        }
    }

    private List<ResultadoAnalise> ProcessarPorDia(string pastaRaiz, int dia)
    {
        var resultados = new List<ResultadoAnalise>();

        // Pastas principais a procurar dentro de cada dia
        string[] pastasPrincipais = { 
            string.Format(Constantes.Pastas.PADRAO_PASTA_CINEMATICA, dia), 
            string.Format(Constantes.Pastas.PADRAO_PASTA_MARCADORES, dia) 
        };

        // Subpastas que queremos processar
        string[] atividades = Constantes.Atividades.Todas;

        List<string> arquivosDireita = new List<string>();
        List<string> arquivosEsquerda = new List<string>();

        foreach (var principal in pastasPrincipais)
        {
            string caminhoPrincipal = Path.Combine(pastaRaiz, principal);
            if (!Directory.Exists(caminhoPrincipal))
            {
                Console.WriteLine(string.Format(Constantes.Mensagens.PASTA_NAO_ENCONTRADA, caminhoPrincipal));
                continue;
            }
            Console.WriteLine(string.Format(Constantes.Mensagens.EXPLORANDO_PASTA, caminhoPrincipal));

            // Procurar somente pelas subpastas válidas
            foreach (var atividade in atividades)
            {
                string caminhoSubpasta = Path.Combine(caminhoPrincipal, atividade);
                if (!Directory.Exists(caminhoSubpasta))
                {
                    Console.WriteLine(string.Format(Constantes.Mensagens.SUBPASTA_NAO_ENCONTRADA, atividade, caminhoSubpasta));
                    continue;
                }
                Console.WriteLine(string.Format(Constantes.Mensagens.SUBPASTA_ENCONTRADA, atividade, caminhoSubpasta));

                arquivosDireita.AddRange(Directory.GetFiles(caminhoSubpasta, Constantes.Pastas.PADRAO_ARQUIVO_DIREITA, SearchOption.TopDirectoryOnly));
                arquivosEsquerda.AddRange(Directory.GetFiles(caminhoSubpasta, Constantes.Pastas.PADRAO_ARQUIVO_ESQUERDA, SearchOption.TopDirectoryOnly));
            }
        }

        // Processar cada tipo de atividade
        foreach (var atividade in atividades)
        {
            Console.WriteLine(string.Format(Constantes.Mensagens.INICIANDO_ATIVIDADE, atividade));
            
            // Filtrar e processar arquivos do lado direito
            var arquivosDireitaAtual = arquivosDireita
                .Where(w => w.Contains(atividade))
                .ToList();
                
            // Filtrar e processar arquivos do lado esquerdo
            var arquivosEsquerdaAtual = arquivosEsquerda
                .Where(w => w.Contains(atividade))
                .ToList();

            switch (atividade)
            {
                case Constantes.Atividades.SLDL:
                    Console.WriteLine(string.Format(Constantes.Mensagens.PROCESSANDO_ATIVIDADE, Constantes.Atividades.SLDL));
                    if (arquivosDireitaAtual.Any()) 
                        resultados.AddRange(ProcessarSLDL(arquivosDireitaAtual, Constantes.Lados.DIREITA, dia));
                    if (arquivosEsquerdaAtual.Any())
                        resultados.AddRange(ProcessarSLDL(arquivosEsquerdaAtual, Constantes.Lados.ESQUERDA, dia));
                    break;

                case Constantes.Atividades.SLHFD:
                    Console.WriteLine(string.Format(Constantes.Mensagens.PROCESSANDO_ATIVIDADE, Constantes.Atividades.SLHFD));
                    if (arquivosDireitaAtual.Any())
                        resultados.AddRange(ProcessarSLHFD(arquivosDireitaAtual, Constantes.Lados.DIREITA, dia));
                    if (arquivosEsquerdaAtual.Any())
                        resultados.AddRange(ProcessarSLHFD(arquivosEsquerdaAtual, Constantes.Lados.ESQUERDA, dia));
                    break;

                case Constantes.Atividades.SLLV:
                    Console.WriteLine(string.Format(Constantes.Mensagens.PROCESSANDO_ATIVIDADE, Constantes.Atividades.SLLV));
                    // Implementar processamento de SLLV quando necessário
                    // if (arquivosDireitaAtual.Any())
                    //     resultados.AddRange(ProcessarSLLV(arquivosDireitaAtual, Constantes.Lados.DIREITA, dia));
                    // if (arquivosEsquerdaAtual.Any())
                    //     resultados.AddRange(ProcessarSLLV(arquivosEsquerdaAtual, Constantes.Lados.ESQUERDA, dia));
                    break;
            }
        }

        Console.WriteLine($"✅ Concluído processamento do Dia {dia}.");
        return resultados;
    }

    private List<ResultadoAnalise> ProcessarSLHFD(List<string> caminhos, string lado, int dia)
    {
        Console.WriteLine($"\n📊 Processando {Constantes.Atividades.SLHFD} - Lado: {lado}");

        var resultados = new List<ResultadoAnalise>();
        var ladoDireito = lado == Constantes.Lados.DIREITA;

        if (caminhos == null || caminhos.Count == 0)
        {
            Console.WriteLine("⚠️ Nenhum arquivo encontrado para processar SLHFD.");
            return new List<ResultadoAnalise>();
        }

        Console.WriteLine($"🧾 Total de arquivos encontrados: {caminhos.Count}");

        var arquivoMarcadores = caminhos.FirstOrDefault(f => f.Contains("Marcadores"));
        var arquivoCinematica = caminhos.FirstOrDefault(f => f.Contains("Cinematica"));


        if (arquivoMarcadores == null || arquivoCinematica == null)
        {
            Console.WriteLine("⚠️ Arquivos de Marcadores ou Cinemática não encontrados para SLHFD.");
            return new List<ResultadoAnalise>();
        }

        using var wbMarcador = new XLWorkbook(arquivoMarcadores);
        var planilhasMarcador = wbMarcador.Worksheets.ToList();

        using var wbCinematica = new XLWorkbook(arquivoCinematica);
        var planilhasCinematica = wbCinematica.Worksheets.ToList();

        for (var i = 1; i <= 5; i++)
        {
            Console.WriteLine(string.Format(Constantes.Mensagens.ANALISANDO_TENTATIVA, i));

            var planilhaM = planilhasMarcador.FirstOrDefault(f => f.Name.Contains(i.ToString()));
            var planilhaC = planilhasCinematica.FirstOrDefault(f => f.Name.Contains(i.ToString()));

            if (planilhaC is null || planilhaM is null)
            {
                Console.WriteLine(string.Format("Planilhas não encontradas na tentativa {0}", i));
                continue;
            }

            Console.WriteLine(Constantes.Mensagens.LENDO_DADOS);

            #region Dados 

            var tempoMarcadores = planilhaM.Column(Constantes.ColunasMarcadores.TEMPO)
                   .CellsUsed()
                   .Skip(1)
                   .Select(c => c.GetValue<double>())
                   .ToList();

            var aducaoQuadril = planilhaC.Column(ladoDireito ? "I" : "P")
                   .CellsUsed()
                   .Skip(1)
                   .Select(c => c.GetValue<double>())
                   .ToList();

            var flexaoQuadril  = planilhaC.Column(ladoDireito ? "H" : "O")
                   .CellsUsed()
                   .Skip(1)
                   .Select(c => c.GetValue<double>())
                   .ToList();

            var rotacaoMedialQuadril = planilhaC.Column(ladoDireito ? "J" : "Q")
                   .CellsUsed()
                   .Skip(1)
                   .Select(c => c.GetValue<double>())
                   .ToList();

            var flexaoJoelho = planilhaC.Column(ladoDireito ? "K" : "R")
                   .CellsUsed()
                   .Skip(1)
                   .Select(c => c.GetValue<double>())
                   .ToList();

            var calcanharMarcador = planilhaM.Column(ladoDireito ? "BI" : "AZ")
                   .CellsUsed()
                   .Skip(1)
                   .Select(c => c.GetValue<double>())
                   .ToList();

            var pelveCinematica = planilhaM.Column(ladoDireito ? "BL" : "BO")
                   .CellsUsed()
                   .Skip(1)
                   .Select(c => c.GetValue<double>())
                   .ToList();

            #endregion Dados 

            Console.WriteLine("🔎 Calculando pontos de corte...");

            #region Primeiro ponto de corte

            var IndexMaximoValorCalcanharMarcador = calcanharMarcador.IndexOf(calcanharMarcador.Max());

            int pontoCorteCalcanhar = -1;
            var calcanharPontoNeutro = calcanharMarcador[0];
            for (int index = 1; index < calcanharMarcador.Count - 1; index++)
            {
                if (calcanharMarcador[index] <= calcanharPontoNeutro && index > IndexMaximoValorCalcanharMarcador)
                {
                    pontoCorteCalcanhar = index;
                    break;
                }
            }
            Console.WriteLine($"📍 Ponto de corte calcanhar: {pontoCorteCalcanhar}");

            #endregion Primeiro ponto de corte

            #region Segundo ponto de corte

            var IndexMaximoValorPelveMarcador = calcanharMarcador.IndexOf(calcanharMarcador.Max());

            // pegar somente os valores após o pico
            var valoresDepoisDoMaximoPelve = pelveCinematica
                .Skip(IndexMaximoValorPelveMarcador + 1)
                .ToList();

            var menorValor = valoresDepoisDoMaximoPelve.Min();

            // índice global do menor valor
            var pontoCortePelve = pelveCinematica.IndexOf(menorValor, IndexMaximoValorPelveMarcador + 1);

            Console.WriteLine($"📍 Ponto de corte pelve: {pontoCortePelve}");

            #endregion Segundo ponto de corte

            // Garante que os índices estão dentro dos limites
            int inicio = Math.Max(0, pontoCorteCalcanhar);
            int fim = Math.Min(pelveCinematica.Count - 1, pontoCortePelve);

            // Caso os índices estejam invertidos (por segurança)
            if (fim <= inicio)
            {
                Console.WriteLine("⚠️ Intervalo inválido entre pontoCorteCalcanhar e pontoCortePelve. Pulando tentativa...");
                continue;
            }

            // === Calcula os máximos entre os pontos de corte ===
            Console.WriteLine("📊 Calculando máximos entre os pontos de corte...");
            var maxAducaoQuadril = MaxEntre(aducaoQuadril, inicio, fim);
            var maxFlexaoQuadril = MaxEntre(flexaoQuadril, inicio, fim);
            var maxFlexaoJoelho = MaxEntre(flexaoJoelho, inicio, fim);
            var maxRotacaoMedialQuadril = MaxEntre(rotacaoMedialQuadril, inicio, fim);

            // === Mostra os resultados ===
            Console.WriteLine(string.Format(Constantes.Mensagens.RESULTADOS_TENTATIVA, i));
            Console.WriteLine($"   - Máx Aducao Quadril: {maxAducaoQuadril}");
            Console.WriteLine($"   - Máx Flexao Quadril: {maxFlexaoQuadril}");
            Console.WriteLine($"   - Máx Rotacao Medial Quadril: {maxRotacaoMedialQuadril}");
            Console.WriteLine($"   - Máx Flexao Joelho: {maxFlexaoJoelho}");

            var resultado = new ResultadoAnalise
            {
                NomeArquivo = $"SLHFD",
                Tentativa = i,
                Lado = lado,
                DiaColeta = dia,
                AducaoQuadrilMax = maxAducaoQuadril,
                FlexaoQuadrilMax = maxFlexaoQuadril,
                RotacaoMedialQuadrilMax = maxRotacaoMedialQuadril,
                FlexaoJoelhoMax = maxFlexaoJoelho
            };
            resultados.Add(resultado);
        }

        Console.WriteLine("🏁 Finalizado processamento SLHFD.");
        return resultados;
    }

    private List<ResultadoAnalise> ProcessarSLDL(List<string> caminhos, string lado, int dia)
    {
        Console.WriteLine($"\n📊 Processando {Constantes.Atividades.SLDL} - Lado: {lado}");

        var resultados = new List<ResultadoAnalise>();
        var ladoDireito = lado == Constantes.Lados.DIREITA;

        if (caminhos == null || caminhos.Count == 0)
        {
            Console.WriteLine("⚠️ Nenhum arquivo encontrado para processar SLDL.");
            return new List<ResultadoAnalise>();
        }

        Console.WriteLine($"🧾 Total de arquivos encontrados: {caminhos.Count}");

        var arquivoMarcadores = caminhos.FirstOrDefault(f => f.Contains("Marcadores"));
        var arquivoCinematica = caminhos.FirstOrDefault(f => f.Contains("Cinematica"));


        if (arquivoMarcadores == null || arquivoCinematica == null)
        {
            Console.WriteLine("⚠️ Arquivos de Marcadores ou Cinemática não encontrados para SLDL.");
            return new List<ResultadoAnalise>();
        }

        using var wbMarcador = new XLWorkbook(arquivoMarcadores);
        var planilhasMarcador = wbMarcador.Worksheets.ToList();

        using var wbCinematica = new XLWorkbook(arquivoCinematica);
        var planilhasCinematica = wbCinematica.Worksheets.ToList();

        for (var i = 1; i <= 5; i++)
        {
            Console.WriteLine(string.Format(Constantes.Mensagens.ANALISANDO_TENTATIVA, i));

            var planilhaM = planilhasMarcador.FirstOrDefault(f => f.Name.Contains(i.ToString()));
            var planilhaC = planilhasCinematica.FirstOrDefault(f => f.Name.Contains(i.ToString()));

            if (planilhaC is null || planilhaM is null)
            {
                Console.WriteLine(string.Format("Planilhas não encontradas na tentativa {0}", i));
                continue;
            }

            Console.WriteLine(Constantes.Mensagens.LENDO_DADOS);

            #region Dados 

            var tempoMarcadores = planilhaM.Column(Constantes.ColunasMarcadores.TEMPO)
                   .CellsUsed()
                   .Skip(1)
                   .Select(c => c.GetValue<double>())
                   .ToList();

            var aducaoQuadril = planilhaC.Column(ladoDireito ? "I" : "P")
                   .CellsUsed()
                   .Skip(1)
                   .Select(c => c.GetValue<double>())
                   .ToList();

            var flexaoQuadril = planilhaC.Column(ladoDireito ? "H" : "O")
                   .CellsUsed()
                   .Skip(1)
                   .Select(c => c.GetValue<double>())
                   .ToList();

            var rotacaoMedialQuadril = planilhaC.Column(ladoDireito ? "J" : "Q")
                   .CellsUsed()
                   .Skip(1)
                   .Select(c => c.GetValue<double>())
                   .ToList();

            var flexaoJoelho = planilhaC.Column(ladoDireito ? "K" : "R")
                   .CellsUsed()
                   .Skip(1)
                   .Select(c => c.GetValue<double>())
                   .ToList();

            var calcanharMarcador = planilhaM.Column(ladoDireito ? "BI" : "AZ")
                   .CellsUsed()
                   .Skip(1)
                   .Select(c => c.GetValue<double>())
                   .ToList();

            var pelveCinematica = planilhaM.Column(ladoDireito ? "BL" : "BO")
                   .CellsUsed()
                   .Skip(1)
                   .Select(c => c.GetValue<double>())
                   .ToList();

            #endregion Dados 

            Console.WriteLine("🔎 Calculando pontos de corte...");

            #region Primeiro ponto de corte

            if (calcanharMarcador.Count == 0)
            {
                Console.WriteLine($"Nenhum ponto de corte da pelve encontrado na tentativa {i}");
                continue;
            }

            var IndexMaximoValorCalcanharMarcador = calcanharMarcador.IndexOf(calcanharMarcador.Max());


            int pontoCorteCalcanhar = -1;
            var calcanharPontoNeutro = calcanharMarcador[0];
            for (int index = 1; index < calcanharMarcador.Count - 1; index++)
            {
                if (calcanharMarcador[index] <= calcanharPontoNeutro && index > IndexMaximoValorCalcanharMarcador)
                {
                    pontoCorteCalcanhar = index;
                    break;
                }
            }
            Console.WriteLine($"📍 Ponto de corte calcanhar: {pontoCorteCalcanhar}");

            #endregion Primeiro ponto de corte

            #region Segundo ponto de corte

            var IndexMaximoValorPelveMarcador = calcanharMarcador.IndexOf(calcanharMarcador.Max());

            // pegar somente os valores após o pico
            var valoresDepoisDoMaximoPelve = pelveCinematica
                .Skip(IndexMaximoValorPelveMarcador + 1)
                .ToList();

            if(valoresDepoisDoMaximoPelve.Count == 0)
            {
                Console.WriteLine($"Nenhum ponto de corte da pelve encontrado na tentativa {i}");
                continue;
            }

            var menorValor = valoresDepoisDoMaximoPelve.Min();

            // índice global do menor valor
            var pontoCortePelve = pelveCinematica.IndexOf(menorValor, IndexMaximoValorPelveMarcador + 1);

            Console.WriteLine($"📍 Ponto de corte pelve: {pontoCortePelve}");

            #endregion Segundo ponto de corte

            // Garante que os índices estão dentro dos limites
            int inicio = Math.Max(0, pontoCorteCalcanhar);
            int fim = Math.Min(pelveCinematica.Count - 1, pontoCortePelve);

            // Caso os índices estejam invertidos (por segurança)
            if (fim <= inicio)
            {
                Console.WriteLine("⚠️ Intervalo inválido entre pontoCorteCalcanhar e pontoCortePelve. Pulando tentativa...");
                continue;
            }

            // === Calcula os máximos entre os pontos de corte ===
            Console.WriteLine("📊 Calculando máximos entre os pontos de corte...");
            var maxAducaoQuadril = MaxEntre(aducaoQuadril, inicio, fim);
            var maxFlexaoQuadril = MaxEntre(flexaoQuadril, inicio, fim);
            var maxFlexaoJoelho = MaxEntre(flexaoJoelho, inicio, fim);
            var maxRotacaoMedialQuadril = MaxEntre(rotacaoMedialQuadril, inicio, fim);

            // === Mostra os resultados ===
            Console.WriteLine(string.Format(Constantes.Mensagens.RESULTADOS_TENTATIVA, i));
            Console.WriteLine($"   - Máx Aducao Quadril: {maxAducaoQuadril}");
            Console.WriteLine($"   - Máx Flexao Quadril: {maxFlexaoQuadril}");
            Console.WriteLine($"   - Máx Rotacao Medial Quadril: {maxRotacaoMedialQuadril}");
            Console.WriteLine($"   - Máx Flexao Joelho: {maxFlexaoJoelho}");

            var resultado = new ResultadoAnalise
            {
                NomeArquivo = $"SLDL",
                Tentativa = i,
                Lado = lado,
                DiaColeta = dia,
                AducaoQuadrilMax = maxAducaoQuadril,
                FlexaoQuadrilMax = maxFlexaoQuadril,
                RotacaoMedialQuadrilMax = maxRotacaoMedialQuadril,
                FlexaoJoelhoMax = maxFlexaoJoelho
            };
            resultados.Add(resultado);
        }

        Console.WriteLine("🏁 Finalizado processamento SLDL.");
        return resultados;
    }

    double MaxEntre(List<double> lista, int start, int end)
    {
        if (lista.Count() == 0)
            return 0;
        return lista.Skip(start).Take(end - start + 1).Max();
    }

    public List<ResultadoAnalise> CalcularMedia(List<ResultadoAnalise> resultados)
    {
        var medias = resultados
            .GroupBy(r => new { r.NomeArquivo, r.DiaColeta, r.Lado })
            .Select(g => new ResultadoAnalise
            {
                NomeArquivo = g.Key.NomeArquivo,
                DiaColeta = g.Key.DiaColeta,
                Lado = g.Key.Lado,
                MediaAducaoQuadril = g.Average(x => x.AducaoQuadrilMax),
                MediaFlexaoQuadril = g.Average(x => x.FlexaoQuadrilMax),
                MediaRotacaoMedialQuadril = g.Average(x => x.RotacaoMedialQuadrilMax),
                MediaFlexaoJoelho = g.Average(x => x.FlexaoJoelhoMax)
            })
            .OrderBy(x => x.DiaColeta)
            .ThenBy(x => x.Lado)
            .ToList();

        return medias;
    }

    public void ExportarResultadosExcel(List<ResultadoAnalise> resultados, List<ResultadoAnalise> medias, string caminhoArquivo)
    {
        using (var wb = new XLWorkbook())
        {
            var ws = wb.Worksheets.Add("Resultados");

            // Cabeçalhos
            ws.Cell(1, 1).Value = "Arquivo";
            ws.Cell(1, 2).Value = "Lado";
            ws.Cell(1, 3).Value = "Dia";
            ws.Cell(1, 4).Value = "Adução Quadril Máx.";
            ws.Cell(1, 5).Value = "Flexão Quadril Máx.";
            ws.Cell(1, 6).Value = "Flexão Joelho Máx.";
            ws.Cell(1, 7).Value = "Rotação Medial Quadril Máx.";

            int linha = 2;
            foreach (var r in resultados)
            {
                ws.Cell(linha, 1).Value = $"{r.NomeArquivo} - Tentativa {r.Tentativa}";
                ws.Cell(linha, 2).Value = r.Lado;
                ws.Cell(linha, 3).Value =  $"Dia {r.DiaColeta}";
                ws.Cell(linha, 4).Value = r.AducaoQuadrilMax;
                ws.Cell(linha, 5).Value = r.FlexaoQuadrilMax;
                ws.Cell(linha, 6).Value = r.FlexaoJoelhoMax;
                ws.Cell(linha, 7).Value = r.RotacaoMedialQuadrilMax;
                linha++;
            }

            linha++;

            foreach(var m in medias)
            {
                ws.Cell(linha, 1).Value = "Média - " + m.NomeArquivo;
                ws.Cell(linha, 2).Value = m.Lado;
                ws.Cell(linha, 3).Value = $"Dia {m.DiaColeta}";
                ws.Cell(linha, 4).Value = m.MediaAducaoQuadril;
                ws.Cell(linha, 5).Value = m.MediaFlexaoQuadril;
                ws.Cell(linha, 6).Value = m.MediaFlexaoJoelho;
                ws.Cell(linha, 7).Value = m.MediaRotacaoMedialQuadril;
                linha++;
            }

            ws.Columns().AdjustToContents();
            wb.SaveAs(caminhoArquivo);
        }
    }
}
