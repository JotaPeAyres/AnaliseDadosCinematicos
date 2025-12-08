namespace AnalisarDados;

/// <summary>
/// Classe para armazenar constantes e configurações do sistema.
/// </summary>
public static class Constantes
{
    #region Configurações de Diretórios
    
    /// <summary>
    /// Padrões de nomenclatura de pastas e arquivos
    /// </summary>
    public static class Pastas
    {
        public const string PADRAO_PASTA_CINEMATICA = "Dia_{0}_Cinematica";
        public const string PADRAO_PASTA_MARCADORES = "Dia_{0}_Marcadores";
        public const string PADRAO_ARQUIVO_DIREITA = "*_Direita.xlsx";
        public const string PADRAO_ARQUIVO_ESQUERDA = "*_Esquerda.xlsx";
    }
    
    #endregion

    #region Nomes de Atividades
    
    /// <summary>
    /// Nomes das atividades de análise
    /// </summary>
    public static class Atividades
    {
        public const string SLDL = "SLDL";
        public const string SLHFD = "SLHFD";
        public const string SLLV = "SLLV";
        
        public static readonly string[] Todas = { SLDL, SLHFD, SLLV };
    }
    
    #endregion
    
    #region Nomes de Lados
    
    /// <summary>
    /// Nomes dos lados do corpo
    /// </summary>
    public static class Lados
    {
        public const string DIREITA = "Direita";
        public const string ESQUERDA = "Esquerda";
    }
    
    #endregion
    
    #region Colunas do Excel - Marcadores
    
    /// <summary>
    /// Nomes das colunas da planilha de Marcadores
    /// </summary>
    public static class ColunasMarcadores
    {
        public const string TEMPO = "B";
        public const string CALCANHAR_DIREITO = "BI";
        public const string CALCANHAR_ESQUERDO = "AZ";
        public const string PELVE_DIREITA = "BL";
        public const string PELVE_ESQUERDA = "BO";
    }
    
    #endregion
    
    #region Colunas do Excel - Cinemática
    
    /// <summary>
    /// Nomes das colunas da planilha de Cinemática
    /// </summary>
    public static class ColunasCinematica
    {
        // Colunas para o lado direito (sufixo _D) e esquerdo (sufixo _E)
        public const string FLEXAO_QUADRIL_D = "H";
        public const string ADUCAO_QUADRIL_D = "I";
        public const string ROTACAO_MEDIAL_QUADRIL_D = "J";
        public const string FLEXAO_JOELHO_D = "K";
        
        public const string FLEXAO_QUADRIL_E = "O";
        public const string ADUCAO_QUADRIL_E = "P";
        public const string ROTACAO_MEDIAL_QUADRIL_E = "Q";
        public const string FLEXAO_JOELHO_E = "R";
    }
    
    #endregion
    
    #region Mensagens de Log
    
    /// <summary>
    /// Mensagens de log do sistema
    /// </summary>
    public static class Mensagens
    {
        public const string INICIO_PROCESSAMENTO = "🚀 Iniciando processamento na pasta raiz: {0}";
        public const string PROCESSANDO_DIA = "📅 Processando arquivos do Dia {0}...";
        public const string PROCESSAMENTO_CONCLUIDO = "✅ Processamento concluído!";
        public const string TEMPO_EXECUCAO = "⏱️ Tempo total de execução: {0:mm\\:ss\\.fff}";
        public const string PASTA_NAO_ENCONTRADA = "⚠️ Pasta não encontrada: {0}";
        public const string SUBPASTA_NAO_ENCONTRADA = "⚠️ Subpasta não encontrada para {0}: {1}";
        public const string EXPLORANDO_PASTA = "📂 Explorando: {0}";
        public const string SUBPASTA_ENCONTRADA = "📁 Encontrada subpasta de {0}: {1}";
        public const string INICIANDO_ATIVIDADE = "⚙️ Iniciando processamento da atividade: {0}";
        public const string PROCESSANDO_ATIVIDADE = "➡️ Processando {0}...";
        public const string ANALISANDO_TENTATIVA = "🔢 Analisando tentativa {0}...";
        public const string LENDO_DADOS = "📈 Lendo colunas de dados...";
        public const string CALCULANDO_PONTOS = "🔎 Calculando pontos de corte...";
        public const string PONTO_CORTE = "📍 Ponto de corte {0}: {1}";
        public const string CALCULANDO_MAXIMOS = "📊 Calculando máximos entre os pontos de corte...";
        public const string RESULTADOS_TENTATIVA = "✅ Resultados tentativa {0}:";
        public const string FINALIZADO_PROCESSAMENTO = "🏁 Finalizado processamento {0}.";
        public const string DIA_CONCLUIDO = "✅ Concluído processamento do Dia {0}.";
    }
    
    #endregion
    
    #region Configurações de Processamento
    
    /// <summary>
    /// Configurações relacionadas ao processamento dos dados
    /// </summary>
    public static class Processamento
    {
        public const int NUMERO_TENTATIVAS = 5;
        public const int INDICE_INICIAL_DADOS = 1; // Linha onde começam os dados nas planilhas
        public const int DIAS_PROCESSAMENTO = 2;   // Número de dias a serem processados (Dia 1 e Dia 2)
        public const int DIA_INICIAL = 1;          // Dia inicial do processamento
    }
    
    #endregion
}
