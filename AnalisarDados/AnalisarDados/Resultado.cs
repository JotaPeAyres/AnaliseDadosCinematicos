namespace ValidadorExcel;

public class Resultado
{
    public string CaminhoArquivo { get; set; } = "";
    public bool Valido { get; set; }
    public string Mensagem { get; set; } = "";
}


public class ResultadoAnalise
{
    public string NomeArquivo { get; set; } = "";
    public int Tentativa { get; set; }
    public string Lado { get; set; } = "";
    public int DiaColeta { get; set; }
    public double AducaoQuadrilMax { get; set; }
    public double FlexaoQuadrilMax { get; set; }
    public double FlexaoJoelhoMax { get; set; }
    public double RotacaoMedialQuadrilMax { get; set; }


    public double MediaAducaoQuadril { get; set; }
    public double MediaFlexaoQuadril { get; set; }
    public double MediaFlexaoJoelho { get; set; }
    public double MediaRotacaoMedialQuadril { get; set; }
}