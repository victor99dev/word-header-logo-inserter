namespace Cabecalho.LogoInserter.Models;

public class InsercaoLogoModel
{
    public string DocumentoOrigem { get; init; } = string.Empty;
    public string Logo { get; init; } = string.Empty;
    public string DocumentoSaida { get; init; } = string.Empty;
    public bool AbrirAoFinal { get; init; }
}