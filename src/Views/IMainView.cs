namespace Cabecalho.LogoInserter.Views;

public interface IMainView
{
    string DocumentoPath { get; set; }
    string LogoPath { get; set; }
    string SaidaPath { get; set; }
    bool AbrirAoFinal { get; }

    event EventHandler? SelecionarDocumento;
    event EventHandler? SelecionarLogo;
    event EventHandler? SelecionarSaida;
    event EventHandler? Processar;

    string? EscolherDocumento();
    string? EscolherLogo();
    string? EscolherSaida(string sugestaoInicial);

    void SetBusy(bool busy);
    void SetStatus(string mensagem);
    void MostrarSucesso(string mensagem);
    void MostrarErro(string mensagem);
}
