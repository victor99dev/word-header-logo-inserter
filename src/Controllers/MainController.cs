using System.Diagnostics;
using Cabecalho.LogoInserter.Models;
using Cabecalho.LogoInserter.Services;
using Cabecalho.LogoInserter.Views;

namespace Cabecalho.LogoInserter.Controllers;

public class MainController
{
    private readonly IMainView _view;
    private readonly IWordLogoService _service;

    public MainController(IMainView view, IWordLogoService service)
    {
        _view = view;
        _service = service;

        _view.SelecionarDocumento += OnSelecionarDocumento;
        _view.SelecionarLogo += OnSelecionarLogo;
        _view.SelecionarSaida += OnSelecionarSaida;
        _view.Processar += OnProcessar;
    }

    private void OnSelecionarDocumento(object? sender, EventArgs e)
    {
        var arquivo = _view.EscolherDocumento();
        if (string.IsNullOrWhiteSpace(arquivo))
            return;

        _view.DocumentoPath = arquivo;

        if (string.IsNullOrWhiteSpace(_view.SaidaPath))
            _view.SaidaPath = CriarSugestaoSaida(arquivo);
    }

    private void OnSelecionarLogo(object? sender, EventArgs e)
    {
        var arquivo = _view.EscolherLogo();
        if (!string.IsNullOrWhiteSpace(arquivo))
            _view.LogoPath = arquivo;
    }

    private void OnSelecionarSaida(object? sender, EventArgs e)
    {
        var sugestao = string.IsNullOrWhiteSpace(_view.SaidaPath)
            ? CriarSugestaoSaida(_view.DocumentoPath)
            : _view.SaidaPath;

        var arquivo = _view.EscolherSaida(sugestao);
        if (!string.IsNullOrWhiteSpace(arquivo))
            _view.SaidaPath = arquivo;
    }

    private void OnProcessar(object? sender, EventArgs e)
    {
        var model = new InsercaoLogoModel
        {
            DocumentoOrigem = _view.DocumentoPath,
            Logo = _view.LogoPath,
            DocumentoSaida = _view.SaidaPath,
            AbrirAoFinal = _view.AbrirAoFinal
        };

        try
        {
            _view.SetBusy(true);
            _view.SetStatus("Processando documento no Microsoft Word...");

            _service.InserirLogo(model);

            _view.SetStatus($"Concluído: {model.DocumentoSaida}");
            _view.MostrarSucesso("Logo inserida com sucesso no cabeçalho do documento.");

            if (model.AbrirAoFinal && File.Exists(model.DocumentoSaida))
            {
                Process.Start(new ProcessStartInfo(model.DocumentoSaida)
                {
                    UseShellExecute = true
                });
            }
        }
        catch (Exception ex)
        {
            _view.SetStatus("Falha no processamento.");
            _view.MostrarErro(ex.Message);
        }
        finally
        {
            _view.SetBusy(false);
        }
    }

    private static string CriarSugestaoSaida(string? documentoOrigem)
    {
        if (string.IsNullOrWhiteSpace(documentoOrigem))
            return $"nome_com_logo.docx";

        var diretorio = Path.GetDirectoryName(documentoOrigem) ?? string.Empty;
        var nome = Path.GetFileNameWithoutExtension(documentoOrigem);
        var extensao = Path.GetExtension(documentoOrigem);

        if (string.IsNullOrWhiteSpace(extensao))
            extensao = ".docx";

        return Path.Combine(diretorio, $"{nome}_com_logo{extensao}");
    }
}
