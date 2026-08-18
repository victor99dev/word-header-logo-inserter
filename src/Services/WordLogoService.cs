using System.Runtime.InteropServices;
using Cabecalho.LogoInserter.Models;

namespace Cabecalho.LogoInserter.Services;

public class WordLogoService : IWordLogoService
{
    private const string MarcadorLogo = "LOGO_AUTOMATICA";
    private const int WdHeaderFooterPrimary = 1;
    private const int WdCollapseStart = 1;
    private const int WdRelativeHorizontalPositionMargin = 0;
    private const int WdRelativeVerticalPositionParagraph = 2;
    private const int WdWrapBehind = 5;
    private const int MsoTrue = -1;

    public void InserirLogo(InsercaoLogoModel model)
    {
        Validar(model);

        var origem = Path.GetFullPath(model.DocumentoOrigem);
        var logo = Path.GetFullPath(model.Logo);
        var saida = Path.GetFullPath(model.DocumentoSaida);

        if (!string.Equals(origem, saida, StringComparison.OrdinalIgnoreCase))
        {
            var pastaSaida = Path.GetDirectoryName(saida);
            if (!string.IsNullOrWhiteSpace(pastaSaida))
                Directory.CreateDirectory(pastaSaida);

            File.Copy(origem, saida, overwrite: true);
        }

        Type? wordType = Type.GetTypeFromProgID("Word.Application") ?? throw new InvalidOperationException("Microsoft Word Desktop não foi encontrado. Este teste usa automação COM e requer o Word instalado no Windows.");

        dynamic? word = null;
        dynamic? document = null;

        try
        {
            word = Activator.CreateInstance(wordType)
                ?? throw new InvalidOperationException("Não foi possível iniciar o Microsoft Word.");

            word.Visible = false;
            word.DisplayAlerts = 0;

            document = word.Documents.Open(saida);

            foreach (dynamic section in document.Sections)
            {
                dynamic? header = null;
                dynamic? rangeLogo = null;
                dynamic? inlineLogo = null;
                dynamic? shapeLogo = null;

                try
                {
                    header = section.Headers.Item(WdHeaderFooterPrimary);

                    // Remove logo previamente inserida como InlineShape.
                    for (int i = header.Range.InlineShapes.Count; i >= 1; i--)
                    {
                        dynamic inlineShape = header.Range.InlineShapes.Item(i);
                        try
                        {
                            if (string.Equals(
                                (string?)inlineShape.AlternativeText,
                                MarcadorLogo,
                                StringComparison.Ordinal))
                            {
                                inlineShape.Delete();
                            }
                        }
                        finally
                        {
                            ReleaseCom(inlineShape);
                        }
                    }

                    // Remove logo previamente inserida como Shape.
                    for (int i = header.Shapes.Count; i >= 1; i--)
                    {
                        dynamic shape = header.Shapes.Item(i);
                        try
                        {
                            if (string.Equals(
                                (string?)shape.AlternativeText,
                                MarcadorLogo,
                                StringComparison.Ordinal))
                            {
                                shape.Delete();
                            }
                        }
                        finally
                        {
                            ReleaseCom(shape);
                        }
                    }

                    rangeLogo = header.Range.Duplicate;
                    rangeLogo.Collapse(WdCollapseStart);

                    var (larguraLogo, alturaLogo) = CalcularDimensoesLogo(logo);

                    shapeLogo = header.Shapes.AddPicture(
                        FileName: logo,
                        LinkToFile: false,
                        SaveWithDocument: true,
                        Left: 0f,
                        Top: 0f,
                        Width: larguraLogo,
                        Height: alturaLogo,
                        Anchor: rangeLogo);

                    shapeLogo.AlternativeText = MarcadorLogo;
                    shapeLogo.LockAspectRatio = MsoTrue;

                    shapeLogo.RelativeHorizontalPosition = WdRelativeHorizontalPositionMargin;
                    shapeLogo.Left = 0;

                    shapeLogo.RelativeVerticalPosition = WdRelativeVerticalPositionParagraph;
                    shapeLogo.Top = 0;

                    shapeLogo.WrapFormat.Type = WdWrapBehind;
                    shapeLogo.WrapFormat.AllowOverlap = MsoTrue;
                }
                catch (COMException ex)
                {
                    throw new InvalidOperationException(
                        $"Erro do Microsoft Word ao inserir a logo. " +
                        $"HRESULT: 0x{ex.HResult:X8}. " +
                        $"Mensagem: {ex.Message}", ex);
                }
                finally
                {
                    ReleaseCom(shapeLogo);
                    ReleaseCom(inlineLogo);
                    ReleaseCom(rangeLogo);
                    ReleaseCom(header);
                    ReleaseCom(section);
                }
            }

            document.Save();
            document.Close(false);
            document = null;
        }
        finally
        {
            if (document is not null)
            {
                try { document.Close(false); } catch { }
                ReleaseCom(document);
            }

            if (word is not null)
            {
                try { word.Quit(); } catch { }
                ReleaseCom(word);
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    private static void Validar(InsercaoLogoModel model)
    {
        if (string.IsNullOrWhiteSpace(model.DocumentoOrigem) || !File.Exists(model.DocumentoOrigem))
            throw new FileNotFoundException("Selecione um documento Word válido.", model.DocumentoOrigem);

        if (string.IsNullOrWhiteSpace(model.Logo) || !File.Exists(model.Logo))
            throw new FileNotFoundException("Selecione uma imagem de logo válida.", model.Logo);

        if (string.IsNullOrWhiteSpace(model.DocumentoSaida))
            throw new ArgumentException("Selecione o caminho do documento de saída.");

        var extensaoDocumento = Path.GetExtension(model.DocumentoOrigem);
        if (!new[] { ".docx", ".docm", ".doc" }.Contains(extensaoDocumento, StringComparer.OrdinalIgnoreCase))
            throw new ArgumentException("O documento deve ser .docx, .docm ou .doc.");

        var extensaoLogo = Path.GetExtension(model.Logo);
        if (!new[] { ".png", ".jpg", ".jpeg", ".bmp" }.Contains(extensaoLogo, StringComparer.OrdinalIgnoreCase))
            throw new ArgumentException("A logo deve ser PNG, JPG, JPEG ou BMP.");
    }

    private static void ReleaseCom(object? value)
    {
        if (value is null || !Marshal.IsComObject(value))
            return;

        try
        {
            Marshal.FinalReleaseComObject(value);
        }
        catch
        {

        }
    }

    private static (float Largura, float Altura) CalcularDimensoesLogo(string caminhoLogo)
    {
        const float largura = 90f;

        using var imagem = Image.FromFile(caminhoLogo);

        var proporcao = (float)imagem.Height / imagem.Width;
        var altura = largura * proporcao;

        return (largura, altura);
    }
}
