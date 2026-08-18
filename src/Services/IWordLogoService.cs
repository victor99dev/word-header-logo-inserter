using Cabecalho.LogoInserter.Models;

namespace Cabecalho.LogoInserter.Services;

public interface IWordLogoService
{
    void InserirLogo(InsercaoLogoModel model);
}
