using CsvHelper.Configuration.Attributes;

namespace FlexImport.Classes
{
    public class OriginProduct
    {
        [Name("ID")]
        public string ID { get; set; } = string.Empty;

        [Name("Código")]
        public string Codigo { get; set; } = string.Empty;

        [Name("Descrição")]
        public string Descricao { get; set; } = string.Empty;

        [Name("Unidade")]
        public string Unidade { get; set; } = string.Empty;

        [Name("NCM")]
        public string NCM { get; set; } = string.Empty;

        [Name("Origem")]
        public string Origem { get; set; } = string.Empty;

        [Name("Preço")]
        public float Preco { get; set; } = 0;

        [Name("Valor IPI fixo")]
        public string ValorIPIFixo { get; set; } = string.Empty;

        [Name("Observações")]
        public string Observacoes { get; set; } = string.Empty;

        [Name("Situação")]
        public string Situacao { get; set; } = string.Empty;

        [Name("Estoque")]
        public int Estoque { get; set; } = 0;

        [Name("Preço de custo")]
        public float PrecoDeCusto { get; set; } = 0;

        [Name("Cód. no fornecedor")]
        public string CodNoFornecedor { get; set; } = string.Empty;

        [Name("Fornecedor")]
        public string Fornecedor { get; set; } = string.Empty;

        [Name("Localização")]
        public string Localizacao { get; set; } = string.Empty;

        [Name("Estoque máximo")]
        public int EstoqueMaximo { get; set; } = 0;

        [Name("Estoque mínimo")]
        public int EstoqueMinimo { get; set; } = 0;

        [Name("Peso líquido (Kg)")]
        public float PesoLiquidoKg { get; set; } = 0;

        [Name("Peso bruto (Kg)")]
        public float PesoBrutoKg { get; set; } = 0;

        [Name("GTIN/EAN")]
        public string GTIN_EAN { get; set; } = string.Empty;

        [Name("GTIN/EAN da Embalagem")]
        public string GTIN_EAN_Embalagem { get; set; } = string.Empty;

        [Name("Largura do produto")]
        public int LarguraProduto { get; set; } = 0;

        [Name("Altura do Produto")]
        public int AlturaProduto { get; set; } = 0;

        [Name("Profundidade do produto")]
        public int ProfundidadeProduto { get; set; } = 0;

        [Name("Data Validade")]
        public string DataValidade { get; set; } = string.Empty;

        [Name("Descrição do Produto no Fornecedor")]
        public string DescricaoProdutoFornecedor { get; set; } = string.Empty;

        [Name("Descrição Complementar")]
        public string DescricaoComplementar { get; set; } = string.Empty;

        [Name("Itens p/ caixa")]
        public int ItensPorCaixa { get; set; } = 0;

        [Name("Produto Variação")]
        public string ProdutoVariacao { get; set; } = string.Empty;

        [Name("Tipo Produção")]
        public string TipoProducao { get; set; } = string.Empty;

        [Name("Classe de enquadramento do IPI")]
        public string ClasseEnquadramentoIPI { get; set; } = string.Empty;

        [Name("Código na Lista de Serviços")]
        public string CodigoListaServicos { get; set; } = string.Empty;

        [Name("Tipo do item")]
        public string TipoDoItem { get; set; } = string.Empty;

        [Name("Grupo de Tags/Tags")]
        public string GrupoTags { get; set; } = string.Empty;

        [Name("Tributos")]
        public string Tributos { get; set; } = string.Empty;

        [Name("Código Pai")]
        public string CodigoPai { get; set; } = string.Empty;

        [Name("Código Integração")]
        public string CodigoIntegracao { get; set; } = string.Empty;

        [Name("Grupo de produtos")]
        public string GrupoProdutos { get; set; } = string.Empty;

        [Name("Marca")]
        public string Marca { get; set; } = string.Empty;

        [Name("CEST")]
        public string CEST { get; set; } = string.Empty;

        [Name("Volumes")]
        public string Volumes { get; set; } = string.Empty;

        [Name("Descrição Curta")]
        public string DescricaoCurta { get; set; } = string.Empty;

        [Name("Cross-Docking")]
        public string CrossDocking { get; set; } = string.Empty;

        [Name("URL Imagens Externas")]
        public string UrlImagensExternas { get; set; } = string.Empty;

        [Name("Link Externo")]
        public string LinkExterno { get; set; } = string.Empty;

        [Name("Meses Garantia no Fornecedor")]
        public string MesesGarantiaFornecedor { get; set; } = string.Empty;

        [Name("Clonar dados do pai")]
        public string ClonarDadosPai { get; set; } = string.Empty;

        [Name("Condição do Produto")]
        public string CondicaoProduto { get; set; } = string.Empty;

        [Name("Frete Grátis")]
        public string FreteGratis { get; set; } = string.Empty;

        [Name("Número FCI")]
        public string NumeroFCI { get; set; } = string.Empty;

        [Name("Vídeo")]
        public string Video { get; set; } = string.Empty;

        [Name("Departamento")]
        public string Departamento { get; set; } = string.Empty;

        [Name("Unidade de Medida")]
        public string UnidadeMedida { get; set; } = string.Empty;

        [Name("Preço de Compra")]
        public float PrecoCompra { get; set; } = 0;

        [Name("Valor base ICMS ST para retenção")]
        public string ValorBaseICMSST { get; set; } = string.Empty;

        [Name("Valor ICMS ST para retenção")]
        public string ValorICMSST { get; set; } = string.Empty;

        [Name("Valor ICMS próprio do substituto")]
        public string ValorICMSProprioSubstituto { get; set; } = string.Empty;

        [Name("Categoria do produto")]
        public string CategoriaProduto { get; set; } = string.Empty;

        [Name("Informações Adicionais")]
        public string InformacoesAdicionais { get; set; } = string.Empty;
    }
}