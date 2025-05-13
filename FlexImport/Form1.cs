using CsvHelper;
using FlexImport.Classes;
using FlexImport.Services;
using OfficeOpenXml;
using System.Text.RegularExpressions;

namespace FlexImport
{
    public partial class Form1 : Form
    {
        private string FileName = string.Empty;
        private string SelectedCategory => (CmbCategory.SelectedIndex > 0) ? CmbCategory.SelectedItem?.ToString() ?? string.Empty : string.Empty;
        private string SelectedSubCategory => (CmbSubCategory.SelectedIndex > 0) ? CmbSubCategory.SelectedItem?.ToString() ?? string.Empty : string.Empty;
        private string SelectedGenre => (CmbGenre.SelectedIndex > 0) ? CmbGenre.SelectedItem?.ToString() ?? string.Empty : string.Empty;

        public Form1()
        {
            InitializeComponent();
            InitializeCategory();
            InitializeGenre();
        }

        private void InitializeCategory()
        {
            // Adiciona item inicial obrigatório para categoria
            CmbCategory.Items.Add("-- Selecione --");
            CmbCategory.Items.AddRange(UtilityService.Categories.Keys.ToArray());

            CmbCategory.SelectedIndexChanged += (s, e) =>
            {
                CmbSubCategory.Items.Clear();
                CmbSubCategory.Items.Add("");

                if (CmbCategory.SelectedIndex > 0 && CmbCategory.SelectedItem is string category && UtilityService.Categories.TryGetValue(category, out var subs))
                {
                    CmbSubCategory.Items.AddRange(subs);
                }

                CmbSubCategory.SelectedIndex = 0;
            };

            CmbCategory.SelectedIndex = 0;
            CmbSubCategory.SelectedIndex = 0;
        }

        private void InitializeGenre()
        {
            // Adiciona item inicial obrigatório para gênero
            CmbGenre.Items.Add("-- Selecione --");
            CmbGenre.Items.AddRange(UtilityService.Genre.ToArray());
            CmbGenre.SelectedIndex = 0;
        }

        private void BtnSelectSheet_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Arquivos CSV (*.csv)|*.csv",
                Title = "Selecione a planilha exportada do Bling"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                if (Path.GetExtension(openFileDialog.FileName).ToLower() != ".csv")
                {
                    MessageBox.Show("Por favor, selecione um arquivo CSV exportado do Bling.",
                        "Formato incorreto", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    TxtSheet.Text = openFileDialog.FileName;
                }
            }
        }

        private void BtnSelectPath_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "CSV Files (*.csv)|*.csv",
                Title = "Salvar planilha convertida como"
            };

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                TxtPath.Text = Path.GetDirectoryName(saveFileDialog.FileName);

                string nomeInformado = Path.GetFileNameWithoutExtension(saveFileDialog.FileName);
                FileName = string.IsNullOrWhiteSpace(nomeInformado)
                    ? "Planilha_Convertida_" + DateTime.Now.ToString("dd-MM-yyyy_HH-mm-ss")
                    : nomeInformado;
            }
        }

        private void ToggleControls(bool enabled)
        {
            TxtSheet.Enabled = enabled;
            BtnSelectSheet.Enabled = enabled;
            TxtPath.Enabled = enabled;
            BtnSelectPath.Enabled = enabled;
            BtnConvert.Enabled = enabled;
            CmbCategory.Enabled = enabled;
            CmbSubCategory.Enabled = enabled;
        }

        private void BtnConvert_Click(object sender, EventArgs e)
        {
            ToggleControls(false);

            if (CmbCategory.SelectedIndex == 0 || CmbGenre.SelectedIndex == 0)
            {
                MessageBox.Show($"É preciso selecionar uma categoria e um gênero para converter.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                ToggleControls(true);
                return;
            }

            if (TxtSheet.Text == string.Empty)
                TxtPath.Text = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");

            if (FileName == string.Empty)
                FileName = "Planilha_Convertida_" + DateTime.Now.ToString("dd-MM-yyyy_HH-mm-ss");

            string originFile = TxtSheet.Text;
            string savePath = Path.Combine(TxtPath.Text, FileName + ".csv");

            if (!File.Exists(originFile) || !Directory.Exists(TxtPath.Text))
            {
                ToggleControls(true);
                MessageBox.Show("Verifique os caminhos selecionados.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                // Configuração do CsvHelper para tabulação
                using var reader = new StreamReader(originFile);
                using var csv = new CsvReader(reader, UtilityService.Config);

                // Registra o conversor customizado
                UtilityService.RegisterPtBrNumberConverters(csv);

                var originProducts = csv.GetRecords<OriginProduct>().ToList();

                // Agrupa por NCM
                var groupedByNCM = originProducts.GroupBy(p => p.NCM);
                List<FinalProduct> finalProducts = new();

                foreach (var group in groupedByNCM)
                {
                    string fatherSku = UtilityService.GenerateRandomSKU(8);
                    int childCount = 1;
                    for (int i = 0; i < group.Count(); i++)
                    {
                        var origin = group.ElementAt(i);

                        string originSize = string.Empty;
                        string originColor = origin.Descricao.Split(";").Where(d => d.Contains("Cor:")).FirstOrDefault()?.Replace("Cor:", string.Empty);

                        var matchSize = Regex.Match(origin.Descricao, @"Tamanho:\s*(RN|P|M|G|GG)\b", RegexOptions.IgnoreCase);
                        if (matchSize.Success)
                            originSize = matchSize.Groups[1].Value.ToUpper();

                        string childSku = fatherSku + 
                            ($"-{SelectedGenre}{(originColor != null ? $"-{originColor.Replace("/", "-").ToLower()}" : string.Empty)}-{originSize}").ToLower();

                        var final = new FinalProduct
                        {
                            Id = origin.ID,
                            Tipo = (i == 0) ? "com-variacao" : "variacao",
                            SkuPai = (i == 0) ? string.Empty : fatherSku,
                            Sku = (i == 0) ? fatherSku : childSku,
                            Ativo = "S",
                            Usado = "N",
                            Destaque = "S",
                            Ncm = origin.NCM,
                            Gtin = origin.GTIN_EAN,
                            Mpn = origin.CodNoFornecedor,
                            Nome = origin.Descricao,
                            SeoTagTitle = origin.Descricao,
                            SeoTagDescription = origin.DescricaoComplementar,
                            DescricaoCompleta = origin.DescricaoCurta,
                            UrlVideoYoutube = origin.Video,
                            EstoqueGerenciado = "S",
                            EstoqueQuantidade = origin.Estoque,
                            EstoqueSituacaoEmEstoque = "2 dias",
                            EstoqueSituacaoSemEstoque = "indisponivel",
                            PrecoSobConsulta = "N",
                            PrecoCusto = origin.PrecoDeCusto,
                            PrecoCheio = origin.Preco,
                            PrecoPromocional = origin.PrecoCompra,
                            Marca = origin.Marca,
                            PesoEmKg = origin.PesoLiquidoKg,
                            AlturaEmCm = origin.AlturaProduto,
                            LarguraEmCm = origin.LarguraProduto,
                            ComprimentoEmCm = origin.ProfundidadeProduto,
                            CategoriaNomeNivel1 = SelectedCategory,
                            CategoriaNomeNivel2 = SelectedSubCategory,
                            CategoriaNomeNivel3 = string.Empty,
                            CategoriaNomeNivel4 = string.Empty,
                            CategoriaNomeNivel5 = string.Empty,
                            Imagem1 = origin.UrlImagensExternas,
                            Imagem2 = string.Empty,
                            Imagem3 = string.Empty,
                            Imagem4 = string.Empty,
                            Imagem5 = string.Empty,
                            GradeGenero = (i == 0) ? string.Empty : SelectedGenre,
                            GradeProdutoComUmaCor = (i == 0) ? string.Empty : originColor ?? string.Empty,
                            GradeProdutoComDuasCores = string.Empty,
                            GradeTamanhoDeAnelAlianca = string.Empty,
                            GradeTamanhoDeCalca = string.Empty,
                            GradeTamanhoDeCamisaCamiseta = string.Empty,
                            GradeTamanhoDeCapacete = string.Empty,
                            GradeTamanhoDeTenis = string.Empty,
                            GradeVoltagem = string.Empty,
                            GradeTamanhoJuvenilInfantil = (i == 0) ? string.Empty : originSize,
                            GradeTamanho = string.Empty,
                            UrlAntiga = string.Empty
                        };
                        finalProducts.Add(final);
                        if (i > 0) childCount++;
                    }
                }


                ExcelPackage.License.SetNonCommercialOrganization("Developer");

                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Planilha");

                    // Cabeçalho
                    for (int col = 0; col < UtilityService.finalHeader.Count; col++)
                    {
                        worksheet.Cells[1, col + 1].Value = UtilityService.finalHeader[col];
                    }

                    // Dados
                    for (int row = 0; row < finalProducts.Count; row++)
                    {
                        var product = finalProducts[row];
                        var values = new Object[]
                        {
                            product.Id,                                 // id
                            product.Tipo,                               // tipo
                            product.SkuPai,                             // sku-pai
                            product.Sku,                                // sku
                            product.Ativo,                              // ativo
                            product.Usado,                              // usado
                            product.Destaque,                           // destaque
                            product.Ncm,                                // ncm
                            product.Gtin,                               // gtin
                            product.Mpn,                                // mpn
                            product.Nome,                               // nome
                            product.SeoTagTitle,                        // seo-tag-title
                            product.SeoTagDescription,                  // seo-tag-description
                            product.DescricaoCompleta,                  // descricao-completa
                            product.UrlVideoYoutube,                    // url-video-youtube
                            product.EstoqueGerenciado,                  // estoque-gerenciado
                            product.EstoqueQuantidade,                  // estoque-quantidade
                            product.EstoqueSituacaoEmEstoque,           // estoque-situacao-em-estoque
                            product.EstoqueSituacaoSemEstoque,          // estoque-situacao-sem-estoque
                            product.PrecoSobConsulta,                   // preco-sob-consulta
                            product.PrecoCusto,                         // preco-custo
                            product.PrecoCheio,                         // preco-cheio
                            product.PrecoPromocional,                   // preco-promocional
                            product.Marca,                              // marca
                            product.PesoEmKg,                           // peso-em-kg
                            product.AlturaEmCm,                         // altura-em-cm
                            product.LarguraEmCm,                        // largura-em-cm
                            product.ComprimentoEmCm,                    // comprimento-em-cm
                            product.CategoriaNomeNivel1,                // categoria-nome-nivel-1
                            product.CategoriaNomeNivel2,                // categoria-nome-nivel-2
                            product.CategoriaNomeNivel3,                // categoria-nome-nivel-3
                            product.CategoriaNomeNivel4,                // categoria-nome-nivel-4
                            product.CategoriaNomeNivel5,                // categoria-nome-nivel-5
                            product.Imagem1,                            // imagem-1
                            product.Imagem2,                            // imagem-2
                            product.Imagem3,                            // imagem-3
                            product.Imagem4,                            // imagem-4
                            product.Imagem5,                            // imagem-5
                            product.GradeGenero,                        // grade-genero
                            product.GradeProdutoComDuasCores,           // grade-produto-com-duas-cores
                            product.GradeProdutoComUmaCor,              // grade-produto-com-uma-cor
                            product.GradeTamanhoDeAnelAlianca,          // grade-tamanho-de-anelalianca
                            product.GradeTamanhoDeCalca,                // grade-tamanho-de-calca
                            product.GradeTamanhoDeCamisaCamiseta,       // grade-tamanho-de-camisacamiseta
                            product.GradeTamanhoDeCapacete,             // grade-tamanho-de-capacete
                            product.GradeTamanhoDeTenis,                // grade-tamanho-de-tenis
                            product.GradeTamanhoJuvenilInfantil,        // grade-tamanho-juvenil-infantil
                            product.GradeVoltagem,                      // grade-voltagem
                            product.GradeTamanho                        // grade-tamanho (ajuste conforme necessário)
                        };

                        for (int col = 0; col < values.Length; col++)
                        {
                            worksheet.Cells[row + 2, col + 1].Value = values[col];
                        }
                    }

                    // Autoajuste de largura das colunas
                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                    // Salva o arquivo Excel
                    package.SaveAs(new FileInfo(savePath.Replace(".csv", ".xlsx")));
                }

                ToggleControls(true);
                MessageBox.Show("Conversão concluída com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ToggleControls(true);
                MessageBox.Show("Erro ao converter: " + ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
