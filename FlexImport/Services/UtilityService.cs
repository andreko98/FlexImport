using CsvHelper;
using CsvHelper.Configuration;
using System.Globalization;
using System.Reflection;
using CsvHelper.TypeConversion;

namespace FlexImport.Services
{
    public static class UtilityService
    {
        public static readonly CsvConfiguration Config = new(CultureInfo.InvariantCulture)
        {
            Delimiter = ";",
            Encoding = System.Text.Encoding.UTF8,
            BadDataFound = null,
            MissingFieldFound = null,
            HeaderValidated = null
        };

        public static readonly List<string> finalHeader = new()
           {
               "id",
               "tipo",
               "sku-pai",
               "sku",
               "ativo",
               "usado",
               "destaque",
               "ncm",
               "gtin",
               "mpn",
               "nome",
               "seo-tag-title",
               "seo-tag-description",
               "descricao-completa",
               "url-video-youtube",
               "estoque-gerenciado",
               "estoque-quantidade",
               "estoque-situacao-em-estoque",
               "estoque-situacao-sem-estoque",
               "preco-sob-consulta",
               "preco-custo",
               "preco-cheio",
               "preco-promocional",
               "marca",
               "peso-em-kg",
               "altura-em-cm",
               "largura-em-cm",
               "comprimento-em-cm",
               "categoria-nome-nivel-1",
               "categoria-nome-nivel-2",
               "categoria-nome-nivel-3",
               "categoria-nome-nivel-4",
               "categoria-nome-nivel-5",
               "imagem-1",
               "imagem-2",
               "imagem-3",
               "imagem-4",
               "imagem-5",
               "grade-genero",
               "grade-produto-com-duas-cores",
               "grade-produto-com-uma-cor",
               "grade-tamanho-de-anelalianca",
               "grade-tamanho-de-calca",
               "grade-tamanho-de-camisacamiseta",
               "grade-tamanho-de-capacete",
               "grade-tamanho-de-tenis",
               "grade-tamanho-juvenil-infantil",
               "grade-voltagem",
               "grade-tamanho"
           };

        public static readonly Dictionary<string, string[]> Categories = new()
           {
               { "Meninas", new[] { "Conjunto Curto", "Conjunto Longo", "Macacão", "Kits" } },
               { "Meninos", new[] { "Conjunto Curto", "Conjunto Longo", "Macacão", "Kits" } },
               { "Unissex", new[] { "Conjunto Curto", "Conjunto Longo", "Macacão", "Kits" } },
               { "Conjunto Curto", Array.Empty<string>() },
               { "Conjunto Longo", Array.Empty<string>() },
               { "Macacão", Array.Empty<string>() },
               { "Kits", Array.Empty<string>() }
           };

        public static readonly List<string> Genre = new()
           {
               "Masculino",
               "Feminino",
               "Unissex"
           };

        public static readonly List<string> Sizes = new()
           {
               "RN", "P", "M", "G", "GG"
           };

        private static readonly string SkuFilePath = Path.Combine(
            Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? "",
            "skus_gerados.txt"
        );

        public static string GenerateRandomSKU(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            HashSet<string> oldSku = new();

            if (!File.Exists(SkuFilePath))
            {
                File.Create(SkuFilePath).Dispose();
            }

            foreach (var sku in File.ReadAllLines(SkuFilePath))
            {
                if (!string.IsNullOrWhiteSpace(sku))
                    oldSku.Add(sku.Trim());
            }

            var random = new Random();
            string newSku;

            // Gera até encontrar um SKU único  
            do
            {
                newSku = new string(Enumerable.Repeat(chars, length)
                    .Select(s => s[random.Next(s.Length)]).ToArray());
            }
            while (oldSku.Contains(newSku));

            // Salva o novo SKU no arquivo  
            File.AppendAllLines(SkuFilePath, new[] { newSku });

            return newSku;
        }

        public static void RegisterPtBrNumberConverters(CsvReader csv)
        {
            csv.Context.TypeConverterCache.AddConverter<int>(new Int32PtBrConverter());
            csv.Context.TypeConverterCache.AddConverter<int?>(new Int32PtBrConverter());
            csv.Context.TypeConverterCache.AddConverter<decimal>(new DecimalPtBrConverter());
            csv.Context.TypeConverterCache.AddConverter<decimal?>(new DecimalPtBrConverter());
            csv.Context.TypeConverterCache.AddConverter<float>(new FloatPtBrConverter());
            csv.Context.TypeConverterCache.AddConverter<float?>(new FloatPtBrConverter());
            csv.Context.TypeConverterCache.AddConverter<double>(new DoublePtBrConverter());
            csv.Context.TypeConverterCache.AddConverter<double?>(new DoublePtBrConverter());
        }
    }

    public class Int32PtBrConverter : Int32Converter
    {
        public override object ConvertFromString(string? text, IReaderRow row, MemberMapData memberMapData)
        {
            if (string.IsNullOrWhiteSpace(text))
                return 0;

            text = text.Trim().Replace(",", ".");

            if (decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out var dec))
                return (int)dec;

            if (int.TryParse(text, out var i))
                return i;

            return base.ConvertFromString(text, row, memberMapData);
        }
    }

    public class DecimalPtBrConverter : DecimalConverter
    {
        public override object ConvertFromString(string? text, IReaderRow row, MemberMapData memberMapData)
        {
            if (string.IsNullOrWhiteSpace(text))
                return 0m;

            text = text.Trim().Replace(",", ".");

            if (decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out var dec))
                return dec;

            return base.ConvertFromString(text, row, memberMapData);
        }
    }

    public class FloatPtBrConverter : SingleConverter
    {
        public override object ConvertFromString(string? text, IReaderRow row, MemberMapData memberMapData)
        {
            if (string.IsNullOrWhiteSpace(text))
                return 0f;

            text = text.Trim().Replace(",", ".");

            if (float.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out var f))
                return f;

            return base.ConvertFromString(text, row, memberMapData);
        }
    }

    public class DoublePtBrConverter : DoubleConverter
    {
        public override object ConvertFromString(string? text, IReaderRow row, MemberMapData memberMapData)
        {
            if (string.IsNullOrWhiteSpace(text))
                return 0d;

            text = text.Trim().Replace(",", ".");

            if (double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out var d))
                return d;

            return base.ConvertFromString(text, row, memberMapData);
        }
    }

}
