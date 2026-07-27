using Raporlama.Models;

namespace Raporlama.Services;

public class ValidationService
{
    // =========================================================
    // ESKİ ÜRETİM RAPORU DOĞRULAMASI
    // =========================================================

    public ValidationResult Validate(UretimKaydi kayit)
    {
        ValidationResult result = new();

        if (string.IsNullOrWhiteSpace(kayit.UrunKodu))
            result.Errors.Add("Ürün kodu boş.");

        if (string.IsNullOrWhiteSpace(kayit.UrunAdi))
            result.Errors.Add("Ürün adı boş.");

        if (string.IsNullOrWhiteSpace(kayit.PartiNo))
            result.Errors.Add("Parti numarası boş.");

        if (!DateTime.TryParse(kayit.UretimTarihi, out _))
            result.Errors.Add("Geçersiz tarih.");

        if (!int.TryParse(kayit.Miktar, out int miktar))
        {
            result.Errors.Add("Miktar sayısal değil.");
        }
        else if (miktar <= 0)
        {
            result.Errors.Add("Miktar 0'dan büyük olmalıdır.");
        }

        string[] depolar = { "A1", "A2", "B1", "B2" };

        if (!depolar.Contains(kayit.Depo))
            result.Errors.Add("Geçersiz depo kodu.");

        result.IsValid = result.Errors.Count == 0;

        return result;
    }


    // =========================================================
    // GENEL AMAÇLI EXCEL DOĞRULAMASI
    // =========================================================

    public void ValidateDynamic(List<ExcelDataRow> rows)
    {
        if (rows.Count == 0)
            return;

        foreach (var row in rows)
            row.Errors.Clear();

        List<string> headers =
            rows.First().Values.Keys.ToList();

        // Boş bırakılmasına izin verdiğimiz bilinen alanlar
        string[] optionalHeaders =
        {
        "İşten Çıkış Tarihi",
        "Açıklama",
        "Not",
        "Ek Açıklama",
        "İkinci Telefon"
    };

        foreach (string header in headers)
        {
            List<string> allValues = rows
                .Select(r => r.Values[header] ?? "")
                .ToList();

            List<string> filledValues = allValues
                .Where(v => !string.IsNullOrWhiteSpace(v))
                .ToList();

            if (filledValues.Count == 0)
                continue;

            // Sütunun yüzde kaçı boş?
            int emptyCount = allValues.Count(v =>
                string.IsNullOrWhiteSpace(v));

            double emptyRatio =
                (double)emptyCount / allValues.Count;

            // Alan bilinen opsiyonel bir alan mı?
            bool knownOptional =
                optionalHeaders.Any(x =>
                    x.Equals(
                        header,
                        StringComparison.OrdinalIgnoreCase));

            // Yarısından fazlası boşsa opsiyonel kabul et
            bool optionalColumn =
                knownOptional || emptyRatio >= 0.50;

            // Sütunun veri tipini tahmin et
            int numericCount =
                filledValues.Count(v =>
                    double.TryParse(v, out _));

            int dateCount =
                filledValues.Count(v =>
                    DateTime.TryParse(v, out _));

            bool numericColumn =
                numericCount >= filledValues.Count * 0.80;

            bool dateColumn =
                dateCount >= filledValues.Count * 0.80;

            foreach (ExcelDataRow row in rows)
            {
                string value =
                    row.Values[header] ?? "";

                // -----------------------------
                // BOŞ DEĞER
                // -----------------------------

                if (string.IsNullOrWhiteSpace(value))
                {
                    // Opsiyonel değilse hata
                    if (!optionalColumn)
                    {
                        row.Errors.Add(
                            $"{header}: Boş değer.");
                    }

                    continue;
                }

                // -----------------------------
                // SAYISAL DEĞER
                // -----------------------------

                if (numericColumn &&
                    !double.TryParse(value, out _))
                {
                    row.Errors.Add(
                        $"{header}: Sayısal değer bekleniyor, " +
                        $"'{value}' bulundu.");
                }

                // -----------------------------
                // TARİH
                // -----------------------------

                if (dateColumn &&
                    !DateTime.TryParse(value, out _))
                {
                    row.Errors.Add(
                        $"{header}: Geçersiz tarih değeri '{value}'.");
                }
            }
        }
    }
}