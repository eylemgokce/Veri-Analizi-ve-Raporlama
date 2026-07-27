using OfficeOpenXml;
using Raporlama.Models;

namespace Raporlama.Services;

public class ExcelService
{
    public List<UretimKaydi> ReadExcel(string filePath)
    {
        ExcelPackage.License.SetNonCommercialPersonal("Eylem");

        using var package = new ExcelPackage(new FileInfo(filePath));

        if (package.Workbook.Worksheets.Count == 0)
        {
            throw new Exception("Excel sayfası bulunamadı.");
        }

        var worksheet = package.Workbook.Worksheets[0];

        List<UretimKaydi> kayitlar = new();

        for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
        {
            kayitlar.Add(new UretimKaydi
            {
                UrunKodu = worksheet.Cells[row, 1].Text,
                UrunAdi = worksheet.Cells[row, 2].Text,
                PartiNo = worksheet.Cells[row, 3].Text,
                UretimTarihi = worksheet.Cells[row, 4].Text,
                Miktar = worksheet.Cells[row, 5].Text,
                Depo = worksheet.Cells[row, 6].Text
            });
        }

        return kayitlar;
    }


    // =========================================================
    // HERHANGİ BİR EXCEL DOSYASINI DİNAMİK OLARAK OKUR
    // =========================================================

    public List<ExcelDataRow> ReadDynamicExcel(string filePath)
    {
        ExcelPackage.License.SetNonCommercialPersonal("Eylem");

        using var package = new ExcelPackage(new FileInfo(filePath));

        if (package.Workbook.Worksheets.Count == 0)
        {
            throw new Exception("Excel sayfası bulunamadı.");
        }

        var worksheet = package.Workbook.Worksheets[0];

        if (worksheet.Dimension == null)
        {
            throw new Exception("Excel dosyası boş.");
        }

        List<string> headers = new();

        // İlk satırdaki sütun başlıklarını oku
        for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
        {
            string header = worksheet.Cells[1, col].Text.Trim();

            // Başlık boşsa otomatik isim oluştur
            if (string.IsNullOrWhiteSpace(header))
            {
                header = $"Sütun_{col}";
            }

            // Aynı isimli sütun varsa benzersiz hale getir
            string originalHeader = header;
            int counter = 2;

            while (headers.Contains(header))
            {
                header = $"{originalHeader}_{counter}";
                counter++;
            }

            headers.Add(header);
        }

        List<ExcelDataRow> rows = new();

        // Excel satırlarını oku
        for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
        {
            ExcelDataRow excelRow = new()
            {
                RowNumber = row
            };

            bool tamamenBos = true;

            for (int col = 1; col <= headers.Count; col++)
            {
                string value = worksheet.Cells[row, col].Text.Trim();

                if (!string.IsNullOrWhiteSpace(value))
                {
                    tamamenBos = false;
                }

                excelRow.Values[headers[col - 1]] = value;
            }

            // Tamamen boş satırları sisteme alma
            if (!tamamenBos)
            {
                rows.Add(excelRow);
            }
        }

        return rows;
    }


    // =========================================================
    // RAPOR OLUŞTUR
    // =========================================================

    public void CreateDynamicReport(
    string filePath,
    
    List<ExcelDataRow> rows)
    {
        ExcelPackage.License.SetNonCommercialPersonal("Eylem");

        using var package = new ExcelPackage();

        var basariliSheet =
            package.Workbook.Worksheets.Add("Başarılı Kayıtlar");

        var hataliSheet =
            package.Workbook.Worksheets.Add("Hatalı Kayıtlar");

        if (rows.Count == 0)
        {
            throw new Exception("Rapor oluşturulacak veri bulunamadı.");
        }

        List<string> headers =
            rows[0].Values.Keys.ToList();

        // Başlıklar
        for (int i = 0; i < headers.Count; i++)
        {
            basariliSheet.Cells[1, i + 1].Value = headers[i];
            hataliSheet.Cells[1, i + 1].Value = headers[i];
        }

        // Hatalı kayıtlar için ekstra sütun
        hataliSheet.Cells[1, headers.Count + 1].Value =
            "Hata Açıklaması";

        int basariliRow = 2;
        int hataliRow = 2;

        foreach (ExcelDataRow row in rows)
        {
            if (row.IsValid)
            {
                WriteDynamicRow(
                    basariliSheet,
                    basariliRow,
                    headers,
                    row);

                basariliRow++;
            }
            else
            {
                WriteDynamicRow(
                    hataliSheet,
                    hataliRow,
                    headers,
                    row);

                hataliSheet.Cells[
                    hataliRow,
                    headers.Count + 1].Value =
                    string.Join(" | ", row.Errors);

                hataliRow++;
            }
        }

        basariliSheet.Cells.AutoFitColumns();
        hataliSheet.Cells.AutoFitColumns();

        package.SaveAs(new FileInfo(filePath));
    }

    private void WriteDynamicRow(
    ExcelWorksheet sheet,
    int rowIndex,
    List<string> headers,
    ExcelDataRow row)
    {
        for (int i = 0; i < headers.Count; i++)
        {
            string header = headers[i];

            sheet.Cells[rowIndex, i + 1].Value =
                row.Values.TryGetValue(header, out string? value)
                    ? value
                    : "";
        }
    }


}