# 📊 Excel Veri Analiz ve Raporlama

Excel dosyalarındaki verileri analiz etmek, hatalı kayıtları tespit etmek ve analiz sonuçlarını raporlamak amacıyla geliştirilmiş bir **C# Windows Forms masaüstü uygulamasıdır.**

Uygulama, yalnızca belirli bir Excel şablonuna bağlı kalmadan farklı sütun yapılarına sahip Excel dosyalarını dinamik olarak okuyabilir. Veriler üzerinde doğrulama işlemleri gerçekleştirerek başarılı ve hatalı kayıtları kullanıcıya sunar.

---

## 🎯 Projenin Amacı

Kurumsal ortamlarda Excel dosyaları üzerinden tutulan verilerde;

- Boş bırakılan alanlar
- Hatalı veri girişleri
- Geçersiz tarih formatları
- Sayısal olması gereken alanlara metin girilmesi
- Uygun olmayan karakter kullanımları

gibi veri kalitesi problemleri oluşabilmektedir.

Bu uygulamanın amacı, **orijinal veriyi değiştirmeden** bu tür problemleri tespit etmek ve kullanıcıya raporlamaktır.

---

## ✨ Özellikler

- 📂 `.xlsx` formatındaki Excel dosyalarını yükleme
- 🔄 Farklı sütun yapılarına sahip Excel dosyalarını dinamik olarak okuma
- 🔍 Satır bazlı veri doğrulama
- ⚠️ Hatalı kayıtların ve hata nedenlerinin görüntülenmesi
- ✅ Başarılı ve hatalı kayıtların ayrıştırılması
- 🎨 DataGridView üzerinde başarılı ve hatalı satırların görsel olarak işaretlenmesi
- 📊 Toplam, başarılı ve hatalı kayıt sayılarının gösterilmesi
- 🍩 Başarı oranının grafik üzerinde görüntülenmesi
- 📑 Analiz sonucunun yeni bir Excel raporu olarak oluşturulması
- 📧 Oluşturulan raporun e-posta üzerinden gönderilebilmesi
- 🖥️ Windows masaüstü uygulaması olarak kullanılabilme

---

## 🔎 Analiz Mantığı

Kullanıcı bir Excel dosyası seçtiğinde uygulama ilk satırdaki sütun başlıklarını otomatik olarak algılar ve verileri dinamik şekilde okur.

Analiz sırasında tespit edilen hatalar ilgili satırlarla birlikte kullanıcıya gösterilir.

Örneğin:

```text
Satır 24: Maaş alanında geçersiz değer bulundu.
Satır 57: İşe giriş tarihi geçerli bir tarih değil.
Satır 103: Zorunlu alan boş bırakılamaz.
```

Uygulama hatalı verileri **otomatik olarak değiştirmez**. Böylece yanlış bir düzeltme yapılarak mevcut verinin bozulmasının önüne geçilir.

---

## 📄 Raporlama

Analiz tamamlandıktan sonra kullanıcı yeni bir Excel raporu oluşturabilir.

Oluşturulan raporda kayıtlar:

- **Başarılı Kayıtlar**
- **Hatalı Kayıtlar**

olmak üzere ayrı çalışma sayfalarına aktarılır.

Hatalı kayıtların yanında ilgili **hata açıklaması** da rapora eklenir.

---

## 🗂️ Farklı Veri Yapıları

Uygulama yalnızca tek bir veri türü için tasarlanmamıştır.

Test sürecinde farklı Excel yapıları kullanılmıştır:

### İnsan Kaynakları

Personel bilgileri, birim, maaş, işe giriş/çıkış tarihleri ve izin bilgileri gibi alanlardan oluşan veriler.

### Muhasebe / Satın Alma

Firma, ürün, miktar, satın alma ve diğer işlem bilgilerinden oluşan veriler.

### Üretim

Ürün kodu, ürün adı, parti numarası, üretim tarihi, miktar ve depo bilgilerinden oluşan veriler.

Bu yapı sayesinde uygulama farklı departmanlardan gelen Excel dosyalarının kontrolünde kullanılabilir.

---

## 🛠️ Kullanılan Teknolojiler

| Teknoloji | Kullanım |
|---|---|
| C# | Uygulama geliştirme |
| .NET | Uygulama altyapısı |
| Windows Forms | Masaüstü kullanıcı arayüzü |
| EPPlus | Excel dosyalarını okuma ve rapor oluşturma |
| MailKit | E-posta gönderme |
| DataGridView | Excel verilerini görüntüleme |

---

## 🚀 Kullanım

1. Uygulamayı başlatın.
2. **Excel Dosyası Seç** butonuyla `.xlsx` dosyanızı yükleyin.
3. **Analiz Et** butonuna basın.
4. Başarılı ve hatalı kayıtları inceleyin.
5. Hata listesinden tespit edilen problemleri görüntüleyin.
6. **Rapor Oluştur** ile analiz sonucunu Excel dosyası olarak kaydedin.
7. İstenirse oluşturulan raporu e-posta üzerinden gönderin.

---

## 📸 Uygulama Görüntüsü

> Bu bölüme uygulamanın ana ekran görüntüsü eklenecektir.

---

## 📁 Proje Yapısı

```text
Raporlama
│
├── Models
│   └── Veri modelleri
│
├── Services
│   ├── ExcelService
│   └── ValidationService
│
├── Form1
├── MailForm
└── Program
```

`ExcelService`, Excel dosyalarının okunması ve rapor oluşturulmasından sorumludur.

`ValidationService`, yüklenen verilerin doğrulama işlemlerini gerçekleştirir.

---

## 🔐 Veri Güvenliği

Uygulama analiz sırasında kaynak Excel dosyasındaki verileri değiştirmez.

Tespit edilen problemler kullanıcıya bildirilir ve analiz sonucu ayrı bir rapor dosyası olarak oluşturulur.

---

## 👩‍💻 Geliştirici

**Eylem Gökçe**

Bilgisayar Mühendisliği

Sistem, ağ teknolojileri ve yazılım geliştirme alanlarında çalışmalar yapıyorum.

---

## 📌 Proje Durumu

Proje aktif olarak geliştirilmektedir.

İlerleyen sürümlerde doğrulama kurallarının genişletilmesi ve raporlama özelliklerinin geliştirilmesi planlanmaktadır.
