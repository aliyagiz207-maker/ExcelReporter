# Excel Reporter

Excel Reporter, işletmeler için Excel ve CSV dosyalarını otomatik olarak analiz eden, KPI hesaplayan ve Excel ile PDF formatında rapor oluşturan Python tabanlı bir otomasyon uygulamasıdır.

---

## Özellikler

- Birden fazla Excel dosyasını otomatik okuma
- CSV dosyalarını destekleme
- Veri doğrulama
- Veri temizleme
- KPI (Temel Performans Göstergeleri) hesaplama
- Excel Dashboard oluşturma
- PDF raporu oluşturma
- Bölgelere göre gelir analizi
- En çok gelir getiren ürünler analizi
- Firma logosu desteği
- Otomatik rapor arşivleme
- JSON ile yapılandırılabilir yapı
- Ay bazlı raporlama (`--month` parametresi)

---

## Kullanılan Teknolojiler

- Python 3
- Pandas
- OpenPyXL
- ReportLab
- Loguru

---

## Proje Yapısı

```text
ExcelReporter/
│
├── assets/
│   └── logo.png
│
├── data/
│   ├── archive/
│   ├── input/
│   └── output/
│
├── logs/
├── src/
├── config.json
├── requirements.txt
└── README.md
```

---

## Kurulum

Projeyi bilgisayarınıza indirin:

```bash
git clone https://github.com/aliyagiz207-maker/ExcelReporter.git
```

Proje klasörüne girin:

```bash
cd ExcelReporter
```

Sanal ortam oluşturun:

```bash
python -m venv .venv
```

Sanal ortamı etkinleştirin:

**Windows**

```bash
.venv\Scripts\activate
```

Gerekli kütüphaneleri yükleyin:

```bash
pip install -r requirements.txt
```

---

## Kullanım

Tüm Excel ve CSV dosyalarını işlemek için:

```bash
python src/main.py
```

Sadece belirli ayları işlemek için:

```bash
python src/main.py --month january
```

Birden fazla ay seçmek için:

```bash
python src/main.py --month january march
```

---

## Oluşturulan Çıktılar

Program çalıştırıldığında aşağıdaki dosyalar oluşturulur:

- Excel Dashboard (`.xlsx`)
- PDF Raporu (`.pdf`)
- Arşivlenmiş Excel raporları
- Uygulama logları

---

## Yapılandırma

Tüm uygulama ayarları `config.json` dosyası üzerinden değiştirilebilir.

Örnek:

```json
{
    "company_name": "Honda Terakki",
    "dashboard_title": "Monthly Sales Dashboard",
    "currency": "$",
    "logo_path": "assets/logo.png",
    "output_file": "data/output/Honda_Report.xlsx"
}
```

---

## Gelecek Sürümler

Planlanan geliştirmeler:

- Dashboard grafiklerinin geliştirilmesi
- Daha fazla grafik türü
- Komut satırı seçeneklerinin artırılması
- Web arayüzü
- E-posta ile otomatik rapor gönderimi
- Bulut depolama desteği

---

## Geliştirici

**Ali Yağız Demir**


GitHub:
https://github.com/aliyagiz207-maker