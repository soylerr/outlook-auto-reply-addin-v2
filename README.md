# Outlook Otomatik Yanıt Eklentisi

Bu Outlook eklentisi, şirket çalışanlarının standart otomatik yanıt mesajlarını kolayca ayarlamalarını sağlar.

## Özellikler

- 🇹🇷 Türkçe ve 🇬🇧 İngilizce mesaj şablonları
- 👥 D365 entegrasyonu ile yetkili kişi seçimi
- 📅 Tarih ve saat seçimi
- 📧 Canlı mesaj önizleme
- 🔒 Sabit şablon yapısı (kullanıcı içeriği değiştiremez)

## Kurulum

1. `manifest.xml` dosyasını Outlook'a yükleyin
2. Eklenti otomatik olarak yüklenecektir

## Kullanım

1. Outlook'ta eklenti butonuna tıklayın
2. Yetkili kişiyi seçin
3. Başlangıç ve bitiş tarihlerini ayarlayın
4. "Otomatik Yanıtı Ayarla" butonuna tıklayın

## Geliştirme

Bu eklenti Office.js API'sini kullanır ve GitHub Pages üzerinde barındırılır.

### Yerel Geliştirme

```bash
# Basit HTTP sunucusu başlatın
python -m http.server 8080
# veya
npx http-server
```

### Deployment

Eklenti GitHub Pages üzerinde otomatik olarak yayınlanır.

## Lisans

MIT License
