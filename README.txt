RECACUST RE-FX SITE - v12 (GitHub Pages)
========================================

v12 Güncellemeleri (GitHub Pages'e taşıma):
- Site GitHub Pages üzerinde yayınlanır; PHP çalışmadığı için iletişim formu artık
  Web3Forms ile gönderilir (assets/js/main.js içindeki WEB3FORMS_KEY). Anahtar boşken
  form, bilgileri hazır bir e-posta olarak açar. contact.php yalnızca PHP destekli
  hosting'e geri dönülürse kullanılır.
- Tailwind CDN kaldırıldı; derlenmiş CSS: assets/css/tailwind.css (bkz. tools/README.md).
- robots.txt, sitemap.xml ve 404.html eklendi; teşekkürler sayfası güncel tasarıma
  taşındı ve noindex yapıldı; yasal sayfalarda dışarıda kalan bölümler düzeltildi.
- Yapısal veri (JSON-LD) zenginleştirildi; boş LinkedIn bağlantıları kaldırıldı.


Teknoloji:
- HTML5
- Tailwind CSS CDN
- Vanilla JavaScript
- PHP mail endpoint

v11 Güncellemeleri:
- İletişim formu FormSubmit üzerinden kaldırıldı.
- Form artık contact.php dosyasına POST eder.
- Linux hosting/cPanel üzerinde PHP mail() fonksiyonu ile e-posta gönderimi yapılır.
- Google Maps iframe adresi Taşdelen Merkez, Çekmeköy, İstanbul olarak güncellendi.
- Görsel placeholder alanları gerçek dosyalarla değiştirildi:
  - assets/images/hero-refx.webp
  - assets/images/leasing-process.webp
- Görseller telif riski oluşturmaması için siteye özel soyut/kurumsal teknoloji görselleri olarak üretildi ve ZIP içine eklendi.

Önemli PHP Ayarı:
contact.php içinde aşağıdaki alanları canlıya almadan önce kontrol edin:

$to = 'info@recacust.com';
$from = 'no-reply@recacust.com';

Daha stabil mail teslimatı için:
- info@recacust.com adresi hosting üzerinde oluşturulmalı.
- no-reply@recacust.com adresi de mümkünse hosting üzerinde oluşturulmalı.
- Hosting firması PHP mail() fonksiyonunu kapattıysa SMTP/PHPMailer çözümüne geçilmelidir.

Dosya Yapısı:
- index.html
- hakkimizda.html
- cozumler.html
- sektorel.html
- kiralama.html
- ifrs16-asc842.html
- sss.html
- notlar.html
- iletisim.html
- tesekkurler.html
- contact.php
- cerez-politikasi.html
- kvkk-aydinlatma-metni.html
- gizlilik-politikasi.html
- kullanim-kosullari.html
- assets/css/custom.css
- assets/js/main.js
- assets/images/

Hosting Kurulum:
1. ZIP dosyasını açın.
2. İçindeki tüm dosya ve klasörleri public_html içine yükleyin.
3. index.html ana sayfa olarak çalışır.
4. İletişim formu için hostingde PHP desteği gerekir.
5. Formu test ederken spam klasörünü de kontrol edin.

Netlify Notu:
Netlify PHP çalıştırmaz. Bu nedenle contact.php Netlify üzerinde çalışmaz.
Form testi için Linux hosting veya PHP destekli test ortamı kullanın.


V12 NOTLARI:
- Footer iki kolonlu kurumsal yapıya alındı.
- 0216 370 24 78 numarası tel: linki olarak eklendi.
- WhatsApp butonu mobilde sağ alt köşeye küçültüldü ve footer linklerini kapatmayacak şekilde yukarı alındı.
- Yasal sayfalardaki not alanı kaldırıldı, içerikler detaylandırıldı.
- Görseller daha açık ve anlaşılır kurumsal illüstrasyonlarla güncellendi.


V13: Footer yapısı v11 düzenine geri alındı; 0216 telefon linki, WhatsApp linki ve yasal linkler korundu.
