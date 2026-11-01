#!/usr/bin/env python3
"""RECACUST revizyonu (Kasım 2026, RECACUST_Websitesi.xlsx).

Tüm sayfalardaki üst menü, mobil menü ve alt bilgi tek kaynaktan üretilir;
Excel'deki ESKİ → YENİ metin değişiklikleri uygulanır; Bilgi Merkezi,
Kiralama & Finans ve Muhasebe & Bütçe Raporlama sayfaları oluşturulur.

Çalıştırma (repo kökünde):  python3 tools/revizyon.py
Betik tekrar çalıştırılabilir; menü/alt bilgi her seferinde yeniden üretilir.
"""
import html, re, pathlib

ROOT = pathlib.Path(__file__).resolve().parent.parent
E = html.escape

PHONE_DISPLAY = '+90 216 370 24 78'
PHONE_TEL = 'tel:+902163702478'
EMAIL = 'info@recacust.com'
WA = ('https://wa.me/905418638613?text=Merhaba%2C%20SAP%20RE-FX%20dan%C4%B1%C5%9Fmanl%C4%B1%C4%9F%C4%B1'
      '%20hakk%C4%B1nda%20bilgi%20almak%20istiyorum.')

ICON = {
    'phone': '<svg class="ci" viewBox="0 0 24 24" aria-hidden="true"><path d="M6.6 10.8a15.1 15.1 0 0 0 6.6 6.6l2.2-2.2a1 1 0 0 1 1-.25 11.4 11.4 0 0 0 3.6.57 1 1 0 0 1 1 1V20a1 1 0 0 1-1 1A17 17 0 0 1 3 4a1 1 0 0 1 1-1h3.5a1 1 0 0 1 1 1c0 1.25.2 2.45.57 3.57a1 1 0 0 1-.25 1z"/></svg>',
    'mail': '<svg class="ci" viewBox="0 0 24 24" aria-hidden="true"><path d="M3 5h18a1 1 0 0 1 1 1v12a1 1 0 0 1-1 1H3a1 1 0 0 1-1-1V6a1 1 0 0 1 1-1zm9 7.2L4 7.3V17h16V7.3zM4.9 7l7.1 4.3L19.1 7z"/></svg>',
    'wa': '<svg class="ci" viewBox="0 0 24 24" aria-hidden="true"><path d="M12 2a10 10 0 0 0-8.6 15.1L2 22l5-1.3A10 10 0 1 0 12 2zm0 18.2a8.2 8.2 0 0 1-4.2-1.15l-.3-.18-3 .78.8-2.9-.2-.3A8.2 8.2 0 1 1 12 20.2zm4.5-6.1c-.25-.12-1.47-.72-1.7-.8s-.39-.12-.56.12-.64.8-.78.97-.29.18-.54.06a6.7 6.7 0 0 1-3.3-2.9c-.25-.43.25-.4.71-1.33a.45.45 0 0 0-.02-.42c-.06-.12-.56-1.34-.77-1.84s-.4-.42-.56-.43h-.48a.92.92 0 0 0-.67.31 2.8 2.8 0 0 0-.87 2.08 4.86 4.86 0 0 0 1.02 2.58 11.1 11.1 0 0 0 4.25 3.76c1.58.68 2.2.74 2.99.62a2.55 2.55 0 0 0 1.68-1.18 2.08 2.08 0 0 0 .14-1.18c-.06-.1-.23-.16-.48-.28z"/></svg>',
}

# ---------------------------------------------------------------------------
# İÇERİK (Excel "BAŞLIKLAR Yeni" + konu sayfaları). Bloklar:
#   ('p', metin) | ('ul', [..]) | ('dl', [(terim, açıklama)]) | ('img', dosya, alt, genişlik, yükseklik)
#   ('table', [başlıklar], [[satır]]) | ('note', metin) | ('h4', metin) | ('links', [(href, metin)])
# ---------------------------------------------------------------------------

SERVICES = {  # cozumler.html
    'kicker': 'Hizmetlerimiz',
    'items': [
        dict(id='sap-re-fx-danismanligi', title='SAP RE-FX Danışmanlığı',
             desc='Kurulum, uyarlama, canlı geçiş ve operasyonel destek süreçleri.', body=[
            ('h4', 'Sıfırdan Proje'),
            ('p', "Hiç SAP kullanmamış ya da SAP'nin diğer modüllerini kullanıp Esnek Gayrimenkul modülü ile yeni tanışacak şirketlere uygulanan proje biçimidir."),
            ('p', 'Hazır ana veri aktarım şablonlarımız ile ya da kaynak 3. parti sistemlerden beslenerek ana veri yaratılması.'),
            ('h4', 'Dönüşüm Projeleri'),
            ('p', 'Hali hazırda ECC ya da HANA kullanan ve S/4HANA ya da RISE dönüşümü yapacak şirketlere uygulanan proje biçimidir.'),
            ('p', 'Greenfield, Brownfield ve Bluefield ayrımı yapmadan analiz ve iyileştirme çalışması yapılıp raporlanması.'),
            ('h4', 'Destek İşleri'),
            ('p', 'Direkt müşteri ile ya da başka bir SAP Danışmanlık firması üzerinden gelen destek biletlerinde yardım temin edilmesidir.'),
            ('p', 'Direkt müşteri üzerinden gelen biletlerde IT departmanına çözümü raporluyor, bir daha tekrarlanmaması için azami gayret gösteriyoruz. SAP danışmanlık firması üzerinden gelen biletler için ise aktarım yapılacak SAP danışmanına yön gösteriyoruz.'),
            ('p', 'Kotalı ya da kotasız destek paketlerimizden size uygun olanını seçebilirsiniz.'),
            ('h4', 'Rehberlik Hizmetleri'),
            ('p', 'SAP Esnek Gayrimenkul projesi ya da destek işi alan danışmanlık firmalarının tercih ettiği bir yöntemdir. Alt yüklenici ya da ilk yüklenici olmanız fark etmeksizin arka tarafta yapılan işleri kontrol edip danışmanlık aktivitelerinizde rehberlik sunuyoruz.'),
            ('h4', 'Eğitim Hizmetleri'),
            ('p', 'SAP danışmanlık şirketi olan firmaların Junior ve Mid seviyedeki danışmanlarına ya da SAP Gayrimenkul modülü kullanan şirketlerin çalışanlarına verilen eğitim hizmetidir.'),
        ]),
        dict(id='egitim-ve-kullanici-destegi', title='Eğitim ve Kullanıcı Desteği',
             desc='Anahtar kullanıcı eğitimleri, süreç dokümantasyonu ve soru-cevap desteği.', body=[
            ('h4', 'Kullanıcı kılavuzu'),
            ('ul', ['Şirketinize özel kullanıcı kılavuzları en ince ayrıntısına kadar hazırlanır.',
                    'İş akışınıza göre farklı kullanıcı dokümanları iletilir.',
                    'Toplu işlemler için iş sürecine özel varyantlar ve kümeler tanımlanarak adım adım “To Do List” şeklinde hazırlanır.',
                    'Püf noktalar belirtilir.',
                    'Her bir senaryo için örnek veri yer alır.',
                    'Tüm kullanıcı kılavuzlarında her bir adım için ekran görüntüsü yer alır. Kullanılan butonlar renklendirilir.']),
            ('h4', 'Eğitim'),
            ('ul', ['Anahtar kullanıcı eğitimlerinde her bir senaryo için süreç çevrilir.',
                    'Tercihen “onsite” yapılması arzu edilir.',
                    'Verilen eğitimlerin ardından mutlaka raporlar üzerinden kontrol edilmesi sağlanır.']),
            ('h4', 'Öngörülen cevap ve çözüm süresi'),
            ('table', ['Seviye', 'Çalışma Zamanı', 'Cevap Süresi', 'Çözüm Süresi'], [
                ['Acil (P1)', '7/24', '30 dakika', '4 saat'],
                ['Yüksek (P2)', 'Mesai saatleri', '1 saat', '8 saat'],
                ['Orta (P3)', 'Mesai saatleri', '4 saat', '1 gün'],
                ['Düşük (P4)', 'Mesai saatleri', '8 saat', '2 gün']]),
        ]),
        dict(id='sektorunuze-gore-analiz', title='Sektörünüze Göre Analiz',
             desc='Sektörel yapınıza göre en doğru SAP RE-FX kapsamını belirleyelim.', body=[
            ('p', 'RECACUST Yazılım Bilişim Danışmanlık, gayrimenkul alanında tüm sektörlerde yapmış olduğu projeler ile iş süreçlerinizi daha önceden tecrübe etmiş ve iş süreçlerinizi optimize etmiştir.'),
            ('h4', 'Hangi şirketler kullanabilir?'),
            ('p', 'SAP RE-FX, birçok sektörde kullanılan ve geniş gayrimenkul portföylerine ya da kiralama sözleşmelerine sahip şirketlerde tercih edilen bir modüldür. Kiralık mülklerin, alanların, sözleşmelerin ve IFRS 16 uyumlu kiralama süreçlerinin yönetimi için uygundur.'),
            ('h4', 'Hangi sektörler tercih eder?'),
            ('ul', ['Gayrimenkul Yatırım Ortaklıkları (GYO)', 'Perakende şirketleri', 'Havayolu şirketleri', 'Leasing şirketleri',
                    'Telekomünikasyon şirketleri', 'Turizm & Otelcilik', 'Hastane ve klinik gibi rezervasyon usulü çalışan şirketler',
                    'Kiraya verme işlemi yapan şirketler (araç, ekipman, demirbaş vb.)', 'Mağazacılık şirketleri',
                    'Restoran ve yemek servisinde bulunan şirketler', 'Holdingler', 'Akaryakıt şirketleri']),
            ('links', [('sektorel.html', 'Sektörel kullanım alanlarını inceleyin')]),
        ]),
    ],
    'feature': dict(kicker='Odak Hizmet', title='SAP RE-FX projelerinde uçtan uca uzmanlık.',
                    text='Kurulumdan canlı geçişe, IFRS16 uyumundan kullanıcı desteğine kadar gayrimenkul süreçlerini SAP standardında yönetilebilir hale getiriyoruz.',
                    href='iletisim.html', link='Ön görüşme planlayın'),
}

BILGI = {  # bilgi-merkezi.html
    'groups': [
        ('Ana Veri Yönetimi', [
            dict(id='kullanim-gorunumu', title='Kullanım Görünümü',
                 desc='Bina, arsa, kat, alan ve kullanım görünümünün doğru modellenmesi.', body=[
                ('p', 'Şirketinizin faaliyet alanına göre gayrimenkul objeleri bir ağaç gibi yapılandırılır. Bir GYO şirketi için başka, perakende zinciri bir şirket için ise farklı bir modelleme olur.'),
                ('dl', [('İşletme Yapısı – DVG (Business Entity)', 'Gayrimenkul yönetimi modülünün en üst birimidir. Şirket kodu ile gayrimenkul yönetimi bağlantısının kurulması bu ana veri öğesi üzerinden gerçekleşir.'),
                        ('Bina (Building)', 'Gayrimenkul yönetimi modülünde kiralanabilir alanların bulunduğu binaların tanımlanması için kullanılır.'),
                        ('Arsa (Land)', 'Gayrimenkul yönetimi modülünde kiralanabilir alanların bulunduğu arazilerin tanımlanması için kullanılır.'),
                        ('Alan Havuzu (Pooled Space)', 'Farklı kiralık alanların bağlanarak genel bir grup altında toplanabileceği gayrimenkul yönetimi modülü objesidir.'),
                        ('Kiralık Alan (Rental Space)', 'Mali yapının en alt birimidir. Bu birim dahilinde kiralanabilir alanlar tanımlanır.'),
                        ('Kiralık Nesne (Rental Unit)', 'Alan havuzundan bağımsız olarak açılabilecek kira alanlarıdır.')]),
                ('ul', ['Buradaki asli amaç, hem en üst yapıda hem de en alt yapıda tüm P&L etkisinin aynı anda raporlanabilmesinin sağlanmasıdır.',
                        'Boşta kalan kiralık nesnelerin anlık ya da gelecek zaman verilerine göre aksiyon alınması sağlanır.',
                        'Organizasyon yapısının kullanım türlerine göre işlem verisi oluşturulur ve asli amacından farklı kullanımı engellenmiş olur.',
                        'Arzu edilen tüm raporların hem en üst düzeyde hem de en alt düzeyde alınabilmesini mümkün kılar; karşılaştırma imkânı sunar.']),
            ]),
            dict(id='mimari-gorunum', title='Mimari Görünüm',
                 desc='Bina kompleksi, bina bölümü, arsa alt alanı, park alanı gibi kullanım görünümü ile entegre modelleme.', body=[
                ('p', 'Mimari görünüm, RE-FX projelerinde kullanım görünümünün aksine opsiyoneldir. Kullanım görünümünün eksik kaldığı noktalarda kullanılması uygundur.'),
                ('p', 'Bir AVM işletmesinde “Outlet” bölümü ile “Lüks Tüketim” bölümlerini ayrı ayrı takip etmek için kullanım görünümüne ek olarak daha üst bir yapı sağlayabildiği gibi, bir katta yer alan odaların bütününün toplantı ve başka hizmetler için ayrılmasını da sağlar.'),
                ('p', 'Bu veriler sistemde ayraç yardımı ile numaralandırılır. Nokta, virgül, tire gibi ayraçlar yardımı ile bir alt düzeye inilir.'),
                ('note', 'Örnek: 10000000-0003.00001-32.B gösterimi; 10000000 numaralı komplekste 3. bloktaki 1. kata ait 32 numaralı balkonu ifade eder. “A” olsaydı iç mimari yapıyı temsil ediyor olabilirdi.'),
                ('p', 'Mimari görünüm şu durumlarda kullanılabilir:'),
                ('ul', ['Oda rezervasyonları ve uzun vadeli oturma düzenlemelerine ait fonksiyonlar kullanılmak ve mimari nesne türlerinden rezervasyon nesneleri oluşturulmak isteniyorsa,',
                        'Grafik sistemlere bağlantı kurulmak isteniyorsa,',
                        'Ortak kullanım alanlarından kiralanabilir alan çıkarılmak isteniyorsa.']),
            ]),
            dict(id='bilgi-sistemi', title='Bilgi Sistemi',
                 desc='Altyapı ve donanım karakteristikleri, bölgesel değerler, ulaşım verileri, ölçüm değerleri.', body=[
                ('dl', [('Bölgesel bilgiler', 'Konumla ilgili bilgiler yer alır; adres bilgileri girilir ve bu adres verileri yazışmalarda kullanılabilir. Konum bilgisine göre arama talepleri yapılabilir.'),
                        ('Donanımsal bilgiler', 'İlgili nesnede yer alan malzeme bilgilerinin tutulmasını sağlar ve bir üst başlıkta toplanır; örneğin yer döşemesinin ve kapının malzeme bilgisi ayrıdır. Yer döşemesi için parke, fayans, halıfleks gibi detaylar seçilebilir. Malzeme bilgisinden ayrı olarak projektör, PC, beyaz tahta gibi ekipman bilgileri de girilebilir. Son olarak VIP oda, havuza bakan oda, standart oda gibi konumsal farklılıklar ve ihtiyaca göre kablo TV gibi diğer özellikler de tanımlanabilir.'),
                        ('Altyapısal bilgiler', 'Ulaşım bağlantıları sayesinde nesnenin toplu taşımaya, hastaneye ve şehir merkezine uzaklığı gibi bilgiler girilebilir. Bu mesafelere yaya, otomobil, toplu taşıma gibi farklı parametrelerle bakılabilir.'),
                        ('Ölçümler', 'Farklı ölçüm birimlerinde (m², m³, adet, % vb.) ölçüm bilgilerinin girildiği yerdir. Kapasite özelliğine göre çalışır; kapasite aşımında sistem girişe izin vermez. Zaman ayarlıdır, tarihsel veriler takip edilebilir. Toplama özelliği bulunur: örneğin açık ve kapalı otopark araç sayılarının toplamı, toplam otopark araç sayısını eş zamanlı gösterir.')]),
                ('note', 'Tüm bu bilgi sistemi özellikleri, gayrimenkul arama taleplerinde arama kriteri olarak kullanılabilir.'),
            ]),
            dict(id='arazi-kullanim-yonetimi', title='Arazi Kullanım Yönetimi (Eklenti)',
                 desc='Parsel, parsel güncelleme, tapu sicili, rayiç değer bildirimi.', body=[
                ('note', 'Bu özellik bir eklentidir. Daha çok arsa yatırımı yapan şirketler için önerilir. Ek fonksiyonların aktifleştirilmesi gerektiğinden yeni kullanıcı rolleri gerekecektir.'),
                ('p', 'SAP RE-FX LUM (Land Use Management) modülü, SAP Esnek Gayrimenkul Yönetimi (Flexible Real Estate Management) içerisinde mülk, arsa ve arazi varlıklarının hukuki, coğrafi ve operasyonel olarak envanterinin tutulmasını ve yönetilmesini sağlayan alt bileşendir.'),
                ('h4', 'Temel özellikleri ve kapsamı'),
                ('dl', [('Parsel Yönetimi (Parcel of Land)', 'Arazilerin yasal/hukuki açıdan takibini, tapu/kadastro bilgilerini ve parsel güncellemelerini yönetir.'),
                        ('Ana Veri Nesneleri', 'Parsel, parsel güncellemeleri, tapu kaydı (land register), müşterek sorumluluk (joint liability) ve değerlendirme bildirimleri gibi nesneleri içerir.'),
                        ('Finansal ve Varlık Entegrasyonu', 'Duran Varlıklar (FI-AA), Muhasebe (FI), Maliyet Muhasebesi (CO) ve Bakım Onarım (PM) modülleriyle entegre çalışır.'),
                        ('Coğrafi Entegrasyon', 'ESRI ve coğrafi bilgi sistemleri (GIS) ile entegre edilerek parsel düzeyinde harita ve coğrafi görüntüleme desteği sunar.'),
                        ('Teminat ve Ortak Girişimler', 'Arazilerin ipotek ve teminat gösterilmesi, ortak girişimlerin (joint ventures) takibi ve arazi holding değişiklikleri (birleşme, bölünme vb.) gibi süreçleri destekler.')]),
            ]),
        ]),
        ('Kurumsal Yönetim', [
            dict(id='yerlesim-ve-ofis-yonetimi', title='Yerleşim ve Ofis Yönetimi',
                 desc='Genel müdürlük ve büroların tüm varlıklarını izleme ve masa detayına kadar çalışan yönetimi.', body=[
                ('ul', ['Genel müdürlük, ofis ve fabrikadaki gayrimenkul objeleri altında mimari görünüm ve oda tanımlama özelliği açılır.',
                        'Çalışma odasının müsaitlik kontrolü sağlanır. Her bir departman için ya da sıra düzeyinde işlem verileri tanımlanır.',
                        'Ofis özelliği çalışma odası olarak tayin edilmelidir.',
                        'Oda içinde çalışma masaları belirlenir.',
                        'Her bir çalışma masasını kullanan çalışanlar tanımlanır.',
                        'Çalışanlar, Business Partner olarak iş tanımı rolleri ile açılır.',
                        'Taşınma yaşanması durumunda “Taşınma planlaması” açılarak terminler belirlenir.']),
                ('dl', [('Maliyetlerin muhasebeleştirilmesi', 'Sistem, son kayıt edilen tarihten girilen bitiş tarihine kadar olan maliyetleri belirtilen sıklıkta ilgili maliyet merkezine tahakkuk ettirir.')]),
                ('p', 'Departman ve/veya çalışan bazında maliyetlerin takibi yapılabilir.'),
            ]),
            dict(id='oda-rezervasyonu', title='Oda Rezervasyonu ve Sürekli Tahsis',
                 desc='Rezervasyona ilişkin hizmetler, fiyatlandırma ve iptal masrafları.', body=[
                ('dl', [('Oda Rezervasyonları (Room Reservations)', 'Toplantı odaları, seminer salonları veya geçici çalışma alanları için kısa süreli/günlük rezervasyon yapılmasına olanak tanır. Rezervasyon takvimi üzerinden gün ve saat bilgisi görüntülenebilir; iptal edilen rezervasyonlar kırmızı renkle görünür. Doktor muayenehanesi gibi amaçlarla da kullanılabilir.')]),
                ('img', 'assets/images/refx-oda-rezervasyon-takvimi.png', 'SAP RE-FX oda rezervasyonu takvim ekranı: aylık görünümde rezervasyonlar', 712, 227),
                ('dl', [('Uzun Süreli Oturma Planlaması (Long-Term Seating Arrangements)', 'Daha çok oturum (residential) amacıyla kullanılır; otel rezervasyonu gibi uzun süreli konaklamalar için idealdir. Müsaitlik durumuna göre arayüz üzerinden takvim planlaması belirlenen parametrelere göre yapılır. Arama yapıldıktan sonra yalnızca uygun saatler, uygunluk kontrolüne göre ekranda belirir.')]),
                ('h4', 'Ücretlendirme'),
                ('dl', [('Rezervasyon fiyatı', 'Farklı odalar, kişiler ve dönemlere istinaden belirlenen kurallara göre hesaplanan fiyat üzerinden rezervasyon yaptıran kişiye kesilen fatura bedelidir.'),
                        ('İptal ücreti', 'Opsiyoneldir. Cayma ücreti faturalandırılır.'),
                        ('Fiyat listesi', 'Otel konaklamasındaki minibar kullanımı buna en güzel örnektir. Farklı ürünler için rezervasyon bedeli üzerine müşteriye kesilecek faturaya ek olarak tahakkuk ettirilir.'),
                        ('Maliyet Dağıtımı (Controlling – CO)', 'Rezervasyon ve uzun süreli kullanım maliyetlerini ilgili maliyet merkezlerine (Cost Center) veya iş ortaklarına (Business Partner) bağlar ve muhasebeleştirir.')]),
            ]),
            dict(id='muhatap-rolleri', title='Muhatap Rolleri',
                 desc='Her bir gayrimenkul nesnesi için sorumlu kişiler.', body=[
                ('p', 'SAP BP (Business Partner), SAP S/4HANA sistemlerinde müşteri, tedarikçi (satıcı) ve çalışan gibi tüm iş ortağı verilerini tek bir merkezi çatı altında toplayan ana veri modelidir. SAP RE-FX modülünde de farklı BP rolleri için her bir gayrimenkul objesinde kullanılmak üzere tanımlamalar yapılabilir.'),
                ('h4', 'Örnek roller'),
                ('chips', ['İnşaat müteahhidi', 'Mimar', 'Sorumlu', 'Mal sahibi', 'Noter', 'İş yazışması muhatabı', 'Başvuru sahibi',
                           'Miras yoluyla geçen üst hakkı sahibi', 'Kadastro dairesi', 'Yönetici', 'Teknik', 'Bina sorumlusu', 'AVM Müdürü',
                           'Teknik Yönetici', 'İşletme Yöneticisi', 'Pazarlama Yöneticisi', 'Tekniker', 'Teknisyen', 'Güvenlik Müdürü',
                           'Güvenlik Personeli', 'Temizlik Müdürü', 'Temizlik Personeli', 'Peyzaj Görevlisi', 'Kiralama Yöneticisi', 'Kefil',
                           'Kiracı (borçlu değil)', 'Teminat için banka', 'Mal sahibi (müşteri)', 'Sözleşmenin borçlu muhatabı', 'Kullanan',
                           'Satıcı hesabı olan yönetici', 'Müşteri hesabı olan ana kiracı', 'Destek veren müşteri', 'Kiralayan (satıcı hesabı)',
                           'Sözleşmenin alacaklı muhatabı', 'Ofis kiracıları', 'Uzman', 'Asıl ihale veren', 'Tapu sicil dairesi',
                           'Yasal mal sahibi', 'Yetkili (arsa hukukuna göre)', 'Yükümlü (arsa hukukuna göre)', 'Vergi dairesi']),
                ('note', 'Uygulama: Bir bağımsız bölümdeki herhangi bir ekipman (klima gibi) arızalandı. Buradaki müdür rolündeki muhatap, teknik yöneticiye haber verir ve bu bağımsız bölümden sorumlu tekniker gelip klimanın tamiratını üstlenir. Bu bağımsız bölümdeki tüm muhataplar, ilgili gayrimenkul objesinde iletişim bilgileri ile yer alır.'),
            ]),
        ]),
        ('Ticari İşlemler', [
            dict(id='olasi-musteri-yonetimi', title='Olası Müşteri Yönetimi',
                 desc='Arama ve geçmiş görüşmelerin yer aldığı, sözleşme oluşması öncesi iş süreci.', body=[
                ('dl', [('Arama Talebi', 'Potansiyel bir müşterinin bölge, fiyat ve karakteristik bilgilerine göre yaptığı arama talebidir. Yapılan görüşmelerin loglandığı yerdir.'),
                        ('Önerilen Nesne', 'Arama talebine istinaden belirlenmiş kriterlere göre, portföydeki uygunluk durumuna göre önerilen gayrimenkul nesneleridir.'),
                        ('Sözleşme Teklifi', 'Potansiyel müşteriye verilen tekliftir. Portföydeki nesne bilgilerine ait form ve sözleşmeye baz oluşturacak fiyatların çıktısının verildiği işlem verisidir.'),
                        ('Sözleşme', 'Arama özelliğini kullanmadan da açılabilen; yapılan anlaşmanın tüm şartlarının tutulduğu, ödeme koşullarının girildiği ve muhasebeleştirilen nakit akışı sağlayan işlem verisidir.')]),
                ('img', 'assets/images/refx-olasi-musteri-sureci.png', 'SAP RE-FX süreç şeması: arama talebi, önerilen nesne, sözleşme teklifi ve sözleşmeden kiralama, IFRS 16 ve yevmiye kayıtlarına uzanan akış ile FI-AA, FI-CO, CO-PS ve PM entegrasyonları', 974, 573),
            ]),
            dict(id='diger-bilesenlere-tayin', title='Diğer Bileşenlere Tayin',
                 desc='Sipariş, duran varlık, bakım-onarım entegrasyonu.', body=[
                ('p', "SAP'de birçok farklı modülde farklı ana veri yapıları vardır; örneğin teknik birimler, duran varlıklar, projeler ve CO siparişleri. Bu ana veriler arzu edilen gayrimenkul objesine tayin edilebilir."),
                ('p', 'Örneğin bir ekipman kiralama firması olarak, yaratılan ekipmanı kiraya vermek istediğinizde bu ekipmanı sözleşmede bilgi amaçlı olarak kullanabilirsiniz.'),
                ('p', 'Fonksiyonel konumlar atandığında, fonksiyonel konum için arıza raporları doğrudan gayrimenkul nesnesinden de oluşturulur.'),
                ('dl', [('Uygulama 1', 'SAP PM (Bakım Onarım) modülündeki bir fonksiyonel konum ana verisi, gayrimenkul objesi ile eşgüdümlü olarak çalıştırılabilir. Gayrimenkulle ilişkili fonksiyonel konumların her iki modülde de izlenmesi otomatik olarak sağlanır.'),
                        ('Uygulama 2', 'Bir fonksiyonel lokasyonda servis emri ile hem periyodik bakımları hem de tamir edilmesi gereken ekipmanları yönetmeyi mümkün kılar. Bu şekilde yalnızca gelirler değil, masraflar da gayrimenkul objesi bazında raporlanır.')]),
            ]),
            dict(id='ucuncu-taraf-islemleri', title='3. Taraf İşlemleri',
                 desc='Hem SAP sisteminde bulunan hem de bulunmayan şirket kodlarında ortaklık ya da yönetimsel yapılandırma.', body=[
                ('dl', [('Yönetici şirket kodu', 'Bir yönetim kurulu tarafından yönetilen ancak tapu sicillerinin ayrı ayrı kaydettirildiği bireysel bağımsız bölüm sahiplerinden oluşan şirketler için uygundur.'),
                        ('Kat mülkiyeti sahibi şirket kodu', 'Bu sözleşmeler, işletici ile kat mülkiyeti sahibi arasındaki ticari ilişkiyi tanımlar.'),
                        ('Yetki şirket kodu', 'Yöneticinin, üçüncü bir tarafın yetkilendirmesi (vekâletnamesi) ile yönetilen gayrimenkul nesnelerinin verilerini kaydettiği ve bunlarla çalıştığı şirket kodudur.')]),
                ('ul', ['Vekâlet tanımlanmasının ardından, vekâleti veren A şirket kodundaki ana veriler sayesinde vekâleti alan B şirket kodu, bu entegre ana veri şeması ile işleticisi olduğu yerde sözleşme yapma yetkisine sahip olur.',
                        'Aynı SAP sisteminde olan yapılarda şirketler arası denkleştirme belgesi ile dönemsel kayıtlar oluşturulur. Dönemsel kayıtlar eş zamanlı ve çift yönlü olarak hem gelir hem de gider muhasebesinde çalışır.',
                        'Ödeme bilgilerinin vekâleti veren şirket tarafından girilmesini mümkün kılar.',
                        'Bu yapılarda sözleşme türleri standart şirket kodundan farklı olarak modellenir.']),
            ]),
        ]),
    ],
    'feature': dict(kicker='Bilgi Merkezi', title='Teknik doğruluk, SAP RE-FX projelerinde lüks değil zorunluluktur.',
                    text='Yanlış ana veri modeli veya eksik muhasebe kurgusu, canlı kullanımda operasyonel yük ve raporlama hatası üretir.',
                    href='sss.html', link='Soruları inceleyin'),
}

KIRALAMA = {  # kiralama.html
    'groups': [
        ('Sözleşme', [
            dict(id='sozlesme-turleri', title='Sözleşme Türleri',
                 desc='Kiralama, kiraya verme, dahili kiralama, hizmet sözleşmeleri vb.', body=[
                ('p', 'Sözleşme türü, gayrimenkul sözleşmesini kontrol eden ana unsurdur ve sözleşmenin amacını belirtir. Kiraya verme ve kiralama sözleşmeleri en yaygın türlerdir.'),
                ('p', 'İç kullanım, üçüncü şahıs kullanımı veya genel muhasebe hesabı sözleşmesi (sözleşme kategorisi: iç sözleşme, dış sözleşme, genel muhasebe hesabı sözleşmesi) gibi farklı sözleşme türlerinin kullanımı da mümkündür. Ayrıca hizmet, satış ve alt kiralama sözleşmeleri de sistemde tanımlanabilir.'),
                ('ul', ['Yapılan tanımlamalara göre sözleşmelerde alan yönetimi farklılaştırılır, karşı taraf bilgisi sınırlandırılır.',
                        'Örneğin araç kiralama sözleşme türünde yalnızca taşıtların plakasını içeren verilerin bulunması, bağımsız bölümlerin yer alamaması gibi sınırlandırmalar yapılarak hataların önüne geçilir.',
                        'Gayrimenkul sözleşmesi tek başına bir obje olarak kullanılabildiği gibi, diğer bileşenlere tayin bölümünde belirtilen nesneler de sözleşmeye konu olabilir.']),
            ]),
            dict(id='depozito-anlasmalari', title='Depozito Anlaşmaları',
                 desc='Teminat sözleşmeleri: rehin, nakit, kefalet gibi farklı depozito türleri.', body=[
                ('ul', ['Depozito anlaşmaları ile teminat altına alınan sözleşmeleri raporlama imkânı sağlar.',
                        'Hem kiralama hem de kiraya verme sözleşme türlerinde bu özellik kullanılabilir.',
                        'Rehin, nakit, kefalet, ipotek gibi farklı depozito türleri modül üzerinde takip edilebilir.',
                        'Depozito tutarının “3 baz kira toplamı” gibi bir formülle otomatik hesaplanması sağlanabilir.',
                        'Yevmiye kayıtlarını yönetmenizi ya da kontrol amaçlı izlemenizi mümkün kılar.',
                        'Geri ödeme tarihlerini takip edip hatırlatma kuralları kullanılmasına olanak sağlar.',
                        'Süre aşımı olması ya da yaklaşması durumunda farklı renklerle raporlama imkânı sunar.',
                        'Ödenen depozito ile kalan bakiye kontrolünü takvim detayı ile raporlar.',
                        'Fiilen verilen ya da alınan teminatların girişinin yapılacağı alanlar kullanılabilir.']),
            ]),
            dict(id='fesih-ve-uzatim', title='Fesih ve Uzatım',
                 desc='Fesih kuralları ve uzatma kuralları.', body=[
                ('dl', [('Fesih', 'Son tarih ya da süreye göre önceden belirlenmiş bildirim zorunluluğuna istinaden simüle edilebilen, onay süreçleri bulunan ve fesih hakkına sahip tarafların belirlenebildiği sözleşme fesih haklarıdır.'),
                        ('Uzatım', 'Otomatik ya da seçerek yapılabilen, belirli bir tarih bildirimi ile periyotların simüle edilerek yapılabildiği uzatımlardır.')]),
            ]),
            dict(id='hatirlatma-dosyalama-ic-denetim', title='Hatırlatma Kuralları, Dosyalama, İç Denetim',
                 desc='Önemli tarihlerde hatırlatma özelliği, belge yönetimi, yapılan değişikliklerin loglanması.', body=[
                ('note', 'Bu özellikler yalnızca sözleşmeler için değil, tüm RE-FX ana verileri için de kullanılabilir.'),
                ('cols', [
                    ('Hatırlatma Kuralları', ['Genel hatırlatma', 'Koşul kontrolü', 'Sözleşme bitişi', 'Yenileme opsiyonu kontrolü',
                                              'Otomatik yenileme kontrolü', 'Fesih kontrolü',
                                              'Bir kez yapılabileceği gibi belirli sıklıklarla tekrarlanabilir.',
                                              'İş birimine ya da karşı tarafa e-posta gönderilmesi için dayanak sağlar.']),
                    ('Dosya Yönetimi', ['Birçok dosya tipinde dokümanlar sözleşmelere eklenebilir; örneğin:', 'Asıl sözleşme', 'Protokol', 'IBAN formu']),
                    ('İç Denetim', ['Sözleşmelere ilişkin tüm değişiklikler sistemde loglanır.', 'Tarih, saat ve kullanıcı bilgileri yer alır.',
                                    'Eski değer ve yeni değer gibi veriler tutulur.',
                                    'Hangi özelliğe ilişkin, hangi tarih aralığında, hangi kullanıcı tarafından ne değişiklik yapıldığı toplu raporlardan izlenebilir.',
                                    'Sözleşme yaratma, değiştirme, fesih, uzatım ve kira artışı gibi kritik durumlarda iş sürecine göre onay mekanizması kurmak mümkündür.']),
                ]),
            ]),
        ]),
        ('Koşullar', [
            dict(id='kosul-turleri', title='Koşul Türleri', desc='Asgari, temel, ciro, ofis, depo, araç kira gibi türler.', body=[]),
            dict(id='odeme-bilgileri', title='Ödeme Bilgileri', desc='Vergi, para birimi çeviri kuralı, yapılacak ödeme için gönderen ve karşı taraf banka bilgisi gibi bilgiler.', body=[]),
            dict(id='donem-yonetimi', title='Dönem Yönetimi', desc='Hesaplama yöntemi, vade tarihi, ödeme sıklığı ve KIST uyumlu dönem bilgisi.', body=[]),
        ]),
        ('Satış Hasılatı', [
            dict(id='bildirim-kurallari', title='Bildirim Kuralları', desc='Ciro verilerinin tutulması, farklı satış türlerine göre ayrıştırılması.', body=[]),
            dict(id='satis-hasilatina-dayali-kiralamalar', title='Satış Hasılatına Dayalı Kiralamalar', desc='Yüzde, baremli hesaplama, minimum kira mukayeseli, azami satış hasılatı gibi ciro hesaplaşma şartları.', body=[]),
            dict(id='karsi-taraf-yazismalari', title='Karşı Taraf Yazışmaları', desc='Ciro bildirimi, hesaplaşma sonuçları, eksik satış hasılatı bildirimi formları.', body=[]),
        ]),
        ('Yan Masraflar', [
            dict(id='ortak-alan-giderleri', title='Ortak Alan Giderleri', desc='Genel giderlerin m² bilgisine göre yansıtılması.', body=[]),
            dict(id='tuketimler', title='Tüketimler', desc='Elektrik, su ve atık su, iklimlendirme gibi sayaç verilerine göre yansıtma.', body=[]),
        ]),
        ('Artış ve Diğer Etkiler', [
            dict(id='artis-kurallari', title='Artış Kuralları', desc='Enflasyon endeks verilerine göre ya da serbest ve bileşik uygulanabilen kira artışlarının işlenmesi. (TTK 344 uyumlu çalışır.)', body=[]),
            dict(id='kira-disi-etki', title='Kira Dışı Etki', desc='Kiralama öncesi masraflar, erken fesih ve diğer ceza bedelleri ile stand ve reyon bedellerinin yer aldığı yükümlülük ve borçlar.', body=[]),
            dict(id='yansitma', title='Yansıtma', desc='Franchise sözleşmelerindeki yansıtmalar ile belirli parametrelere göre yansıtılan 3. parti (işletmeci gibi) bedeller.', body=[]),
            dict(id='vade-farki-faturasi', title='Vade Farkı Faturası', desc='Ödeme gecikmesi durumunda temerrüde düşen fark faturası.', body=[]),
        ]),
    ],
    'feature': dict(kicker='Sektörel Bakış', title='Her portföy aynı değildir; modelleme sektöre göre kurulmalıdır.',
                    text='GYO, perakende, holding ve tesis yönetimi yapılarında RE-FX kurgusu farklı öncelikler ister. Menü ve süreç tasarımı buna göre yapılmalıdır.',
                    href='sektorel.html', link='Sektörel kullanım alanlarını inceleyin'),
}

MUHASEBE = {  # muhasebe-raporlama.html
    'groups': [
        ('IFRS 16 & ASC 842', [
            dict(id='kira-siniflama', title='Kira Sınıflama ve Değerleme Türü', desc='Değerleme defterine göre soru setinin cevaplanarak “Lease Classification” yapılması; “Low Value”, “Short Term” gibi etkilerin belirlenmesi.', body=[]),
            dict(id='degisiklik-yonetimi', title='Değişiklik Yönetimi', desc='Değişiklik (amendment), ekleme (addendum), uzatım (renewal) ve fesih (termination) işlemlerinin otomatik işlenmesi.', body=[]),
            dict(id='diger-faktorler', title='Diğer Faktörler', desc='Teşvikler, ceza ödemeleri, plan dışı amortisman, başlangıç masrafları, kalıntı değer (residual value), ARO (Asset Retirement Obligation) ve zorunlu satın alma opsiyonu (bargain purchase option).', body=[]),
            dict(id='devir-alim', title='Devir Alım', desc='Mevcut IFRS 16 hesaplamalarınızdaki RoU (kullanım hakkı varlığı) ve Lease Liability (kira yükümlülüğü) kapanış bakiyelerinize bakılarak, devir alım tarihinden itibaren açılış bakiyesi olarak veri aktarımı yapılır.', body=[]),
        ]),
        ('Muhasebe', [
            dict(id='periyodik', title='Periyodik', desc='Kira giderlerinin muhasebeleştirilmesi.', body=[]),
            dict(id='donemsellestirme', title='Dönemselleştirme', desc='Gelir/gider tahakkukları, gelecek aylara ait gelirler/giderler, provizyon ayırma.', body=[]),
            dict(id='degerleme-kayitlari', title='Değerleme Kayıtları', desc='IFRS 16 değerleme kayıtları.', body=[]),
            dict(id='iade-surecleri-kayitlari', title='İade Süreçleri Kayıtları', desc='İade süreçlerine ilişkin muhasebe kayıtları.', body=[]),
        ]),
        ('Raporlama', [
            dict(id='kosul-ve-nakit-akisi', title='Koşul ve Nakit Akışı', desc='Kira koşulları, ödeme planları ve değerleme etkileri.', body=[]),
            dict(id='fi-co-entegrasyonu', title='FI/CO Entegrasyonu', desc='Muhasebe, masraf yeri, hesap tayini ve dönemsel kayıt akışları.', body=[]),
            dict(id='fon-yonetimi', title='Fon Yönetimi', desc='Nakit akış fon yönetimi ve bütçeleme.', body=[]),
        ]),
    ],
    'feature': dict(kicker='RECACUST', title='Kurumsal disiplin, butik danışmanlık çevikliği.',
                    text='Tek odaklı uzmanlık modeliyle SAP RE-FX projelerinde sade, ölçülebilir ve sürdürülebilir çözümler hedeflenir.',
                    href='hakkimizda.html', link='Kurumsal yaklaşımı inceleyin'),
}

PANELS = [  # (anahtar, menü adı, sayfa, veri)
    ('services', 'Hizmetlerimiz', 'cozumler.html', SERVICES),
    ('knowledge', 'Bilgi Merkezi', 'bilgi-merkezi.html', BILGI),
    ('leasing', 'Kiralama & Finans', 'kiralama.html', KIRALAMA),
    ('accounting', 'Muhasebe & Bütçe Raporlama', 'muhasebe-raporlama.html', MUHASEBE),
]

def groups_of(data):
    return [(data['kicker'], data['items'])] if 'items' in data else data['groups']

# ---------------------------------------------------------------------------
# HTML üreticiler
# ---------------------------------------------------------------------------

def render_blocks(blocks):
    out = []
    for b in blocks:
        k = b[0]
        if k == 'p': out.append(f'<p>{E(b[1])}</p>')
        elif k == 'h4': out.append(f'<h4>{E(b[1])}</h4>')
        elif k == 'note': out.append(f'<p class="doc-note">{E(b[1])}</p>')
        elif k == 'ul': out.append('<ul class="check-list-v5">' + ''.join(f'<li>{E(x)}</li>' for x in b[1]) + '</ul>')
        elif k == 'chips': out.append('<ul class="doc-chips">' + ''.join(f'<li>{E(x)}</li>' for x in b[1]) + '</ul>')
        elif k == 'dl': out.append('<dl class="doc-dl">' + ''.join(f'<div><dt>{E(t)}</dt><dd>{E(d)}</dd></div>' for t, d in b[1]) + '</dl>')
        elif k == 'img': out.append(f'<figure class="doc-figure"><img src="{b[1]}" alt="{E(b[2])}" width="{b[3]}" height="{b[4]}" loading="lazy"></figure>')
        elif k == 'table':
            out.append('<div class="doc-table-wrap"><table class="doc-table"><thead><tr>' + ''.join(f'<th scope="col">{E(h)}</th>' for h in b[1]) +
                       '</tr></thead><tbody>' + ''.join('<tr>' + ''.join(f'<td>{E(c)}</td>' for c in r) + '</tr>' for r in b[2]) + '</tbody></table></div>')
        elif k == 'cols':
            out.append('<div class="doc-cols">' + ''.join(f'<div><h4>{E(t)}</h4><ul>' + ''.join(f'<li>{E(x)}</li>' for x in items) + '</ul></div>' for t, items in b[1]) + '</div>')
        elif k == 'links': out.append('<p class="doc-links">' + ' '.join(f'<a href="{h}">{E(t)} →</a>' for h, t in b[1]) + '</p>')
    return '\n'.join(out)

def mega_panel(key, page, data):
    cols = []
    for kicker, items in groups_of(data):
        links = ''.join(f'<a href="{page}#{it["id"]}"><strong>{E(it["title"])}</strong><span>{E(it["desc"])}</span></a>' for it in items)
        cols.append(f'<div class="mega-col"><p class="mega-kicker">{E(kicker)}</p>{links}</div>')
    f = data['feature']
    n = len(cols)
    return (f'<div class="mega-panel" data-panel="{key}" aria-hidden="true"><div class="mega-inner mega-n{n}">' + ''.join(cols) +
            f'<div class="mega-feature"><p class="mega-kicker">{E(f["kicker"])}</p><h3>{E(f["title"])}</h3><p>{E(f["text"])}</p>'
            f'<a href="{f["href"]}">{E(f["link"])}</a></div></div></div>')

ACTIVE = {'index.html': 'index', 'cozumler.html': 'services', 'bilgi-merkezi.html': 'knowledge', 'sss.html': 'knowledge',
          'kiralama.html': 'leasing', 'sektorel.html': 'leasing', 'muhasebe-raporlama.html': 'accounting', 'ifrs16-asc842.html': 'accounting',
          'hakkimizda.html': 'corporate', 'notlar.html': 'corporate', 'iletisim.html': 'contact'}

def header_html(fname):
    act = ACTIVE.get(fname, '')
    def cls(k): return 'nav-item active' if act == k else 'nav-item'
    nav = (f'<a class="{cls("index")}" href="index.html" data-mega="0">Ana Sayfa</a>' +
           ''.join(f'<a class="{cls(k)}" href="{p}" data-mega="{k}">{E(n)}</a>' for k, n, p, _ in PANELS) +
           f'<a class="{cls("corporate")}" href="hakkimizda.html" data-mega="0">Kurumsal</a>'
           f'<a class="{cls("contact")}" href="iletisim.html" data-mega="0">İletişim</a>')
    panels = '\n'.join(mega_panel(k, p, d) for k, n, p, d in PANELS)
    drawer = ''
    for i, (k, n, p, d) in enumerate(PANELS):
        links = ''.join(f'<a href="{p}#{it["id"]}">{E(it["title"])}</a>' for _, items in groups_of(d) for it in items)
        drawer += f'<details{" open" if i == 0 else ""}><summary>{E(n)}</summary><a class="drawer-all" href="{p}">Tümünü görüntüle</a>{links}</details>'
    drawer += '<details><summary>Kurumsal</summary><a href="hakkimizda.html">Hakkımızda</a><a href="sektorel.html">Sektörel</a><a href="sss.html">Sık Sorulan Sorular</a><a href="iletisim.html">İletişim</a></details>'
    return f'''<div class="top-strip">
    <div class="container top-strip-inner">
      <span class="top-legal"><a href="cerez-politikasi.html">Çerez Politikası</a><a href="kvkk-aydinlatma-metni.html">KVKK Metni</a><a href="gizlilik-politikasi.html">Gizlilik</a></span>
      <div class="top-contact"><a href="{PHONE_TEL}">{ICON['phone']}{PHONE_DISPLAY}</a><a href="mailto:{EMAIL}">{ICON['mail']}{EMAIL}</a><a href="{WA}" target="_blank" rel="noopener">{ICON['wa']}WhatsApp ile ulaşın</a></div>
    </div>
  </div>
  <header class="main-header" id="mainHeader">
    <div class="container header-row">
      <a href="index.html" class="brand" aria-label="RECACUST Ana Sayfa"><img class="brand-logo" src="assets/images/logo-recacust.webp" alt="RECACUST SAP RE-FX Danışmanlığı logosu"><span class="brand-copy"><strong>RECACUST</strong><small>SAP RE-FX Danışmanlığı</small></span></a>
      <nav class="desktop-nav" id="desktopNav">{nav}</nav>
      <div class="header-actions"><a href="iletisim.html" class="consult-btn">Ön Görüşme</a><button id="mobileOpen" class="mobile-open" aria-label="Menüyü aç">☰</button></div>
    </div>
{panels}
  </header>
  <div id="mobileDrawer" class="mobile-drawer" aria-hidden="true">
    <div class="drawer-head"><a href="index.html" class="brand"><img class="brand-logo" src="assets/images/logo-recacust.webp" alt="RECACUST SAP RE-FX Danışmanlığı logosu"><span class="brand-copy"><strong>RECACUST</strong><small>SAP RE-FX Danışmanlığı</small></span></a><button id="mobileClose" aria-label="Menüyü kapat">×</button></div>
    <div class="drawer-body">
      {drawer}
      <a class="drawer-cta" href="iletisim.html">Ön Görüşme</a>
      <p class="drawer-contact"><a href="{PHONE_TEL}">{ICON['phone']}{PHONE_DISPLAY}</a><a href="mailto:{EMAIL}">{ICON['mail']}{EMAIL}</a><a href="{WA}" target="_blank" rel="noopener">{ICON['wa']}WhatsApp</a></p>
    </div>
  </div>
  <div id="drawerBackdrop" class="drawer-backdrop"></div>'''

FOOTER_BRAND = 'SAP Esnek Gayrimenkul Modülü üzerinde proje, dönüşüm, destek, eğitim uzman danışmanlık hizmetleri.'

def footer_cols():
    return (f'<div><h3>Hizmetler</h3><a href="cozumler.html">Hizmetlerimiz</a><a href="bilgi-merkezi.html">Bilgi Merkezi</a>'
            f'<a href="kiralama.html">Kiralama &amp; Finans</a><a href="muhasebe-raporlama.html">Muhasebe &amp; Bütçe Raporlama</a><a href="sss.html">SSS</a></div>'
            f'<div><h3>Kurumsal</h3><a href="hakkimizda.html">Hakkımızda</a><a href="sektorel.html">Sektörel</a><a href="ifrs16-asc842.html">IFRS16 / ASC842</a><a href="iletisim.html">İletişim</a><a href="notlar.html">Notlar</a></div>'
            f'<div><h3>İletişim</h3><a href="{PHONE_TEL}">{ICON["phone"]}{PHONE_DISPLAY}</a><a href="{WA}" target="_blank" rel="noopener">{ICON["wa"]}WhatsApp ile ulaşın</a>'
            f'<a href="mailto:{EMAIL}">{ICON["mail"]}{EMAIL}</a><small>Pazartesi-Cuma: 09:00 - 17:00</small></div>')

# ---------------------------------------------------------------------------
# İçerik sayfaları
# ---------------------------------------------------------------------------

def doc_main(crumb, eyebrow, h1, lead, data, intro=''):
    groups = groups_of(data)
    toc = ''.join(f'<p class="toc-group">{E(g)}</p>' + ''.join(f'<a href="#{it["id"]}">{E(it["title"])}</a>' for it in items) for g, items in groups)
    body = ''
    for g, items in groups:
        body += f'<section class="doc-group"><p class="eyebrow-v5">{E(g)}</p>'
        for it in items:
            inner = render_blocks(it['body'])
            short = ' doc-item-short' if not it['body'] else ''
            body += (f'<article class="doc-item{short}" id="{it["id"]}"><h2>{E(it["title"])}</h2><p class="doc-lead">{E(it["desc"])}</p>{inner}</article>')
        body += '</section>'
    return (f'<section class="subhero"><div class="container"><p class="breadcrumb-v5">{crumb}</p><span class="eyebrow-v5">{E(eyebrow)}</span>'
            f'<h1>{E(h1)}</h1><p>{E(lead)}</p></div></section>'
            f'<section class="container page-grid doc-grid"><div class="doc-body">{intro}{body}</div>'
            f'<aside class="side-rail doc-toc" aria-label="İçindekiler"><h3>İçindekiler</h3>{toc}<a class="toc-cta" href="iletisim.html">Ön Görüşme planlayın →</a></aside></section>')

def make_page(fname, title, desc, main):
    tpl = (ROOT / 'sektorel.html').read_text()
    head = tpl[:tpl.index('<body')]
    head = re.sub(r'<title>.*?</title>', f'<title>{E(title)}</title>', head)
    head = re.sub(r'(<meta name="description" content=")[^"]*', lambda m: m.group(1) + E(desc), head)
    head = re.sub(r'(<meta property="og:description" content=")[^"]*', lambda m: m.group(1) + E(desc), head)
    head = re.sub(r'(<meta name="twitter:description" content=")[^"]*', lambda m: m.group(1) + E(desc), head)
    head = re.sub(r'(<meta property="og:title" content=")[^"]*', lambda m: m.group(1) + E(title), head)
    head = re.sub(r'(<meta name="twitter:title" content=")[^"]*', lambda m: m.group(1) + E(title), head)
    head = head.replace('sektorel.html', fname)
    rest = tpl[tpl.index('<body'):]
    rest = re.sub(r'<main>.*?</main>', '<main>\n' + main + '\n</main>', rest, flags=re.S)
    (ROOT / fname).write_text(head + rest)

def replace_main(fname, main):
    p = ROOT / fname
    t = p.read_text()
    t = re.sub(r'<main>.*?</main>', lambda m: '<main>\n' + main + '\n</main>', t, flags=re.S)
    p.write_text(t)

def set_meta(fname, title, desc):
    p = ROOT / fname
    t = p.read_text()
    t = re.sub(r'<title>.*?</title>', f'<title>{E(title)}</title>', t)
    for pat in [r'(<meta name="description" content=")[^"]*', r'(<meta property="og:description" content=")[^"]*', r'(<meta name="twitter:description" content=")[^"]*']:
        t = re.sub(pat, lambda m: m.group(1) + E(desc), t)
    for pat in [r'(<meta property="og:title" content=")[^"]*', r'(<meta name="twitter:title" content=")[^"]*']:
        t = re.sub(pat, lambda m: m.group(1) + E(title), t)
    p.write_text(t)

# ---------------------------------------------------------------------------
# Excel ESKİ → YENİ metin değişiklikleri
# ---------------------------------------------------------------------------

TEXT = {
    'index.html': [
        ('<span class="eyebrow-v5">SAP RE-FX Danışmanlığı</span>\n      <h1>SAP gayrimenkul yazılım çözümleri</h1>',
         '<span class="eyebrow-v5">SAP RE-FX Yazılım ve Bilişim Danışmanlığı</span>\n      <h1>SAP Gayrimenkul Yazılım ve Bilişim Danışmanlığı</h1>'),
        ('<figcaption>RE-FX süreçleri için kurumsal danışmanlık yaklaşımı</figcaption>', '<figcaption>SAP RE-FX uzman danışmanlık hizmetleri</figcaption>'),
        ('<span>Sıfırdan Proje</span>', '<span>Proje</span>'),
        ('<h2>SAP RE-FX odağıyla gayrimenkul, kiralama ve muhasebe süreçleri</h2>', '<h2>SAP RE-FX modülü ile gayrimenkul, kiralama ve muhasebe süreçleri</h2>'),
        ('tek merkezden yönetmek isteyen kurumlar için güçlü bir standart sunar.', 'tek merkezden yönetmek isteyen kurumlar için güçlü bir ERP modülüdür.'),
        ('<p>Ana veri, sözleşme, kira koşulu, değerleme ve FI/CO entegrasyonu farklı ekipler arasında kopuk ilerleyebilir.</p>',
         '<p>Ana veri, işlem verisi, kiralama standartları takip ve raporlama zorluğu.</p>'),
        ('<span class="eyebrow-v5">IFRS 16 & SAP RE-FX</span>', '<span class="eyebrow-v5">SAP RE-FX</span>'),
        ('SAP RE-FX özellikle gerçek gayrimenkul kiralamaları, bina/alan yönetimi, sözleşme yönetimi, kullanım hakkı varlığı ve kira yükümlülüğü hesaplamaları gibi alanlarda kullanılır. SAP FI, FI-AA ve CO entegrasyonu sayesinde finans ve gayrimenkul süreçlerini aynı sistemsel akışta buluşturur.',
         'SAP RE-FX özellikle gerçek gayrimenkul kiralamaları, bina/alan yönetimi, sözleşme yönetimi, ticari ve kurumsal faaliyet alanlarında kullanılır. SAP FI, AA ve CO entegrasyonu sayesinde finans ve gayrimenkul süreçlerini aynı belge akışında buluşturur.'),
        ('<ul class="check-list-v5"><li>Sözleşmelerin kiralama olarak sınıflandırılması</li><li>Kullanım hakkı varlığı ve kira yükümlülüğü hesaplamaları</li><li>FI, FI-AA ve CO entegrasyonu</li><li>Sözleşme değişiklikleri, uzatım ve fesih süreçleri</li><li>IFRS16 uyumlu raporlama</li></ul>',
         '<ul class="check-list-v5"><li>Gayrimenkul ana verilerinin işletme yapısına göre sınıflandırılması ve tek merkezden izlenmesi</li><li>Olası müşteri yönetimi; arama talebi, teklif nesnesi, sözleşme teklifi</li><li>Sözleşmelerin kiralama türüne göre sınıflandırılması</li><li>Oda rezervasyonu ve tahsisi</li><li>Sözleşme değişiklikleri, kira artışları, uzatım ve fesih süreçleri</li><li>IFRS16 / ASC 842 uluslararası muhasebe standartları ve raporlaması</li></ul>'),
        ('<aside class="side-rail"><h3>İçindekiler</h3><a href="cozumler.html">SAP RE-FX Çözümleri</a><a href="kiralama.html">Kiralama Süreçleri</a><a href="ifrs16-asc842.html">IFRS16 / ASC842</a><a href="sss.html">SSS</a></aside>',
         '<aside class="side-rail"><h3>İçindekiler</h3><a href="cozumler.html">Hizmetlerimiz</a><a href="bilgi-merkezi.html">Bilgi Merkezi</a><a href="kiralama.html">Kiralama &amp; Finans</a><a href="muhasebe-raporlama.html">Muhasebe &amp; Bütçe Raporlama</a><a href="sss.html">SSS</a></aside>'),
    ],
    'cozumler.html': [
        ('Mevcut SAP yapısı, şirket kodu, gayrimenkul ana verileri, sözleşme tipleri ve entegrasyon gereksinimleri analiz edilir.',
         'Mevcut şirket yapısı, esnek gayrimenkul modülü üzerinde yapılandırılması ve entegrasyon gereksinimleri analiz edilir.'),
        ('RE-FX standartlarıyla uyumlu ana veri, sözleşme, koşul, nakit akışı ve muhasebe senaryoları hazırlanır.',
         'RE-FX standartlarıyla uyumlu ana veri, sözleşme, koşul ve muhasebe dizaynı hazırlanır.'),
        ('<h2>RE-FX projelerinde sadece ekran uyarlaması değil, süreç mimarisi kurulur</h2>',
         '<h2>RE-FX projelerinde devamlılık esas alınarak sistem mimarisi kurulur</h2>'),
    ],
    'sektorel.html': [
        ('<div>Kiraya verme işlemi yapan şirketler</div>', '<div>Ekipman kiralama şirketleri</div>'),
        ('<span class="eyebrow-v5">Sector expertise</span>', '<span class="eyebrow-v5">Sektörel uzmanlık</span>'),
    ],
    'iletisim.html': [
        ('<p>SAP RE-FX danışmanlığı, kiralama süreçleri veya IFRS16 ihtiyaçlarınız için talebinizi iletin; uzman değerlendirmesi için en kısa sürede sizinle iletişime geçelim.</p>',
         '<p>SAP RE-FX danışmanlığı ihtiyaçlarınız için talebinizi iletin; uzman değerlendirmesi için en kısa sürede sizinle iletişime geçelim.</p>'),
        ('''<option>SAP RE-FX Danışmanlığı</option>
            <option>IFRS16 / ASC842</option>
            <option>Kiralama ve Sözleşme Süreçleri</option>
            <option>Süreç Analizi ve Yol Haritası</option>
            <option>Diğer</option>''',
         '''<option>Sıfırdan Proje</option>
            <option>Dönüşüm Projeleri</option>
            <option>Destek</option>
            <option>Rehberlik Hizmetleri</option>
            <option>Eğitim Hizmetleri</option>'''),
        ('<a href="tel:+902163702478">Tel: 0216 370 24 78</a>\n      <a href="mailto:info@recacust.com">info@recacust.com</a>\n      <a href="https://wa.me/905418638613?text=Merhaba%2C%20SAP%20RE-FX%20dan%C4%B1%C5%9Fmanl%C4%B1%C4%9F%C4%B1%20hakk%C4%B1nda%20bilgi%20almak%20istiyorum.">WhatsApp ile yazın</a>',
         f'<a class="icon-link" href="{PHONE_TEL}">{ICON["phone"]}{PHONE_DISPLAY}</a>\n      <a class="icon-link" href="mailto:{EMAIL}">{ICON["mail"]}{EMAIL}</a>\n      <a class="icon-link" href="{WA}" target="_blank" rel="noopener">{ICON["wa"]}WhatsApp ile yazın</a>'),
    ],
    'hakkimizda.html': [
        ('SAP sisteminde RE-FX uyarlamalarının yer aldığı işlem kodu RECACUST’tur.', 'SAP sisteminde RE-FX uyarlamalarının bulunduğu işlem kodu RECACUST’tur.'),
    ],
    'kiralama.html': [],
}

GLOBAL = [  # tüm sayfalarda
    ('<span>Ücretsiz ön görüşme</span>', '<span>Ön Görüşme</span>'),
    ('>Ücretsiz ön görüşme planlayın<', '>Ön görüşme planlayın<'),
    ('>Ücretsiz ön görüşme<', '>Ön Görüşme<'),
    ('0216 370 24 78', PHONE_DISPLAY),
]

def apply_text(fname, t):
    for old, new in TEXT.get(fname, []):
        if old in t: t = t.replace(old, new)
        elif new not in t: print(f'  ! {fname}: bulunamadı → {old[:70]}…')
    for old, new in GLOBAL: t = t.replace(old, new)
    return t

# ---------------------------------------------------------------------------
# SSS (Excel "FAQ" sayfası)
# ---------------------------------------------------------------------------

FAQ_ANA = [
    ('Gayrimenkullerin yönetimi nasıl yapılır?', 'İki ayrı ana veri görünümü vardır. Kullanım görünümünde, şirket kodunun altındaki gayrimenkul yapılandırmasının bir ağaç gibi izlenmesi sağlanır. Mimari görünüm ise kullanım görünümü ile aynı düzeyde ya da alt ve üst düzeyinde daha çok ana verinin takip edilebilmesine yarar.'),
    ('Ana verilerde nelerin takibi yapılır?', ['Merkezi noktalara olan uzaklıkların yer aldığı altyapısal karakteristikler.', 'Araç gereçlerin sayısal takibinin yapıldığı ve kullanılan malzemelere ilişkin bilgilerin yer aldığı donanımsal karakteristikler.', 'Yüzeysel ve sayısal takibi yapılan ölçüm bilgileri.', 'Her bir özelleştirilmiş gayrimenkul nesnesi için spesifik genel ve ek veriler.']),
    ('AutoCAD ile ana veriler eşlenebilir mi?', 'Mimari ve kullanım görünümlerine AutoCAD entegrasyonu yapılması mümkündür.'),
    ('Arsa ve arazi aynı ana veri şemasında mı takip edilir?', 'Arsaların alanları modül yapısında farklılık göstermektedir. Belediyeye ait bilgiler, emlak vergisi verileri, değer verileri, üst hakkı ve kullanım verileri gibi birçok veri izlenebilir.'),
    ('Arsa ve diğer ana verilerimizin haritadan görüntülenmesi mümkün mü?', 'Haritalama eklentisinin etkinleştirilmesi ile mümkündür.'),
    ('Ana verilerde revize ve değişiklik yapılabilir mi?', 'Yapılacak tadilat işlemlerine istinaden tarih duyarlı planlama yapıp yenilenmiş şeklini bu tarihten itibaren kullanabilirsiniz.'),
    ('Öğrenci yurdu gibi vergi teşviki bulunan ana verilerin takibi nasıl yapılır?', 'Bu ana veriler, yukarıda bahsedilen mimari ve kullanım ana verisi haricinde Esnek Gayrimenkul Modülü altında farklı ana verilerde takip edilebilir.'),
]
FAQ_DIGER = [
    ('Ölçüm türlerini ana verilerde tutabilir miyiz?', 'Metrekare, sayı, metreküp, yüzde, kişi gibi ölçümler tutulabilir ve üst nesneye güncellenebilir.'),
    ('SAP PM entegrasyonu var mı?', 'Evet, bakım-onarım süreçleri SAP PM ile entegre edilerek ana veri bazında izlenebilir.'),
    ('Toplantı odası rezervasyonu yapılabilir mi?', 'Uygunluk kontrolü, saat görüntüleme, oda tutma ve e-posta bilgilendirmesi yapılabilir.'),
    ('Yazışma ve form gönderimi yönetilebilir mi?', 'Ana veri ve işlem verilerine özel form yapıları düzenlenip e-posta ile gönderilebilir.'),
]

def faq_html(items):
    out = ''
    for q, a in items:
        ans = f'<p>{E(a)}</p>' if isinstance(a, str) else '<ul class="check-list-v5 faq-list">' + ''.join(f'<li>{E(x)}</li>' for x in a) + '</ul>'
        out += f'<details><summary>{E(q)}</summary>{ans}</details>'
    return out

def faq_ld(items):
    import json
    ents = [{'@type': 'Question', 'name': q, 'acceptedAnswer': {'@type': 'Answer', 'text': a if isinstance(a, str) else ' '.join(a)}} for q, a in items]
    return '<script type="application/ld+json">' + json.dumps({'@context': 'https://schema.org', '@type': 'FAQPage', 'mainEntity': ents}, ensure_ascii=False) + '</script>'

# ---------------------------------------------------------------------------

def main():
    # 1) Yeni / yeniden kurulan içerik sayfaları
    make_page('bilgi-merkezi.html', 'SAP RE-FX Bilgi Merkezi | Ana Veri, Kurumsal ve Ticari Yönetim',
              'SAP RE-FX kullanım ve mimari görünüm, bilgi sistemi, arazi kullanım yönetimi, ofis ve oda rezervasyonu, muhatap rolleri, olası müşteri yönetimi ve 3. taraf işlemleri.',
              doc_main('Ana Sayfa » Bilgi Merkezi', 'Bilgi Merkezi', 'SAP RE-FX Bilgi Merkezi',
                       'Ana veri yönetiminden kurumsal ve ticari işlemlere kadar SAP Esnek Gayrimenkul modülünün temel yapı taşları.', BILGI,
                       '<p class="doc-intro-links"><a href="sss.html">Sık sorulan sorular →</a></p>'))
    make_page('muhasebe-raporlama.html', 'SAP RE-FX Muhasebe ve Bütçe Raporlama | IFRS 16 & ASC 842',
              'SAP RE-FX ile IFRS 16 / ASC 842 kira sınıflama ve değerleme, değişiklik yönetimi, devir alım, dönemselleştirme, değerleme kayıtları, FI/CO entegrasyonu ve fon yönetimi.',
              doc_main('Ana Sayfa » Muhasebe &amp; Bütçe Raporlama', 'Muhasebe & Bütçe Raporlama', 'Muhasebe ve bütçe raporlama',
                       'IFRS 16 ve ASC 842 uyumundan dönemsel muhasebe kayıtlarına, koşul ve nakit akışından fon yönetimine kadar finansal süreçler.', MUHASEBE,
                       '<p class="doc-intro-links"><a href="ifrs16-asc842.html">IFRS16 / ASC842 danışmanlığı hakkında →</a></p>'))
    replace_main('kiralama.html', doc_main('Ana Sayfa » Kiralama &amp; Finans', 'Kiralama & Finans', 'Kiralama ve finans süreçleri',
                 'Sözleşme türlerinden koşullara, satış hasılatına dayalı kiralamalardan yan masraflara ve kira artışlarına kadar RE-FX kiralama süreçleri.', KIRALAMA))
    set_meta('kiralama.html', 'SAP RE-FX Kiralama ve Finans | Sözleşme, Koşullar, Satış Hasılatı',
             'SAP RE-FX sözleşme türleri, depozito anlaşmaları, fesih ve uzatım, hatırlatma kuralları, koşul türleri, satış hasılatına dayalı kiralamalar, yan masraflar ve kira artışları.')

    # Hizmetlerimiz: mevcut içeriğin altına 3 hizmet bölümü
    p = ROOT / 'cozumler.html'; t = p.read_text()
    t = re.sub(r'\s*<section class="container page-grid doc-grid">.*?</section>(?=\s*</main>)', '', t, flags=re.S)
    svc = ''.join(f'<article class="doc-item" id="{it["id"]}"><h2>{E(it["title"])}</h2><p class="doc-lead">{E(it["desc"])}</p>{render_blocks(it["body"])}</article>' for it in SERVICES['items'])
    toc = ''.join(f'<a href="#{it["id"]}">{E(it["title"])}</a>' for it in SERVICES['items'])
    t = t.replace('</main>', f'<section class="container page-grid doc-grid"><div class="doc-body"><section class="doc-group"><p class="eyebrow-v5">Hizmetlerimiz</p>{svc}</section></div><aside class="side-rail doc-toc" aria-label="İçindekiler"><h3>Hizmetlerimiz</h3>{toc}<a class="toc-cta" href="iletisim.html">Ön Görüşme planlayın →</a></aside></section>\n</main>', 1)
    p.write_text(t)

    # SSS
    p = ROOT / 'sss.html'; t = p.read_text()
    faq = (f'<section class="container article-panel"><h2>Ana veri yapısı hakkında sorular ve cevapları</h2><div class="faq-v5">{faq_html(FAQ_ANA)}</div>'
           f'<h2 class="faq-h2">Diğer sorular</h2><div class="faq-v5">{faq_html(FAQ_DIGER)}</div></section>')
    t = re.sub(r'<section class="container article-panel">.*?</section>(?=\s*</main>)', lambda m: faq, t, flags=re.S)
    t = t.replace('<span class="eyebrow-v5">Questions & answers</span>', '<span class="eyebrow-v5">Sorular &amp; Cevaplar</span>')
    t = re.sub(r'<script type="application/ld\+json">\{"@context": "https://schema.org", "@type": "FAQPage".*?</script>', '', t, flags=re.S)
    t = t.replace('</head>', faq_ld(FAQ_ANA + FAQ_DIGER) + '</head>', 1)
    p.write_text(t)

    # 2) Tüm sayfalar: üst menü, alt bilgi, metinler
    for p in sorted(ROOT.glob('*.html')):
        t = p.read_text()
        if '<div class="top-strip">' not in t: continue
        s = t.index('<div class="top-strip">')
        mark = '<div id="drawerBackdrop" class="drawer-backdrop"></div>'
        e = t.index(mark) + len(mark) if mark in t else t.index('</header>') + len('</header>')
        t = t[:s] + header_html(p.name) + t[e:]
        t = re.sub(r'(<div class="footer-brand">.*?<p>).*?(</p></div>)', lambda m: m.group(1) + FOOTER_BRAND + m.group(2), t, flags=re.S)
        t = re.sub(r'(</div>\s*)<div><h3>Hizmetler</h3>.*?<small>Pazartesi-Cuma: 09:00 - 17:00</small></div>', lambda m: m.group(1) + footer_cols(), t, flags=re.S)
        t = apply_text(p.name, t)
        p.write_text(t)
        print('güncellendi', p.name)

    # 3) Site haritası
    sm = ROOT / 'sitemap.xml'; s = sm.read_text()
    for f in ['bilgi-merkezi.html', 'muhasebe-raporlama.html']:
        if f not in s:
            s = s.replace('</urlset>', f'  <url><loc>https://www.recacust.com/{f}</loc><lastmod>2026-11-01</lastmod></url>\n</urlset>')
    sm.write_text(s)

if __name__ == '__main__':
    main()
