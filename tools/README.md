# Tailwind CSS derleme

Site, Tailwind'i CDN yerine derlenmiş `assets/css/tailwind.css` dosyasından kullanır.
HTML'e yeni bir Tailwind sınıfı eklediğinizde CSS'i yeniden derleyin:

```bash
npx tailwindcss@3 -c tools/tailwind.config.js -i tools/tailwind.input.css -o assets/css/tailwind.css --minify
```

Komut repo kök klasöründen çalıştırılmalıdır.
