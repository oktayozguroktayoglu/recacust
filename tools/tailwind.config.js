// RECACUST Tailwind yapılandırması (eski CDN içindeki tailwind.config ile aynı)
// Derleme: tools/README.md
module.exports = {
  content: ['./*.html', './assets/js/*.js'],
  theme: { extend: { colors: { brand: { purple: '#b369fe', black: '#040814', violet: '#8b35ff', dark: '#0b1020', ice: '#f7f2ff' } } } },
}
