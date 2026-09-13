<?php
// RECACUST PHP Mail Endpoint
// Yayına almadan önce $to ve $from alanlarını kendi hosting/mail yapınıza göre kontrol edin.

if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
    header('Location: iletisim.html');
    exit;
}

function clean_input($value) {
    $value = trim((string)$value);
    $value = str_replace(["\r", "\n"], ' ', $value);
    return htmlspecialchars($value, ENT_QUOTES, 'UTF-8');
}

$name = clean_input($_POST['name'] ?? '');
$company = clean_input($_POST['company'] ?? '');
$email = trim((string)($_POST['email'] ?? ''));
$phone = clean_input($_POST['phone'] ?? '');
$topic = clean_input($_POST['topic'] ?? '');
$message = trim((string)($_POST['message'] ?? ''));
$honeypot = trim((string)($_POST['website'] ?? ''));

// Basit spam tuzağı. Botlar bu gizli alanı doldurursa form gönderilmiş gibi davranılır.
if ($honeypot !== '') {
    header('Location: tesekkurler.html');
    exit;
}

if ($name === '' || $message === '' || !filter_var($email, FILTER_VALIDATE_EMAIL)) {
    http_response_code(422);
    echo 'Eksik veya hatalı bilgi girdiniz. Lütfen forma dönüp gerekli alanları kontrol edin.';
    exit;
}

$to = 'info@recacust.com';
$subject = 'RECACUST Web Sitesi İletişim Formu';

$body = "Yeni iletişim formu mesajı:\n\n";
$body .= "Ad Soyad: {$name}\n";
$body .= "Şirket: {$company}\n";
$body .= "E-posta: {$email}\n";
$body .= "Telefon: {$phone}\n";
$body .= "Talep Konusu: {$topic}\n\n";
$body .= "Mesaj:\n{$message}\n\n";
$body .= "Kaynak: recacust.com iletişim formu\n";
$body .= "Tarih: " . date('Y-m-d H:i:s') . "\n";

// From adresi mümkünse hosting üzerinde tanımlı gerçek bir domain maili olmalı.
$from = 'no-reply@recacust.com';
$headers = [];
$headers[] = 'MIME-Version: 1.0';
$headers[] = 'Content-Type: text/plain; charset=UTF-8';
$headers[] = 'From: RECACUST Web <' . $from . '>';
$headers[] = 'Reply-To: ' . $name . ' <' . $email . '>';
$headers[] = 'X-Mailer: PHP/' . phpversion();

if (mail($to, '=?UTF-8?B?' . base64_encode($subject) . '?=', $body, implode("\r\n", $headers))) {
    header('Location: tesekkurler.html');
    exit;
}

http_response_code(500);
echo 'Mesaj gönderilemedi. Lütfen daha sonra tekrar deneyin veya WhatsApp üzerinden iletişime geçin.';
