# Classroom Sync

Синхронізація журналів nz.ua з Google Classroom.

## Налаштування

1. Створіть проєкт у [Google Cloud Console](https://console.cloud.google.com/)
2. Увімкніть **Google Classroom API** та **Admin SDK API**
3. Створіть OAuth 2.0 credentials (тип: Web application)
4. Додайте Authorized redirect URI: `http://localhost:5173/callback`
5. Скопіюйте `.env.example` у `.env` та вставте Client ID та Client Secret

## Запуск

```bash
npm install
npm run install:browser
npm run dev
```

Якщо встановлено Chrome в Program Files, `install:browser` можна пропустити.

Відкрийте http://localhost:5173

## Розгортання на Fly.io

1. Додайте в Google Cloud Console redirect URI: `https://nz-to-classroom.fly.dev/callback`
2. Встановіть секрети:
   ```bash
   flyctl secrets set GOOGLE_REDIRECT_URI=https://nz-to-classroom.fly.dev/callback
   flyctl secrets set GOOGLE_CLIENT_ID=ваш_client_id
   flyctl secrets set GOOGLE_CLIENT_SECRET=ваш_client_secret
   ```
3. Розгорніть: `flyctl deploy`

## Використання

1. **Увійти через Google** — обліковий запис з доступом до Classroom та Admin (адміністратор домену)
2. **Авторизація nz.ua** — логін та пароль від порталу
3. **Завантажити журнали** — парсинг таблиці journal-choose з nz.ua
4. **Вчителі з Google Admin** — завантаження з організаційного підрозділу «Вчителі» (admin.google.com)
5. **Запустити синхронізацію** — створення курсів, співставлення вчителів за іменем, додавання за email

## Усунення неполадок

**«Не знайдено журнали»** — перевірте `debug-nz-page.html` (зберігається при помилці):
- Cloudflare — додайте `HEADLESS=false` в `.env` і спробуйте знову (відкриється вікно Chrome)
- Stealth — використовується puppeteer-extra-plugin-stealth для обходу захисту
# CLASSROOM-SYNC
