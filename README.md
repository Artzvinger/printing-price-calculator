# 🖨️ Калькулятор типографии

Веб-приложение для мгновенного расчёта стоимости полиграфической продукции.  
Интегрирован с **Google Sheets** через **Google Apps Script** — все формулы и цены хранятся и обновляются в таблице.
---
[Копия интегрированной таблицы на которую приходят запросы](https://docs.google.com/spreadsheets/d/1saDkcLveXOZj-cxM80uwQ6ukp1aTi5wmMLis46mkqro/edit?gid=892978193#gid=892978193)

## 🚀 Демо

[👉 Открыть демо](https://printing-price-calculator-iota.vercel.app/)

---

## ✨ Возможности

- 📋 **Пошаговый ввод данных** — заказчик, заказ, материалы, операции, результат
- 💱 **Автообновление курсов валют** — USD и EUR напрямую с сайта ЦБ РФ
- 💾 **Автосохранение** — все данные сохраняются в `localStorage`
- 📊 **Интеграция с Google Sheets** — расчёт происходит по формулам в таблице
- 🛡️ **Защита от ошибок** — валидация ввода и понятные подсказки
- 📱 **Адаптивный интерфейс** — удобно работать с любого устройства

---

## 🛠️ Стек технологий

- **Frontend:** React, JavaScript, CSS3
- **Backend (прокси):** Node.js (Vercel Serverless Functions)
- **API:** Google Apps Script, Google Sheets API
- **Деплой:** Vercel
- **Тесты:** Jest + React Testing Library

---

## 📦 Установка и запуск

```bash
# Клонируй репозиторий
git clone https://github.com/Artzvinger/printing-price-calculator.git

# Перейди в папку проекта
cd printing-price-calculator

# Установи зависимости
npm install

# Запусти локально (включая прокси до Google Apps Script)
npm start
```

Приложение открывается на `http://localhost:3000`. Локальный прокси
(`src/setupProxy.js`) пробрасывает запросы `/api/proxy` на Google Apps Script,
поэтому расчёт работает и без Vercel.

### 🔧 Переменные окружения

Скопируй `.env.example` в `.env.local` и при необходимости укажи свой URL:

```bash
cp .env.example .env.local   # Windows: copy .env.example .env.local
```

| Переменная        | Где задаётся          | Описание                              |
| ----------------- | --------------------- | ------------------------------------- |
| `GAS_WEB_APP_URL` | `.env.local` / Vercel | URL развёрнутого Apps Script          |
| `SPREADSHEET_ID`  | Свойства скрипта Apps Script | ID Google-таблицы (см. ниже) |

`SPREADSHEET_ID` задаётся не в `.env`, а в самом скрипте:
**Расширения → Apps Script → Файл → Настройки проекта → Свойства скрипта →**
ключ `SPREADSHEET_ID`, значение — ID вашей таблицы.

---

## 🧪 Тесты

```bash
npm test                 # запуск в интерактивном режиме
CI=true npm test         # однократный прогон всех тестов
```

Покрывают: загрузку курсов валют, компонент шага «Заказ», корневой компонент
`App` и хук `useLocalStorage`.

---

## 📦 Деплой

```bash
npm run build            # production-сборка в папку build
```

На Vercel задай переменную `GAS_WEB_APP_URL` в настройках проекта.
