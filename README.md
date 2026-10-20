# 🖨️ Калькулятор типографии

Веб-приложение для мгновенного расчёта стоимости полиграфической продукции.  
Интегрирован с **Google Sheets** через **Google Apps Script** — все формулы и цены хранятся и обновляются в таблице.
---
[Копия интегрированной таблицы на которую приходят запросы](https://docs.google.com/spreadsheets/d/1saDkcLveXOZj-cxM80uwQ6ukp1aTi5wmMLis46mkqro/edit?gid=892978193#gid=892978193)

## 🚀 Демо

[👉 Открыть демо](https://printing-price-calculator-iota.vercel.app/)

---

## 📸 Скриншоты
  
### Информация о заказчике
<img width="1079" height="453" alt="Screenshot_2" src="https://github.com/user-attachments/assets/ec6354d5-7326-4390-b378-b9015ca68d2e" />

### Параметры заказа 
<img width="1091" height="831" alt="Screenshot_3" src="https://github.com/user-attachments/assets/e8547615-9522-45c6-a257-78338a866bad" />

### Информация о заказе и курс ЦБ
<img width="1089" height="650" alt="Screenshot_4" src="https://github.com/user-attachments/assets/8fb329c1-beae-490b-9423-532b4e5da879" />

### Результат расчёта
<img width="1359" height="940" alt="Screenshot_1" src="https://github.com/user-attachments/assets/8f678d36-3090-471c-b5da-eb2331f933b9" />


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
