# GPT Помощник для Google Docs 🚀

Мощное расширение для Google Docs, которое использует возможности GPT для улучшения работы с текстом. Расширение предлагает различные инструменты обработки текста с выводом на русском языке и возможностью перевода на английский.

## ✨ Возможности

### 🎯 Обработка текста
- **Краткое содержание**: Автоматическое создание сжатого изложения выбранного текста
- **Улучшение текста**: Повышение качества и профессионализма написанного
- **Исправление грамматики**: Автоматическая проверка и исправление ошибок
- **Стилистика**: Преобразование в формальный или разговорный стиль
- **Перевод**: Перевод текста на английский язык

### 💻 Интерфейс
- **Боковая панель**: Удобный интерфейс для работы с текстом
- **Контекстное меню**: Быстрый доступ ко всем функциям
- **Панель настроек**: Гибкая настройка параметров API и модели

### ⚙️ Настройки
- Базовый URL для API OpenAI
- Выбор модели:
  - GPT-3.5-turbo
  - GPT-4
  - GPT-4-turbo-preview
  - Claude-3-Sonnet
  - Пользовательская модель
- Настройка температуры (0-1)
- Установка максимальной длины ответа (от 150 токенов)

## 🚀 Установка

1. Откройте редактор Google Apps Script
2. Создайте новый проект
3. Скопируйте следующие файлы в проект:
   - `Code.gs`
   - `Sidebar.html`
   - `Settings.html`
4. Настройте ключ API OpenAI:
   - Перейдите в Настройки проекта
   - Нажмите "Script Properties"
   - Добавьте новое свойство:
     - Имя: `OPENAI_API_KEY`
     - Значение: Ваш ключ API OpenAI
5. Сохраните и опубликуйте как дополнение для Google Docs

## 📝 Использование

### Основные функции
1. Откройте документ Google Docs
2. Найдите меню "GPT Помощник"
3. Выберите нужную функцию:
   - Используйте боковую панель для интерактивной работы
   - Используйте меню для быстрых действий
   - Настройте параметры в разделе настроек

### Боковая панель
1. Нажмите "Показать панель"
2. Введите или вставьте текст
3. Выберите нужное действие
4. Просмотрите результат
5. Нажмите "Вставить в документ"

### Настройки
1. Откройте "Настройки" в меню GPT Помощник
2. Настройте:
   - URL API
   - Модель
   - Температуру
   - Максимальное количество токенов
3. Нажмите "Сохранить"

## 🔧 Структура проекта

```
├── Code.gs                  # Основной файл с логикой
├── Sidebar.html            # Интерфейс боковой панели
├── Settings.html           # Интерфейс настроек
└── appsscript.json         # Конфигурация проекта
```

## 📋 Требования

- Аккаунт Google Workspace
- Ключ API OpenAI
- Браузер Google Chrome (рекомендуется)

## 🔐 Безопасность

Ключ API OpenAI хранится в защищенном хранилище Script Properties. Никогда не передавайте свой ключ API третьим лицам и не включайте его напрямую в код.

## 🤝 Участие в разработке

Мы приветствуем ваши предложения по улучшению проекта! Создавайте issues и pull requests.

## 📄 Лицензия

Этот проект лицензирован под [MIT License](https://opensource.org/licenses/MIT):

```
MIT License

Copyright (c) 2024 GPT Помощник для Google Docs

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

---
Сделано с ❤️ для удобной работы с текстом 