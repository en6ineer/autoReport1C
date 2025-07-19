#include <algorithm> // для std::find_if_not
#include <cstddef>
#include <fstream>
#include <iostream>
#include <locale.h>
#include <stack>
#include <string>
#include <vector>
#include <xlsxwriter.h>

// Используем стандартное пространство имен для удобства
using namespace std;

// =================================================================================
// 1. СТРУКТУРЫ И УТИЛИТАРНЫЕ ФУНКЦИИ
// =================================================================================

/**
 * @brief Структура для хранения полной информации о разобранном объекте.
 */
struct ChangedObject {
  string fullPath;   // Полный иерархический путь (Конфигурация.Справочник.Имя)
  string objectType; // Тип объекта (Справочник, Документ и т.д.)
  string objectName; // Конечное имя объекта
  string state;      // Состояние (Изменен, Новый, Удален)
  string
      changeDescription; // Описание (например, "Модуль - Различаются значения")
  string codeDetails;    // Собранный и очищенный код изменений

  // Конструктор для очистки полей при создании
  ChangedObject()
      : fullPath(""), objectType(""), objectName(""), state(""),
        changeDescription(""), codeDetails("") {}

  bool isEmpty() const { return state.empty(); }
};

/**
 * @brief Удаляет начальные и конечные пробелы из строки.
 * @param s Исходная строка.
 * @return Очищенная строка.
 */
string trim(const string &s) {
  auto start = s.find_first_not_of(" \t\n\r");
  if (string::npos == start)
    return "";
  auto end = s.find_last_not_of(" \t\n\r");
  return s.substr(start, end - start + 1);
}

/**
 * @brief Определяет уровень вложенности строки по количеству отступов.
 * @param line Строка из файла отчета.
 * @return Целочисленный уровень вложенности (0, 1, 2...).
 */
size_t getIndentLevel(const string &line) {
  size_t indent = 0;
  for (char c : line) {
    if (c == '\t') {
      indent++;
    } else {
      break;
    }
  }
  return indent;
}

/**
 * @brief Очищает строку кода от служебных символов отчета.
 * @param line "Грязная" строка кода.
 * @return "Чистая" строка кода.
 */
string cleanCodeLine(string line) {
  line = trim(line);
  if (line.rfind("< ", 0) == 0 || line.rfind("> ", 0) == 0) {
    line = line.substr(2);
  }
  if (!line.empty() && line.front() == '"' && line.back() == '"') {
    line = line.substr(1, line.length() - 2);
  }
  // Удаляем символ '·'
  line.erase(remove(line.begin(), line.end(), L'·'), line.end());
  return trim(line);
}

// =================================================================================
// 2. ОСНОВНАЯ ФУНКЦИЯ АНАЛИЗА ФАЙЛА
// =================================================================================

void analyzeFile(const string &filename, const string &excelFilename) {
    ifstream file(filename);
    if (!file.is_open()) {
        cout << "Не удалось открыть файл: " << filename << endl;
        return;
    }

    // --- Инициализация Excel ---
    lxw_workbook *workbook = workbook_new(excelFilename.c_str());
    if (!workbook) {
        cout << "Не удалось создать Excel файл: " << excelFilename << endl;
        file.close();
        return;
    }
    lxw_worksheet *worksheet = workbook_add_worksheet(workbook, NULL);
    worksheet_write_string(worksheet, 0, 0, "Тип объекта", NULL);
    worksheet_write_string(worksheet, 0, 1, "Имя объекта", NULL);
    worksheet_write_string(worksheet, 0, 2, "Состояние", NULL);
    worksheet_write_string(worksheet, 0, 3, "Описание", NULL);
    worksheet_write_string(worksheet, 0, 4, "Детали кода", NULL);

    // --- Основные переменные для парсинга ---
    vector<ChangedObject> allObjects; // Финальный список всех разобранных объектов
    stack<string> contextStack;       // Стек для отслеживания иерархии
    ChangedObject currentObject;      // Текущий собираемый объект
    string line;

    // Пропускаем заголовок/легенду отчета
    while (getline(file, line) && line.find("- ***Конфигурация") == string::npos) {
        continue;
    }
    
    // Главный цикл разбора
    do {
        string trimmedLine = trim(line);
        if (trimmedLine.empty()) continue;

        size_t indent = getIndentLevel(line);
        size_t markerPos = trimmedLine.find("- ");
        string marker = "";
        string content = "";

        // Ищем маркеры состояния (***, -->, <---, ^-)
        if (markerPos != string::npos) {
            size_t contentStart = markerPos + 2;
            if (trimmedLine.find("***", contentStart) == contentStart) marker = "***";
            else if (trimmedLine.find("-->", contentStart) == contentStart) marker = "-->";
            else if (trimmedLine.find("<---", contentStart) == contentStart) marker = "<---";
            else if (trimmedLine.find("^-", contentStart) == contentStart) marker = "^-";
        }
        
        // ========================================
        // ЭТО НАЧАЛО НОВОГО ОБЪЕКТА
        // ========================================
        if (!marker.empty()) {
            // 1. Если уже был объект в обработке, сохраняем его в общий список
            if (!currentObject.isEmpty()) {
                allObjects.push_back(currentObject);
            }
            // 2. Создаем и инициализируем новый объект
            currentObject = ChangedObject();

            // 3. Управляем стеком иерархии
            while (contextStack.size() > indent) {
                contextStack.pop();
            }
            
            content = trim(trimmedLine.substr(markerPos + 2 + marker.length()));
            contextStack.push(content);

            // 4. Заполняем поля нового объекта
            if (marker == "***") currentObject.state = "Изменен";
            else if (marker == "-->") currentObject.state = "Новый";
            else if (marker == "<---") currentObject.state = "Удален";
            else if (marker == "^-") currentObject.state = "Порядок изменен";
            
            // Собираем полный путь из стека
            stack<string> tempStack = contextStack;
            string fullPathStr = "";
            while(!tempStack.empty()) {
                fullPathStr = tempStack.top() + (fullPathStr.empty() ? "" : "." + fullPathStr);
                tempStack.pop();
            }
            currentObject.fullPath = fullPathStr;

            // Извлекаем тип и имя
            size_t firstDot = content.find('.');
            if (firstDot != string::npos) {
                currentObject.objectType = content.substr(0, firstDot);
                currentObject.objectName = content.substr(firstDot + 1);
            } else {
                currentObject.objectType = "Конфигурация"; // Корень
                currentObject.objectName = content;
            }

        } 
        // ========================================
        // ЭТО ДЕТАЛИ ДЛЯ ТЕКУЩЕГО ОБЪЕКТА
        // ========================================
        else if (!currentObject.isEmpty()) {
            if (trimmedLine.find("Изменено:") != string::npos || 
                trimmedLine.find("Объект присутствует только") != string::npos ||
                trimmedLine.find("Различаются значения") != string::npos) {
                // Если уже есть описание, добавляем через новую строку
                if (!currentObject.changeDescription.empty()) {
                    currentObject.changeDescription += "\n" + trimmedLine;
                } else {
                    currentObject.changeDescription = trimmedLine;
                }
            } else {
                // Иначе это, скорее всего, строка кода
                string cleaned = cleanCodeLine(line);
                if (!cleaned.empty()) {
                    currentObject.codeDetails += cleaned + "\n";
                }
            }
        }

    } while (getline(file, line));

    // Не забываем сохранить самый последний объект
    if (!currentObject.isEmpty()) {
        allObjects.push_back(currentObject);
    }

    // --- Запись результатов в Excel ---
    int row = 1;
    for (const auto& obj : allObjects) {
        worksheet_write_string(worksheet, row, 0, obj.objectType.c_str(), NULL);
        worksheet_write_string(worksheet, row, 1, obj.objectName.c_str(), NULL);
        worksheet_write_string(worksheet, row, 2, obj.state.c_str(), NULL);
        worksheet_write_string(worksheet, row, 3, obj.changeDescription.c_str(), NULL);
        worksheet_write_string(worksheet, row, 4, obj.codeDetails.c_str(), NULL);
        row++;
    }

    file.close();
    workbook_close(workbook);
    cout << "Отчет успешно сохранен в " << excelFilename << endl;
    cout << "Объектов обработано: " << allObjects.size() << endl;
}


// =================================================================================
// 3. ФУНКЦИИ МЕНЮ И ТОЧКА ВХОДА
// =================================================================================

void menu() {
  int choice;
  do {
    cout << "Анализ отчетов 1С" << endl;
    cout << endl
         << "1. Разобрать отчет\n"
         << "0. Выйти из программы"
         << "\nВыберите действие: ";
    cin >> choice;

   
    switch (choice) {
    case 1: {
      cout << "Введите имя файла отчета (например, report.txt):\n";
      string filename;
      cin >> filename;
      string excelFilename = "report_new.xlsx";
      analyzeFile(filename, excelFilename);
      break;
    }
    case 0: {
      break;
    }
    default:
      cout << "\nВы ввели неверное значение! Повторите выбор команды.\n";
    }

    if (choice != 0) {
      cout << "\nНажмите Enter, чтобы продолжить...\n";
      cin.get();
    }

  } while (choice != 0);
}

int main() {
  setlocale(LC_ALL, "Russian");
  menu();
  return 0;
}
