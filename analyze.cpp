#include <cstddef>
#include <fstream>
#include <iostream>
#include <locale.h>
#include <string>
#include <xlsxwriter.h>
using namespace std;

void analyzeFile(const std::string &filename,
                 const std::string &excelFilename) {

  std::ifstream file(filename);
  if (!file.is_open()) {
    std::cout << "Не удалось открыть файл: " << filename << std::endl;
    return;
  }

  // Создаем новый Excel файл
  lxw_workbook *workbook = workbook_new(excelFilename.c_str());
  if (!workbook) {
    std::cout << "Не удалось создать Excel файл: " << excelFilename
              << std::endl;
    file.close();
    return;
  }
  lxw_worksheet *worksheet = workbook_add_worksheet(workbook, NULL);

  // Заголовки колонок
  worksheet_write_string(worksheet, 0, 0, "Тип объекта", NULL);
  worksheet_write_string(worksheet, 0, 1, "Имя объекта", NULL);
  worksheet_write_string(worksheet, 0, 2, "Состояние объекта", NULL);
  worksheet_write_string(worksheet, 0, 3, "Описание изменений", NULL);
  worksheet_write_string(worksheet, 0, 4, "Детали кода", NULL);

  int row = 1;                   // Текущая строка в Excel
  std::string objectType;        // Тип объекта
  std::string objectName;        // Имя объекта
  std::string objectState;       // Изменен / Новый
  std::string changeDescription; // Описание изменений
  std::string codeDetails;       // Детали кода
  bool inCodeBlock = false;

  while (file.peek() != EOF) {
    std::string line;
    std::getline(file, line);

    cout << line << "\n"; // for debug

    // Пропускаем пояснения в начале отчета
    if (line.find("Объект изменен") != std::string::npos ||
        line.find("Объект присутствует только") != std::string::npos ||
        line.find("Порядок объекта изменен") != std::string::npos) {
      continue;
    }

    // Определяем состояние и тип объекта
    if (line.find("***") != std::string::npos) {
      // Записываем предыдущий объект, если он был
      if (!objectName.empty()) {
        worksheet_write_string(worksheet, row, 0, objectType.c_str(), NULL);
        worksheet_write_string(worksheet, row, 1, objectName.c_str(), NULL);
        worksheet_write_string(worksheet, row, 2, objectState.c_str(), NULL);
        worksheet_write_string(worksheet, row, 3, changeDescription.c_str(),
                               NULL);
        worksheet_write_string(worksheet, row, 4, codeDetails.c_str(), NULL);
        ++row;
        changeDescription.clear();
        codeDetails.clear();
        inCodeBlock = false;
      }

      objectState = "Изменен";
      size_t nameStart = line.find("***") + 3;
      objectName = line.substr(nameStart);

      if (line.find("Справочник") != std::string::npos) {
        objectType = "Справочник";
      } else if (line.find("Документ") != std::string::npos) {
        objectType = "Документ";
      } else if (line.find("Отчет") != std::string::npos) {
        objectType = "Отчет";
      } else if (line.find("ОбщийМодуль") != std::string::npos) {
        objectType = "ОбщийМодуль";
      } else {
        objectType = "Неопределен";
      }
    } else if (line.find("-->") != std::string::npos) {
      if (!objectName.empty()) {
        worksheet_write_string(worksheet, row, 0, objectType.c_str(), NULL);
        worksheet_write_string(worksheet, row, 1, objectName.c_str(), NULL);
        worksheet_write_string(worksheet, row, 2, objectState.c_str(), NULL);
        worksheet_write_string(worksheet, row, 3, changeDescription.c_str(),
                               NULL);
        worksheet_write_string(worksheet, row, 4, codeDetails.c_str(), NULL);
        ++row;
        changeDescription.clear();
        codeDetails.clear();
      }

      objectState = "Новый";
      size_t nameStart = line.find("-->") + 3;
      objectName = line.substr(nameStart);

      if (line.find("ОбщийМодуль") != std::string::npos) {
        objectType = "ОбщийМодуль";
      } else {
        objectType = "Неопределен";
      }
    } else if (line.find("<---") != std::string::npos) {
      if (!objectName.empty()) {
        worksheet_write_string(worksheet, row, 0, objectType.c_str(), NULL);
        worksheet_write_string(worksheet, row, 1, objectName.c_str(), NULL);
        worksheet_write_string(worksheet, row, 2, objectState.c_str(), NULL);
        worksheet_write_string(worksheet, row, 3, changeDescription.c_str(),
                               NULL);
        worksheet_write_string(worksheet, row, 4, codeDetails.c_str(), NULL);
        ++row;
        changeDescription.clear();
        codeDetails.clear();
      }

      objectState = "Удален";
      size_t nameStart = line.find("<---") + 4;
      objectName = line.substr(nameStart);
      objectType = "Неопределен";
    } else if (line.find("^-") != std::string::npos) {
      if (!objectName.empty()) {
        worksheet_write_string(worksheet, row, 0, objectType.c_str(), NULL);
        worksheet_write_string(worksheet, row, 1, objectName.c_str(), NULL);
        worksheet_write_string(worksheet, row, 2, objectState.c_str(), NULL);
        worksheet_write_string(worksheet, row, 3, changeDescription.c_str(),
                               NULL);
        worksheet_write_string(worksheet, row, 4, codeDetails.c_str(), NULL);
        ++row;
        changeDescription.clear();
        codeDetails.clear();
      }

      objectState = "Порядок изменен";
      size_t nameStart = line.find("^-") + 2;
      objectName = line.substr(nameStart);
      objectType = "Неопределен";
    }
    // Описание изменений
    else if (line.find("Изменено:") != std::string::npos ||
             line.find("Объект присутствует только") != std::string::npos) {
      changeDescription = line;
      inCodeBlock = true;
      codeDetails.clear();
    }
    // Собираем детали кода
    else if (inCodeBlock && !line.empty()) {
      if (line.find("< ") != std::string::npos ||
          line.find("> ") != std::string::npos) {
        codeDetails += line + "\n";
      } else if (codeDetails.empty()) {
        codeDetails = line + "\n"; // Первая строка кода
      } else {
        codeDetails += line + "\n";
        inCodeBlock = false; // Код закончился
      }
    }
  }

  // Записываем последний объект, если он есть
  if (!objectName.empty()) {
    worksheet_write_string(worksheet, row, 0, objectType.c_str(), NULL);
    worksheet_write_string(worksheet, row, 1, objectName.c_str(), NULL);
    worksheet_write_string(worksheet, row, 2, objectState.c_str(), NULL);
    worksheet_write_string(worksheet, row, 3, changeDescription.c_str(), NULL);
    worksheet_write_string(worksheet, row, 4, codeDetails.c_str(), NULL);
  }

  file.close();
  workbook_close(workbook);
  std::cout << "Отчет успешно сохранен в " << excelFilename << std::endl;
}

void menu() {

  int choice;
  do {
    cout << "Записная книжка" << endl;
    cout << endl
         << "1. Разобрать отчет\n"
         //  << "2. Добавить контакт\n"
         //  << "3. Редактировать контакт\n"
         //  << "4. Удалить контакт\n"
         << "0. Выйти из программы"
         << "\nВыберите действие: ";
    cin >> choice;

    switch (choice) {
    case 1: {
      std::cout << "Введите название файла с расширением .txt\n";
      std::string filename;
      std::string excelFilename = "report.xlsx";
      std::cin >> filename;

      analyzeFile(filename, excelFilename);
      break;
    }

    case 0: {
      break;
    }

    default:
      cout << "\n Вы ввели неверное значение! Повторите выбор команды.\n";
    }

    cout << "\nНажмите Enter, чтобы продолжить...\n";
    std::cin.ignore();
    std::cin.get();

  } while (choice != 0);
}

int main() {

  setlocale(LC_ALL, "Russian");
  menu();
  return 0;
}
