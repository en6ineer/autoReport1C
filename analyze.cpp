#include <fstream>
#include <iostream>
#include <string>
#include <xlsxwriter.h>

void analyzeFile(const std::string &filename,
                 const std::string &excelFilename) {
  std::ifstream file(filename);

  if (!file.is_open()) {
    std::cerr << "Не удалось открыть файл: " << filename << std::endl;
    return;
  }

  // Создаем новый Excel файл
  lxw_workbook *workbook = workbook_new(excelFilename.c_str());
  lxw_worksheet *worksheet = workbook_add_worksheet(workbook, NULL);

  // Заполняем заголовки колонок
  worksheet_write_string(worksheet, 0, 0, "Тип объекта", NULL);
  worksheet_write_string(worksheet, 0, 1, "Имя объекта", NULL);
  worksheet_write_string(worksheet, 0, 2, "Описание изменений", NULL);
  worksheet_write_string(worksheet, 0, 3, "Детали кода", NULL);

  int row = 1;

  std::string currentObjectName;
  std::string objectType;
  std::string changeDescription;
  std::string codeDetails;

  while (file.peek() != EOF) {
    std::string line;
    std::getline(file, line);

    if (line.find("Объект изменен") != std::string::npos) {
      // Определение типа объекта
      if (line.find("Справочник") != std::string::npos) {
        objectType = "Справочник";
      } else if (line.find("Документ") != std::string::npos) {
        objectType = "Документ";
      } else if (line.find("Отчет") != std::string::npos) {
        objectType = "Отчет";
      }

      // Имя объекта
      size_t objNameStart = line.find(":") + 1;
      currentObjectName = line.substr(objNameStart);
    } else if (line.find("Изменили") != std::string::npos ||
               line.find("Добавили") != std::string::npos) {
      // Описание изменений
      changeDescription = line;
    }

    // Записываем данные в Excel файл
    worksheet_write_string(worksheet, row, 0, objectType.c_str(), NULL);
    worksheet_write_string(worksheet, row, 1, currentObjectName.c_str(), NULL);
    worksheet_write_string(worksheet, row, 2, changeDescription.c_str(), NULL);
    worksheet_write_string(worksheet, row, 3, codeDetails.c_str(), NULL);

    ++row;
  }

  file.close();

  // Закрываем Excel файл
  workbook_close(workbook);
}

int main() {
  std::string filename = "report.txt"; // Имя файла с отчетом
  std::string excelFilename =
      "report.xlsx"; // Имя файла для сохранения данных в Excel

  analyzeFile(filename, excelFilename);
  return 0;
}
