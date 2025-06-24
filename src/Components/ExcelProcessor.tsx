/* eslint-disable jsx-a11y/label-has-associated-control */
import React, { useState } from 'react';
import ExcelJS from 'exceljs';

import styles from './ExcelProcessor.module.scss';

function replaceMonthAndYearWithCurrent(text: string): string {
  const months = [
    'січень',
    'лютий',
    'березень',
    'квітень',
    'травень',
    'червень',
    'липень',
    'серпень',
    'вересень',
    'жовтень',
    'листопад',
    'грудень',
  ];

  const now = new Date();
  const currentMonthIndex = now.getMonth(); // 0–11
  const currentMonthName = months[currentMonthIndex];
  const currentYear = now.getFullYear().toString();

  // Замінюємо місяць
  const monthRegex = new RegExp(months.join('|'), 'i');
  let updatedText = text.replace(monthRegex, currentMonthName);

  // Замінюємо рік (будь-яке число з 4 цифр)
  updatedText = updatedText.replace(/\b\d{4}\b/, currentYear);

  return updatedText;
}

export const ExcelEditor: React.FC = () => {
  const [month, setMonth] = useState('07');
  const [monthName, setMonthName] = useState('липень');
  const [loading, setLoading] = useState(false);
  const [progress, setProgress] = useState(0);
  const [message, setMessage] = useState('');

  const monthOptions = [
    { value: '01', label: 'січень' },
    { value: '02', label: 'лютий' },
    { value: '03', label: 'березень' },
    { value: '04', label: 'квітень' },
    { value: '05', label: 'травень' },
    { value: '06', label: 'червень' },
    { value: '07', label: 'липень' },
    { value: '08', label: 'серпень' },
    { value: '09', label: 'вересень' },
    { value: '10', label: 'жовтень' },
    { value: '11', label: 'листопад' },
    { value: '12', label: 'грудень' },
  ];

  const handleUploadAndEdit = async (
    e: React.ChangeEvent<HTMLInputElement>
  ) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    const daysInMonthMap: Record<string, number> = {
      '01': 31,
      '02': 28,
      '03': 31,
      '04': 30,
      '05': 31,
      '06': 30,
      '07': 31,
      '08': 31,
      '09': 30,
      '10': 31,
      '11': 30,
      '12': 31,
    };

    const daysInMonth: number = daysInMonthMap[month];
    if (!daysInMonth) {
      alert('Некоректне значення місяця.');
      return;
    }

    const clearColumns = ['B', 'E', 'F', 'G', 'H', 'K', 'L', 'M'];
    const startRow = 13;
    const endRow = 50;

    setLoading(true);
    setProgress(0);
    setMessage('');

    try {
      for (let i = 0; i < files.length; i++) {
        const file = files[i];
        const buffer = await file.arrayBuffer();
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(buffer);

        const worksheet = workbook.getWorksheet('Картки обліку роботи машин');
        if (!worksheet) continue;

        // Взяти значення з D44 і додати в F6 і G6
        const speedometer = worksheet.getCell('D44').value;
        // Взяти значення з I44 і додати в K49
        const fuel = worksheet.getCell('I44').value;
        // Взяти значення з I44 і додати в K49
        const str = worksheet.getCell('A5').value;

        worksheet.getCell('K63').value = daysInMonth;
        worksheet.getCell('K65').value = daysInMonth;
        worksheet.getCell('A5').value = replaceMonthAndYearWithCurrent(
          String(str)
        );

        let sourceValue: number | undefined;
        let sourceValueTwo: number | undefined;

        if (typeof speedometer === 'number') {
          sourceValue = speedometer;
        } else if (
          speedometer &&
          typeof speedometer === 'object' &&
          'result' in speedometer
        ) {
          sourceValue = (speedometer as ExcelJS.CellFormulaValue)
            .result as number;
        }

        if (sourceValue !== undefined) {
          worksheet.getCell('F6').value = sourceValue;
          worksheet.getCell('G6').value = sourceValue;
        }
        //!-----------------------------------------------------------------
        if (typeof fuel === 'number') {
          sourceValueTwo = fuel;
        } else if (fuel && typeof fuel === 'object' && 'result' in fuel) {
          sourceValueTwo = (fuel as ExcelJS.CellFormulaValue).result as number;
        }

        if (sourceValueTwo !== undefined) {
          worksheet.getCell('K49').value = sourceValueTwo;
        }

        let formulaTemplate = '';
        for (let r = startRow; r <= endRow; r++) {
          const cell = worksheet.getCell(`J${r}`);
          if (cell.formula) {
            formulaTemplate = cell.formula;
            break;
          }
        }

        const existingDays: number[] = [];

        for (let row = startRow; row <= endRow; row++) {
          const cell = worksheet.getCell(`A${row}`);
          const raw = cell.value;

          if (typeof raw === 'string') {
            const match = raw.match(/^(\d{2})\.(\d{2})/);
            if (match) {
              const day = parseInt(match[1]);
              existingDays.push(day);
              cell.value = `${match[1]}.${month}`;

              clearColumns.forEach((col) => {
                const targetCell = worksheet.getCell(`${col}${row}`);
                if (!targetCell.formula) {
                  targetCell.value = null;
                }
              });

              const jCell = worksheet.getCell(`J${row}`);
              if (!jCell.formula && formulaTemplate) {
                const relativeFormula = formulaTemplate.replace(
                  /([A-Z]+)(\d+)/g,
                  (_, col) => `${col}${row}`
                );
                jCell.value = { formula: relativeFormula };
              }
            }
          }
        }

        const missingDays = [];
        for (let d = 1; d <= daysInMonth; d++) {
          if (!existingDays.includes(d)) missingDays.push(d);
        }

        let insertRow = startRow + existingDays.length;
        missingDays.forEach((day) => {
          if (insertRow > endRow) return;
          const formattedDay = `${String(day).padStart(2, '0')}.${month}`;
          worksheet.getCell(`A${insertRow}`).value = formattedDay;

          clearColumns.forEach((col) => {
            worksheet.getCell(`${col}${insertRow}`).value = null;
          });

          if (formulaTemplate) {
            const formula = formulaTemplate.replace(
              /([A-Z]+)(\d+)/g,
              (_, col) => `${col}${insertRow}`
            );
            worksheet.getCell(`J${insertRow}`).value = { formula };
          }

          insertRow++;
        });

        if (existingDays.length > daysInMonth) {
          for (
            let row = startRow + daysInMonth;
            row <= startRow + existingDays.length;
            row++
          ) {
            ['A', 'B', 'E', 'F', 'G', 'H', 'J', 'K', 'L', 'M'].forEach(
              (col) => {
                worksheet.getCell(`${col}${row}`).value = null;
              }
            );
          }
        }

        const headerCell = worksheet.getCell('A5');
        if (typeof headerCell.value === 'string') {
          headerCell.value = headerCell.value.replace(
            /(січень|лютий|березень|квітень|травень|червень|липень|серпень|вересень|жовтень|листопад|грудень)/,
            monthName
          );
        }

        const updatedBuffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([updatedBuffer], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        });

        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `${file.name.replace('.xlsx', '').replace(/\+.*/, '')}.xlsx`;
        link.click();

        // Оновити прогрес
        setProgress(Math.round(((i + 1) / files.length) * 100));
      }

      setMessage(`✅ Успішно оброблено ${files.length} файл(и).`);
    } catch (err) {
      console.error(err);
      setMessage('❌ Помилка під час обробки.');
    } finally {
      setLoading(false);
    }
  };

  const handleMonthChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const selected = monthOptions.find((opt) => opt.value === e.target.value);
    if (selected) {
      setMonth(selected.value);
      setMonthName(selected.label);
    }
  };

  return (
    <div className={styles.container}>
      <h3 className={styles.title}>Завантаж картки Excel</h3>

      <label className={styles.label}>Оберіть місяць: </label>
      <select
        value={month}
        onChange={handleMonthChange}
        className={styles.select}
      >
        {monthOptions.map((opt) => (
          <option key={opt.value} value={opt.value}>
            {opt.label}
          </option>
        ))}
      </select>

      <input
        type="file"
        accept=".xlsx"
        multiple
        onChange={handleUploadAndEdit}
        className={styles.fileInput}
      />

      {loading && (
        <div className={styles.progressContainer}>
          <div className={styles.progressBarBackground}>
            <div
              className={styles.progressBar}
              style={{ width: `${progress}%` }}
            />
          </div>
          <p>{progress}%</p>
        </div>
      )}

      {!loading && message && <p className={styles.message}>{message}</p>}
    </div>
  );
};
