/* eslint-disable no-useless-escape */
/* eslint-disable jsx-a11y/label-has-associated-control */
import React, { useState } from 'react';
import ExcelJS, { CellRichTextValue, CellValue } from 'exceljs';

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

        const cellA44Value = worksheet.getCell('A44').value;
        const lastRow =
          typeof cellA44Value === 'string' && cellA44Value.startsWith('31.')
            ? 45
            : 44;

        const speedometer = worksheet.getCell(`D${lastRow}`).value;
        const fuel = worksheet.getCell(`I${lastRow}`).value;
        const rawValue: CellValue = worksheet.getCell('A5').value;

        console.log(112, rawValue);

        let str: string;

        if (typeof rawValue === 'string') {
          str = rawValue;
        } else if (
          typeof rawValue === 'object' &&
          rawValue !== null &&
          'richText' in rawValue
        ) {
          str = (rawValue as CellRichTextValue).richText
            .map((el) => el.text)
            .join('');
        } else {
          str = ''; // fallback, якщо тип не підтримується
        }

        worksheet.getCell('K64').value = daysInMonth;
        worksheet.getCell('K66').value = daysInMonth;
        worksheet.getCell('A5').value = replaceMonthAndYearWithCurrent(str);
        worksheet.getCell('A78').value = str.replace(/\b\d{4}\b/, '2026');

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
          // worksheet.getCell('A6').value = Math.round(sourceValue); // 🟩 тут нове
        }
        //!-----------------------------------------------------------------
        if (typeof fuel === 'number') {
          sourceValueTwo = fuel;
        } else if (fuel && typeof fuel === 'object' && 'result' in fuel) {
          sourceValueTwo = (fuel as ExcelJS.CellFormulaValue).result as number;
        }

        if (sourceValueTwo !== undefined) {
          worksheet.getCell('K50').value = Math.round(sourceValueTwo); // 🟩 тут нове
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
              const currentFormula = jCell.formula;

              if (currentFormula) {
                // Якщо формула існує, але не має ROUND(..., 0)
                if (!/^ROUND\(.+,\s*0\)$/.test(currentFormula)) {
                  jCell.value = {
                    formula: `ROUND(${currentFormula}, 0)`,
                  };
                }
              } else if (formulaTemplate) {
                // Якщо формули немає — вставити нову з ROUND
                const relativeFormula = formulaTemplate.replace(
                  /([A-Z]+)(\d+)/g,
                  (_, col) => `${col}${row}`
                );
                jCell.value = {
                  formula: `ROUND(${relativeFormula}, 0)`,
                };
              }
            }
          }
        }

        const missingDays = [];
        for (let d = 1; d <= daysInMonth; d++) {
          if (!existingDays.includes(d)) missingDays.push(d);
        }

        let insertRow =
          Math.max(
            ...existingDays.map((d) => {
              // Знайти реальний рядок для кожного дня (перевірити де він стоїть)
              for (let row = startRow; row <= endRow; row++) {
                const cellValue = worksheet.getCell(`A${row}`).value;
                if (
                  typeof cellValue === 'string' &&
                  cellValue.startsWith(`${String(d).padStart(2, '0')}.`)
                ) {
                  return row;
                }
              }
              return startRow;
            })
          ) + 1;
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
            worksheet.getCell(`J${insertRow}`).value = {
              formula: `ROUND(${formula}, 0)`,
            };
          }

          insertRow++;
        });

        const lastRowCell = worksheet.getCell('A44');
        const lastRowValue = lastRowCell.value;

        const cellA45 = worksheet.getCell('A45');
        const cellB45 = worksheet.getCell('B45');
        const cellC45 = worksheet.getCell('C45');
        const cellD45 = worksheet.getCell('D45');
        const cellE45 = worksheet.getCell('E45');
        const cellF45 = worksheet.getCell('F45');
        const cellG45 = worksheet.getCell('G45');
        const cellH45 = worksheet.getCell('H45');
        const cellI45 = worksheet.getCell('I45');
        const cellJ45 = worksheet.getCell('J45');
        const cellK45 = worksheet.getCell('K45');
        const cellL45 = worksheet.getCell('L45');
        const cellM45 = worksheet.getCell('M45');
        const array = [
          cellA45,
          cellB45,
          cellC45,
          cellD45,
          cellE45,
          cellF45,
          cellG45,
          cellH45,
          cellI45,
          cellJ45,
          cellK45,
          cellL45,
          cellM45,
        ];

        if (
          typeof lastRowValue === 'string' &&
          lastRowValue.startsWith('31.') &&
          daysInMonth === 30
        ) {
          worksheet.getCell('K51').value = { formula: '=K44+L44' };
          worksheet.getCell('K52').value = { formula: '=I44' };
          worksheet.getCell('K53').value = { formula: '=J44' };
          worksheet.getCell('K54').value = { formula: '=K53' };
          // 🟩 Додаємо формули у 44-й рядок
          worksheet.getCell('D44').value = { formula: '=D43+E43' };
          worksheet.getCell('E44').value = { formula: '=SUM(E14:E43)' };
          worksheet.getCell('F44').value = { formula: '=SUM(F14:F43)' };
          worksheet.getCell('G44').value = { formula: '=SUM(G14:G43)' };
          worksheet.getCell('H44').value = { formula: '=SUM(H14:H43)' };
          worksheet.getCell('I44').value = { formula: '=I43' };
          worksheet.getCell('J44').value = { formula: '=SUM(J14:J43)' };
          worksheet.getCell('K44').value = { formula: '=SUM(K14:K43)' };
          worksheet.getCell('L44').value = { formula: '=SUM(L14:L43)' };
          worksheet.getCell('M44').value = { formula: '=SUM(M14:M43)' };
          worksheet.getCell('K65').value = { formula: '=COUNTA(E14:E43)' };
          worksheet.getCell('F7').value = { formula: '=D44' };

          worksheet.mergeCells('A44:B44');
          const cell = worksheet.getCell('A44');

          cell.value = 'Всього:';
          cell.style = {
            font: { bold: true, name: 'Times New Roman', size: 12 },
            alignment: { horizontal: 'center', vertical: 'middle' },
            border: {
              top: { style: 'thin' },
              left: { style: 'thin' },
              bottom: { style: 'thin' },
              right: { style: 'thin' },
            },
          };

          for (const c of array) {
            c.style = {
              font: { bold: false, name: 'Times New Roman', size: 12 },
              alignment: { horizontal: 'center', vertical: 'middle' },
              border: {
                top: { style: 'thin' },
              },
            };
          }

          // 🧹 Очищаємо 45-й рядок
          ['A', 'B', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'].forEach(
            (col) => {
              worksheet.getCell(`${col}45`).value = null;
            }
          );
        }

        if (daysInMonth === 31) {
          worksheet.getCell('K51').value = { formula: '=K45+L45' };
          worksheet.getCell('K52').value = { formula: '=I45' };
          worksheet.getCell('K53').value = { formula: '=J45' };
          worksheet.getCell('K54').value = { formula: '=K53' };

          worksheet.getCell('D45').value = { formula: '=D44+E44' };
          worksheet.getCell('E45').value = { formula: '=SUM(E14:E44)' };
          worksheet.getCell('F45').value = { formula: '=SUM(F14:F44)' };
          worksheet.getCell('G45').value = { formula: '=SUM(G14:G44)' };
          worksheet.getCell('H45').value = { formula: '=SUM(H14:H44)' };
          worksheet.getCell('I45').value = { formula: '=I44' };
          worksheet.getCell('J45').value = { formula: '=SUM(J14:J44)' };
          worksheet.getCell('K45').value = { formula: '=SUM(K14:K44)' };
          worksheet.getCell('L45').value = { formula: '=SUM(L14:L44)' };
          worksheet.getCell('M45').value = { formula: '=SUM(M14:M44)' };
          worksheet.getCell('K65').value = { formula: '=COUNTA(E14:E44)' };

          worksheet.getCell('F7').value = { formula: '=D45' };

          const cellA45 = worksheet.getCell('A45');

          if (!cellA45.isMerged) {
            worksheet.mergeCells('A45:B45');
          }

          cellA45.value = 'Всього:';
          cellA45.style = {
            font: { bold: true, name: 'Times New Roman', size: 12 },
            alignment: { horizontal: 'center', vertical: 'middle' },
          };

          // Очистити 44-й рядок під 31 число
          ['A', 'B', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'].forEach(
            (col) => {
              worksheet.getCell(`${col}44`).value = null;
            }
          );

          worksheet.unMergeCells('A44:B44');
          worksheet.getCell('A44').value = `31.${month}`;
          const cell = worksheet.getCell('A44');
          const cell45 = worksheet.getCell('A45');
          cell.style = {
            font: { bold: false, name: 'Times New Roman', size: 11 },
            alignment: { horizontal: 'center', vertical: 'middle' },
            border: {
              top: { style: 'thin' },
              left: { style: 'thin' },
              bottom: { style: 'thin' },
              right: { style: 'thin' },
            },
          };

          worksheet.getCell('D44').value = { formula: '=D43+E43' };
          worksheet.getCell('I44').value = { formula: '=I43+K44+L44-J44' };
          const j44 = worksheet.getCell('J44');

          // Перевіряємо: чи формула вже обгорнута рівно один раз у ROUND(..., 0)
          if (j44.formula) {
            const existing = j44.formula.trim();

            const alreadyRounded = /^ROUND\([^\)]*,\s*0\)$/.test(existing);

            if (!alreadyRounded) {
              j44.value = {
                formula: `ROUND(${existing}, 0)`,
              };
            }
          } else if (formulaTemplate) {
            const dynamicFormula = formulaTemplate.replace(
              /([A-Z]+)(\d+)/g,
              (_, col) => `${col}44`
            );

            j44.value = {
              formula: dynamicFormula,
            };
          }

          for (const c of array) {
            c.style = {
              font: { bold: false, name: 'Times New Roman', size: 12 },
              alignment: { horizontal: 'center', vertical: 'middle' },
              border: {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' },
              },
            };
          }

          cell45.style = {
            font: { bold: true, name: 'Times New Roman', size: 12 },
            alignment: { horizontal: 'center', vertical: 'middle' },
            border: {
              top: { style: 'thin' },
              left: { style: 'thin' },
              bottom: { style: 'thin' },
              right: { style: 'thin' },
            },
          };
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
