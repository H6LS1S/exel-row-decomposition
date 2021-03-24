import { stream } from 'exceljs';

import * as inquirerFileTree from 'inquirer-file-tree-selection-prompt';
import * as inquirer from 'inquirer';
import * as figlet from 'figlet';
import * as chalk from 'chalk';
import * as path from 'path';

export function logger(context: string, value: string): void {
  console.info(
    ` ${chalk.green('~')}`,
    new Date().toLocaleDateString(),
    new Date().toLocaleTimeString(),
    chalk.yellow(`  [${context}]`),
    chalk.green(value),
  );
}

export async function bootstrap(): Promise<void> {
  inquirer.registerPrompt('path', inquirerFileTree);

  console.clear();
  console.log(
    chalk.yellow(
      figlet.textSync(process.env['npm_package_name'], {
        horizontalLayout: 'fitted',
      }),
    ),
  );

  const { inputPath } = await inquirer.prompt<Record<string, string>>({
    name: 'inputPath',
    // eslint-disable-next-line @typescript-eslint/ban-ts-comment
    // @ts-ignore
    type: 'path',
    pageSize: '25',
    hideRoot: true,
    onlyShowValid: true,
    default: 'static/input.xlsx',
    message: 'Select a target input file:',
    validate: (value: string) => {
      return !(value.includes('/.') || value.includes('node_modules'));
    },
  });

  if (!path.parse(inputPath).ext) return bootstrap();

  const { prefix } = await inquirer.prompt<Record<string, string>>({
    name: 'prefix',
    type: 'input',
    default: '-output',
    message: 'Input prefix for transformed file:',
  });

  const { root, dir, name, ext } = path.parse(inputPath);
  const filename = path.format({ root, dir, name: name + prefix, ext });

  const { WorkbookReader, WorkbookWriter } = stream.xlsx;
  const workbookWriter = new WorkbookWriter({ filename });
  const workbookReader = new WorkbookReader(inputPath, {
    worksheets: 'emit',
  });

  for await (const worksheetReader of workbookReader) {
    // eslint-disable-next-line @typescript-eslint/ban-ts-comment
    // @ts-ignore
    const { model, name } = worksheetReader;
    const { transform } = await inquirer.prompt({
      name: 'transform',
      type: 'confirm',
      message: `Transform Worksheet - ${name} ?`,
    });

    if (!transform) continue;

    const { headersRow } = await inquirer.prompt({
      name: 'headersRow',
      type: 'number',
      default: 1,
      message: `Input headers row:`,
      validate: (value: number) => value > 0,
    });

    const { separator } = await inquirer.prompt({
      name: 'separator',
      type: 'input',
      default: ';',
      message: `Input data separator:`,
    });

    const worksheet = workbookWriter.addWorksheet();
    worksheet.model = model;
    worksheet.name = name;

    let targetCell = 0;
    for await (const row of worksheetReader) {
      if (!(row.values instanceof Array)) continue;
      if (row.number === headersRow) {
        worksheet.addRow(row.values);
        const { data } = await inquirer.prompt<Record<string, string>>({
          name: 'data',
          // eslint-disable-next-line @typescript-eslint/ban-ts-comment
          // @ts-ignore
          type: 'list',
          message: `Select a target column:`,
          choices: row.values.filter((value) => value),
        });
        targetCell = row.values.indexOf(data);
        continue;
      }

      const cell = row.getCell(targetCell);
      if (cell.value === null) {
        worksheet.addRow(row.values);
        logger(worksheet.name, `Copied row ${row.number}`);
        continue;
      }

      const values = cell.value.toString().split(separator);
      if (values.length <= 1) {
        worksheet.addRow(row.values);
        logger(worksheet.name, `Copied row ${row.number}`);
        continue;
      }

      for (const value of values) {
        cell.value = value;
        worksheet.addRow(row.values);
        logger(worksheet.name, `Created row from row ${row.number}`);
      }
    }
  }

  await workbookWriter.commit();
  logger('Finished writing to file', filename);
}

bootstrap();
