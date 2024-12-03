import { Workbook, Column } from "exceljs";
import dayjs from "dayjs";
import path from "path";
import fs from "fs";

interface TurnAroundTime {
  hr: number;
  min: number;
  identifier: string;
  hr_to_sec: number;
  min_to_sec: number;
  total_sec: number;
  total_minutes: number;
  grouping: "A" | "B" | "C";
}

const columns: Partial<Column>[] = [
  { header: "Identifier", key: "identifier", width: 30 },
  { header: "Hour", key: "hr", width: 15 },
  { header: "Minutes", key: "min", width: 15 },
  { header: "Hour to Sec", key: "hr_to_sec", width: 20 },
  { header: "Minutes to Sec", key: "min_to_sec", width: 20 },
  { header: "Total Sec", key: "total_sec", width: 20 },
  { header: "Total Minutes", key: "total_minutes", width: 20 },
  { header: "Grouping", key: "grouping", width: 10 },
];

const processFile = async (filepath: string): Promise<TurnAroundTime[]> => {
  const workbook = new Workbook();
  const turnAroundTime: TurnAroundTime[] = [];

  try {
    const sheet = await workbook.csv.readFile(filepath);

    sheet.eachRow((row, rowIndex) => {
      if (rowIndex === 1) return; // Skip the header row

      const createdAt = dayjs(row.getCell("AA").value as string);
      const updatedAt = dayjs(row.getCell("X").value as string);

      if (!createdAt.isValid() || !updatedAt.isValid()) {
        console.warn(`Invalid date at row ${rowIndex}`);
        return;
      }

      const total_minutes = updatedAt.diff(createdAt, "minutes");
      const hr = Math.floor(total_minutes / 60);
      const min = total_minutes % 60;
      const identifier = `${hr} hours, ${min} minutes`;
      const hr_to_sec = hr * 3600;
      const min_to_sec = min * 60;
      const total_sec = total_minutes * 60;

      const grouping: TurnAroundTime["grouping"] =
        total_minutes <= 120 ? "A" : total_minutes <= 240 ? "B" : "C";

      turnAroundTime.push({
        hr,
        min,
        identifier,
        hr_to_sec,
        min_to_sec,
        total_sec,
        total_minutes,
        grouping,
      });
    });
  } catch (error) {
    console.error(`Error processing file ${filepath}:`, error);
  }

  return turnAroundTime;
};

const runner = async () => {
  const inputDir = path.resolve(__dirname, "../assets");
  const outputDir = path.resolve(__dirname, "../exports");

  // Ensure output directory exists
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir);
  }

  const files = fs
    .readdirSync(inputDir)
    .filter((file) => file.endsWith(".csv"));
  if (files.length === 0) {
    console.warn("No CSV files found in the input directory.");
    return;
  }

  const workbook = new Workbook();
  const timestamp = dayjs().format("YYYY-MM-DD_HH-mm-ss");
  const outputFileName = `processed_${timestamp}.xlsx`;

  for (const file of files) {
    console.log(`Processing file: ${file}`);
    const filepath = path.join(inputDir, file);
    const turnAroundTimes = await processFile(filepath);

    if (turnAroundTimes.length > 0) {
      const sheet = workbook.addWorksheet(file);
      sheet.columns = columns;
      sheet.addRows(turnAroundTimes);
    } else {
      console.warn(`No valid data found in file: ${file}`);
    }
  }

  const outputPath = path.join(outputDir, outputFileName);
  await workbook.xlsx.writeFile(outputPath);

  console.log(`Processing complete. Output saved to: ${outputPath}`);
};

runner();
