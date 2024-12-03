import { Workbook, Column } from "exceljs";
import dayjs from "dayjs";
import path from "path";
import fs from "fs";

interface TurnAroundTime {
  grouping: "A" | "B" | "C";
  median_total_minutes: number;
  median_total_seconds: number;
  count: number; // Number of items in this group
}

const columns: Partial<Column>[] = [
  { header: "Grouping", key: "grouping", width: 30 },
  { header: "Median Total Minutes", key: "median_total_minutes", width: 30 },
  { header: "Median Total Seconds", key: "median_total_seconds", width: 30 },
  { header: "Count", key: "count", width: 15 },
];

// Function to calculate the median of an array of numbers
const calculateMedian = (values: number[]): number => {
  if (!values.length) return 0;
  values.sort((a, b) => a - b); // Sort in ascending order
  const mid = Math.floor(values.length / 2);
  return values.length % 2 !== 0
    ? values[mid] // Odd length
    : (values[mid - 1] + values[mid]) / 2; // Even length
};

// Updated `getExcelD` function to compute median and count for each file by group
const getExcelD = async (filename: string): Promise<TurnAroundTime[]> => {
  let wb: Workbook = new Workbook();
  let datafile = path.join(__dirname, "../assets", filename);
  const groupData: { [key: string]: number[] } = { A: [], B: [], C: [] }; // Store total_minutes for each group

  console.log(`Processing file: ${datafile}`);

  await wb.csv.readFile(datafile).then((sh) => {
    for (let i = 2; i <= sh.actualRowCount; i++) {
      const row = sh.getRow(i);
      const createdAt = dayjs(row.getCell("AA").value as string);
      const updatedAt = dayjs(row.getCell("X").value as string);
      const total_minutes = updatedAt.diff(createdAt, "minutes");

      // Group data based on thresholds
      if (total_minutes <= 120) groupData.A.push(total_minutes);
      else if (total_minutes > 120 && total_minutes <= 240)
        groupData.B.push(total_minutes);
      else groupData.C.push(total_minutes);
    }
  });

  // Compute median and count for each group
  return Object.entries(groupData).map(([group, values]) => ({
    grouping: group as TurnAroundTime["grouping"],
    median_total_minutes: calculateMedian(values),
    median_total_seconds: calculateMedian(values) * 60,
    count: values.length, // Count of items in the group
  }));
};

const runner = async () => {
  console.log("Start...");
  const files = fs.readdirSync(path.join(__dirname, "../assets"));
  let wb: Workbook = new Workbook();

  const promises = files.map(async (file) => {
    const turnAroundTimes = await getExcelD(file);

    // Create a worksheet for each partner (file)
    const sh = wb.addWorksheet(file.replace(".csv", "")); // Use file name (without extension) as sheet name
    sh.columns = columns;
    sh.addRows(turnAroundTimes);
  });

  await Promise.all(promises);

  // Write the workbook with results by partner
  await wb.xlsx.writeFile(
    path.join(__dirname, "../exports", `partners_median_summary.xlsx`)
  );
  console.log("Median summary with counts saved by partner!");
};

runner();
