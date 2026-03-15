import "dotenv/config";
import fs from "node:fs/promises";
import path from "node:path";
import xlsx from "xlsx";
import mongoose from "mongoose";
import { connectToMongoDB } from "../mongodb";
import { User } from "../models";

const DEFAULT_PASSWORD = process.env.SEED_DEFAULT_PASSWORD || "123";

async function run() {
  const connected = await connectToMongoDB();
  if (!connected) {
    console.error("[Export] MongoDB not connected. Check MONGODB_URI.");
    process.exit(1);
  }

  const users = await User.find({ role: "user" })
    .sort({ employeeId: 1 })
    .lean();

  const rows = [
    ["Employee ID", "Name", "Email", "Department", "Position", "Role", "Password"],
    ...users.map((u: any) => [
      u.employeeId || "",
      u.name || "",
      u.email || "",
      u.department || "",
      u.position || "",
      u.role || "user",
      DEFAULT_PASSWORD,
    ]),
  ];

  const workbook = xlsx.utils.book_new();
  const worksheet = xlsx.utils.aoa_to_sheet(rows);
  xlsx.utils.book_append_sheet(workbook, worksheet, "Credentials");

  const outDir = path.resolve("exports");
  await fs.mkdir(outDir, { recursive: true });

  const xlsxPath = path.join(outDir, "employee-credentials.xlsx");
  const csvPath = path.join(outDir, "employee-credentials.csv");

  xlsx.writeFile(workbook, xlsxPath);
  const csv = xlsx.utils.sheet_to_csv(worksheet);
  await fs.writeFile(csvPath, csv, "utf8");

  console.log(
    `[Export] Wrote ${users.length} rows to ${xlsxPath} and ${csvPath}`
  );

  await mongoose.connection.close();
}

run().catch((error) => {
  console.error("[Export] Failed:", error);
  process.exit(1);
});
