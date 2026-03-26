import "dotenv/config";
import fs from "node:fs/promises";
import path from "node:path";
import xlsx from "xlsx";
import mongoose from "mongoose";
import { connectToMongoDB } from "../mongodb";
import { EmployeeProfile, User } from "../models";

const DEFAULT_PASSWORD = process.env.SEED_DEFAULT_PASSWORD || "123";

async function run() {
  const connected = await connectToMongoDB();
  if (!connected) {
    console.error("[Export] MongoDB not connected. Check MONGODB_URI.");
    process.exit(1);
  }

  const admins = await User.find({ role: "admin" })
    .sort({ employeeId: 1 })
    .lean();

  const adminIds = admins.map((u: any) => u._id);
  const profiles = await EmployeeProfile.find({ userId: { $in: adminIds } }).lean();
  const addressMap = new Map(
    profiles.map((p: any) => [
      String(p.userId),
      p.currentAddress || p.permanentAddress || "",
    ])
  );

  const rows = [
    ["Employee ID", "Name", "Email", "Address", "Password"],
    ...admins.map((u: any) => [
      u.employeeId || "",
      u.name || "",
      u.email || "",
      addressMap.get(String(u._id)) || "",
      DEFAULT_PASSWORD,
    ]),
  ];

  const workbook = xlsx.utils.book_new();
  const worksheet = xlsx.utils.aoa_to_sheet(rows);
  xlsx.utils.book_append_sheet(workbook, worksheet, "Admins");

  const outDir = path.resolve("exports");
  await fs.mkdir(outDir, { recursive: true });

  const xlsxPath = path.join(outDir, "admin-credentials.xlsx");
  const csvPath = path.join(outDir, "admin-credentials.csv");

  xlsx.writeFile(workbook, xlsxPath);
  const csv = xlsx.utils.sheet_to_csv(worksheet);
  await fs.writeFile(csvPath, csv, "utf8");

  console.log(
    `[Export] Wrote ${admins.length} rows to ${xlsxPath} and ${csvPath}`
  );

  await mongoose.connection.close();
}

run().catch((error) => {
  console.error("[Export] Failed:", error);
  process.exit(1);
});
