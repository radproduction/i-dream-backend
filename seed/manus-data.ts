export type ManusUserSeed = {
  openId: string;
  name: string;
  email: string;
  loginMethod: "custom";
  role: "user" | "admin";
  employeeId: string;
  passwordHash: string;
  department: string;
  position: string;
};

export type ManusProjectSeed = {
  name: string;
  description: string;
  status: "active" | "on_hold" | "completed" | "cancelled";
  priority: "low" | "medium" | "high";
  createdByOpenId: string;
  source?: "team_lead" | "employee";
};

export type ManusProjectAssignmentSeed = {
  projectName: string;
  userOpenId: string;
  role: string;
};

// Seed users captured from Manus DB logs (.manus/db/*.json).
// Password hash corresponds to the "123" password in the original seed.
const SEED_PASSWORD_HASH =
  "$2b$10$WGscasGNs.ojsYycPzM9IuXiRySSrrGB26fDKuTUnRYpz9Tlrci6.";

export const manusUsers: ManusUserSeed[] = [
  {
    openId: "admin-aamir-001",
    name: "Aamir",
    email: "aamir@rad.com",
    loginMethod: "custom",
    role: "admin",
    employeeId: "ADMIN001",
    passwordHash: SEED_PASSWORD_HASH,
    department: "Management",
    position: "Founder & COO",
  },
  {
    openId: "hrms-hassan",
    name: "Hassan",
    email: "hassan@rad.com",
    loginMethod: "custom",
    role: "user",
    employeeId: "EMP001",
    passwordHash: SEED_PASSWORD_HASH,
    department: "Engineering",
    position: "Software Developer",
  },
  {
    openId: "hrms-talha",
    name: "Talha",
    email: "talha@rad.com",
    loginMethod: "custom",
    role: "user",
    employeeId: "EMP002",
    passwordHash: SEED_PASSWORD_HASH,
    department: "Engineering",
    position: "Software Developer",
  },
];

export const manusProjects: ManusProjectSeed[] = [
  {
    name: "IDream Entertainment HRMS Development",
    description:
      "Complete HRMS portal with time tracking, leave management, and project management",
    status: "active",
    priority: "high",
    createdByOpenId: "hrms-hassan",
    source: "team_lead",
  },
  {
    name: "Website Redesign Project",
    description: "Redesign company website with modern UI/UX",
    status: "active",
    priority: "medium",
    createdByOpenId: "hrms-hassan",
    source: "team_lead",
  },
  {
    name: "Mobile App Development",
    description: "Develop mobile application for employee portal",
    status: "active",
    priority: "high",
    createdByOpenId: "hrms-hassan",
    source: "team_lead",
  },
];

export const manusProjectAssignments: ManusProjectAssignmentSeed[] = [
  {
    projectName: "IDream Entertainment HRMS Development",
    userOpenId: "hrms-hassan",
    role: "Lead Developer",
  },
  {
    projectName: "IDream Entertainment HRMS Development",
    userOpenId: "hrms-talha",
    role: "Frontend Developer",
  },
  {
    projectName: "Website Redesign Project",
    userOpenId: "hrms-hassan",
    role: "UI/UX Designer",
  },
  {
    projectName: "Website Redesign Project",
    userOpenId: "hrms-talha",
    role: "Developer",
  },
  {
    projectName: "Mobile App Development",
    userOpenId: "hrms-hassan",
    role: "Backend Developer",
  },
  {
    projectName: "Mobile App Development",
    userOpenId: "hrms-talha",
    role: "Mobile Developer",
  },
];

