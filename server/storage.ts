import { db } from "./db";
import { reports, executiveSummaries, users } from "@shared/schema";
import type { User, InsertUser, Report, InsertReport, ExecutiveSummary, InsertExecutiveSummary } from "@shared/schema";
import { eq, desc } from "drizzle-orm";

export interface IStorage {
  getUser(id: string): Promise<User | undefined>;
  getUserByUsername(username: string): Promise<User | undefined>;
  createUser(user: InsertUser): Promise<User>;

  getReports(): Promise<Report[]>;
  getReport(id: number): Promise<Report | undefined>;
  createReport(report: InsertReport): Promise<Report>;
  deleteReport(id: number): Promise<void>;

  getExecutiveSummary(reportId: number): Promise<ExecutiveSummary | undefined>;
  createExecutiveSummary(summary: InsertExecutiveSummary): Promise<ExecutiveSummary>;
}

export class DatabaseStorage implements IStorage {
  async getUser(id: string): Promise<User | undefined> {
    const [user] = await db.select().from(users).where(eq(users.id, id));
    return user;
  }

  async getUserByUsername(username: string): Promise<User | undefined> {
    const [user] = await db.select().from(users).where(eq(users.username, username));
    return user;
  }

  async createUser(insertUser: InsertUser): Promise<User> {
    const [user] = await db.insert(users).values(insertUser).returning();
    return user;
  }

  async getReports(): Promise<Report[]> {
    return db.select().from(reports).orderBy(desc(reports.createdAt));
  }

  async getReport(id: number): Promise<Report | undefined> {
    const [report] = await db.select().from(reports).where(eq(reports.id, id));
    return report;
  }

  async createReport(report: InsertReport): Promise<Report> {
    const [created] = await db.insert(reports).values(report).returning();
    return created;
  }

  async deleteReport(id: number): Promise<void> {
    await db.delete(executiveSummaries).where(eq(executiveSummaries.reportId, id));
    await db.delete(reports).where(eq(reports.id, id));
  }

  async getExecutiveSummary(reportId: number): Promise<ExecutiveSummary | undefined> {
    const [summary] = await db.select().from(executiveSummaries).where(eq(executiveSummaries.reportId, reportId));
    return summary;
  }

  async createExecutiveSummary(summary: InsertExecutiveSummary): Promise<ExecutiveSummary> {
    const [created] = await db.insert(executiveSummaries).values(summary).returning();
    return created;
  }
}

export const storage = new DatabaseStorage();
