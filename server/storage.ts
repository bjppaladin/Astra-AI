import { db } from "./db";
import { reports, executiveSummaries, loginHistory } from "@shared/schema";
import type { Report, InsertReport, ExecutiveSummary, InsertExecutiveSummary, InsertLoginHistory, LoginHistory } from "@shared/schema";
import { eq, desc } from "drizzle-orm";

export interface IStorage {
  getReports(): Promise<Report[]>;
  getReport(id: number): Promise<Report | undefined>;
  createReport(report: InsertReport): Promise<Report>;
  deleteReport(id: number): Promise<void>;

  getExecutiveSummary(reportId: number): Promise<ExecutiveSummary | undefined>;
  createExecutiveSummary(summary: InsertExecutiveSummary): Promise<ExecutiveSummary>;

  recordLogin(entry: InsertLoginHistory): Promise<LoginHistory>;
  getLoginHistory(userEmail: string): Promise<LoginHistory[]>;
  getLoginCount(userEmail: string): Promise<number>;
}

export class DatabaseStorage implements IStorage {
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

  async recordLogin(entry: InsertLoginHistory): Promise<LoginHistory> {
    const [created] = await db.insert(loginHistory).values(entry).returning();
    return created;
  }

  async getLoginHistory(userEmail: string): Promise<LoginHistory[]> {
    return db.select().from(loginHistory).where(eq(loginHistory.userEmail, userEmail)).orderBy(desc(loginHistory.loginAt));
  }

  async getLoginCount(userEmail: string): Promise<number> {
    const rows = await db.select().from(loginHistory).where(eq(loginHistory.userEmail, userEmail));
    return rows.length;
  }
}

export const storage = new DatabaseStorage();
