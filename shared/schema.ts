import { sql } from "drizzle-orm";
import { pgTable, text, varchar, serial, integer, real, timestamp, jsonb } from "drizzle-orm/pg-core";
import { createInsertSchema } from "drizzle-zod";
import { z } from "zod";

export const users = pgTable("users", {
  id: varchar("id").primaryKey().default(sql`gen_random_uuid()`),
  username: text("username").notNull().unique(),
  password: text("password").notNull(),
});

export const insertUserSchema = createInsertSchema(users).pick({
  username: true,
  password: true,
});

export type InsertUser = z.infer<typeof insertUserSchema>;
export type User = typeof users.$inferSelect;

export const reports = pgTable("reports", {
  id: serial("id").primaryKey(),
  name: text("name").notNull(),
  strategy: text("strategy").notNull().default("current"),
  commitment: text("commitment").notNull().default("monthly"),
  userData: jsonb("user_data").notNull(),
  customRules: jsonb("custom_rules"),
  createdAt: timestamp("created_at").default(sql`CURRENT_TIMESTAMP`).notNull(),
});

export const insertReportSchema = createInsertSchema(reports).omit({
  id: true,
  createdAt: true,
});

export type Report = typeof reports.$inferSelect;
export type InsertReport = z.infer<typeof insertReportSchema>;

export const executiveSummaries = pgTable("executive_summaries", {
  id: serial("id").primaryKey(),
  reportId: integer("report_id").notNull().references(() => reports.id, { onDelete: "cascade" }),
  content: text("content").notNull(),
  costCurrent: real("cost_current").notNull(),
  costSecurity: real("cost_security").notNull(),
  costSaving: real("cost_saving").notNull(),
  costBalanced: real("cost_balanced").notNull(),
  costCustom: real("cost_custom"),
  commitment: text("commitment").notNull(),
  createdAt: timestamp("created_at").default(sql`CURRENT_TIMESTAMP`).notNull(),
});

export const insertExecutiveSummarySchema = createInsertSchema(executiveSummaries).omit({
  id: true,
  createdAt: true,
});

export type ExecutiveSummary = typeof executiveSummaries.$inferSelect;
export type InsertExecutiveSummary = z.infer<typeof insertExecutiveSummarySchema>;

export const microsoftTokens = pgTable("microsoft_tokens", {
  id: serial("id").primaryKey(),
  sessionId: text("session_id").notNull().unique(),
  accessToken: text("access_token").notNull(),
  refreshToken: text("refresh_token"),
  expiresAt: timestamp("expires_at").notNull(),
  tenantId: text("tenant_id"),
  userEmail: text("user_email"),
  userName: text("user_name"),
  createdAt: timestamp("created_at").default(sql`CURRENT_TIMESTAMP`).notNull(),
});

export const insertMicrosoftTokenSchema = createInsertSchema(microsoftTokens).omit({
  id: true,
  createdAt: true,
});

export type MicrosoftToken = typeof microsoftTokens.$inferSelect;
export type InsertMicrosoftToken = z.infer<typeof insertMicrosoftTokenSchema>;

export const loginHistory = pgTable("login_history", {
  id: serial("id").primaryKey(),
  userEmail: text("user_email").notNull(),
  userName: text("user_name"),
  tenantId: text("tenant_id"),
  loginAt: timestamp("login_at").default(sql`CURRENT_TIMESTAMP`).notNull(),
});

export const insertLoginHistorySchema = createInsertSchema(loginHistory).omit({
  id: true,
  loginAt: true,
});

export type LoginHistory = typeof loginHistory.$inferSelect;
export type InsertLoginHistory = z.infer<typeof insertLoginHistorySchema>;

export { pgTable };
