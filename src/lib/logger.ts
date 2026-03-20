import pino from "pino";

export const logger = pino({
  name: "teacheractive-exec-reporting",
  level: process.env.LOG_LEVEL ?? "info",
});
