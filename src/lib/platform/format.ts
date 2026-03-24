const platformDateTimeFormatter = new Intl.DateTimeFormat("en-GB", {
  day: "2-digit",
  month: "short",
  year: "numeric",
  hour: "2-digit",
  minute: "2-digit",
  hour12: false,
  timeZone: "Europe/London",
});

export function formatPlatformDateTime(value: string | null | undefined): string {
  if (!value) {
    return "Not published";
  }

  return platformDateTimeFormatter.format(new Date(value));
}
