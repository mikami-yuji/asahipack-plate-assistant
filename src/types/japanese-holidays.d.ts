declare module 'japanese-holidays' {
  export function isHoliday(date: Date): string | undefined;
  export function getHolidaysOf(year: number): Array<{ date: Date; name: string }>;
}
