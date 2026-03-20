export interface DefaultOfficeLocation {
  officeName: string;
  region: string;
  displayOrder: number;
  mapX: number;
  mapY: number;
}

export const DEFAULT_OFFICE_LOCATIONS: DefaultOfficeLocation[] = [
  { officeName: "Birmingham", region: "Midlands", displayOrder: 1, mapX: 238, mapY: 255 },
  { officeName: "Manchester", region: "North West", displayOrder: 2, mapX: 196, mapY: 195 },
  { officeName: "Leeds", region: "Yorkshire", displayOrder: 3, mapX: 228, mapY: 185 },
  { officeName: "London City", region: "London", displayOrder: 4, mapX: 298, mapY: 355 },
  { officeName: "London East", region: "London", displayOrder: 5, mapX: 308, mapY: 360 },
  { officeName: "Bristol", region: "South West", displayOrder: 6, mapX: 148, mapY: 375 },
  { officeName: "Sheffield", region: "Yorkshire", displayOrder: 7, mapX: 228, mapY: 198 },
  { officeName: "Nottingham", region: "East Midlands", displayOrder: 8, mapX: 242, mapY: 232 },
  { officeName: "Leicester", region: "East Midlands", displayOrder: 9, mapX: 244, mapY: 248 },
  { officeName: "Cardiff", region: "Wales", displayOrder: 10, mapX: 128, mapY: 378 },
  { officeName: "Swansea", region: "Wales", displayOrder: 11, mapX: 102, mapY: 370 },
  { officeName: "Liverpool", region: "North West", displayOrder: 12, mapX: 182, mapY: 208 },
  { officeName: "Newcastle", region: "North East", displayOrder: 13, mapX: 228, mapY: 108 },
  { officeName: "Derby", region: "East Midlands", displayOrder: 14, mapX: 238, mapY: 238 },
  { officeName: "Coventry", region: "Midlands", displayOrder: 15, mapX: 248, mapY: 262 },
  { officeName: "Worcester", region: "Midlands", displayOrder: 16, mapX: 218, mapY: 272 },
  { officeName: "Oxford", region: "South East", displayOrder: 17, mapX: 248, mapY: 308 },
  { officeName: "Cambridge", region: "East", displayOrder: 18, mapX: 292, mapY: 295 },
  { officeName: "Southampton", region: "South", displayOrder: 19, mapX: 228, mapY: 398 },
  { officeName: "Plymouth", region: "South West", displayOrder: 20, mapX: 105, mapY: 455 },
  { officeName: "Norwich", region: "East", displayOrder: 21, mapX: 328, mapY: 295 },
  { officeName: "Exeter", region: "South West", displayOrder: 22, mapX: 128, mapY: 448 },
];
