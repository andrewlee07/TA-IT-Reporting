export var MAP_VIEWBOX = {
  width: 675,
  height: 1180,
};

export var OFFICE_MAP_REGISTRY = [
  { name: "Birmingham", x: 448, y: 792 },
  { name: "Manchester", x: 433, y: 680 },
  { name: "Leeds", x: 468, y: 645 },
  { name: "London City", x: 529, y: 905 },
  { name: "London East", x: 543, y: 901 },
  { name: "Bristol", x: 419, y: 912 },
  { name: "Sheffield", x: 470, y: 689 },
  { name: "Nottingham", x: 483, y: 736 },
  { name: "Leicester", x: 485, y: 771 },
  { name: "Cardiff", x: 394, y: 905 },
  { name: "Swansea", x: 362, y: 901 },
  { name: "Liverpool", x: 404, y: 685 },
  { name: "Newcastle", x: 486, y: 577 },
  { name: "Derby", x: 469, y: 741 },
  { name: "Coventry", x: 465, y: 798 },
  { name: "Worcester", x: 438, y: 824 },
  { name: "Oxford", x: 481, y: 871 },
  { name: "Cambridge", x: 541, y: 825 },
  { name: "Southampton", x: 472, y: 969 },
  { name: "Plymouth", x: 351, y: 1030 },
  { name: "Norwich", x: 586, y: 778 },
  { name: "Exeter", x: 380, y: 990 },
];

function normalizeOfficeName(value) {
  return String(value)
    .trim()
    .toLowerCase()
    .replace(/&/g, "and")
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .replace(/-{2,}/g, "-");
}

var OFFICE_LOOKUP = OFFICE_MAP_REGISTRY.reduce(function buildLookup(map, office) {
  map[office.name] = office;
  map[normalizeOfficeName(office.name)] = office;

  if (Array.isArray(office.aliases)) {
    office.aliases.forEach(function addAlias(alias) {
      map[alias] = office;
      map[normalizeOfficeName(alias)] = office;
    });
  }

  return map;
}, Object.create(null));

export function resolveOfficeMapPoint(officeName) {
  if (!officeName) {
    return null;
  }

  return OFFICE_LOOKUP[officeName] || OFFICE_LOOKUP[normalizeOfficeName(officeName)] || null;
}
