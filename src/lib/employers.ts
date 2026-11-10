// Built-in profiles of the candidate's past employers. Each profile says what
// the company actually does, which kinds of work are realistic there, and which
// industries must never appear in bullets under it. These override whatever the
// AI guesses, because a bullet that could not exist at the employer is the most
// visible mistake on a tailored resume.
//
// Sources: Northern Trust 2025 Form 10-K (Item 1, Business) — reporting segments
// Asset Servicing and Wealth Management, with Asset Management supporting both.

export interface EmployerProfile {
  /** Matches the company line of the template. */
  match: RegExp;
  name: string;
  /** One line on what the company does. */
  business: string;
  /** Kinds of work a bullet under this employer may describe. */
  allowedAreas: string[];
  /** Industries/settings that must never appear in a bullet under this employer. */
  forbidden: Array<[RegExp, string]>;
}

// Shared "wrong industry" patterns.
const HEALTHCARE: [RegExp, string] = [
  /\b(health ?care|healthcare|medical|clinical|clinicians?|patients?|hospitals?|physicians?|nurs(?:e|es|ing)|EHR|EMR|electronic health records?|pharmac(?:y|ies|eutical)|HIPAA|claims adjudication|care delivery)\b/i,
  "healthcare/medical",
];
const AUTOMOTIVE: [RegExp, string] = [
  /\b(automotive|vehicles?|car makers?|carmakers?|dealerships?|auto parts|powertrain|telematics|assembly lines?)\b/i,
  "automotive/manufacturing",
];
const RETAIL: [RegExp, string] = [
  /\b(retail|retailers?|in-store|storefronts?|shoppers?|e-?commerce|checkout|point[- ]of[- ]sale|POS terminals?|merchandis(?:e|ing)|grocery|groceries)\b/i,
  "retail/e-commerce",
];
const FOOD_DELIVERY: [RegExp, string] = [
  /\b(food delivery|restaurants?|delivery partners?|riders?|couriers?|dark stores?|quick commerce|menus?|order dispatch)\b/i,
  "food delivery",
];
const ENERGY: [RegExp, string] = [
  /\b(batter(?:y|ies)|energy storage|power grid|grid operators?|substations?|utility-scale|oil and gas|refiner(?:y|ies))\b/i,
  "energy/utilities",
];
const NETWORK_HW: [RegExp, string] = [
  /\b(routers?|switches|IoT gateways?|network devices?|firmware|chipsets?)\b/i,
  "networking hardware",
];
const BANK_CUSTODY: [RegExp, string] = [
  /\b(custody|custodial|wealth management|trust accounts?|trust and estate|fund administration|securities lending|asset servicing|private banking|high-net-worth)\b/i,
  "custody/wealth management",
];

export const EMPLOYER_PROFILES: EmployerProfile[] = [
  {
    match: /\bnorthern\s+trust\b/i,
    name: "Northern Trust",
    business:
      "global custody bank and wealth manager — Asset Servicing for institutional investors (pension funds, fund managers, insurers, sovereign wealth funds) and Wealth Management for high-net-worth families, with Asset Management supporting both; supervised by the Federal Reserve",
    allowedAreas: [
      "banking and payment systems (wire transfers, cash management, treasury)",
      "financial operations and reconciliations (custody and cash reconciliations, breaks, NAV)",
      "internal audit and IT controls (SOX, ITGC, access reviews, change management)",
      "risk, regulatory compliance and cybersecurity (AML, sanctions, Volcker, GLBA, Federal Reserve exams)",
      "asset servicing, wealth management and investment operations (custody, fund administration, securities lending, trust and estate administration)",
      "transaction processing, settlements and financial reporting (trade settlement, corporate actions, regulatory and client reporting)",
    ],
    forbidden: [HEALTHCARE, AUTOMOTIVE, RETAIL, FOOD_DELIVERY, ENERGY, NETWORK_HW],
  },
  {
    match: /\bswiggy\b/i,
    name: "Swiggy",
    business:
      "Indian food-delivery and quick-commerce platform (Instamart) connecting customers, restaurant partners and delivery partners, with in-app payments",
    allowedAreas: [
      "order, payment and refund processing",
      "restaurant and delivery-partner payouts and settlements",
      "finance operations, reconciliations and reporting",
      "fraud and risk controls on orders, coupons and payments",
      "platform engineering, data pipelines and internal tools",
      "internal audit, IT controls and compliance for the platform",
    ],
    forbidden: [HEALTHCARE, AUTOMOTIVE, ENERGY, BANK_CUSTODY, NETWORK_HW],
  },
  {
    match: /\bcisco\b/i,
    name: "Cisco",
    business:
      "networking, security, collaboration (Webex) and observability (Splunk) technology company selling to enterprises and service providers",
    allowedAreas: [
      "networking software and device platforms",
      "security products and internal security operations",
      "collaboration and cloud platforms",
      "observability, data and AI platforms",
      "internal IT, finance systems, audit and compliance",
    ],
    forbidden: [HEALTHCARE, AUTOMOTIVE, RETAIL, FOOD_DELIVERY, ENERGY, BANK_CUSTODY],
  },
];

export function findEmployerProfile(company: string): EmployerProfile | null {
  return EMPLOYER_PROFILES.find((p) => p.match.test(company)) ?? null;
}

/** One-line context used in prompts and rewrite messages. */
export function profileContext(p: EmployerProfile): string {
  return `${p.name}: ${p.business}. Realistic work: ${p.allowedAreas.join("; ")}`;
}

/** Prompt block with hard rules for every past employer that has a built-in profile. */
export function employerBlock(profiles: Array<EmployerProfile | null>): string {
  const lines = profiles
    .map((p, k) =>
      p
        ? `- COMPANY ${k + 1} = ${p.name}: ${p.business}. EVERY bullet under it must be realistic work that ${p.name} really does, drawn from: ${p.allowedAreas.join("; ")}. Tailor the JD's skills INTO these areas. NEVER describe ${p.forbidden.map(([, label]) => label).join(", ")} work there.`
        : "",
    )
    .filter(Boolean);
  return lines.length ? "\n\nEMPLOYER RULES (checked in code — a bullet that breaks them is rewritten):\n" + lines.join("\n") + "\n" : "";
}
