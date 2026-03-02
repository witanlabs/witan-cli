import { PDFDocument, StandardFonts, rgb } from 'pdf-lib';
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  AlignmentType,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle,
} from 'docx';
import PptxGenJS from 'pptxgenjs';
import { writeFile, mkdir } from 'fs/promises';
import { join } from 'path';
import { parseArgs } from 'util';

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

export const ACME_PDF_FILENAME = 'acme-annual-report-fy2025.pdf';
export const ACME_DOCX_FILENAME = 'acme-board-minutes-jan2026.docx';
export const ACME_PPTX_FILENAME = 'acme-investor-deck-q1-2026.pptx';

export const ACME_QUESTION =
  "Using all three documents, answer: What was Acme Corp's total FY2025 revenue, " +
  'what growth target did the board approve for FY2026, what is the projected Q1 FY2026 revenue, ' +
  'and how many new hires were authorized? Cite the specific document each fact comes from.';

// Shared financial data (interlocking numbers)
const REVENUE_Q1 = 10.8;
const REVENUE_Q2 = 11.4;
const REVENUE_Q3 = 12.3;
const REVENUE_Q4 = 13.7;
const TOTAL_REVENUE = 48.2; // Q1+Q2+Q3+Q4
const NET_INCOME = 9.7;
const HEADCOUNT = 312;
const GROWTH_TARGET_PCT = 15;
const NEW_HIRES = 45;
const PROJECTED_REVENUE_FY2026 = 55.4; // 48.2 * 1.15 ≈ 55.4
const PROJECTED_Q1_FY2026 = 12.8;
const ARR = 42.6;
const RESOLUTION_NUMBER = '2026-001';

// ---------------------------------------------------------------------------
// PDF — Annual Report (3 pages)
// ---------------------------------------------------------------------------

export async function createAcmePdf(filePath: string): Promise<void> {
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  const fontBold = await doc.embedFont(StandardFonts.HelveticaBold);
  const pageWidth = 612;
  const pageHeight = 792;
  const margin = 72;

  // Helper to add text
  const addText = (
    page: ReturnType<typeof doc.addPage>,
    text: string,
    x: number,
    y: number,
    size: number,
    fontFace = font,
    color = rgb(0, 0, 0),
  ) => {
    page.drawText(text, { x, y, size, font: fontFace, color });
  };

  // --- Page 1: Cover ---
  const p1 = doc.addPage([pageWidth, pageHeight]);
  addText(p1, 'ACME CORPORATION', margin, 600, 28, fontBold, rgb(0.1, 0.2, 0.5));
  addText(p1, 'Annual Report — Fiscal Year 2025', margin, 560, 18, font, rgb(0.3, 0.3, 0.3));
  addText(p1, 'Confidential — For Shareholder Distribution Only', margin, 520, 11, font, rgb(0.5, 0.5, 0.5));
  addText(p1, 'Prepared by the Office of the CFO', margin, 490, 12);
  addText(p1, 'January 2026', margin, 470, 12);

  // --- Page 2: Financial Summary ---
  const p2 = doc.addPage([pageWidth, pageHeight]);
  addText(p2, 'Financial Summary — FY2025', margin, 700, 20, fontBold);

  const summaryLines = [
    `Total Revenue: $${TOTAL_REVENUE} million`,
    `Net Income: $${NET_INCOME} million`,
    `Net Margin: ${((NET_INCOME / TOTAL_REVENUE) * 100).toFixed(1)}%`,
    `Full-Time Employees: ${HEADCOUNT}`,
    '',
    'Quarterly Revenue Breakdown:',
    `  Q1 FY2025: $${REVENUE_Q1} million`,
    `  Q2 FY2025: $${REVENUE_Q2} million`,
    `  Q3 FY2025: $${REVENUE_Q3} million`,
    `  Q4 FY2025: $${REVENUE_Q4} million`,
    '',
    `Year-over-year revenue grew 22% from $39.5 million in FY2024.`,
    `The company ended the year with $18.3 million in cash and equivalents.`,
  ];

  let y2 = 660;
  for (const line of summaryLines) {
    addText(p2, line, margin, y2, 12);
    y2 -= 20;
  }

  // --- Page 3: Business Segments ---
  const p3 = doc.addPage([pageWidth, pageHeight]);
  addText(p3, 'Business Segments', margin, 700, 20, fontBold);

  const segmentLines = [
    'Enterprise Software (62% of revenue)',
    `  Revenue: $${(TOTAL_REVENUE * 0.62).toFixed(1)} million`,
    '  Customers: 847 enterprise accounts',
    '',
    'Professional Services (25% of revenue)',
    `  Revenue: $${(TOTAL_REVENUE * 0.25).toFixed(1)} million`,
    '  Utilization rate: 78%',
    '',
    'Managed Services (13% of revenue)',
    `  Revenue: $${(TOTAL_REVENUE * 0.13).toFixed(1)} million`,
    '  MRR growth: 31% year-over-year',
    '',
    'Outlook:',
    `Management expects continued momentum in FY2026, targeting $${PROJECTED_REVENUE_FY2026} million`,
    `in total revenue, representing ${GROWTH_TARGET_PCT}% year-over-year growth.`,
  ];

  let y3 = 660;
  for (const line of segmentLines) {
    addText(p3, line, margin, y3, 12);
    y3 -= 20;
  }

  const pdfBytes = await doc.save();
  await writeFile(filePath, pdfBytes);
}

// ---------------------------------------------------------------------------
// DOCX — Board Minutes
// ---------------------------------------------------------------------------

export async function createAcmeDocx(filePath: string): Promise<void> {
  const doc = new Document({
    sections: [
      {
        children: [
          // Title
          new Paragraph({
            children: [
              new TextRun({
                text: 'ACME CORPORATION',
                bold: true,
                size: 36,
                color: '1A3366',
              }),
            ],
            alignment: AlignmentType.CENTER,
            spacing: { after: 100 },
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: 'Minutes of the Board of Directors Meeting',
                bold: true,
                size: 28,
              }),
            ],
            alignment: AlignmentType.CENTER,
            spacing: { after: 100 },
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: 'January 15, 2026 — 10:00 AM EST',
                size: 22,
                color: '666666',
              }),
            ],
            alignment: AlignmentType.CENTER,
            spacing: { after: 400 },
          }),

          // Attendees
          new Paragraph({
            text: '1. Attendees',
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 200, after: 100 },
          }),
          new Paragraph({
            children: [
              new TextRun('Present: J. Smith (Chair), M. Johnson (CEO), R. Williams (CFO), '),
              new TextRun('L. Brown (COO), K. Davis (CTO), P. Martinez (General Counsel), '),
              new TextRun('A. Thompson (Independent Director), S. Garcia (Independent Director)'),
            ],
            spacing: { after: 100 },
          }),
          new Paragraph({
            children: [new TextRun('Quorum established with 8 of 8 directors present.')],
            spacing: { after: 200 },
          }),

          // FY2025 Review
          new Paragraph({
            text: '2. FY2025 Financial Review',
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 200, after: 100 },
          }),
          new Paragraph({
            children: [
              new TextRun(
                `CFO Williams presented the FY2025 financial results. Total revenue reached $${TOTAL_REVENUE} million, `,
              ),
              new TextRun(
                `representing 22% year-over-year growth. Net income was $${NET_INCOME} million. `,
              ),
              new TextRun(
                `The company ended the fiscal year with ${HEADCOUNT} full-time employees.`,
              ),
            ],
            spacing: { after: 200 },
          }),

          // FY2026 Growth Target
          new Paragraph({
            text: '3. FY2026 Strategic Plan and Growth Target',
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 200, after: 100 },
          }),
          new Paragraph({
            children: [
              new TextRun(
                `CEO Johnson outlined the FY2026 strategic plan, proposing a ${GROWTH_TARGET_PCT}% revenue growth target. `,
              ),
              new TextRun(
                `This would bring projected FY2026 revenue to approximately $${PROJECTED_REVENUE_FY2026} million. `,
              ),
              new TextRun(
                'The plan includes expansion into the APAC market and launch of the Enterprise Platform v3.0.',
              ),
            ],
            spacing: { after: 200 },
          }),

          // Hiring Plan
          new Paragraph({
            text: '4. Hiring Plan',
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 200, after: 100 },
          }),
          new Paragraph({
            children: [
              new TextRun(
                `COO Brown presented the FY2026 hiring plan requesting approval for ${NEW_HIRES} new positions: `,
              ),
              new TextRun('20 in Engineering, 12 in Sales, 8 in Customer Success, and 5 in Operations. '),
              new TextRun(
                `This would bring total headcount from ${HEADCOUNT} to ${HEADCOUNT + NEW_HIRES} by end of FY2026.`,
              ),
            ],
            spacing: { after: 200 },
          }),

          // Resolution
          new Paragraph({
            text: '5. Resolutions',
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 200, after: 100 },
          }),
          new Paragraph({
            children: [
              new TextRun({ text: `Resolution ${RESOLUTION_NUMBER}: `, bold: true }),
              new TextRun(
                `Approval of FY2026 Strategic Plan including ${GROWTH_TARGET_PCT}% revenue growth target `,
              ),
              new TextRun(
                `and authorization of ${NEW_HIRES} new hires as presented by the COO.`,
              ),
            ],
            spacing: { after: 100 },
          }),

          // Vote table
          new Table({
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: 'Motion', bold: true })] })],
                    width: { size: 25, type: WidthType.PERCENTAGE },
                  }),
                  new TableCell({
                    children: [
                      new Paragraph({
                        children: [
                          new TextRun(
                            `Approve FY2026 strategic plan with ${GROWTH_TARGET_PCT}% growth target and ${NEW_HIRES} new hires`,
                          ),
                        ],
                      }),
                    ],
                    width: { size: 75, type: WidthType.PERCENTAGE },
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: 'Moved by', bold: true })] })],
                  }),
                  new TableCell({
                    children: [new Paragraph('A. Thompson')],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: 'Seconded by', bold: true })] })],
                  }),
                  new TableCell({
                    children: [new Paragraph('S. Garcia')],
                  }),
                ],
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: 'Vote', bold: true })] })],
                  }),
                  new TableCell({
                    children: [
                      new Paragraph({
                        children: [
                          new TextRun({ text: 'Approved unanimously', bold: true }),
                          new TextRun(' (8-0)'),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
            ],
            width: { size: 100, type: WidthType.PERCENTAGE },
          }),

          // Adjournment
          new Paragraph({
            text: '6. Adjournment',
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 300, after: 100 },
          }),
          new Paragraph({
            children: [
              new TextRun('Meeting adjourned at 12:45 PM EST. Next meeting scheduled for April 16, 2026.'),
            ],
            spacing: { after: 200 },
          }),
          new Paragraph({
            children: [
              new TextRun({ text: 'Respectfully submitted,', italics: true }),
            ],
            spacing: { after: 50 },
          }),
          new Paragraph({
            children: [new TextRun('P. Martinez, General Counsel & Corporate Secretary')],
          }),
        ],
      },
    ],
  });

  const buffer = await Packer.toBuffer(doc);
  await writeFile(filePath, buffer);
}

// ---------------------------------------------------------------------------
// PPTX — Investor Deck (4 slides)
// ---------------------------------------------------------------------------

export async function createAcmePptx(filePath: string): Promise<void> {
  const pptx = new PptxGenJS();
  pptx.author = 'Acme Corp IR Team';
  pptx.title = 'Acme Corp — Q1 FY2026 Investor Update';
  pptx.subject = 'Quarterly investor presentation';

  const DARK_BLUE = '1A3366';
  const MEDIUM_BLUE = '2D5AA0';
  const LIGHT_GRAY = 'F5F5F5';

  // --- Slide 1: Title ---
  const s1 = pptx.addSlide();
  s1.background = { color: DARK_BLUE };
  s1.addText('ACME CORPORATION', {
    x: 0.5,
    y: 1.5,
    w: 9,
    h: 1,
    fontSize: 36,
    bold: true,
    color: 'FFFFFF',
  });
  s1.addText('Q1 FY2026 Investor Update', {
    x: 0.5,
    y: 2.5,
    w: 9,
    h: 0.8,
    fontSize: 24,
    color: 'CCCCCC',
  });
  s1.addText('Presented to Institutional Investors — February 2026', {
    x: 0.5,
    y: 4.5,
    w: 9,
    h: 0.5,
    fontSize: 14,
    color: '999999',
  });

  // --- Slide 2: FY2025 Quarterly Breakdown ---
  const s2 = pptx.addSlide();
  s2.addText('FY2025 Revenue by Quarter', {
    x: 0.5,
    y: 0.3,
    w: 9,
    h: 0.6,
    fontSize: 24,
    bold: true,
    color: DARK_BLUE,
  });

  s2.addTable(
    [
      [
        { text: 'Quarter', options: { bold: true, color: 'FFFFFF', fill: { color: MEDIUM_BLUE } } },
        { text: 'Revenue ($M)', options: { bold: true, color: 'FFFFFF', fill: { color: MEDIUM_BLUE } } },
        { text: '% of Total', options: { bold: true, color: 'FFFFFF', fill: { color: MEDIUM_BLUE } } },
      ],
      [
        { text: 'Q1 FY2025' },
        { text: `$${REVENUE_Q1.toFixed(1)}` },
        { text: `${((REVENUE_Q1 / TOTAL_REVENUE) * 100).toFixed(1)}%` },
      ],
      [
        { text: 'Q2 FY2025' },
        { text: `$${REVENUE_Q2.toFixed(1)}` },
        { text: `${((REVENUE_Q2 / TOTAL_REVENUE) * 100).toFixed(1)}%` },
      ],
      [
        { text: 'Q3 FY2025' },
        { text: `$${REVENUE_Q3.toFixed(1)}` },
        { text: `${((REVENUE_Q3 / TOTAL_REVENUE) * 100).toFixed(1)}%` },
      ],
      [
        { text: 'Q4 FY2025', options: { bold: true } },
        { text: `$${REVENUE_Q4.toFixed(1)}`, options: { bold: true } },
        {
          text: `${((REVENUE_Q4 / TOTAL_REVENUE) * 100).toFixed(1)}%`,
          options: { bold: true },
        },
      ],
      [
        { text: 'Full Year', options: { bold: true, fill: { color: LIGHT_GRAY } } },
        { text: `$${TOTAL_REVENUE.toFixed(1)}`, options: { bold: true, fill: { color: LIGHT_GRAY } } },
        { text: '100.0%', options: { bold: true, fill: { color: LIGHT_GRAY } } },
      ],
    ],
    {
      x: 0.5,
      y: 1.2,
      w: 9,
      fontSize: 14,
      border: { type: 'solid', pt: 0.5, color: 'CCCCCC' },
      colW: [3, 3, 3],
      rowH: [0.5, 0.45, 0.45, 0.45, 0.45, 0.5],
    },
  );

  s2.addText(`Annual Recurring Revenue (ARR): $${ARR} million`, {
    x: 0.5,
    y: 4.5,
    w: 9,
    h: 0.5,
    fontSize: 16,
    bold: true,
    color: MEDIUM_BLUE,
  });

  // --- Slide 3: FY2026 Projections ---
  const s3 = pptx.addSlide();
  s3.addText('FY2026 Revenue Projections', {
    x: 0.5,
    y: 0.3,
    w: 9,
    h: 0.6,
    fontSize: 24,
    bold: true,
    color: DARK_BLUE,
  });

  const projectionLines = [
    `Board-approved growth target: ${GROWTH_TARGET_PCT}% year-over-year`,
    `FY2026 revenue target: $${PROJECTED_REVENUE_FY2026} million`,
    '',
    `Q1 FY2026 projected revenue: $${PROJECTED_Q1_FY2026} million`,
    `  - Enterprise Software: $${(PROJECTED_Q1_FY2026 * 0.62).toFixed(1)} million`,
    `  - Professional Services: $${(PROJECTED_Q1_FY2026 * 0.25).toFixed(1)} million`,
    `  - Managed Services: $${(PROJECTED_Q1_FY2026 * 0.13).toFixed(1)} million`,
    '',
    'Key growth drivers:',
    '  1. APAC market entry (targeting $4.2M incremental)',
    '  2. Enterprise Platform v3.0 launch',
    '  3. Expansion of managed services offering',
  ];

  let projY = 1.2;
  for (const line of projectionLines) {
    if (line === '') {
      projY += 0.15;
      continue;
    }
    s3.addText(line, {
      x: 0.8,
      y: projY,
      w: 8.5,
      h: 0.4,
      fontSize: 14,
      color: '333333',
    });
    projY += 0.35;
  }

  // --- Slide 4: Key Metrics ---
  const s4 = pptx.addSlide();
  s4.addText('Key Metrics & KPIs', {
    x: 0.5,
    y: 0.3,
    w: 9,
    h: 0.6,
    fontSize: 24,
    bold: true,
    color: DARK_BLUE,
  });

  const metrics = [
    ['FY2025 Total Revenue', `$${TOTAL_REVENUE} million`],
    ['Net Income', `$${NET_INCOME} million`],
    ['ARR', `$${ARR} million`],
    ['Net Revenue Retention', '118%'],
    ['Gross Margin', '72.4%'],
    ['Headcount (EOY FY2025)', `${HEADCOUNT}`],
    ['Planned New Hires (FY2026)', `${NEW_HIRES}`],
    ['Q1 FY2026 Projection', `$${PROJECTED_Q1_FY2026} million`],
  ];

  s4.addTable(
    [
      [
        { text: 'Metric', options: { bold: true, color: 'FFFFFF', fill: { color: MEDIUM_BLUE } } },
        { text: 'Value', options: { bold: true, color: 'FFFFFF', fill: { color: MEDIUM_BLUE } } },
      ],
      ...metrics.map(([metric, value], i) => [
        { text: metric, options: { fill: { color: i % 2 === 0 ? LIGHT_GRAY : 'FFFFFF' } } },
        { text: value, options: { bold: true, fill: { color: i % 2 === 0 ? LIGHT_GRAY : 'FFFFFF' } } },
      ]),
    ],
    {
      x: 0.5,
      y: 1.2,
      w: 9,
      fontSize: 14,
      border: { type: 'solid', pt: 0.5, color: 'CCCCCC' },
      colW: [5, 4],
      rowH: [0.5, ...metrics.map(() => 0.4)],
    },
  );

  await pptx.writeFile({ fileName: filePath });
}

// ---------------------------------------------------------------------------
// Generate all fixtures
// ---------------------------------------------------------------------------

export async function generateAll(outDir: string): Promise<void> {
  await mkdir(outDir, { recursive: true });
  await Promise.all([
    createAcmePdf(join(outDir, ACME_PDF_FILENAME)),
    createAcmeDocx(join(outDir, ACME_DOCX_FILENAME)),
    createAcmePptx(join(outDir, ACME_PPTX_FILENAME)),
  ]);
}

// ---------------------------------------------------------------------------
// CLI entry point: tsx lib/acme-fixtures.ts --outdir <path>
// ---------------------------------------------------------------------------

const isMainModule =
  typeof process !== 'undefined' &&
  process.argv[1] &&
  (process.argv[1].endsWith('acme-fixtures.ts') || process.argv[1].endsWith('acme-fixtures'));

if (isMainModule) {
  const { values } = parseArgs({
    args: process.argv.slice(2),
    options: {
      outdir: { type: 'string', default: './fixtures/acme' },
    },
  });

  const outDir = values.outdir!;
  console.log(`Generating Acme Corp FY2025 fixtures in ${outDir} ...`);
  generateAll(outDir).then(() => {
    console.log(`  ✓ ${ACME_PDF_FILENAME}`);
    console.log(`  ✓ ${ACME_DOCX_FILENAME}`);
    console.log(`  ✓ ${ACME_PPTX_FILENAME}`);
    console.log('Done.');
  });
}
