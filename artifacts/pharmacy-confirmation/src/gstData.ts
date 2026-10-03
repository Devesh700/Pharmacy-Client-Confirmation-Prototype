// gstData.ts - Store master, customers, batches, stock ledger, seed invoices, and GST helpers
import { products, type Product } from './data';

export type StoreInfo = {
  id: string;
  name: string;
  legalName: string;
  address: string;
  state: string;
  stateCode: string;
  gstin: string;
  drugLicence: string;
  invoicePrefix: string;
  phone: string;
};

export const storeMasters: StoreInfo[] = [
  {
    id: 'S1',
    name: 'Main Store',
    legalName: 'Sampada Community Pharmacy LLP',
    address: '12th Main Road, Indiranagar, Bengaluru, Karnataka 560038',
    state: 'Karnataka',
    stateCode: '29',
    gstin: '29AABCS1429B1Z2',
    drugLicence: 'KA-B1-204918, KA-B1-204919',
    invoicePrefix: 'MS/2026-27/',
    phone: '+91 80 4123 6880'
  },
  {
    id: 'S2',
    name: 'Branch 2',
    legalName: 'Sampada Community Pharmacy (Branch 2)',
    address: '4th Block, Jayanagar, Bengaluru, Karnataka 560011',
    state: 'Karnataka',
    stateCode: '29',
    gstin: '29AABCS1429B1Z2',
    drugLicence: 'KA-B2-109281, KA-B2-109282',
    invoicePrefix: 'B2/2026-27/',
    phone: '+91 80 2664 1190'
  },
  {
    id: 'S3',
    name: 'Branch 3',
    legalName: 'Sampada Community Pharmacy (Branch 3)',
    address: '8th Cross, Malleshwaram, Bengaluru, Karnataka 560003',
    state: 'Karnataka',
    stateCode: '29',
    gstin: '29AABCS1429B1Z2',
    drugLicence: 'KA-B3-301928, KA-B3-301929',
    invoicePrefix: 'B3/2026-27/',
    phone: '+91 80 2334 7721'
  }
];

export type Customer = {
  id: string;
  name: string;
  phone: string;
  address: string;
  state: string;
  gstin?: string;
  type: 'B2C' | 'B2B';
};

export const seedCustomers: Customer[] = [
  { id: 'C01', name: 'Walk-in / Cash Customer', phone: '—', address: 'Counter Sale', state: 'Karnataka', type: 'B2C' },
  { id: 'C02', name: 'Meera Krishnan', phone: '+91 98450 18421', address: 'Indiranagar, Bengaluru', state: 'Karnataka', type: 'B2C' },
  { id: 'C03', name: 'Raghav Menon', phone: '+91 99867 22310', address: 'Jayanagar, Bengaluru', state: 'Karnataka', type: 'B2C' },
  { id: 'C04', name: 'Lakshmi Nair', phone: '+91 99011 64509', address: 'Malleshwaram, Bengaluru', state: 'Karnataka', type: 'B2C' },
  { id: 'C05', name: 'Sanjay Kulkarni', phone: '+91 98441 75320', address: 'Banaswadi, Bengaluru', state: 'Karnataka', type: 'B2C' },
  { id: 'C06', name: 'Dr. Ramesh Heart Clinic', phone: '+91 98801 44210', address: 'Koramangala, Bengaluru', state: 'Karnataka', gstin: '29AAACR3918M1Z8', type: 'B2B' },
  { id: 'C07', name: 'Sunrise Multispeciality Hospital', phone: '+91 80 4912 3000', address: 'Old Airport Road, Bengaluru', state: 'Karnataka', gstin: '29AABCS9912K1ZB', type: 'B2B' },
  { id: 'C08', name: 'Cauvery Care Nursing Home', phone: '+91 98440 22119', address: 'Mysuru, Karnataka', state: 'Karnataka', gstin: '29AAACC1029P1ZK', type: 'B2B' },
  { id: 'C09', name: 'Astra Polyclinic & Diagnostics', phone: '+91 99000 81234', address: 'Whitefield, Bengaluru', state: 'Karnataka', gstin: '29AABCA4419J1Z4', type: 'B2B' },
  { id: 'C10', name: 'Apex MediCorp Enterprises', phone: '+91 98452 77112', address: 'Hosur Road, Bengaluru', state: 'Karnataka', gstin: '29AAAAP8819L1ZG', type: 'B2B' },
  { id: 'C11', name: 'Apollo First Aid Station (Hosur)', phone: '+91 94432 11990', address: 'Sipcot Industrial Area, Hosur, Tamil Nadu', state: 'Tamil Nadu', gstin: '33AAACA9921B1ZD', type: 'B2B' }
];

export type ProductGstMeta = {
  hsn: string;
  gstRate: number; // 0, 5, 12, 18
  schedule: 'None' | 'Schedule H' | 'Schedule H1' | 'Schedule X';
  unit: string;
  mrp: number;
};

// Map realistic GST rates, HSN and Schedule info to each product
export const productGstMap: Record<string, ProductGstMeta> = {};

products.forEach((p, idx) => {
  let rate = 12; // Standard for most formulation medicaments under HSN 3004
  let schedule: 'None' | 'Schedule H' | 'Schedule H1' | 'Schedule X' = 'None';
  let unit = 'Strip 10s';

  if (p.group === 'Antibiotics') {
    rate = 12;
    schedule = p.name.includes('Amoxicillin') || p.name.includes('Cefixime') ? 'Schedule H1' : 'Schedule H';
  } else if (p.group === 'Analgesics') {
    rate = 12;
    if (p.name.includes('Aceclofenac')) schedule = 'Schedule H';
  } else if (p.group === 'Diabetes' || p.group === 'Cardiac') {
    rate = 12;
    schedule = 'Schedule H';
  } else if (p.group === 'Vitamins') {
    rate = 18; // Many nutraceuticals/vitamins fall under 18%
    schedule = 'None';
  } else if (p.name.includes('ORS') || p.name.includes('Glucose')) {
    rate = 12;
    unit = 'Sachet';
  } else if (p.name.includes('syrup') || p.name.includes('drops') || p.name.includes('solution')) {
    unit = 'Bottle';
  } else if (p.name.includes('gel') || p.name.includes('cream') || p.name.includes('ointment')) {
    unit = 'Tube';
  } else if (p.name.includes('diaper') || p.name.includes('mask') || p.name.includes('sanitiser')) {
    rate = p.name.includes('mask') ? 5 : 18;
    unit = 'Pack';
  } else if (p.name.includes('Pregabalin')) {
    schedule = 'Schedule H';
  }

  const baseMrp = 35 + (idx * 17) % 240 + ((idx % 3) * 12.5);

  productGstMap[p.id] = {
    hsn: '3004',
    gstRate: rate,
    schedule,
    unit,
    mrp: Math.round(baseMrp * 100) / 100
  };
});

export type StockMovementType =
  | 'Opening / Stock import'
  | 'Purchase (in)'
  | 'Sale (invoice)'
  | 'Sales return / credit note'
  | 'Purchase return'
  | 'Transfer out'
  | 'Transfer in'
  | 'Adjustment';

export type StockMovement = {
  id: string;
  dateTime: string;
  storeId: string;
  productId: string;
  batch: string;
  expiry: string; // YYYY-MM
  qty: number; // positive for additions, negative for deductions
  type: StockMovementType;
  ref: string;
  user: string;
  reason?: string;
  source: 'ERP Import' | 'App Billing';
};

export type BatchStock = {
  productId: string;
  storeId: string;
  batch: string;
  expiry: string;
  qty: number;
  mrp: number;
};

// Seed 2 to 3 realistic batches per product per store
export const seedBatches: BatchStock[] = [];
products.forEach((p, i) => {
  storeMasters.forEach((s, j) => {
    const meta = productGstMap[p.id] || { mrp: 65 };
    // Batch 1: regular fresh batch
    seedBatches.push({
      productId: p.id,
      storeId: s.id,
      batch: `BT-${2400 + (i * 3 + j) % 500}`,
      expiry: '2026-11',
      qty: 45 + ((i * 11 + j * 7) % 60),
      mrp: meta.mrp
    });

    // Batch 2: for some products, expiring soon or slightly older
    if (i % 2 === 0) {
      const isExpiringSoon = i % 6 === 0;
      const isExpired = i === 12; // deliberate expired batch (e.g. Cough syrup or ORS)
      seedBatches.push({
        productId: p.id,
        storeId: s.id,
        batch: `BT-${2200 + (i * 5 + j) % 400}`,
        expiry: isExpired ? '2025-04' : isExpiringSoon ? '2025-07' : '2026-03',
        qty: isExpired ? 12 : 25 + ((i * 3) % 30),
        mrp: meta.mrp
      });
    }

    // Batch 3: for high-rotation items (like Paracetamol, Dolo, Metformin)
    if (i < 8) {
      seedBatches.push({
        productId: p.id,
        storeId: s.id,
        batch: `BT-25${String((i * 4 + j) % 90).padStart(2, '0')}`,
        expiry: '2027-04',
        qty: 60 + j * 15,
        mrp: meta.mrp
      });
    }
  });
});

// Seed Stock Ledger Movements
export const initialMovements: StockMovement[] = [];

// 1. Initial Opening from imported snapshot as of 2025-06-01
seedBatches.forEach((b, idx) => {
  initialMovements.push({
    id: `MOV-OP-${idx + 1001}`,
    dateTime: '2025-06-01 08:00',
    storeId: b.storeId,
    productId: b.productId,
    batch: b.batch,
    expiry: b.expiry,
    qty: b.qty + 15, // opening was slightly higher before June sales
    type: 'Opening / Stock import',
    ref: 'IMPORT-SNAPSHOT-20250601',
    user: 'System Import',
    source: 'ERP Import'
  });
});

// 2. Add some purchases in June
products.slice(0, 15).forEach((p, idx) => {
  const batch = seedBatches.find(b => b.productId === p.id && b.storeId === 'S1') || seedBatches[0];
  initialMovements.push({
    id: `MOV-PUR-${idx + 1}`,
    dateTime: `2025-06-05 11:${10 + (idx % 40)}`,
    storeId: 'S1',
    productId: p.id,
    batch: batch.batch,
    expiry: batch.expiry,
    qty: 30,
    type: 'Purchase (in)',
    ref: `INV-2025-098${idx}`,
    user: 'Ananya Rao',
    source: 'ERP Import'
  });
});

// 3. One deliberate inter-store transfer
initialMovements.push({
  id: 'MOV-TR-OUT-01',
  dateTime: '2025-06-12 14:30',
  storeId: 'S1',
  productId: 'P001',
  batch: 'BT-2400',
  expiry: '2026-11',
  qty: -20,
  type: 'Transfer out',
  ref: 'TR-2025-0014',
  user: 'Ananya Rao',
  reason: 'Replenishment for Branch 2',
  source: 'App Billing'
});

initialMovements.push({
  id: 'MOV-TR-IN-01',
  dateTime: '2025-06-12 16:15',
  storeId: 'S2',
  productId: 'P001',
  batch: 'BT-2400',
  expiry: '2026-11',
  qty: 20,
  type: 'Transfer in',
  ref: 'TR-2025-0014',
  user: 'Nikhil Shetty',
  reason: 'Received from Main Store',
  source: 'App Billing'
});

// 4. One deliberate stock adjustment (Damage / Expiry write-off)
initialMovements.push({
  id: 'MOV-ADJ-01',
  dateTime: '2025-06-15 10:20',
  storeId: 'S1',
  productId: 'P013', // Cough syrup
  batch: 'BT-2439',
  expiry: '2026-11',
  qty: -3,
  type: 'Adjustment',
  ref: 'ADJ-2025-0008',
  user: 'Dev Shah',
  reason: 'Damage: 3 glass bottles broken during shelf handling',
  source: 'App Billing'
});

// 5. One credit note reversal
initialMovements.push({
  id: 'MOV-CN-01',
  dateTime: '2025-06-18 17:05',
  storeId: 'S1',
  productId: 'P002', // Paracetamol 650 mg
  batch: 'BT-2403',
  expiry: '2026-11',
  qty: 1,
  type: 'Sales return / credit note',
  ref: 'CN/2026-27/00001',
  user: 'Ananya Rao',
  reason: 'Customer returned unneeded sealed strip',
  source: 'App Billing'
});

export type InvoiceLine = {
  productId: string;
  productName: string;
  hsn: string;
  batch: string;
  expiry: string;
  qty: number;
  unit: string;
  mrp: number;
  rate: number;
  discountPct: number;
  discountAmt: number;
  taxableValue: number;
  gstRate: number;
  cgst: number;
  sgst: number;
  igst: number;
  lineTotal: number;
};

export type Invoice = {
  id: string;
  date: string;
  storeId: string;
  customerId: string;
  customerName: string;
  customerPhone: string;
  customerAddress?: string;
  customerGstin?: string;
  customerState: string;
  isB2B: boolean;
  doctorName?: string;
  prescriptionId?: string;
  lines: InvoiceLine[];
  subtotal: number;
  discountTotal: number;
  taxableTotal: number;
  cgstTotal: number;
  sgstTotal: number;
  igstTotal: number;
  roundOff: number;
  grandTotal: number;
  paymentMode: 'Cash' | 'UPI' | 'Card' | 'Credit';
  amountReceived: number;
  status: 'Final' | 'Draft' | 'Cancelled';
  cancelReason?: string;
  creditNoteId?: string;
  user: string;
  source: 'App Billing';
};

// Indian Rupees to words conversion
export function numberToWordsIndian(num: number): string {
  const a = [
    '', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten',
    'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'
  ];
  const b = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];

  const n = Math.floor(Math.abs(num));
  if (n === 0) return 'Rupees Zero Only';

  function convertTens(n: number): string {
    if (n < 20) return a[n];
    return `${b[Math.floor(n / 10)]}${n % 10 ? ' ' + a[n % 10] : ''}`;
  }

  function convertGroup(n: number): string {
    let str = '';
    if (Math.floor(n / 100) > 0) {
      str += `${a[Math.floor(n / 100)]} Hundred `;
      n %= 100;
    }
    if (n > 0) {
      str += convertTens(n);
    }
    return str.trim();
  }

  let words = '';
  const crore = Math.floor(n / 10000000);
  const lakh = Math.floor((n % 10000000) / 100000);
  const thousand = Math.floor((n % 100000) / 1000);
  const remainder = n % 1000;

  if (crore > 0) words += `${convertGroup(crore)} Crore `;
  if (lakh > 0) words += `${convertGroup(lakh)} Lakh `;
  if (thousand > 0) words += `${convertGroup(thousand)} Thousand `;
  if (remainder > 0) words += convertGroup(remainder);

  const paise = Math.round((Math.abs(num) - n) * 100);
  let result = `Rupees ${words.trim()}`;
  if (paise > 0) {
    result += ` and ${convertTens(paise)} Paise`;
  }
  return `${result} Only`;
}

// Generate 25 realistic seed invoices across stores
export const seedInvoices: Invoice[] = [];

const sampleDates = [
  '2025-06-10 10:15', '2025-06-10 11:42', '2025-06-11 14:20', '2025-06-11 16:55',
  '2025-06-12 09:30', '2025-06-12 13:10', '2025-06-13 15:40', '2025-06-14 11:05',
  '2025-06-14 18:22', '2025-06-15 10:45', '2025-06-16 12:15', '2025-06-16 17:30',
  '2025-06-17 09:50', '2025-06-17 14:10', '2025-06-18 11:25', '2025-06-18 16:40',
  '2025-06-19 10:10', '2025-06-19 15:20', '2025-06-20 09:40', '2025-06-20 12:30',
  '2025-06-20 14:15', '2025-06-20 16:10', '2025-06-20 17:45', '2025-06-20 18:30',
  '2025-06-20 19:10'
];

for (let i = 0; i < 25; i++) {
  const storeId = i % 5 === 0 ? 'S2' : i % 7 === 0 ? 'S3' : 'S1';
  const prefix = storeMasters.find(s => s.id === storeId)?.invoicePrefix || 'MS/2026-27/';
  const invNo = `${prefix}${String(i + 1).padStart(5, '0')}`;
  const dt = sampleDates[i] || '2025-06-20 12:00';

  let customer = seedCustomers[0]; // default walk-in
  if (i === 4) customer = seedCustomers[5]; // B2B
  else if (i === 9) customer = seedCustomers[6]; // B2B
  else if (i === 14) customer = seedCustomers[7]; // B2B
  else if (i === 18) customer = seedCustomers[8]; // B2B
  else if (i === 22) customer = seedCustomers[10]; // Inter-state Hosur Tamil Nadu (IGST!)
  else if (i % 3 === 0) customer = seedCustomers[(i % 4) + 1];

  const isB2B = customer.type === 'B2B';
  const isInterstate = customer.state !== 'Karnataka';

  // Pick 1 to 3 products
  const p1 = products[(i * 3) % products.length];
  const p2 = products[(i * 3 + 1) % products.length];
  const selectedProds = i % 2 === 0 ? [p1, p2] : [p1];

  const lines: InvoiceLine[] = selectedProds.map((prod, lIdx) => {
    const meta = productGstMap[prod.id] || { mrp: 60, hsn: '3004', gstRate: 12, unit: 'Strip 10s' };
    const batch = seedBatches.find(b => b.productId === prod.id && b.storeId === storeId) || seedBatches[0];
    const qty = (i % 3) + 1;
    const mrp = meta.mrp;
    const rate = Math.round((mrp / (1 + meta.gstRate / 100)) * 100) / 100;
    const discountPct = isB2B ? 10 : 0;
    const discountAmt = Math.round((rate * qty * discountPct / 100) * 100) / 100;
    const taxableValue = Math.round((rate * qty - discountAmt) * 100) / 100;
    const taxAmt = Math.round((taxableValue * meta.gstRate / 100) * 100) / 100;

    const cgst = isInterstate ? 0 : Math.round((taxAmt / 2) * 100) / 100;
    const sgst = isInterstate ? 0 : Math.round((taxAmt / 2) * 100) / 100;
    const igst = isInterstate ? taxAmt : 0;
    const lineTotal = Math.round((taxableValue + cgst + sgst + igst) * 100) / 100;

    return {
      productId: prod.id,
      productName: prod.name,
      hsn: meta.hsn,
      batch: batch.batch,
      expiry: batch.expiry,
      qty,
      unit: meta.unit,
      mrp,
      rate,
      discountPct,
      discountAmt,
      taxableValue,
      gstRate: meta.gstRate,
      cgst,
      sgst,
      igst,
      lineTotal
    };
  });

  const taxableTotal = Math.round(lines.reduce((n, l) => n + l.taxableValue, 0) * 100) / 100;
  const cgstTotal = Math.round(lines.reduce((n, l) => n + l.cgst, 0) * 100) / 100;
  const sgstTotal = Math.round(lines.reduce((n, l) => n + l.sgst, 0) * 100) / 100;
  const igstTotal = Math.round(lines.reduce((n, l) => n + l.igst, 0) * 100) / 100;
  const rawGrand = taxableTotal + cgstTotal + sgstTotal + igstTotal;
  const grandTotal = Math.round(rawGrand);
  const roundOff = Math.round((grandTotal - rawGrand) * 100) / 100;

  // Add Sale movement for each line to stock ledger
  lines.forEach((line, lIdx) => {
    initialMovements.push({
      id: `MOV-INV-${i + 1}-${lIdx + 1}`,
      dateTime: dt,
      storeId,
      productId: line.productId,
      batch: line.batch,
      expiry: line.expiry,
      qty: -line.qty,
      type: 'Sale (invoice)',
      ref: invNo,
      user: 'Ananya Rao',
      source: 'App Billing'
    });
  });

  // Scheduled drug check: invoice 6 & 16 have prescription link and doctor name
  const isScheduled = lines.some(l => (productGstMap[l.productId]?.schedule || 'None') !== 'None');
  const doctorName = isScheduled || i === 0 || i === 7 ? 'Dr. Priya Sharma, MBBS' : undefined;
  const prescriptionId = isScheduled ? 'RX-24080' : undefined;

  let status: 'Final' | 'Draft' | 'Cancelled' = 'Final';
  let cancelReason: string | undefined;
  if (i === 12) {
    status = 'Cancelled';
    cancelReason = 'Customer changed mind before counter payment';
  } else if (i === 24) {
    status = 'Draft';
  }

  seedInvoices.push({
    id: invNo,
    date: dt,
    storeId,
    customerId: customer.id,
    customerName: customer.name,
    customerPhone: customer.phone,
    customerAddress: customer.address,
    customerGstin: customer.gstin,
    customerState: customer.state,
    isB2B,
    doctorName,
    prescriptionId,
    lines,
    subtotal: lines.reduce((n, l) => n + l.mrp * l.qty, 0),
    discountTotal: lines.reduce((n, l) => n + l.discountAmt, 0),
    taxableTotal,
    cgstTotal,
    sgstTotal,
    igstTotal,
    roundOff,
    grandTotal,
    paymentMode: i % 4 === 0 ? 'UPI' : i % 3 === 0 ? 'Card' : isB2B ? 'Credit' : 'Cash',
    amountReceived: grandTotal,
    status,
    cancelReason,
    creditNoteId: i === 1 ? 'CN/2026-27/00001' : undefined,
    user: i % 2 === 0 ? 'Ananya Rao' : 'Dev Shah',
    source: 'App Billing'
  });
}

export type CreditNote = {
  id: string;
  date: string;
  storeId: string;
  originalInvoiceId: string;
  customerName: string;
  lines: InvoiceLine[];
  taxableTotal: number;
  cgstTotal: number;
  sgstTotal: number;
  igstTotal: number;
  grandTotal: number;
  reason: string;
  user: string;
};

export const seedCreditNotes: CreditNote[] = [
  {
    id: 'CN/2026-27/00001',
    date: '2025-06-18 17:05',
    storeId: 'S1',
    originalInvoiceId: 'MS/2026-27/00002',
    customerName: 'Meera Krishnan',
    lines: [
      {
        productId: 'P002',
        productName: 'Paracetamol 650 mg',
        hsn: '3004',
        batch: 'BT-2403',
        expiry: '2026-11',
        qty: 1,
        unit: 'Tablet 15s',
        mrp: 45,
        rate: 40.18,
        discountPct: 0,
        discountAmt: 0,
        taxableValue: 40.18,
        gstRate: 12,
        cgst: 2.41,
        sgst: 2.41,
        igst: 0,
        lineTotal: 45
      }
    ],
    taxableTotal: 40.18,
    cgstTotal: 2.41,
    sgstTotal: 2.41,
    igstTotal: 0,
    grandTotal: 45,
    reason: 'Customer returned unneeded sealed strip',
    user: 'Ananya Rao'
  }
];

export type InterStoreTransfer = {
  id: string;
  date: string;
  fromStoreId: string;
  toStoreId: string;
  productId: string;
  batch: string;
  expiry: string;
  qty: number;
  status: 'Draft' | 'Dispatched' | 'Received';
  user: string;
};

export const seedTransfers: InterStoreTransfer[] = [
  {
    id: 'TR-2025-0014',
    date: '2025-06-12 14:30',
    fromStoreId: 'S1',
    toStoreId: 'S2',
    productId: 'P001',
    batch: 'BT-2400',
    expiry: '2026-11',
    qty: 20,
    status: 'Received',
    user: 'Ananya Rao'
  },
  {
    id: 'TR-2025-0015',
    date: '2025-06-19 11:20',
    fromStoreId: 'S1',
    toStoreId: 'S3',
    productId: 'P003',
    batch: 'BT-2406',
    expiry: '2026-11',
    qty: 10,
    status: 'Dispatched',
    user: 'Dev Shah'
  }
];

export type ReconciliationDiscrepancy = {
  id: string;
  storeId: string;
  productId: string;
  batch: string;
  systemQty: number;
  importedErpQty: number;
  difference: number;
  suggestedReason: string;
  status: 'Pending' | 'Accepted' | 'Adjusted';
};

// One deliberate stock discrepancy per store for the Reconciliation screen
export const seedReconciliation: ReconciliationDiscrepancy[] = [
  {
    id: 'REC-01',
    storeId: 'S1',
    productId: 'P001', // Paracetamol 500 mg
    batch: 'BT-2400',
    systemQty: 48,
    importedErpQty: 50,
    difference: -2,
    suggestedReason: 'Unrecorded sale or minor counter breakage',
    status: 'Pending'
  },
  {
    id: 'REC-02',
    storeId: 'S2',
    productId: 'P005', // Metformin 500 mg
    batch: 'BT-2412',
    systemQty: 62,
    importedErpQty: 58,
    difference: 4,
    suggestedReason: 'ERP sales return not yet synced to app',
    status: 'Pending'
  },
  {
    id: 'REC-03',
    storeId: 'S3',
    productId: 'P010', // Cetirizine 10 mg
    batch: 'BT-2427',
    systemQty: 30,
    importedErpQty: 35,
    difference: -5,
    suggestedReason: 'Count error during physical stock audit',
    status: 'Pending'
  }
];

// Compatibility aliases
export const seedStores = storeMasters;
export const productGstMeta = productGstMap;
export const seedMovements = initialMovements;
export const seedReconciliations = seedReconciliation;
export type ReconciliationRow = ReconciliationDiscrepancy;
