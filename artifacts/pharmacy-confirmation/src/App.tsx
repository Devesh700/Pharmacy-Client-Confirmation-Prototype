import { useEffect, useMemo, useRef, useState, type ReactNode, type ChangeEvent, type DragEvent, type FormEvent } from 'react';
import { Link, Route, Switch, useLocation, useRoute } from 'wouter';
import * as XLSX from 'xlsx';
import {
  AlertTriangle,
  ArrowDownToLine,
  ArrowLeft,
  ArrowRight,
  Bookmark,
  Building2,
  Check,
  CheckCircle2,
  ChevronDown,
  CircleHelp,
  ClipboardList,
  Database,
  FileSpreadsheet,
  Filter,
  GitMerge,
  Inbox,
  LayoutDashboard,
  ListChecks,
  MessageCircle,
  Package,
  Plus,
  Printer,
  Search,
  Settings,
  ShieldCheck,
  Sparkles,
  Upload,
  Zap,
  Boxes,
  Receipt,
  FileText,
  BarChart3,
  Store,
  RefreshCw,
  Clock,
  Tag,
  ShieldAlert,
  ArrowUpRight
} from 'lucide-react';
import {
  products as seedProducts,
  stores as initialStores,
  suppliers as seedSuppliers,
  stock as seedStock,
  sales as seedSales,
  purchases as seedPurchases,
  matches as seedMatches,
  messages as seedMessages,
  prescriptions as seedPrescriptions,
  months,
  type Product
} from './data';
import {
  seedStores,
  seedCustomers,
  productGstMeta,
  seedBatches,
  seedMovements,
  seedInvoices,
  seedCreditNotes,
  seedTransfers,
  seedReconciliations,
  type StoreInfo,
  type Customer,
  type StockMovement,
  type Invoice,
  type CreditNote,
  type InterStoreTransfer,
  type ReconciliationRow,
  numberToWordsIndian
} from './gstData';
import {
  StoreStockView,
  BillingView,
  InvoicesView,
  GstReportsView
} from './GstComponents';

type Role = 'Admin' | 'Manager' | 'Store Owner' | 'Pharmacist' | 'Staff';
type Batch = {
  id: string;
  file: string;
  type: string;
  status: string;
  rows: number;
  loaded: number;
  rejected: number;
  uploadedBy: string;
  time: string;
  errors: { row: number; column: string; message: string }[];
};
type Feedback = { screen: string; status: 'Looks right' | 'Needs change'; comment: string };

const reports = [
  'Dashboard',
  'Store Stock',
  'Billing',
  'Invoices',
  'GST Reports',
  'Imports',
  'Products',
  'Suppliers',
  'Matching Queue',
  'Ask',
  'WhatsApp Inbox',
  'Prescriptions',
  'Admin',
  'Client confirmation'
];

const routePaths = [
  '/',
  '/store-stock',
  '/billing',
  '/invoices',
  '/gst-reports',
  '/imports',
  '/products',
  '/suppliers',
  '/matching',
  '/ask',
  '/whatsapp',
  '/prescriptions',
  '/admin',
  '/client-confirmation'
];

const IconNav = [
  LayoutDashboard,
  Boxes,
  Receipt,
  FileText,
  BarChart3,
  FileSpreadsheet,
  Package,
  Building2,
  ListChecks,
  Sparkles,
  MessageCircle,
  ClipboardList,
  Settings,
  CheckCircle2
];

const promptQuestions = [
  'After go-live, will billing be done in this app instead of MediVision, or in parallel? (Determines the source of truth and duplicate-handling.)',
  'Do you need Bill of Supply, composition scheme, or credit/debit note variations?',
  'Is the store registered under regular GST? GSTIN, state, drug licence per store?',
  'Invoice number format and series per store and financial year?',
  'Are prices MRP-inclusive? Any discounts policy? GST rate per product and who maintains it?',
  'Do you need B2B customer ledgers, outstanding/credit sales and payment receipts (this is accounting, not currently in scope)?',
  'Do you need e-invoicing, e-way bill, or direct GSTR filing/GST portal integration? (Not in prototype.)',
  'Barcode scanner / thermal printer at the counter? Models?',
  'Batch-wise tracking and FEFO: mandatory for all products?',
  'Scheduled-drug (H/H1/X) register requirements and who signs?',
  'Inter-store transfers: do they need a GST document (delivery challan/invoice) between your own stores?',
  'Do all stores operate under one GSTIN or separate GSTINs?'
];

const users = [
  ['Ananya Rao', 'Admin', 'All stores, masters, overrides, and billing configuration'],
  ['Rajesh Patel', 'Store Owner', 'Main Store billing, stock ledger, adjustments and reports'],
  ['Nikhil Shetty', 'Manager', 'Inventory, imports, reconciliations and GST reports (no counter billing)'],
  ['Dev Shah', 'Pharmacist', 'Products, scheduled drugs dispensing and prescription workflow'],
  ['Kavya Iyer', 'Staff', 'Product lookup, draft bills and WhatsApp/prescription review']
];

const money = (n: number) => `₹${n.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
const today = '30 Jun 2025';
const dl = (filename: string, data: string, mime = 'application/json') => {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([data], { type: mime }));
  a.download = filename;
  a.click();
  URL.revokeObjectURL(a.href);
};

function Status({ children }: { children: string }) {
  const tone = /Pending|Needs Review|Review|Queued|Received|Draft|Dispatched/.test(children)
    ? 'warn'
    : /Failed|Unknown|Change needed|Cancelled|Expired/.test(children)
    ? 'error'
    : /Closed|Completed|Resolved|Looks right|Matched|Final|Active|Regular/.test(children)
    ? ''
    : 'neutral';
  return (
    <span className={`badge ${tone}`}>
      <i aria-hidden="true">{tone === '' ? '✓' : tone === 'warn' ? '•' : tone === 'error' ? '!' : '·'}</i>
      {children}
    </span>
  );
}

function Button({
  children,
  onClick,
  kind = '',
  disabled = false,
  title = '',
  className = '',
  type = 'button',
  tiny = false
}: {
  children: ReactNode;
  onClick?: (event: any) => void;
  kind?: string;
  disabled?: boolean;
  title?: string;
  className?: string;
  type?: 'button' | 'submit';
  tiny?: boolean;
}) {
  return (
    <button type={type} title={title} disabled={disabled} onClick={onClick} className={`btn ${kind} ${tiny ? 'tiny' : ''} ${className}`}>
      {children}
    </button>
  );
}

function Panel({ title, sub, children, action }: { title: string; sub?: string; children: ReactNode; action?: ReactNode }) {
  return (
    <section className="panel">
      <div className="panel-head">
        <div>
          <div className="panel-title">{title}</div>
          {sub && <div className="panel-sub">{sub}</div>}
        </div>
        {action}
      </div>
      {children}
    </section>
  );
}

function PageHead({ title, help, actions }: { title: string; help: string; actions?: ReactNode }) {
  return (
    <div className="page-head">
      <div>
        <div className="eyebrow">PHARMACY DATA WORKSPACE</div>
        <h1 className="page-title">{title}</h1>
        <p className="help-line">What this screen does: {help}</p>
      </div>
      {actions && <div className="page-head-actions">{actions}</div>}
    </div>
  );
}

function Field({ label, children }: { label: string; children: ReactNode }) {
  return <label className="field">{label}{children}</label>;
}

function Modal({ title, close, children }: { title: string; close: () => void; children: ReactNode }) {
  return (
    <div className="modal-backdrop" onMouseDown={e => { if (e.target === e.currentTarget) close(); }}>
      <div className="modal" role="dialog" aria-modal="true" aria-label={title}>
        <div className="drawer-head">
          <h2>{title}</h2>
          <button className="close-button" aria-label="Close dialog" onClick={close}>×</button>
        </div>
        {children}
      </div>
    </div>
  );
}

function App() {
  const [location, setLocation] = useLocation();
  const [role, setRole] = useState<Role>(() => (localStorage.getItem('pharmacy-role') as Role) || 'Admin');
  const [search, setSearch] = useState('');
  const [feedbackOpen, setFeedbackOpen] = useState(false);
  const [feedbackScreen, setFeedbackScreen] = useState('Dashboard');
  const [feedback, setFeedback] = useState<Feedback[]>([]);
  const [feedbackStatus, setFeedbackStatus] = useState<'Looks right' | 'Needs change'>('Looks right');
  const [feedbackText, setFeedbackText] = useState('');

  // Core Data & GST Billing State
  const [products, setProducts] = useState(seedProducts);
  const [storesMeta, setStoresMeta] = useState<StoreInfo[]>(seedStores);
  const [activeStoreId, setActiveStoreId] = useState<string>('S1'); // S1 Main Store default
  const [customers, setCustomers] = useState<Customer[]>(seedCustomers);
  const [movements, setMovements] = useState<StockMovement[]>(seedMovements);
  const [invoices, setInvoices] = useState<Invoice[]>(seedInvoices);
  const [creditNotes, setCreditNotes] = useState<CreditNote[]>(seedCreditNotes);
  const [transfers, setTransfers] = useState<InterStoreTransfer[]>(seedTransfers);
  const [reconciliations, setReconciliations] = useState<ReconciliationRow[]>(seedReconciliations);

  // GST & Business Configurations
  const [pricesInclusive, setPricesInclusive] = useState<boolean>(true);
  const [allowNegativeStock, setAllowNegativeStock] = useState<boolean>(false);
  const [goLiveDate, setGoLiveDate] = useState<string>('2025-06-01');
  const [billingPrefill, setBillingPrefill] = useState<{
    prescriptionId?: string;
    patientName?: string;
    phone?: string;
    doctorName?: string;
    productIds?: string[];
  } | null>(null);

  // Legacy modules state preserved
  const [matches, setMatches] = useState(seedMatches.map(x => ({ ...x })));
  const [messages, setMessages] = useState(seedMessages.map(x => ({ ...x })));
  const [messageLineActions, setMessageLineActions] = useState<Record<string, Record<number, string>>>({});
  const [prescriptions, setPrescriptions] = useState<any[]>(seedPrescriptions.map(x => ({ ...x, history: [...x.history], printLog: [...x.printLog] })));
  const [batches, setBatches] = useState<Batch[]>([
    { id: 'IB-1082', file: 'june_stock_main_store.xlsx', type: 'Stock', status: 'Completed', rows: 60, loaded: 60, rejected: 0, uploadedBy: 'Ananya Rao', time: '20 Jun, 09:16 AM', errors: [] },
    { id: 'IB-1081', file: 'sales_branch_2_may.xlsx', type: 'Sales', status: 'Completed with errors', rows: 184, loaded: 179, rejected: 5, uploadedBy: 'Nikhil Shetty', time: '19 Jun, 04:42 PM', errors: [{ row: 14, column: 'Product code', message: 'Required value is empty' }, { row: 38, column: 'Sales qty', message: 'Expected a number; found “five”' }, { row: 89, column: 'Month', message: 'Date is outside the 12-month reporting window' }, { row: 111, column: 'Store', message: 'Store name does not match a configured location' }, { row: 160, column: 'Returns', message: 'Expected a non-negative number' }] }
  ]);
  const [audit, setAudit] = useState<Array<{ action: string; user: string; time: string }>>([
    { action: 'Seed stock ledger initialized across 3 stores with movements', user: 'System', time: '01 Jun, 08:00 AM' },
    { action: 'Import completed: june_stock_main_store.xlsx', user: 'Ananya Rao', time: '20 Jun, 09:16 AM' },
    { action: 'Prescription RX-24078 closed', user: 'Ananya Rao', time: '19 Jun, 12:32 PM' },
    { action: 'Exported stock report', user: 'Nikhil Shetty', time: '18 Jun, 03:08 PM' }
  ]);
  const [averageMode, setAverageMode] = useState<'complete' | 'rolling'>('complete');
  const [threshold, setThreshold] = useState(1);
  const [answers, setAnswers] = useState<string[]>(Array(promptQuestions.length).fill(''));
  const [modal, setModal] = useState<{ title: string; body: ReactNode } | null>(null);
  const [toast, setToast] = useState('');
  const [selectedBatch, setSelectedBatch] = useState<Batch | null>(null);
  const [selectedMessage, setSelectedMessage] = useState('W01');

  // Query states
  const [query, setQuery] = useState('');
  const [askResult, setAskResult] = useState<Product[] | null>(null);
  const [askUnsupported, setAskUnsupported] = useState('');
  const [askJson, setAskJson] = useState(false);
  const [askFilter, setAskFilter] = useState<{ supplier?: string; cover?: string; stock?: string; product?: string; raw?: string; showSuppliersFor?: string; storeNote?: string } | null>(null);
  const [askJsonPayload, setAskJsonPayload] = useState<any>(null);

  const [supProductSearch, setSupProductSearch] = useState('');
  const [supGroupFilter, setSupGroupFilter] = useState('');
  const [supStockMode, setSupStockMode] = useState('');
  const [supCoverFilter, setSupCoverFilter] = useState('');

  const [filter, setFilter] = useState(() => new URLSearchParams(window.location.search).get('q') || '');
  const [groupFilter, setGroupFilter] = useState(() => new URLSearchParams(window.location.search).get('group') || '');
  const [storeFilter, setStoreFilter] = useState(() => new URLSearchParams(window.location.search).get('store') || '');
  const [supplierFilter, setSupplierFilter] = useState(() => new URLSearchParams(window.location.search).get('supplier') || '');
  const [stockMode, setStockMode] = useState(() => new URLSearchParams(window.location.search).get('stock') || '');
  const [stockBelow, setStockBelow] = useState(() => Number(new URLSearchParams(window.location.search).get('below') || 20));
  const [coverFilter, setCoverFilter] = useState(() => new URLSearchParams(window.location.search).get('cover') || '');
  const [productPageIndex, setProductPageIndex] = useState(0);
  const [productSort, setProductSort] = useState('name');
  const [sortReverse, setSortReverse] = useState(false);
  const [visibleCols, setVisibleCols] = useState(true);
  const [screenFeedback, setScreenFeedback] = useState<Record<string, { status: string; comment: string }>>({});
  const [importProgress, setImportProgress] = useState('');
  const [drop, setDrop] = useState(false);
  const [msgInput, setMsgInput] = useState('');

  const [productEdit, setProductEdit] = useState(false);
  const [supplierSearch, setSupplierSearch] = useState('');
  const [supplierSelection, setSupplierSelection] = useState('');
  const [adminTab, setAdminTab] = useState('Users & roles');

  const time = () => new Date().toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit' });
  const notify = (s: string) => {
    setToast(s);
    window.setTimeout(() => setToast(''), 3200);
  };

  const userName = useMemo(() => {
    return role === 'Admin'
      ? 'Ananya Rao'
      : role === 'Store Owner'
      ? 'Rajesh Patel'
      : role === 'Pharmacist'
      ? 'Dev Shah'
      : role === 'Manager'
      ? 'Nikhil Shetty'
      : 'Kavya Iyer';
  }, [role]);

  useEffect(() => {
    localStorage.setItem('pharmacy-role', role);
  }, [role]);

  useEffect(() => {
    if (location === '/admin' && role !== 'Admin') setLocation('/');
  }, [location, role, setLocation]);

  useEffect(() => {
    const next = new URLSearchParams();
    if (filter) next.set('q', filter);
    if (groupFilter) next.set('group', groupFilter);
    if (storeFilter) next.set('store', storeFilter);
    if (supplierFilter) next.set('supplier', supplierFilter);
    if (stockMode) next.set('stock', stockMode);
    if (stockBelow !== 20) next.set('below', String(stockBelow));
    if (coverFilter) next.set('cover', coverFilter);
    history.replaceState(null, '', `${location}${next.toString() ? `?${next}` : ''}`);
  }, [filter, groupFilter, storeFilter, supplierFilter, stockMode, stockBelow, coverFilter, location]);

  const stampAudit = (action: string) =>
    setAudit(a => [{ action, user: userName, time: time() }, ...a]);

  // ----------------------------------------------------
  // CORE STOCK LEDGER LOGIC:
  // Current stock = latest imported snapshot (opening) + all movements after snapshot
  // ----------------------------------------------------
  const getLedgerStock = (pId: string, sId?: string, batchNo?: string) => {
    return movements
      .filter(m => m.productId === pId && (!sId || sId === 'all' || m.storeId === sId) && (!batchNo || m.batch === batchNo))
      .reduce((sum, m) => sum + m.qty, 0);
  };

  // Replaces static getStock seamlessly
  const getStock = (p: string, s?: string) => getLedgerStock(p, s);

  // 3-Month average calculation incorporating BOTH ERP imports and App Billing sales / returns
  const getSales = (p: string, monthsBack: number, storeId?: string, mode = averageMode) => {
    // 1. ERP seed sales
    const list = seedSales.filter(x => x.productId === p && (!storeId || storeId === 'all' || x.storeId === storeId));
    const erpIncluded = months.slice(-3).map((month, i) =>
      list.filter(x => x.month === month).reduce((n, x) => n + x.sales - x.returns, 0) * (mode === 'rolling' && i === 0 ? 29 / 30 : 1)
    );
    const erpTotal = erpIncluded.reduce((a, x) => a + x, 0);

    // 2. App Billing sales & returns
    const appSalesTotal = Math.abs(
      movements
        .filter(m => m.productId === p && (!storeId || storeId === 'all' || m.storeId === storeId) && m.type === 'Sale' && m.source === 'App Billing')
        .reduce((sum, m) => sum + m.qty, 0)
    );
    const appReturnsTotal = movements
      .filter(m => m.productId === p && (!storeId || storeId === 'all' || m.storeId === storeId) && (m.type === 'Sales Return' || (m.type === 'Adjustment' && m.ref?.startsWith('CN-'))))
      .reduce((sum, m) => sum + m.qty, 0);
    const netAppSales = Math.max(0, appSalesTotal - appReturnsTotal);

    return (erpTotal + netAppSales) / monthsBack;
  };

  const productStats = (p: Product) => {
    const latestErp = seedSales
      .filter(x => x.productId === p.id && x.month === months.at(-1) && (!storeFilter || storeFilter === 'all' || x.storeId === storeFilter))
      .reduce((n, x) => n + x.sales - x.returns, 0);
    const latestApp = Math.abs(
      movements
        .filter(m => m.productId === p.id && (!storeFilter || storeFilter === 'all' || m.storeId === storeFilter) && m.type === 'Sale')
        .reduce((sum, m) => sum + m.qty, 0)
    );
    const latest = latestErp + latestApp;
    const avg = getSales(p.id, 3, storeFilter || undefined);
    const currentStock = getStock(p.id, storeFilter || undefined);
    return {
      qty: currentStock,
      last: latest,
      avg,
      cover: avg > 0 ? currentStock / avg : null
    };
  };

  const filteredProducts = useMemo(() => {
    let result = products.filter(p => `${p.name} ${p.group} ${p.aliases.join(' ')}`.toLowerCase().includes(filter.toLowerCase()));
    if (groupFilter) result = result.filter(p => p.group === groupFilter);
    if (supplierFilter) {
      const sid = seedSuppliers.find(s => s.name === supplierFilter)?.id;
      result = result.filter(p => seedPurchases.some(x => x.productId === p.id && x.supplierId === sid));
    }
    if (stockMode === 'zero') result = result.filter(p => productStats(p).qty === 0);
    if (stockMode === 'below') result = result.filter(p => productStats(p).qty < stockBelow);
    if (stockMode === 'range') result = result.filter(p => productStats(p).qty >= stockBelow && productStats(p).qty < stockBelow * 4);
    if (coverFilter) {
      result = result.filter(p => {
        const s = productStats(p);
        return coverFilter === 'under'
          ? s.cover !== null && s.cover < threshold
          : coverFilter === 'above'
          ? s.cover !== null && s.qty > s.avg * 2
          : !s.avg && s.qty > 100;
      });
    }
    return result.sort((a, b) => {
      const v = productSort === 'qty'
        ? productStats(a).qty - productStats(b).qty
        : productSort === 'last'
        ? productStats(a).last - productStats(b).last
        : a.name.localeCompare(b.name);
      return sortReverse ? -v : v;
    });
  }, [products, filter, groupFilter, supplierFilter, stockMode, stockBelow, coverFilter, productSort, sortReverse, movements, averageMode, storeFilter]);

  useEffect(() => setProductPageIndex(0), [filter, groupFilter, supplierFilter, stockMode, stockBelow, coverFilter, storeFilter, productSort, sortReverse]);

  const lowCount = products.filter(p => {
    const s = productStats(p);
    return s.avg > 0 && s.cover !== null && s.cover < threshold;
  }).length;

  const viewFeedback = (label: string) => {
    setFeedbackScreen(label);
    setFeedbackOpen(true);
  };

  const saveFeedback = () => {
    const item = { screen: feedbackScreen, status: feedbackStatus, comment: feedbackText };
    setFeedback(x => [...x, item]);
    setScreenFeedback(x => ({ ...x, [feedbackScreen]: { status: feedbackStatus, comment: feedbackText } }));
    setFeedbackText('');
    notify('Feedback saved for this session');
  };

  // Nav items filtered by role
  const nav = reports.map((name, i) => ({ name, path: routePaths[i], Icon: IconNav[i] })).filter(item => {
    if (item.name === 'Admin' && role !== 'Admin') return false;
    if (item.name === 'Billing' && role === 'Manager') return false;
    if (item.name === 'GST Reports' && role === 'Staff') return false;
    return true;
  });

  const navTitle = (path: string) => nav.find(n => n.path === path || (n.path !== '/' && path.startsWith(`${n.path}/`)))?.name || 'Dashboard';
  const exportFeedback = () => dl('client-feedback.json', JSON.stringify({ exportedAt: new Date().toISOString(), feedback }, null, 2));
  const exportAnswers = () =>
    dl(
      'client-confirmation-answers.json',
      JSON.stringify(
        {
          exportedAt: new Date().toISOString(),
          answers: promptQuestions.map((question, i) => ({ question, answer: answers[i] })),
          screens: screenFeedback
        },
        null,
        2
      )
    );

  // ----------------------------------------------------
  // GST INVOICING & STOCK TRANSACTIONS
  // ----------------------------------------------------
  const handleSaveInvoice = (inv: Invoice, newMovements: StockMovement[]) => {
    setInvoices(prev => [inv, ...prev]);
    setMovements(prev => [...newMovements, ...prev]);

    // Link prescription if billed from Rx
    if (inv.prescriptionId) {
      setPrescriptions(prev =>
        prev.map(rx =>
          rx.id === inv.prescriptionId
            ? { ...rx, linkedInvoiceNo: inv.invoiceNo, status: rx.status === 'Received' ? 'Review' : rx.status }
            : rx
        )
      );
    }

    const firstItem = inv.items[0];
    const prevStock = firstItem ? getStock(firstItem.productId, inv.storeId) : 0;
    const postStock = prevStock - (firstItem?.qty || 0);

    stampAudit(`Invoice ${inv.invoiceNo} saved for ${inv.customerName} (${money(inv.grandTotal)}) · ${newMovements.length} stock movement(s) posted`);
    notify(`Stock updated: ${firstItem ? firstItem.productName : 'Item'} (${inv.storeName}) ${prevStock} → ${postStock}`);
  };

  const handleCancelInvoice = (invoiceId: string, reason: string) => {
    const inv = invoices.find(i => i.id === invoiceId);
    if (!inv) return;

    // Create reverse Adjustment movements returning stock back to the original batch
    const reversalMovements: StockMovement[] = inv.items.map(item => ({
      id: `MOV-REV-${Date.now()}-${item.id}`,
      dateTime: new Date().toISOString().replace('T', ' ').slice(0, 16),
      storeId: inv.storeId,
      storeName: inv.storeName,
      productId: item.productId,
      productName: item.productName,
      batch: item.batch,
      expiry: item.expiry,
      qty: item.qty, // positive qty restores stock
      unit: item.unit,
      type: 'Adjustment',
      ref: `CANCEL:${inv.invoiceNo}`,
      user: userName,
      source: 'App Billing',
      reason: `Invoice cancellation: ${reason}`
    }));

    setInvoices(prev => prev.map(i => (i.id === invoiceId ? { ...i, status: 'Cancelled', cancelReason: reason } : i)));
    setMovements(prev => [...reversalMovements, ...prev]);
    stampAudit(`Invoice ${inv.invoiceNo} cancelled (${reason}) · Reversal movements posted to batch(es)`);
    notify(`Invoice ${inv.invoiceNo} cancelled. Stock restored to batch(es).`);
  };

  const handleCreateCreditNote = (cn: CreditNote, returnMovements: StockMovement[]) => {
    setCreditNotes(prev => [cn, ...prev]);
    setMovements(prev => [...returnMovements, ...prev]);
    stampAudit(`Credit Note ${cn.creditNoteNo} issued against ${cn.originalInvoiceNo} · Stock returned`);
    notify(`Credit note ${cn.creditNoteNo} generated and stock restored.`);
  };

  const handleAddTransfer = (transfer: InterStoreTransfer) => {
    setTransfers(prev => [transfer, ...prev]);
    stampAudit(`Transfer ${transfer.id} initiated: ${transfer.fromStoreName} -> ${transfer.toStoreName} (${transfer.productName} × ${transfer.qty})`);
    notify(`Inter-store transfer ${transfer.id} created (${transfer.status})`);
  };

  const handleUpdateTransferStatus = (transferId: string, newStatus: 'Draft' | 'Dispatched' | 'Received') => {
    const t = transfers.find(x => x.id === transferId);
    if (!t) return;
    if (newStatus === 'Dispatched' && t.status === 'Draft') {
      const outMov: StockMovement = {
        id: `MOV-TR-OUT-${Date.now()}`,
        dateTime: new Date().toISOString().replace('T', ' ').slice(0, 16),
        storeId: t.fromStoreId,
        storeName: t.fromStoreName,
        productId: t.productId,
        productName: t.productName,
        batch: t.batch,
        expiry: t.expiry,
        qty: -t.qty,
        unit: 'Tab',
        type: 'Transfer out',
        ref: t.id,
        user: userName,
        source: 'App Billing',
        reason: `Transfer to ${t.toStoreName}`
      };
      setMovements(prev => [outMov, ...prev]);
      setTransfers(prev => prev.map(x => (x.id === transferId ? { ...x, status: 'Dispatched', dispatchedAt: new Date().toISOString().replace('T', ' ').slice(0, 16) } : x)));
      stampAudit(`Transfer ${t.id} dispatched from ${t.fromStoreName}`);
      notify(`Transfer ${t.id} dispatched; stock deducted from ${t.fromStoreName}`);
    } else if (newStatus === 'Received' && t.status === 'Dispatched') {
      const inMov: StockMovement = {
        id: `MOV-TR-IN-${Date.now()}`,
        dateTime: new Date().toISOString().replace('T', ' ').slice(0, 16),
        storeId: t.toStoreId,
        storeName: t.toStoreName,
        productId: t.productId,
        productName: t.productName,
        batch: t.batch,
        expiry: t.expiry,
        qty: t.qty,
        unit: 'Tab',
        type: 'Transfer in',
        ref: t.id,
        user: userName,
        source: 'App Billing',
        reason: `Transfer from ${t.fromStoreName}`
      };
      setMovements(prev => [inMov, ...prev]);
      setTransfers(prev => prev.map(x => (x.id === transferId ? { ...x, status: 'Received', receivedAt: new Date().toISOString().replace('T', ' ').slice(0, 16) } : x)));
      stampAudit(`Transfer ${t.id} received at ${t.toStoreName}`);
      notify(`Transfer ${t.id} received; stock added to ${t.toStoreName}`);
    }
  };

  const handleAcceptReconciliation = (id: string) => {
    setReconciliations(prev => prev.map(r => (r.id === id ? { ...r, status: 'Accepted' } : r)));
    stampAudit(`Reconciliation item ${id} marked accepted`);
    notify(`Reconciliation variance marked as accepted`);
  };

  const handleAdjustReconciliation = (row: ReconciliationRow, reason: string) => {
    const adjMov: StockMovement = {
      id: `MOV-RECON-${Date.now()}`,
      dateTime: new Date().toISOString().replace('T', ' ').slice(0, 16),
      storeId: row.storeId,
      storeName: row.storeName,
      productId: row.productId,
      productName: row.productName,
      batch: row.batch,
      expiry: row.expiry,
      qty: row.difference,
      unit: 'Tab',
      type: 'Adjustment',
      ref: `RECON:${row.snapshotDate}`,
      user: userName,
      source: 'App Billing',
      reason: `Reconciliation adjustment: ${reason}`
    };
    setMovements(prev => [adjMov, ...prev]);
    setReconciliations(prev => prev.map(r => (r.id === row.id ? { ...r, status: 'Adjusted' } : r)));
    stampAudit(`Reconciliation adjustment applied for ${row.productName} (${row.difference > 0 ? '+' : ''}${row.difference})`);
    notify(`Adjustment movement posted. System stock now aligns with ERP snapshot.`);
  };

  const handleAddMovement = (m: StockMovement) => {
    setMovements(prev => [m, ...prev]);
    stampAudit(`Stock movement ${m.type} posted: ${m.productName} (${m.batch}) ${m.qty > 0 ? '+' : ''}${m.qty}`);
    notify(`Stock ledger updated: ${m.type} recorded.`);
  };

  const handleAddCustomer = (c: Customer) => {
    setCustomers(prev => [...prev, c]);
    stampAudit(`Customer added: ${c.name} (${c.type})`);
    notify(`Customer ${c.name} registered locally.`);
  };

  // ----------------------------------------------------
  // EXCEL SAMPLES & EXPORTS
  // ----------------------------------------------------
  const downloadWorkbook = (type: string) => {
    const rows =
      type === 'Sales'
        ? products.slice(0, 8).map((p, i) => ({ 'Product code': p.id, 'Product name': p.name, Store: 'Main Store', Month: '2025-06', 'Sales qty': 12 + i, Returns: i % 3 }))
        : type === 'Stock'
        ? products.slice(0, 8).map(p => ({ 'Product code': p.id, 'Product name': p.name, Store: 'Main Store', 'As of': '2025-06-30', 'Stock qty': getStock(p.id, 'S1') }))
        : type === 'Purchase'
        ? products.slice(0, 8).map((p, i) => ({ 'Product code': p.id, 'Product name': p.name, Supplier: 'Sun Distributors', Voucher: `INV-2025-${1041 + i}`, 'Purchase date': '2025-06-12', Qty: 30 + i, MRP: 68.5 + i }))
        : products.slice(0, 8).map((p, i) => ({ 'Product code': p.id, 'Product name': p.name, Store: 'Main Store', Month: '2025-06', 'Return qty': i + 1 }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, type);
    XLSX.writeFile(wb, `sample-${type.toLowerCase().replaceAll(' ', '-')}.xlsx`);
  };
  const downloadAllSamples = () => ['Sales', 'Purchase', 'Sales Return', 'Stock'].forEach((x, i) => window.setTimeout(() => downloadWorkbook(x), i * 450));

  const exportXlsx = (filename: string, sheetName: string, rows: Record<string, unknown>[]) => {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), sheetName);
    XLSX.writeFile(wb, filename);
  };

  const exportProducts = () =>
    exportXlsx(
      'filtered-products.xlsx',
      'Products',
      filteredProducts.map(p => {
        const s = productStats(p);
        const last = seedPurchases.filter(x => x.productId === p.id).sort((a, b) => b.date.localeCompare(a.date))[0];
        return {
          'Product code': p.id,
          'Product name': p.name,
          Strength: p.strength,
          Pack: p.pack,
          Group: p.group,
          'Ledger Stock': s.qty,
          'Stock as of': '2025-06-30',
          'Last-month net sales': s.last,
          '3-month average': Number(s.avg.toFixed(2)),
          'Stock cover months': s.cover === null ? 'No recent sales' : Number(s.cover.toFixed(2)),
          'Last supplier': seedSuppliers.find(x => x.id === last?.supplierId)?.name || '',
          'Last purchase date': last?.date || ''
        };
      })
    );

  const exportSupplierWorkbook = (supplierId?: string) => {
    if (!supplierId) {
      const rows = seedSuppliers
        .filter(s => s.name.toLowerCase().includes(supplierSearch.toLowerCase()))
        .map(s => {
          const related = products.filter(p => seedPurchases.some(x => x.productId === p.id && x.supplierId === s.id));
          return {
            'Supplier code': s.id,
            Supplier: s.name,
            'Products supplied': related.length,
            'Current stock across products': related.reduce((n, p) => n + productStats(p).qty, 0),
            'Most recent purchase': seedPurchases.filter(x => x.supplierId === s.id).sort((a, b) => b.date.localeCompare(a.date))[0]?.date || ''
          };
        });
      exportXlsx('filtered-suppliers.xlsx', 'Suppliers', rows);
      return;
    }
    const related = products.filter(p => seedPurchases.some(x => x.productId === p.id && x.supplierId === supplierId));
    const rows = related.map(p => {
      const recent = seedPurchases.filter(x => x.productId === p.id && x.supplierId === supplierId).sort((a, b) => b.date.localeCompare(a.date))[0];
      const prices = Object.fromEntries(
        seedSuppliers.map(s => {
          const latest = seedPurchases.filter(x => x.productId === p.id && x.supplierId === s.id).sort((a, b) => b.date.localeCompare(a.date))[0];
          return [`Last MRP · ${s.name}`, latest?.mrp ?? ''];
        })
      );
      return {
        'Product code': p.id,
        Product: p.name,
        Stock: productStats(p).qty,
        'Last-month net sales': productStats(p).last,
        '3-month average': Number(productStats(p).avg.toFixed(2)),
        'Supplier purchase qty': recent?.qty ?? '',
        'Last purchase date': recent?.date ?? '',
        ...prices
      };
    });
    exportXlsx(`${seedSuppliers.find(s => s.id === supplierId)?.name.toLowerCase().replaceAll(' ', '-') || 'supplier'}-product-details.xlsx`, 'Products & supplier MRP', rows);
  };

  const processFile = async (file: File, override = false) => {
    if (role !== 'Admin' && role !== 'Manager') {
      notify('Not available for this role');
      return;
    }
    const duplicate = batches.find(b => b.file === file.name);
    if (duplicate && !override) {
      setModal({
        title: 'Duplicate file detected',
        body: (
          <div>
            <p><b>{file.name}</b> already appears in batch history.</p>
            <p className="small-note">Block this upload, or override as Admin to create a separate batch.</p>
            <div className="modal-actions">
              <Button onClick={() => { setModal(null); notify('Duplicate upload blocked'); }}>Block upload</Button>
              <Button kind="primary" disabled={role !== 'Admin'} title={role !== 'Admin' ? 'Not available for this role' : ''} onClick={() => { setModal(null); void processFile(file, true); }}>Override as Admin</Button>
            </div>
          </div>
        )
      });
      return;
    }
    let wb: XLSX.WorkBook;
    try {
      wb = XLSX.read(await file.arrayBuffer(), { type: 'array' });
    } catch {
      const fail: Batch = {
        id: `IB-${Date.now()}`,
        file: file.name,
        type: 'Unrecognised',
        status: 'Failed',
        rows: 0,
        loaded: 0,
        rejected: 0,
        uploadedBy: userName,
        time: time(),
        errors: [{ row: 0, column: 'Workbook', message: 'Could not read this file as an Excel workbook. Upload a valid .xlsx file.' }]
      };
      setBatches(x => [fail, ...x]);
      return;
    }
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: '' });
    const headers = Object.keys(rows[0] || {}).map(x => x.toLowerCase());
    const identify = () => {
      if (headers.some(x => /return/.test(x))) return 'Sales Return';
      if (headers.some(x => /supplier|voucher/.test(x)) && headers.some(x => /qty|quantity/.test(x))) return 'Purchase';
      if (headers.some(x => /stock|balance/.test(x))) return 'Stock';
      if (headers.some(x => /sales/.test(x))) return 'Sales';
      return 'Unrecognised';
    };
    const type = identify();
    const required =
      type === 'Stock' ? ['product', 'store', 'stock', 'as of'] : type === 'Purchase' ? ['product', 'supplier', 'qty', 'date'] : type === 'Sales Return' ? ['product', 'store', 'month', 'return'] : type === 'Sales' ? ['product', 'store', 'month', 'sales'] : [];
    const errs: { row: number; column: string; message: string }[] = [];
    if (type === 'Unrecognised') {
      errs.push({ row: 1, column: 'Headers', message: 'Could not identify report type from the column headers. Include product plus Sales, Stock, Supplier/Qty or Return columns.' });
    } else {
      required.filter(x => !headers.some(h => h.includes(x))).forEach(x => errs.push({ row: 1, column: x, message: `Required column containing “${x}” is missing` }));
    }
    const requiredHeaders = required.map(token => headers.find(h => h.includes(token))).filter((x): x is string => !!x);
    rows.forEach((row, index) => {
      for (const key of Object.keys(row)) {
        const value = row[key];
        const lower = key.toLowerCase();
        if (requiredHeaders.includes(lower) && (value === null || value === undefined || String(value).trim() === '')) {
          errs.push({ row: index + 2, column: key, message: 'Required value is empty' });
        }
        if (/qty|quantity|sales|stock|return|mrp|price|amount/i.test(key) && value !== '' && value !== null && value !== undefined && ((typeof value === 'string' && value.trim() === '') || !Number.isFinite(Number(value)))) {
          errs.push({ row: index + 2, column: key, message: 'Expected a numeric value' });
        }
      }
    });

    // Check for post go-live sales duplicate warning
    if (type === 'Sales') {
      errs.push({
        row: 1,
        column: 'Source of Truth Check',
        message: `ERP sales imported after Go-Live Date (${goLiveDate}) are flagged as "Possible duplicate of app bills" and are not double deducted.`
      });
    }

    const loaded = Math.max(0, rows.length - errs.filter(e => e.row > 1).length);
    const batch: Batch = {
      id: `IB-${Date.now()}`,
      file: file.name,
      type,
      status: type === 'Unrecognised' ? 'Failed' : errs.length ? 'Completed with notes' : 'Queued',
      rows: rows.length,
      loaded: type === 'Unrecognised' ? 0 : loaded,
      rejected: errs.filter(e => e.row > 1).length,
      uploadedBy: userName,
      time: time(),
      errors: errs
    };
    setBatches(x => [batch, ...x]);
    setSelectedBatch(batch);
    if (type !== 'Unrecognised') {
      setImportProgress(batch.id);
      let step = 0;
      const stages = ['Parse', 'Validate', 'Normalize', 'Match', 'Load'];
      const tick = () => {
        step++;
        if (step < stages.length) {
          setTimeout(tick, 500);
        } else {
          setBatches(old => old.map(b => (b.id === batch.id ? { ...b, status: errs.length ? 'Completed with notes' : 'Completed', loaded: rows.length - errs.filter(e => e.row > 1).length } : b)));
          setImportProgress('');
          stampAudit(`Import processed: ${file.name}`);
          notify(`Processed ${rows.length} rows with ${errs.length} validation notes`);
        }
      };
      setTimeout(tick, 400);
    }
  };

  const filesInput = useRef<HTMLInputElement>(null);
  const onFiles = (e: ChangeEvent<HTMLInputElement>) => {
    Array.from(e.target.files || []).forEach(f => void processFile(f));
    e.target.value = '';
  };
  const onDrop = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setDrop(false);
    Array.from(e.dataTransfer.files).forEach(f => void processFile(f));
  };

  // ----------------------------------------------------
  // ENHANCED ASK (SIMULATED NL QUERY)
  // Supports new prompt examples:
  // - "Show stock of Dolo 650 in Branch 2"
  // - "Today's sales for Main Store"
  // - "Items expiring in 60 days"
  // ----------------------------------------------------
  const calcFilter = (text: string) => {
    const l = text.trim().toLowerCase();
    setQuery(text);
    setAskUnsupported('');
    setAskResult(null);
    setAskFilter(null);
    setAskJsonPayload(null);
    if (!l) return;

    if (/fever|headache|pain|cough|cold|infection|disease|symptom|illness|cure|dosage|what medicine|recommend|prescribe|should i give|take for/.test(l)) {
      setAskUnsupported('This prototype only queries existing pharmacy business and inventory data. It gives no medical advice, symptom evaluation, or medicine recommendations.');
      return;
    }
    if (/predict|forecast|trend|projection|next month|sales next|future/.test(l)) {
      setAskUnsupported('Forecasting and future sales predictions are outside the scope of this query tool. The tool filters existing historical sales and current stock snapshot data only.');
      return;
    }

    // 1. "Show stock of Dolo 650 in Branch 2"
    if ((/dolo|paracetamol 650/i.test(l) && /branch 2/i.test(l)) || (/stock of dolo 650 in branch 2/i.test(l))) {
      const p = products.find(x => x.id === 'P002')!;
      const branch2Stock = getStock('P002', 'S2');
      const b2Batches = seedBatches.filter(b => b.productId === 'P002' && b.storeId === 'S2');
      setAskResult([p]);
      setAskFilter({ product: 'Paracetamol 650 mg', storeNote: 'Branch 2 (Indiranagar)', raw: text });
      setAskJsonPayload({
        parser: 'Deterministic Keyword Rule Parser (Simulated AI Gateway)',
        naturalQuery: text,
        queryIntent: 'STOCK_QUERY_BY_STORE',
        product: 'Paracetamol 650 mg (Dolo 650)',
        store: 'Branch 2 (Indiranagar)',
        currentStockInBranch2: `${branch2Stock} units`,
        batchBreakdown: b2Batches.map(b => ({ batch: b.batch, expiry: b.expiry, stock: b.stock })),
        status: 'VALID_BUSINESS_FILTER'
      });
      setAskJson(true);
      return;
    }

    // 2. "Today's sales for Main Store"
    if (/today.*sales.*main store/i.test(l) || /main store.*sales/i.test(l)) {
      const s1Invoices = invoices.filter(i => i.storeId === 'S1' && i.status !== 'Cancelled');
      const tot = s1Invoices.reduce((sum, i) => sum + i.grandTotal, 0);
      setAskResult(products.slice(0, 5));
      setAskFilter({ storeNote: 'Main Store (MG Road)', raw: text });
      setAskJsonPayload({
        parser: 'Deterministic Keyword Rule Parser (Simulated AI Gateway)',
        naturalQuery: text,
        queryIntent: 'TODAYS_SALES_BY_STORE',
        store: 'Main Store (MG Road)',
        todayDate: '30 Jun 2025',
        totalSalesValue: money(tot),
        billsGenerated: s1Invoices.length,
        invoices: s1Invoices.map(i => ({ no: i.invoiceNo, customer: i.customerName, total: money(i.grandTotal), mode: i.paymentMode })),
        status: 'VALID_BUSINESS_FILTER'
      });
      setAskJson(true);
      return;
    }

    // 3. "Items expiring in 60 days"
    if (/expir(ing|y).*60/i.test(l) || /expiring in 60 days/i.test(l)) {
      const expBatches = seedBatches.filter(b => b.expiry <= '2025-08-31');
      const prodIds = Array.from(new Set(expBatches.map(b => b.productId)));
      const matchingProds = products.filter(p => prodIds.includes(p.id));
      setAskResult(matchingProds);
      setAskFilter({ raw: text });
      setAskJsonPayload({
        parser: 'Deterministic Keyword Rule Parser (Simulated AI Gateway)',
        naturalQuery: text,
        queryIntent: 'EXPIRING_INVENTORY_FILTER',
        window: 'Next 60 days (≤ 31 Aug 2025)',
        affectedBatches: expBatches.map(b => ({
          product: b.productName,
          batch: b.batch,
          store: b.storeName,
          expiry: b.expiry,
          availableStock: b.stock
        })),
        status: 'VALID_BUSINESS_FILTER'
      });
      setAskJson(true);
      return;
    }

    // Standard business data filters
    let sup: string | undefined = seedSuppliers.find(s => l.includes(s.name.toLowerCase()))?.name;
    let prod: string | undefined;
    let showSup: string | undefined;
    if (/dolo.? ?650|suppliers? (supplied|for) dolo/.test(l)) {
      prod = 'Paracetamol 650 mg';
      showSup = 'P002';
    } else if (/paracetamol|pcm/.test(l)) {
      prod = 'Paracetamol';
    } else if (/amoxyclav|amoxicillin/.test(l)) {
      prod = 'Amoxicillin';
    } else if (/metformin/.test(l)) {
      prod = 'Metformin';
    } else if (/cetirizine/.test(l)) {
      prod = 'Cetirizine';
    } else if (/vitamin d3/.test(l)) {
      prod = 'Vitamin D3';
    }

    let cv: string | undefined;
    if (/less than 1 month|under 1 month|cover < 1|< 1 month|low stock/.test(l)) {
      cv = 'under';
    } else if (/stock > 2|above 2x|overstocked/.test(l)) {
      cv = 'above';
    } else if (/no sales|zero sales/.test(l)) {
      cv = 'no-sales';
    }

    let stk: string | undefined;
    if (/stock above 100|stock > 100/.test(l)) {
      stk = 'above-100';
    } else if (/zero stock|out of stock/.test(l)) {
      stk = 'zero';
    } else if (/stock below|less than 20 units/.test(l)) {
      stk = 'below';
    }

    if (!sup && !prod && !cv && !stk && !showSup) {
      setAskUnsupported(
        'I could not map that to a supported business-data filter. Try one of the example queries below: “Show stock of Dolo 650 in Branch 2”, “Today’s sales for Main Store”, or “Items expiring in 60 days”.'
      );
      return;
    }

    const currentFilter = { supplier: sup, cover: cv, stock: stk, product: prod, raw: text, showSuppliersFor: showSup };
    setAskFilter(currentFilter);
    const payload = {
      parser: 'Deterministic Keyword Rule Parser (Simulated AI Gateway)',
      naturalQuery: text,
      interpretedAt: new Date().toISOString(),
      filterRules: {
        productMatch: prod || null,
        supplierFilter: sup || null,
        stockCoverMode: cv === 'under' ? '< 1.0 month cover' : cv === 'above' ? '> 2× 3-month average' : cv === 'no-sales' ? 'No sales in 3 months' : null,
        stockQuantityMode: stk === 'above-100' ? 'Stock > 100 units' : stk === 'zero' ? 'Stock = 0' : stk === 'below' ? 'Stock < 20 units' : null,
        supplierLookupMode: showSup ? 'Lookup suppliers for product' : null
      },
      safetyGuardrails: { medicalAdvice: false, prescriptiveInterpretation: false, salesForecasting: false },
      status: 'VALID_BUSINESS_FILTER'
    };
    setAskJsonPayload(payload);

    let res = [...products];
    if (prod) res = res.filter(p => `${p.name} ${p.aliases.join(' ')}`.toLowerCase().includes(prod!.toLowerCase()));
    if (sup) {
      const sid = seedSuppliers.find(s => s.name === sup)?.id;
      res = res.filter(p => seedPurchases.some(x => x.productId === p.id && x.supplierId === sid));
    }
    if (cv === 'under') res = res.filter(p => { const s = productStats(p); return s.avg > 0 && s.cover !== null && s.cover < threshold; });
    else if (cv === 'above') res = res.filter(p => { const s = productStats(p); return s.avg > 0 && s.cover !== null && s.qty > s.avg * 2; });
    else if (cv === 'no-sales') res = res.filter(p => { const s = productStats(p); return s.avg === 0 && s.qty > 100; });
    if (stk === 'above-100') res = res.filter(p => productStats(p).qty > 100);
    else if (stk === 'zero') res = res.filter(p => productStats(p).qty === 0);
    else if (stk === 'below') res = res.filter(p => productStats(p).qty < 20);

    setAskResult(res);
    setAskJson(true);
  };

  const simulateMessage = () => {
    if (!msgInput.trim()) return;
    const lines = msgInput.split('\n').filter(Boolean);
    const item = {
      id: `W${Date.now()}`,
      sender: 'Simulated incoming message',
      time: time(),
      status: 'Needs Review',
      type: 'Text',
      text: msgInput,
      audit: [`${time()} Received in prototype inbox`, `${time()} Parsed ${lines.length} line${lines.length === 1 ? '' : 's'}; pharmacist confirmation required`]
    };
    setMessages(m => [item, ...m]);
    setSelectedMessage(item.id);
    setMsgInput('');
    notify('Simulated message added to inbox');
  };

  const parseLines = (text: string) =>
    text.split('\n').filter(Boolean).map(line => {
      const qty = Number(line.match(/\d+/)?.[0] || 0);
      const clean = line.replace(/[-x×]?\s*\d+\s*(strip|strips|tab|tabs|tablet|tablets|pack|packs)?/gi, '').trim();
      const exact = products.find(p => p.aliases.some(a => a.toLowerCase() === clean.toLowerCase()) || p.name.toLowerCase() === clean.toLowerCase() || p.name.toLowerCase().startsWith(clean.toLowerCase()));
      const partial = !exact && products.find(p => p.aliases.some(a => a.toLowerCase().includes(clean.toLowerCase())) || p.name.toLowerCase().includes(clean.toLowerCase()) || (clean.length >= 3 && p.name.toLowerCase().includes(clean.slice(0, 3).toLowerCase())));
      const status: string = exact ? 'Matched' : partial ? 'Needs review' : 'Unknown';
      return { raw: line, qty, product: exact ? exact.name : partial ? partial.name : '', status };
    });

  // ----------------------------------------------------
  // SCREEN: IMPORTS (WITH RECONCILIATION & GO-LIVE WARNING)
  // ----------------------------------------------------
  const importPage = () => (
    <>
      <PageHead
        title="Imports"
        help="Inspect Excel workbooks, validate report rows and review simulated reconciliation batches."
        actions={
          <Button onClick={downloadAllSamples}>
            <ArrowDownToLine size={15} />
            Download 4 sample workbooks
          </Button>
        }
      />

      {/* Prominent Source-of-Truth banner required by section 1 & 4 */}
      <div className="notice warn" style={{ marginBottom: 16, display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <div>
          <strong>Source of Truth Rule:</strong> Sales made through this app deduct stock immediately. ERP sales imports dated on or after <b>{goLiveDate}</b> (Go-live date) are flagged <i>“Possible duplicate of app bills”</i> and are not double deducted.
        </div>
        <Button tiny kind="primary" onClick={() => setLocation('/store-stock')}>
          <RefreshCw size={13} /> Reconciliation available ({reconciliations.filter(r => r.status === 'Pending').length} variances)
        </Button>
      </div>

      <div className="content-grid">
        <div>
          <div className="dropzone" onDragOver={e => { e.preventDefault(); setDrop(true); }} onDragLeave={() => setDrop(false)} onDrop={onDrop} style={{ background: drop ? '#eaf3ed' : undefined }}>
            <Upload size={22} />
            <strong>Drop Excel files here, or browse</strong>
            <p>Multiple .xlsx files supported · actual workbook parsing in your browser</p>
            <Button kind="primary" onClick={() => filesInput.current?.click()}>
              <Plus size={15} />
              Choose files
            </Button>
            <input ref={filesInput} hidden type="file" multiple accept=".xlsx,.xls" onChange={onFiles} />
          </div>

          {importProgress && (
            <Panel title="Processing batch" sub="Simulated reconciliation stages">
              <div style={{ padding: 18 }}>
                <div className="stepper">
                  {['Parse', 'Validate', 'Normalize', 'Match', 'Load'].map((x, i) => (
                    <div className={`step ${i < 3 ? 'done' : ''}`} key={x}>
                      <b>{i < 3 ? '✓' : i + 1}</b>
                      {x}
                      {i < 4 && <div className="step-line" />}
                    </div>
                  ))}
                </div>
                <div className="progress">
                  <i style={{ width: '64%' }} />
                </div>
              </div>
            </Panel>
          )}

          <div className="notice info section-gap">
            <b>Simulated in prototype.</b> Excel parsing and validation run locally in this browser. Nothing is sent to an external server.
          </div>

          <Panel title="Import history" sub={`${batches.length} sample and uploaded batches · select a row to inspect`}>
            <div className="table-wrap">
              <table className="data-table">
                <thead>
                  <tr>
                    <th>File</th>
                    <th>Type</th>
                    <th>Status</th>
                    <th>Rows loaded / notes</th>
                    <th>Uploaded by</th>
                    <th>Time</th>
                  </tr>
                </thead>
                <tbody>
                  {batches.map(b => (
                    <tr key={b.id} onClick={() => setSelectedBatch(b)}>
                      <td className="product-cell">{b.file}</td>
                      <td>{b.type}</td>
                      <td><Status>{b.status}</Status></td>
                      <td>{b.loaded} / {b.rejected}</td>
                      <td>{b.uploadedBy}</td>
                      <td>{b.time}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </Panel>
        </div>

        <aside>
          <Panel title="Recognised report types" sub="Header keywords guide report identification">
            <div className="data-list">
              {['Sales', 'Purchase', 'Sales Return', 'Stock'].map(x => (
                <div className="list-row" key={x}>
                  <div>
                    <div className="row-primary">{x}</div>
                    <div className="row-secondary">Expected columns include product and {x.toLowerCase()} details</div>
                  </div>
                  <FileSpreadsheet size={17} color="#63877a" />
                </div>
              ))}
            </div>
          </Panel>
          <div className="notice section-gap">
            <b>Duplicate files are checked by filename.</b> Admin can override a duplicate upload. Other roles can block it.
          </div>
        </aside>
      </div>

      {selectedBatch && (
        <div className="drawer-shade" onMouseDown={e => { if (e.target === e.currentTarget) setSelectedBatch(null); }}>
          <div className="drawer">
            <div className="drawer-head">
              <h2>Batch details</h2>
              <button className="close-button" onClick={() => setSelectedBatch(null)}>×</button>
            </div>
            <div className="eyebrow">{selectedBatch.id}</div>
            <h3>{selectedBatch.file}</h3>
            <Status>{selectedBatch.status}</Status>
            <div className="stats-grid" style={{ gridTemplateColumns: 'repeat(3,1fr)', marginTop: 18 }}>
              {[['Rows read', selectedBatch.rows], ['Loaded', selectedBatch.loaded], ['Notes/Rejected', selectedBatch.rejected]].map(([l, v]) => (
                <div className="panel stat-card" key={l}>
                  <div className="stat-label">{l}</div>
                  <div className="stat-value">{v}</div>
                </div>
              ))}
            </div>
            <p className="small-note">Report type: {selectedBatch.type} · Uploaded by {selectedBatch.uploadedBy} at {selectedBatch.time}</p>
            <Panel title="Validation errors & notes" sub="Row-level findings">
              <div className="data-list">
                {selectedBatch.errors.length ? (
                  selectedBatch.errors.map((e, i) => (
                    <div className="list-row" key={i}>
                      <div>
                        <div className="row-primary">Row {e.row} · {e.column}</div>
                        <div className="row-secondary">{e.message}</div>
                      </div>
                    </div>
                  ))
                ) : (
                  <div className="empty-state">No validation errors recorded.</div>
                )}
              </div>
            </Panel>
            <div className="modal-actions">
              <Button onClick={() => dl(`${selectedBatch.file}-errors.csv`, ['row,column,message', ...selectedBatch.errors.map(e => `${e.row},${e.column},"${e.message}"`)].join('\n'), 'text/csv')}>
                <ArrowDownToLine size={14} />
                Download notes
              </Button>
              <Button kind="primary" onClick={() => { setSelectedBatch(null); filesInput.current?.click(); }}>
                Retry / Reprocess
              </Button>
            </div>
          </div>
        </div>
      )}
    </>
  );

  // ----------------------------------------------------
  // SCREEN: PRODUCTS
  // ----------------------------------------------------
  const productPage = () => {
    const pageSize = 15;
    const pageCount = Math.max(1, Math.ceil(filteredProducts.length / pageSize));
    const shown = filteredProducts.slice(productPageIndex * pageSize, (productPageIndex + 1) * pageSize);
    return (
      <>
        <PageHead
          title="Products"
          help="Search and filter inventory, compare ledger stock with net sales, and inspect product details."
          actions={
            <>
              <Button onClick={exportProducts}><ArrowDownToLine size={14} />Export filtered Excel</Button>
              <Button onClick={() => window.print()}><Printer size={14} />Print / PDF</Button>
            </>
          }
        />
        <div className="notice info" style={{ marginBottom: 13 }}>
          Stock derived dynamically from <b>Stock Ledger (Opening + Movements)</b> as of <b>{today}</b> · 3-month average combines ERP imports + App Billing sales.
        </div>
        <div className="panel">
          <div className="toolbar">
            <div className="global-search" style={{ width: 220 }}>
              <Search size={15} />
              <input data-testid="product-search" placeholder="Search names or aliases" value={filter} onChange={e => setFilter(e.target.value)} />
            </div>
            <select className="control" value={groupFilter} onChange={e => setGroupFilter(e.target.value)}>
              <option value="">All groups</option>
              {Array.from(new Set(products.map(p => p.group))).map(g => <option key={g}>{g}</option>)}
            </select>
            <select className="control" value={supplierFilter} onChange={e => setSupplierFilter(e.target.value)}>
              <option value="">All suppliers</option>
              {seedSuppliers.map(s => <option key={s.id}>{s.name}</option>)}
            </select>
            <select className="control" value={storeFilter} onChange={e => setStoreFilter(e.target.value)}>
              <option value="">All stores</option>
              {storesMeta.map(s => <option key={s.id} value={s.id}>{s.name}</option>)}
            </select>
            <select className="control" value={stockMode} onChange={e => setStockMode(e.target.value)}>
              <option value="">Any stock level</option>
              <option value="zero">Zero stock</option>
              <option value="below">Below threshold</option>
              <option value="range">Threshold to 4× threshold</option>
            </select>
            {(stockMode === 'below' || stockMode === 'range') && (
              <label className="small-note">
                Stock threshold{' '}
                <input className="control" type="number" min="1" value={stockBelow} onChange={e => setStockBelow(Math.max(1, Number(e.target.value) || 1))} style={{ width: 76 }} />
              </label>
            )}
            <select className="control" value={coverFilter} onChange={e => setCoverFilter(e.target.value)}>
              <option value="">Any cover</option>
              <option value="under">Cover under {threshold} month</option>
              <option value="above">Stock &gt; 2× avg</option>
              <option value="no-sales">No recent sales; stock &gt; 100</option>
            </select>
            <Button tiny kind="ghost" onClick={() => setVisibleCols(!visibleCols)}>
              <Filter size={14} />{visibleCols ? 'Hide extra columns' : 'Show all columns'}
            </Button>
            <Button tiny kind="ghost" onClick={() => { setFilter(''); setGroupFilter(''); setStoreFilter(''); setSupplierFilter(''); setStockMode(''); setStockBelow(20); setCoverFilter(''); }}>
              Clear
            </Button>
          </div>
          <div className="table-wrap">
            <table className="data-table">
              <thead>
                <tr>
                  <th onClick={() => { setProductSort('name'); setSortReverse(!sortReverse); }}>Product ↕</th>
                  <th>Group</th>
                  <th>HSN / GST</th>
                  <th>Ledger Stock</th>
                  <th onClick={() => { setProductSort('last'); setSortReverse(!sortReverse); }}>Net sales (ERP+App) ↕</th>
                  <th>3-mo avg</th>
                  <th>Cover</th>
                  {visibleCols && (
                    <>
                      <th>Last supplier</th>
                      <th>Last purchase</th>
                    </>
                  )}
                </tr>
              </thead>
              <tbody>
                {shown.map(p => {
                  const s = productStats(p);
                  const meta = productGstMeta[p.id];
                  const po = seedPurchases.filter(x => x.productId === p.id).sort((a, b) => b.date.localeCompare(a.date))[0];
                  const cover = s.cover === null ? 'No recent sales' : `${s.cover.toFixed(1)} mo`;
                  return (
                    <tr key={p.id} onClick={() => setLocation(`/products/${p.id}`)}>
                      <td className="product-cell">
                        {p.name}
                        <div className="row-secondary">
                          {p.pack} {meta?.schedule && meta.schedule !== 'None' ? <span className="badge warn" style={{ fontSize: 9 }}>{meta.schedule}</span> : null}
                        </div>
                      </td>
                      <td>{p.group}</td>
                      <td><span className="mono" style={{ fontSize: 11 }}>{meta?.hsn || '3004'}</span> · {meta?.gstRate ?? 12}%</td>
                      <td><b>{s.qty} units</b></td>
                      <td>{s.last}</td>
                      <td>{s.avg.toFixed(1)}</td>
                      <td>
                        <Status>{s.cover === null ? 'No recent sales' : s.cover < threshold ? 'Low' : 'Sufficient'}</Status>
                        <div className="row-secondary">{cover}</div>
                      </td>
                      {visibleCols && (
                        <>
                          <td>{seedSuppliers.find(x => x.id === po?.supplierId)?.name || '—'}</td>
                          <td>{po?.date || '—'}</td>
                        </>
                      )}
                    </tr>
                  );
                })}
              </tbody>
            </table>
            {shown.length === 0 && (
              <div className="empty-state">
                <Package size={22} />
                <strong>No products match these filters</strong>
                <p>Clear one or more filters to see inventory.</p>
              </div>
            )}
          </div>
          <div className="panel-head">
            <span className="small-note">
              Showing {shown.length ? productPageIndex * pageSize + 1 : 0}–{Math.min((productPageIndex + 1) * pageSize, filteredProducts.length)} of {filteredProducts.length} results · Page {productPageIndex + 1} of {pageCount}
            </span>
            <span>
              <Button tiny disabled={productPageIndex === 0} onClick={() => setProductPageIndex(v => Math.max(0, v - 1))}>
                <ArrowLeft size={12} />Previous
              </Button>{' '}
              <Button tiny disabled={productPageIndex + 1 >= pageCount} onClick={() => setProductPageIndex(v => Math.min(pageCount - 1, v + 1))}>
                Next <ArrowRight size={12} />
              </Button>
            </span>
          </div>
        </div>
      </>
    );
  };

  // ----------------------------------------------------
  // SCREEN: PRODUCT DETAIL (WITH STORE-WISE STOCK TAB & LEDGER LINK)
  // ----------------------------------------------------
  const productDetail = () => {
    const [, params] = useRoute('/products/:id');
    const p = products.find(x => x.id === params?.id);
    if (!p) return null;
    const meta = productGstMeta[p.id];
    const st = productStats(p);

    // Sales history showing both ERP imports and App Billing with badges
    const appSalesRows = movements.filter(m => m.productId === p.id && m.type === 'Sale');

    return (
      <>
        <PageHead
          title={p.name}
          help="Product identity, real-time stock ledger across stores, GST metadata, and combined sales history."
          actions={
            <>
              <Button onClick={() => setLocation('/products')}><ArrowLeft size={14} />All products</Button>
              <Button disabled={role !== 'Admin'} title={role !== 'Admin' ? 'Not available for this role' : ''} onClick={() => setProductEdit(true)}>
                <Settings size={14} />Edit product
              </Button>
            </>
          }
        />

        <div className="notice info">
          Stock derived dynamically from <b>Stock Ledger (Opening + Movements)</b>. Current stock reflects all in-app bills, returns, transfers and write-offs.
        </div>

        <div className="stats-grid section-gap">
          {[
            ['Ledger Stock', `${st.qty} units`, `Across ${storeFilter ? 1 : 3} store${storeFilter ? '' : 's'}`],
            ['Last-month sales', `${st.last} units`, 'Gross sales less returns (ERP + App)'],
            ['3-month average', `${st.avg.toFixed(1)} units`, 'Net sales ÷ 3'],
            ['Stock cover', st.avg ? `${st.cover?.toFixed(1)} months` : 'No recent sales', st.avg && st.cover! < threshold ? 'Low cover' : 'Calculated from on-hand']
          ].map(([l, v, f]) => (
            <div className="panel stat-card" key={l}>
              <div className="stat-label">
                {l}
                <span title="Formula: net sales = sales − returns + App sales. Average = 3-month net sales ÷ 3. Cover = stock ÷ average." style={{ marginLeft: 6, cursor: 'help' }}>ⓘ</span>
              </div>
              <div className="stat-value" style={{ fontSize: 21 }}>{v}</div>
              <div className="stat-foot">{f}</div>
            </div>
          ))}
        </div>

        <div className="two-col section-gap">
          <Panel title="Product GST & Identity" sub={`${p.group} · ${p.family}`} action={<Status>{st.avg === 0 ? 'No recent sales' : st.cover! < threshold ? 'Low' : 'Sufficient'}</Status>}>
            <div className="data-list">
              <div className="list-row"><span className="row-primary">Pack size</span><span>{p.pack}</span></div>
              <div className="list-row"><span className="row-primary">HSN Code</span><span className="mono">{meta?.hsn || '3004'}</span></div>
              <div className="list-row"><span className="row-primary">GST Rate</span><span><b>{meta?.gstRate ?? 12}%</b> (Illustrative: verify with CA)</span></div>
              <div className="list-row"><span className="row-primary">MRP</span><span>{money(meta?.mrp || 85)}</span></div>
              <div className="list-row">
                <span className="row-primary">Schedule Class</span>
                <span>
                  {meta?.schedule && meta.schedule !== 'None' ? (
                    <span className="badge warn">{meta.schedule} (Doctor name required for billing)</span>
                  ) : (
                    'General / OTC'
                  )}
                </span>
              </div>
              <div className="list-row"><span className="row-primary">Aliases</span><span>{p.aliases.join(' · ')}</span></div>
              <div className="list-row"><span className="row-primary">Product code</span><span className="mono">{p.id}</span></div>
            </div>
          </Panel>

          <Panel title="Store-wise Stock & Batch Inventory" sub="Physical stock across stores from the Stock Ledger" action={<Button tiny kind="primary" onClick={() => setLocation('/store-stock')}><Boxes size={13}/>Open Stock Ledger</Button>}>
            <div className="data-list">
              {storesMeta.map(s => {
                const storeStock = getStock(p.id, s.id);
                const batchesInStore = seedBatches.filter(b => b.productId === p.id && b.storeId === s.id);
                return (
                  <div className="list-row" key={s.id} style={{ display: 'grid', gridTemplateColumns: '1fr auto auto', gap: 12, alignItems: 'center' }}>
                    <div>
                      <div className="row-primary">{s.name}</div>
                      <div className="row-secondary">
                        {batchesInStore.map(b => `${b.batch} (${b.expiry})`).join(', ') || 'Standard batch'}
                      </div>
                    </div>
                    <div style={{ font: '700 15px Manrope' }}>{storeStock} units</div>
                    <Button tiny onClick={() => setLocation('/store-stock')}>View Ledger</Button>
                  </div>
                );
              })}
            </div>
          </Panel>
        </div>

        {/* Sales history with ERP Import vs App Billing source badges */}
        <Panel title="Combined Sales History (ERP Imports + App Billing)" sub="Displays origin source badge for each recorded transaction">
          <div className="table-wrap">
            <table className="data-table">
              <thead>
                <tr>
                  <th>Period / Ref</th>
                  <th>Store</th>
                  <th>Source</th>
                  <th>Quantity</th>
                  <th>Type</th>
                </tr>
              </thead>
              <tbody>
                {appSalesRows.slice(0, 5).map(m => (
                  <tr key={m.id}>
                    <td className="product-cell">{m.ref} ({m.dateTime})</td>
                    <td>{m.storeName}</td>
                    <td><span className="badge" style={{ background: '#eaf4ee', color: '#276749' }}>App Billing</span></td>
                    <td><b>{Math.abs(m.qty)} units</b></td>
                    <td>Sale (POS)</td>
                  </tr>
                ))}
                {seedSales.filter(x => x.productId === p.id).slice(-6).reverse().map(x => (
                  <tr key={`${x.storeId}-${x.month}`}>
                    <td className="product-cell">ERP Month: {x.month}</td>
                    <td>{storesMeta.find(s => s.id === x.storeId)?.name}</td>
                    <td><span className="badge blue">ERP Import</span></td>
                    <td>{x.sales} sold {x.returns > 0 ? `(${x.returns} returns)` : ''}</td>
                    <td>Monthly Net Batch</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </Panel>
      </>
    );
  };

  // ----------------------------------------------------
  // SCREEN: SUPPLIERS
  // ----------------------------------------------------
  const supplierPage = () => {
    const [, params] = useRoute('/suppliers/:id');
    const id = params?.id;
    const selected = seedSuppliers.find(x => x.id === id) || seedSuppliers.find(x => x.id === supplierSelection);
    if (id && selected) {
      const supplied = products.filter(p => seedPurchases.some(x => x.productId === p.id && x.supplierId === selected.id));
      const recent = seedPurchases.filter(x => x.supplierId === selected.id).sort((a, b) => b.date.localeCompare(a.date));
      const shared = supplied.filter(p => new Set(seedPurchases.filter(x => x.productId === p.id).map(x => x.supplierId)).size > 1);
      let filteredSupplied = supplied.filter(p => `${p.name} ${p.group} ${p.aliases.join(' ')}`.toLowerCase().includes(supProductSearch.toLowerCase()));
      if (supGroupFilter) filteredSupplied = filteredSupplied.filter(p => p.group === supGroupFilter);
      if (supStockMode === 'zero') filteredSupplied = filteredSupplied.filter(p => productStats(p).qty === 0);
      if (supStockMode === 'below') filteredSupplied = filteredSupplied.filter(p => productStats(p).qty < 20);
      if (supStockMode === 'range') filteredSupplied = filteredSupplied.filter(p => productStats(p).qty >= 20 && productStats(p).qty < 80);
      if (supCoverFilter === 'under') filteredSupplied = filteredSupplied.filter(p => { const s = productStats(p); return s.cover !== null && s.cover < threshold; });
      if (supCoverFilter === 'above') filteredSupplied = filteredSupplied.filter(p => { const s = productStats(p); return s.cover !== null && s.qty > s.avg * 2; });
      if (supCoverFilter === 'no-sales') filteredSupplied = filteredSupplied.filter(p => { const s = productStats(p); return s.avg === 0 && s.qty > 100; });
      return (
        <>
          <PageHead
            title={selected.name}
            help="Inspect products, stock and purchases; compare last recorded MRPs where multiple suppliers serve the same product."
            actions={
              <>
                <Button onClick={() => exportSupplierWorkbook(selected.id)}><ArrowDownToLine size={14} />Export supplier detail</Button>
                <Button onClick={() => window.print()}><Printer size={14} />Print / PDF</Button>
                <Button onClick={() => setLocation('/suppliers')}><ArrowLeft size={14} />All suppliers</Button>
              </>
            }
          />
          <div className="notice info" style={{ marginBottom: 14 }}>
            Purchase prices are historical reference values only, not a purchase recommendation. Stock as of <b>{today}</b>.
          </div>
          <Panel title="Products supplied" sub={`${filteredSupplied.length} of ${supplied.length} products match filters · stock as of ${today}`}>
            <div className="toolbar" style={{ padding: '10px 14px', borderBottom: '1px solid #f0eee8' }}>
              <div className="global-search" style={{ width: 220 }}>
                <Search size={15} />
                <input placeholder="Search supplied products" value={supProductSearch} onChange={e => setSupProductSearch(e.target.value)} />
              </div>
              <select className="control" value={supGroupFilter} onChange={e => setSupGroupFilter(e.target.value)}>
                <option value="">All groups</option>
                {Array.from(new Set(supplied.map(p => p.group))).map(g => <option key={g}>{g}</option>)}
              </select>
              <select className="control" value={supStockMode} onChange={e => setSupStockMode(e.target.value)}>
                <option value="">Any stock level</option>
                <option value="zero">Zero stock</option>
                <option value="below">Below 20 units</option>
                <option value="range">20 to 80 units</option>
              </select>
              <select className="control" value={supCoverFilter} onChange={e => setCoverFilter(e.target.value)}>
                <option value="">Any cover</option>
                <option value="under">Cover under {threshold} month</option>
                <option value="above">Stock &gt; 2× avg</option>
                <option value="no-sales">No recent sales; stock &gt; 100</option>
              </select>
              <Button tiny kind="ghost" onClick={() => { setSupProductSearch(''); setSupGroupFilter(''); setSupStockMode(''); setSupCoverFilter(''); }}>Clear</Button>
            </div>
            <div className="table-wrap">
              <table className="data-table">
                <thead>
                  <tr>
                    <th>Product</th>
                    <th>Stock</th>
                    <th>Last-month sales</th>
                    <th>3-month avg</th>
                    <th>Cover</th>
                    <th>Last purchase</th>
                    <th>MRP</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredSupplied.map(p => {
                    const buy = recent.find(x => x.productId === p.id);
                    const s = productStats(p);
                    const cover = s.cover === null ? 'No recent sales' : `${s.cover.toFixed(1)} mo`;
                    return (
                      <tr key={p.id} onClick={() => setLocation(`/products/${p.id}`)}>
                        <td className="product-cell">{p.name}<div className="row-secondary">{p.pack}</div></td>
                        <td>{s.qty} units</td>
                        <td>{s.last}</td>
                        <td>{s.avg.toFixed(1)}</td>
                        <td>
                          <Status>{s.cover === null ? 'No recent sales' : s.cover < threshold ? 'Low' : 'Sufficient'}</Status>
                          <div className="row-secondary">{cover}</div>
                        </td>
                        <td>{buy?.date || '—'}</td>
                        <td>{buy ? money(buy.mrp) : '—'}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
              {filteredSupplied.length === 0 && (
                <div className="empty-state">
                  <Package size={22} />
                  <strong>No products match these filters</strong>
                  <p>Clear or adjust filters above to view products supplied by {selected.name}.</p>
                </div>
              )}
            </div>
          </Panel>
        </>
      );
    }

    const visible = seedSuppliers.filter(s => s.name.toLowerCase().includes(supplierSearch.toLowerCase()));
    return (
      <>
        <PageHead
          title="Suppliers"
          help="Browse supplier relationships with stocked products and recorded purchases."
          actions={
            <>
              <Button onClick={() => exportSupplierWorkbook()}><ArrowDownToLine size={14} />Export filtered suppliers</Button>
              <Button onClick={() => window.print()}><Printer size={14} />Print / PDF</Button>
            </>
          }
        />
        <div className="toolbar panel">
          <div className="global-search" style={{ width: 300 }}>
            <Search size={15} />
            <input placeholder="Search suppliers" value={supplierSearch} onChange={e => setSupplierSearch(e.target.value)} />
          </div>
          <span className="small-note">Select a supplier to inspect products, purchase history and multi-supplier MRP comparisons.</span>
        </div>
        <div className="two-col section-gap">
          {visible.map(s => {
            const list = products.filter(p => seedPurchases.some(x => x.productId === p.id && x.supplierId === s.id));
            const latest = seedPurchases.filter(x => x.supplierId === s.id).sort((a, b) => b.date.localeCompare(a.date))[0];
            return (
              <button
                key={s.id}
                className="panel"
                style={{ textAlign: 'left', padding: 18 }}
                onClick={() => {
                  setSupplierSelection(s.id);
                  setLocation(`/suppliers/${s.id}`);
                }}
              >
                <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                  <Building2 size={18} color="#528775" />
                  <ArrowRight size={15} />
                </div>
                <h3 style={{ font: '700 14px Manrope', margin: '13px 0 5px' }}>{s.name}</h3>
                <div className="small-note">{list.length} products supplied · {list.reduce((n, p) => n + productStats(p).qty, 0)} units in stock</div>
                <div className="row-secondary">Most recent purchase: {latest?.date || 'No records'}</div>
              </button>
            );
          })}
        </div>
      </>
    );
  };

  // ----------------------------------------------------
  // SCREEN: MATCHING QUEUE
  // ----------------------------------------------------
  const matchingPage = () => (
    <>
      <PageHead title="Matching queue" help="Resolve uncertain names explicitly. The system never maps uncertain products automatically." />
      <div className="notice info" style={{ marginBottom: 14 }}>
        Every mapping below is awaiting a person. Selecting an existing product adds the supplied raw text as an alias; Create new product adds a new local master record; Skip leaves the source unresolved.
      </div>
      <div className="panel">
        <div className="table-wrap">
          <table className="data-table">
            <thead>
              <tr>
                <th>Raw text</th>
                <th>Source</th>
                <th>Possible matches</th>
                <th>Status</th>
                <th>Actions</th>
              </tr>
            </thead>
            <tbody>
              {matches.map(m => (
                <tr key={m.id}>
                  <td className="product-cell">{m.rawText}</td>
                  <td>{m.source}</td>
                  <td>{m.candidates.filter(Boolean).map(id => products.find(p => p.id === id)?.name || '').filter(Boolean).join(' · ') || 'No candidates'}</td>
                  <td><Status>{m.status}</Status></td>
                  <td>
                    {m.status === 'Pending' ? (
                      <>
                        <Button
                          tiny
                          onClick={() =>
                            setModal({
                              title: `Resolve “${m.rawText}”`,
                              body: (
                                <MatchResolve
                                  m={m}
                                  products={products}
                                  onChoose={id => {
                                    setProducts(old => old.map(p => (p.id === id ? { ...p, aliases: Array.from(new Set([...p.aliases, m.rawText])) } : p)));
                                    setMatches(old => old.map(q => (q.id === m.id ? { ...q, status: 'Resolved' } : q)));
                                    stampAudit(`Match resolved: ${m.rawText}`);
                                    setModal(null);
                                    notify('Alias saved. Future imports will map automatically.');
                                  }}
                                  onCreate={name => {
                                    const p: Product = {
                                      id: `P${String(products.length + 1).padStart(3, '0')}`,
                                      name,
                                      strength: 'Not specified',
                                      pack: 'Not specified',
                                      group: 'Uncategorised',
                                      family: 'Uncategorised',
                                      aliases: [m.rawText]
                                    };
                                    setProducts(old => [...old, p]);
                                    setMatches(old => old.map(q => (q.id === m.id ? { ...q, status: 'New product created' } : q)));
                                    stampAudit(`Product created from match: ${m.rawText}`);
                                    setModal(null);
                                    notify('New product created locally');
                                  }}
                                  onSkip={() => {
                                    setMatches(old => old.map(q => (q.id === m.id ? { ...q, status: 'Skipped' } : q)));
                                    setModal(null);
                                  }}
                                />
                              )
                            })
                          }
                        >
                          Resolve
                        </Button>
                        <Button
                          tiny
                          onClick={() => {
                            setMatches(old => old.map(q => (q.id === m.id ? { ...q, status: 'Skipped' } : q)));
                            notify('Item skipped; no product was linked');
                          }}
                        >
                          Skip
                        </Button>
                      </>
                    ) : null}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </>
  );

  // ----------------------------------------------------
  // SCREEN: ASK (SIMULATED NL QUERY)
  // ----------------------------------------------------
  const askPage = () => (
    <>
      <PageHead title="Ask" help="Turn supported business-data questions into editable filters. Rule-based parser only; no LLM connection." />
      <div className="panel" style={{ padding: 20 }}>
        <div className="eyebrow">ASK ABOUT YOUR SEEDED DATA</div>
        <div style={{ display: 'flex', gap: 9, margin: '12px 0' }}>
          <input
            className="control"
            style={{ flex: 1, fontSize: 13, padding: 12 }}
            placeholder="e.g. Show stock of Dolo 650 in Branch 2, Today's sales for Main Store, or Items expiring in 60 days"
            value={query}
            onChange={e => setQuery(e.target.value)}
            onKeyDown={e => {
              if (e.key === 'Enter') calcFilter(query);
            }}
          />
          <Button kind="primary" onClick={() => calcFilter(query)}>
            <Search size={14} />Run query
          </Button>
        </div>
        <div className="filter-chips">
          {[
            'Show stock of Dolo 650 in Branch 2',
            "Today's sales for Main Store",
            'Items expiring in 60 days',
            'Items with less than 1 month of stock',
            'Low stock from supplier Sun Distributors',
            'Which suppliers supplied Dolo 650?'
          ].map(x => (
            <button className="chip" key={x} onClick={() => { setQuery(x); calcFilter(x); }}>
              {x}
            </button>
          ))}
        </div>
      </div>

      {askUnsupported && (
        <div className="panel section-gap" style={{ padding: 19, borderColor: '#e8c9bd' }}>
          <div style={{ display: 'flex', gap: 10, alignItems: 'flex-start' }}>
            <AlertTriangle color="#a55248" />
            <div>
              <b>Unsupported question</b>
              <p className="help-line">{askUnsupported}</p>
            </div>
          </div>
        </div>
      )}

      {askFilter && (
        <div className="section-gap">
          <Panel
            title={askFilter.storeNote ? `Store Query: ${askFilter.storeNote}` : askFilter.showSuppliersFor ? 'Supplier lookup result' : 'Filtered inventory query'}
            sub={`${askResult ? askResult.length : 0} product(s) matching interpreted criteria`}
            action={
              <Button tiny onClick={() => setAskJson(!askJson)}>
                <ChevronDown size={13} />
                {askJson ? 'Hide' : 'Show'} interpreted JSON
              </Button>
            }
          >
            {askJson && askJsonPayload && (
              <pre style={{ margin: '14px 18px', background: '#213b40', color: '#dcf0ea', padding: 14, borderRadius: 8, fontSize: 11, lineHeight: 1.5, overflow: 'auto', fontFamily: 'monospace' }}>
                {JSON.stringify(askJsonPayload, null, 2)}
              </pre>
            )}

            {askResult && (
              <div className="table-wrap">
                <table className="data-table">
                  <thead>
                    <tr>
                      <th>Product</th>
                      <th>Group</th>
                      <th>Ledger Stock</th>
                      <th>Last-month sales</th>
                      <th>3-month average</th>
                      <th>Cover</th>
                    </tr>
                  </thead>
                  <tbody>
                    {askResult.slice(0, 30).map(p => {
                      const s = productStats(p);
                      const cover = s.cover === null ? 'No recent sales' : `${s.cover.toFixed(1)} mo`;
                      return (
                        <tr key={p.id} onClick={() => setLocation(`/products/${p.id}`)}>
                          <td className="product-cell">{p.name}<div className="row-secondary">{p.pack}</div></td>
                          <td>{p.group}</td>
                          <td>{s.qty} units</td>
                          <td>{s.last}</td>
                          <td>{s.avg.toFixed(1)}</td>
                          <td>
                            <Status>{s.cover === null ? 'No recent sales' : s.cover < threshold ? 'Low' : 'Sufficient'}</Status>
                            <div className="row-secondary">{cover}</div>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}
          </Panel>
        </div>
      )}

      <div className="notice info section-gap">
        <Sparkles size={14} /> Simulated in prototype · query parsing is deterministic keyword matching, not a language model. No medical advice, interpretation or forecasting.
      </div>
    </>
  );

  // ----------------------------------------------------
  // SCREEN: WHATSAPP INBOX
  // ----------------------------------------------------
  const whatsappPage = () => {
    const active = messages.find(x => x.id === selectedMessage) || messages[0];
    const parsed = parseLines(active.text);
    const actions = messageLineActions[active.id] || {};
    const actOnLine = (index: number, action: 'Confirmed' | 'Ignored') => {
      const next = { ...actions, [index]: action };
      setMessageLineActions(old => ({ ...old, [active.id]: next }));
      const completed = Object.keys(next).length >= parsed.length;
      const event = `${time()} · Line ${index + 1} ${action.toLowerCase()} by ${userName}`;
      setMessages(old => old.map(m => (m.id === active.id ? { ...m, status: completed ? 'Resolved' : 'Needs Review', audit: [...m.audit, event] } : m)));
      stampAudit(`WhatsApp ${active.id} line ${index + 1} ${action.toLowerCase()}`);
      notify(`Line ${action.toLowerCase()} in local review`);
    };

    return (
      <>
        <PageHead
          title="WhatsApp inbox"
          help="Review simulated incoming lists and confirm or ignore each parsed line locally."
          actions={
            <Button
              onClick={() =>
                setModal({
                  title: 'Simulate incoming message',
                  body: (
                    <div className="feedback-form">
                      <p className="small-note">Paste a product list. Lines are split locally and quantities are extracted with a basic rule.</p>
                      <textarea rows={5} className="control" value={msgInput} onChange={e => setMsgInput(e.target.value)} placeholder={'Paracetamol 500 - 3 strips\nORS - 2'} />
                      <div className="modal-actions">
                        <Button onClick={() => setModal(null)}>Cancel</Button>
                        <Button kind="primary" onClick={() => { simulateMessage(); setModal(null); }}>Add simulated message</Button>
                      </div>
                    </div>
                  )
                })
              }
            >
              <Plus size={15} />Simulate incoming message
            </Button>
          }
        />
        <div className="notice info" style={{ marginBottom: 14 }}>Single central WhatsApp number (simulated) · no messages are sent or received externally.</div>
        <div className="content-grid">
          <Panel title="Messages" sub="Local sample inbox">
            <div className="data-list">
              {messages.map(m => (
                <button
                  key={m.id}
                  className="list-row"
                  style={{ width: '100%', textAlign: 'left', background: selectedMessage === m.id ? '#f4f7f3' : 'transparent', border: 0, borderBottom: '1px solid #f0eee8' }}
                  onClick={() => setSelectedMessage(m.id)}
                >
                  <div>
                    <div className="row-primary">{m.sender}</div>
                    <div className="row-secondary">{m.time} · {m.type}</div>
                  </div>
                  <Status>{m.status}</Status>
                </button>
              ))}
            </div>
          </Panel>
          <div>
            <Panel title="Message detail" sub={`${active.sender} · ${active.time}`}>
              <div style={{ padding: '14px 18px', whiteSpace: 'pre-line', fontSize: 12, color: '#526268', background: '#f8f7f2', margin: 14, borderRadius: 8 }}>{active.text}</div>
              <div className="table-wrap">
                <table className="data-table">
                  <thead>
                    <tr>
                      <th>Raw line</th>
                      <th>Qty</th>
                      <th>Matched product</th>
                      <th>Status</th>
                      <th>Action</th>
                    </tr>
                  </thead>
                  <tbody>
                    {parsed.map((line, i) => {
                      const action = actions[i];
                      return (
                        <tr key={i}>
                          <td>{line.raw}</td>
                          <td>{line.qty || '—'}</td>
                          <td>{line.product || '—'}</td>
                          <td><Status>{action || line.status}</Status></td>
                          <td>
                            {line.status === 'Matched' ? (
                              <Button tiny disabled={!!action} onClick={() => actOnLine(i, 'Confirmed')}>
                                {action === 'Confirmed' ? 'Confirmed' : 'Confirm'}
                              </Button>
                            ) : null}
                            <Button tiny disabled={!!action} onClick={() => actOnLine(i, 'Ignored')}>
                              {action === 'Ignored' ? 'Ignored' : 'Ignore'}
                            </Button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </Panel>
          </div>
        </div>
      </>
    );
  };

  // ----------------------------------------------------
  // SCREEN: PRESCRIPTIONS (WITH "CREATE BILL" HANDOFF)
  // ----------------------------------------------------
  const prescriptionPage = () => {
    const [rxTab, setRxTab] = useState('Pending');
    const [rxSearch, setRxSearch] = useState('');
    const [selectedId, setSelectedId] = useState<string | null>(null);
    const [rxDialog, setRxDialog] = useState<'detail' | 'form' | null>(null);
    const [draft, setDraft] = useState<any>(null);
    const current = prescriptions.find(p => p.id === selectedId);

    const filtered = prescriptions.filter(p => {
      const q = rxSearch.toLowerCase();
      const matchesText = `${p.patient} ${p.phone} ${p.reference}`.toLowerCase().includes(q);
      const matchesTab = rxTab === 'Pending' ? ['Received', 'Review'].includes(p.status) : rxTab === 'Open' ? p.status !== 'Closed' : p.status === 'Closed';
      return matchesText && matchesTab;
    });

    if (role === 'Manager') {
      return (
        <>
          <PageHead title="Prescriptions" help="Prescription records are restricted for the Manager role." />
          <div className="notice"><b>Not available for this role.</b> Switch to Admin, Pharmacist or Store Owner.</div>
        </>
      );
    }

    return (
      <>
        <PageHead
          title="Prescriptions"
          help="Track received records through pharmacist review. Link prescriptions directly to counter GST bills."
          actions={
            <Button
              kind="primary"
              onClick={() => {
                setDraft({
                  patient: '',
                  phone: '',
                  address: '',
                  information: '',
                  dateTime: new Date().toISOString().slice(0, 16),
                  source: 'Walk-in',
                  reference: '',
                  status: 'Received',
                  assignedUser: userName,
                  linkedProducts: [],
                  history: []
                });
                setRxDialog('form');
              }}
            >
              <Plus size={15} />New prescription record
            </Button>
          }
        />

        <div className="notice warn" style={{ marginBottom: 14, fontWeight: 600 }}>
          Notice: The system does not interpret prescriptions or recommend or substitute medicines. The pharmacist evaluates dispensing. Billing does not suggest or substitute medicines.
        </div>

        <div className="toolbar panel">
          <div className="pill-group">
            {['Pending', 'Open', 'Closed'].map(t => (
              <button className={rxTab === t ? 'selected' : ''} key={t} onClick={() => setRxTab(t)}>
                {t}
              </button>
            ))}
          </div>
          <div className="global-search" style={{ width: 220 }}>
            <Search size={14} />
            <input placeholder="Patient, phone or reference" value={rxSearch} onChange={e => setRxSearch(e.target.value)} />
          </div>
        </div>

        <div className="panel section-gap">
          <div className="table-wrap">
            <table className="data-table">
              <thead>
                <tr>
                  <th>Ref / Patient</th>
                  <th>Phone</th>
                  <th>Date & time</th>
                  <th>Linked Invoice</th>
                  <th>Source</th>
                  <th>Status</th>
                  <th>Action</th>
                </tr>
              </thead>
              <tbody>
                {filtered.map(p => (
                  <tr key={p.id} onClick={() => { setSelectedId(p.id); setRxDialog('detail'); }}>
                    <td className="product-cell">{p.id}<div className="row-secondary">{p.patient}</div></td>
                    <td>{p.phone}</td>
                    <td>{p.dateTime.replace('T', ' ')}</td>
                    <td>{p.linkedInvoiceNo ? <span className="mono badge" style={{ background: '#e2f4ea', color: '#235937' }}>{p.linkedInvoiceNo}</span> : <span className="small-note">Not billed yet</span>}</td>
                    <td>{p.source}</td>
                    <td><Status>{p.status}</Status></td>
                    <td>
                      <div style={{ display: 'flex', gap: 6 }}>
                        <Button tiny onClick={e => { e.stopPropagation(); setSelectedId(p.id); setRxDialog('detail'); }}>Review</Button>
                        <Button
                          tiny
                          kind="primary"
                          onClick={e => {
                            e.stopPropagation();
                            setBillingPrefill({
                              prescriptionId: p.id,
                              patientName: p.patient,
                              phone: p.phone,
                              doctorName: 'Dr. Ramesh Kumar (MBBS)',
                              productIds: p.linkedProducts
                            });
                            setLocation('/billing');
                          }}
                        >
                          <Receipt size={12} />Create Bill
                        </Button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        {rxDialog === 'detail' && current && (
          <Modal title={`${current.id} · ${current.patient}`} close={() => setRxDialog(null)}>
            <div className="two-col">
              <div>
                <div className="row-secondary">PATIENT</div>
                <b>{current.patient}</b>
                <div className="row-secondary">{current.phone} · {current.address || 'No address'}</div>
              </div>
              <div>
                <div className="row-secondary">LINKED INVOICE</div>
                {current.linkedInvoiceNo ? (
                  <b>{current.linkedInvoiceNo}</b>
                ) : (
                  <span className="small-note">No bill created yet</span>
                )}
              </div>
            </div>
            <div className="notice info section-gap" style={{ whiteSpace: 'pre-wrap' }}>
              {current.information || 'No typed notes'}
            </div>
            <div className="modal-actions">
              <Button onClick={() => setRxDialog(null)}>Close</Button>
              <Button
                kind="primary"
                onClick={() => {
                  setBillingPrefill({
                    prescriptionId: current.id,
                    patientName: current.patient,
                    phone: current.phone,
                    doctorName: 'Dr. Ramesh Kumar (MBBS)',
                    productIds: current.linkedProducts
                  });
                  setRxDialog(null);
                  setLocation('/billing');
                }}
              >
                <Receipt size={14} />Create Bill for this Prescription
              </Button>
            </div>
          </Modal>
        )}
      </>
    );
  };

  // ----------------------------------------------------
  // SCREEN: ADMIN (ROLES, STORE MASTER, BILLING SETTINGS)
  // ----------------------------------------------------
  const adminPage = () => (
    <>
      <PageHead title="Admin" help="Store master data, GST billing configurations, role matrix, and prototype audit trail." />
      <div className="toolbar panel">
        <div className="pill-group">
          {['Users & roles', 'Stores Master', 'Billing & GST Settings', 'Audit log'].map(x => (
            <button className={adminTab === x ? 'selected' : ''} key={x} onClick={() => setAdminTab(x)}>
              {x}
            </button>
          ))}
        </div>
      </div>

      {adminTab === 'Users & roles' && (
        <>
          <Panel title="Configured Personas & Permissions" sub="Role switcher controls available views and actions">
            <div className="table-wrap">
              <table className="data-table">
                <thead>
                  <tr>
                    <th>Name</th>
                    <th>Role</th>
                    <th>Store Scope</th>
                    <th>Billing</th>
                    <th>Invoices & Cancel</th>
                    <th>Store Stock</th>
                    <th>Imports</th>
                    <th>GST Reports</th>
                  </tr>
                </thead>
                <tbody>
                  {users.map(u => (
                    <tr key={u[0]}>
                      <td className="product-cell">{u[0]}</td>
                      <td><Status>{u[1]}</Status></td>
                      <td>{u[1] === 'Admin' ? 'All stores' : 'Assigned store'}</td>
                      <td>{u[1] === 'Manager' ? 'Restricted' : u[1] === 'Staff' ? 'Draft bills' : 'Allowed'}</td>
                      <td>{u[1] === 'Admin' || u[1] === 'Store Owner' ? 'Full (Cancel & Credit Notes)' : 'View only'}</td>
                      <td>Allowed</td>
                      <td>{u[1] === 'Admin' || u[1] === 'Manager' ? 'Full' : 'Restricted'}</td>
                      <td>{u[1] === 'Staff' ? 'Restricted' : 'Allowed'}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </Panel>
          <div className="notice info section-gap">
            Admin: full access across all stores · Store Owner: full billing, adjustments, cancellation and reports for their store · Manager: reports, reconciliations, stock ledger, no billing · Pharmacist: billing prescription & scheduled items · Staff: draft bills and stock view.
          </div>
        </>
      )}

      {adminTab === 'Stores Master' && (
        <Panel title="Store Master Data & Series (Admin → Settings → Stores)" sub="Legal entity details, GSTIN, Drug Licence, and invoice number sequences">
          <div className="table-wrap">
            <table className="data-table">
              <thead>
                <tr>
                  <th>Store Code</th>
                  <th>Legal Name & Location</th>
                  <th>GSTIN</th>
                  <th>Drug Licence No.</th>
                  <th>Invoice Series Prefix</th>
                  <th>State</th>
                </tr>
              </thead>
              <tbody>
                {storesMeta.map(s => (
                  <tr key={s.id}>
                    <td className="mono">{s.id}</td>
                    <td className="product-cell">
                      {s.name}
                      <div className="row-secondary">{s.address}</div>
                    </td>
                    <td><span className="mono">{s.gstin}</span></td>
                    <td><span className="mono">{s.drugLicence}</span></td>
                    <td><span className="mono badge">{s.invoicePrefix}</span></td>
                    <td>{s.state} ({s.stateCode})</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div className="notice info section-gap">
            Sequential, unique invoice numbering resets each financial year (e.g. MS/2026-27/00001). Cancelled numbers remain in sequence as Cancelled for GST compliance.
          </div>
        </Panel>
      )}

      {adminTab === 'Billing & GST Settings' && (
        <div className="two-col">
          <Panel title="Tax Calculation Mode" sub="Configure whether rates are inclusive of MRP">
            <div style={{ padding: 18, display: 'grid', gap: 14 }}>
              <label style={{ display: 'flex', gap: 10, alignItems: 'center', cursor: 'pointer' }}>
                <input
                  type="checkbox"
                  checked={pricesInclusive}
                  onChange={e => {
                    setPricesInclusive(e.target.checked);
                    stampAudit(`Changed price tax mode to: ${e.target.checked ? 'MRP-Inclusive' : 'Exclusive'}`);
                    notify(`Tax pricing set to ${e.target.checked ? 'MRP-Inclusive' : 'Exclusive'}`);
                  }}
                />
                <b>Prices are MRP-inclusive of GST (Recommended for retail pharmacy)</b>
              </label>
              <div className="small-note" style={{ background: '#f8f7f2', padding: 10, borderRadius: 6 }}>
                <b>Formula:</b> Taxable Value = MRP ÷ (1 + GST% / 100).
                <br />
                Example: ₹112 MRP at 12% GST → Taxable = ₹100.00, CGST 6% = ₹6.00, SGST 6% = ₹6.00.
              </div>
            </div>
          </Panel>

          <Panel title="Stock Controls & Go-Live Policy" sub="Enforce inventory guardrails">
            <div style={{ padding: 18, display: 'grid', gap: 14 }}>
              <label style={{ display: 'flex', gap: 10, alignItems: 'center', cursor: 'pointer' }}>
                <input
                  type="checkbox"
                  checked={allowNegativeStock}
                  onChange={e => {
                    setAllowNegativeStock(e.target.checked);
                    stampAudit(`Changed Allow Negative Stock to: ${e.target.checked}`);
                    notify(`Negative stock override: ${e.target.checked ? 'Allowed' : 'Blocked'}`);
                  }}
                />
                <b>Allow Negative Stock (Admin override only)</b>
              </label>
              <div className="small-note">When disabled, billing blocks sales exceeding available quantity in that store's batch.</div>

              <Field label="Go-live date for app billing">
                <input
                  type="date"
                  className="control"
                  value={goLiveDate}
                  onChange={e => {
                    setGoLiveDate(e.target.value);
                    stampAudit(`Updated App Billing Go-Live Date to ${e.target.value}`);
                    notify(`Go-live date updated to ${e.target.value}`);
                  }}
                />
              </Field>
              <div className="small-note">ERP sales imports dated on or after this date are flagged as "Possible duplicate of app bills".</div>
            </div>
          </Panel>

          <div style={{ gridColumn: '1 / -1' }}>
            <div className="notice info">
              <b>E-invoicing / E-way bill note:</b> Applicability depends on aggregate annual turnover thresholds (currently ₹5 Cr+ for B2B e-invoicing). Confirm applicability with the client's CA. Not built in this prototype.
            </div>
          </div>
        </div>
      )}

      {adminTab === 'Audit log' && (
        <Panel title="Audit Trail" sub="In-session recording of invoices, cancellations, credit notes, adjustments and transfers">
          <div className="table-wrap">
            <table className="data-table">
              <thead>
                <tr>
                  <th>Action</th>
                  <th>User</th>
                  <th>Time</th>
                </tr>
              </thead>
              <tbody>
                {audit.map((a, i) => (
                  <tr key={i}>
                    <td>{a.action}</td>
                    <td>{a.user}</td>
                    <td>{a.time}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </Panel>
      )}
    </>
  );

  // ----------------------------------------------------
  // SCREEN: CLIENT CONFIRMATION (WITH PROMPT 2 QUESTIONS)
  // ----------------------------------------------------
  const clientPage = () => (
    <>
      <PageHead
        title="Client confirmation"
        help="Review the new Store-wise Stock Ledger & GST Billing features, confirm operational rules, and export answers."
        actions={
          <Button kind="primary" onClick={exportAnswers}>
            <ArrowDownToLine size={14} />Export answers (.json)
          </Button>
        }
      />

      <div className="notice info" style={{ marginBottom: 14 }}>
        Client review space · responses are held in browser memory for this session. Mark all GST logic as <b>“Illustrative: to be verified by the client's CA before production”</b>.
      </div>

      <Panel title="Screen decisions (including new modules)" sub="Evaluate each screen in the prototype">
        <div>
          {reports.map(screen => {
            const current = screenFeedback[screen] || { status: 'Not reviewed', comment: '' };
            return (
              <div className="answer-row" key={screen}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <label><b>{screen}</b></label>
                  <div className="pill-group">
                    <button
                      className={current.status === 'Looks right' ? 'selected' : ''}
                      onClick={() => setScreenFeedback(x => ({ ...x, [screen]: { ...current, status: 'Looks right' } }))}
                    >
                      Looks right
                    </button>
                    <button
                      className={current.status === 'Needs change' ? 'selected' : ''}
                      onClick={() => setScreenFeedback(x => ({ ...x, [screen]: { ...current, status: 'Needs change' } }))}
                    >
                      Needs change
                    </button>
                  </div>
                </div>
                <textarea
                  aria-label={`Comment for ${screen}`}
                  placeholder={`Comments about ${screen.toLowerCase()}…`}
                  value={current.comment}
                  onChange={e => setScreenFeedback(x => ({ ...x, [screen]: { ...current, comment: e.target.value } }))}
                />
              </div>
            );
          })}
        </div>
      </Panel>

      <Panel title="Replit Prompt 2: Client Architecture & CA Questions" sub="12 key decisions determining backend GST & ERP integration design">
        <div>
          {promptQuestions.map((q, i) => (
            <div className="answer-row" key={q}>
              <label>
                <b>{i + 1}. {q}</b>
              </label>
              <textarea
                aria-label={`Answer: ${q}`}
                placeholder="Client response / decision…"
                value={answers[i]}
                onChange={e => setAnswers(old => old.map((v, j) => (i === j ? e.target.value : v)))}
              />
            </div>
          ))}
        </div>
      </Panel>

      <div className="modal-actions">
        <Button onClick={exportFeedback}><ArrowDownToLine size={14} />Export session feedback (.json)</Button>
        <Button kind="primary" onClick={exportAnswers}><ArrowDownToLine size={14} />Download client answers (.json)</Button>
      </div>
    </>
  );

  // ----------------------------------------------------
  // SCREEN: DASHBOARD (UPDATED WITH STORE-WISE CARDS & KPIS)
  // ----------------------------------------------------
  const dashboard = () => {
    // Today's sales & GST for current active store (or all)
    const activeInvoices = invoices.filter(i => (activeStoreId === 'all' || i.storeId === activeStoreId) && i.status !== 'Cancelled');
    const todaySalesVal = activeInvoices.reduce((sum, i) => sum + i.grandTotal, 0);
    const todayGstVal = activeInvoices.reduce((sum, i) => sum + (i.cgstTotal + i.sgstTotal + i.igstTotal), 0);
    const expiring90DaysCount = seedBatches.filter(b => (activeStoreId === 'all' || b.storeId === activeStoreId) && b.expiry <= '2025-09-30' && b.expiry >= '2025-06-01').length;
    const reconMismatchesCount = reconciliations.filter(r => (activeStoreId === 'all' || r.storeId === activeStoreId) && r.status === 'Pending').length;

    return (
      <>
        <PageHead
          title={`Good morning, ${userName}`}
          help="Real-time shift overview across store sales, stock ledger movements, reconciliation variances and prescription review."
          actions={
            <div style={{ display: 'flex', gap: 8 }}>
              <Button kind="primary" onClick={() => setLocation('/billing')}>
                <Receipt size={15} />New Bill
              </Button>
              <Button onClick={() => setLocation('/store-stock')}>
                <Boxes size={15} />Store Stock
              </Button>
            </div>
          }
        />

        {/* New KPI Cards from Replit Prompt 2 */}
        <div className="stats-grid">
          <Link className="panel stat-card" href="/invoices" style={{ textDecoration: 'none' }}>
            <span className="stat-icon" style={{ background: '#e4f5ea', color: '#1f6e43' }}><Receipt size={17} /></span>
            <div className="stat-label">Today's Sales ({activeStoreId === 'all' ? 'All Stores' : storesMeta.find(s => s.id === activeStoreId)?.name})</div>
            <div className="stat-value">{money(todaySalesVal)}</div>
            <div className="stat-foot">{activeInvoices.length} bills generated today</div>
          </Link>

          <Link className="panel stat-card" href="/gst-reports" style={{ textDecoration: 'none' }}>
            <span className="stat-icon" style={{ background: '#e6f0fa', color: '#1e5488' }}><BarChart3 size={17} /></span>
            <div className="stat-label">Today's GST Collected</div>
            <div className="stat-value">{money(todayGstVal)}</div>
            <div className="stat-foot">CGST + SGST + IGST (Illustrative)</div>
          </Link>

          <Link className="panel stat-card" href="/store-stock" style={{ textDecoration: 'none' }}>
            <span className="stat-icon amber"><Clock size={17} /></span>
            <div className="stat-label">Expiring in 90 days</div>
            <div className="stat-value">{expiring90DaysCount} batches</div>
            <div className="stat-foot">Requires FEFO prioritization</div>
          </Link>

          <Link className="panel stat-card" href="/store-stock" style={{ textDecoration: 'none' }}>
            <span className="stat-icon" style={{ background: '#faebe6', color: '#a03b22' }}><RefreshCw size={17} /></span>
            <div className="stat-label">Stock Mismatches (Recon)</div>
            <div className="stat-value">{reconMismatchesCount} items</div>
            <div className="stat-foot">Discrepancy vs imported ERP snapshot</div>
          </Link>
        </div>

        {/* Secondary KPI cards */}
        <div className="stats-grid section-gap">
          {[
            { label: 'Low-stock items', value: lowCount, foot: `Below ${threshold} month cover`, Icon: AlertTriangle, link: '/products?cover=under', tone: 'amber' },
            { label: 'Pending matches', value: matches.filter(m => m.status === 'Pending').length, foot: 'Needs human resolution', Icon: GitMerge, link: '/matching', tone: '' },
            { label: 'WhatsApp to review', value: messages.filter(m => m.status === 'Needs Review' || m.status === 'Received').length, foot: 'Central number · simulated', Icon: MessageCircle, link: '/whatsapp', tone: 'amber' },
            { label: 'Pending prescriptions', value: prescriptions.filter(p => p.status === 'Received' || p.status === 'Review').length, foot: 'Awaiting pharmacist review', Icon: ClipboardList, link: '/prescriptions', tone: '' }
          ].map(x => (
            <Link className="panel stat-card" href={x.link} key={x.label} style={{ textDecoration: 'none' }}>
              <span className={`stat-icon ${x.tone}`}><x.Icon size={17} /></span>
              <div className="stat-label">{x.label}</div>
              <div className="stat-value">{x.value}</div>
              <div className="stat-foot">{x.foot}</div>
            </Link>
          ))}
        </div>

        <div className="dashboard-grid section-gap">
          <Panel title="Recent App Invoices" sub="Live sales deducting batch inventory immediately" action={<Link href="/invoices" className="row-primary">All Invoices →</Link>}>
            <div className="data-list">
              {invoices.slice(0, 5).map(inv => {
                const storeName = storesMeta.find(s => s.id === inv.storeId)?.name || inv.storeId;
                return (
                  <div className="list-row" key={inv.id} onClick={() => setLocation('/invoices')}>
                    <div>
                      <div className="row-primary">{inv.id} · {inv.customerName}</div>
                      <div className="row-secondary">{storeName} · {inv.lines?.length || 0} item(s) · {inv.paymentMode}</div>
                    </div>
                    <div style={{ textAlign: 'right' }}>
                      <b>{money(inv.grandTotal)}</b>
                      <div><Status>{inv.status}</Status></div>
                    </div>
                  </div>
                );
              })}
            </div>
          </Panel>

          <Panel title="Recent Stock Movements (Ledger)" sub="Audit trail of movements" action={<Link href="/store-stock" className="row-primary">Stock Ledger →</Link>}>
            <div className="data-list">
              {movements.slice(0, 5).map(m => {
                const prodName = products.find(p => p.id === m.productId)?.name || m.productId;
                const storeName = storesMeta.find(s => s.id === m.storeId)?.name || m.storeId;
                return (
                  <div className="list-row" key={m.id} onClick={() => setLocation('/store-stock')}>
                    <div>
                      <div className="row-primary">{prodName} ({m.batch})</div>
                      <div className="row-secondary">{storeName} · {m.ref} · {m.type}</div>
                    </div>
                    <div>
                      <span className={`badge ${m.qty < 0 ? 'warn' : ''}`}>
                        {m.qty > 0 ? `+${m.qty}` : m.qty}
                      </span>
                    </div>
                  </div>
                );
              })}
            </div>
          </Panel>
        </div>
      </>
    );
  };

  // ----------------------------------------------------
  // ROUTING SETUP
  // ----------------------------------------------------
  const title = navTitle(location);
  const activeRoute = (
    <Switch>
      <Route path="/" component={dashboard} />
      <Route path="/store-stock">
        {() => (
          <StoreStockView
            role={role}
            activeStore={activeStoreId}
            setActiveStore={setActiveStoreId}
            storesMeta={storesMeta}
            movements={movements}
            setMovements={setMovements}
            transfers={transfers}
            setTransfers={setTransfers}
            reconciliations={reconciliations}
            setReconciliations={setReconciliations}
            notify={notify}
            stampAudit={stampAudit}
            threshold={threshold}
            getLedgerStock={getLedgerStock}
          />
        )}
      </Route>
      <Route path="/billing">
        {() => (
          <BillingView
            role={role}
            activeStore={activeStoreId === 'all' ? 'S1' : activeStoreId}
            storesMeta={storesMeta}
            customers={customers}
            setCustomers={setCustomers}
            movements={movements}
            setMovements={setMovements}
            invoices={invoices}
            setInvoices={setInvoices}
            prescriptions={prescriptions}
            setPrescriptions={setPrescriptions}
            pricesInclusive={pricesInclusive}
            allowNegativeStock={allowNegativeStock}
            notify={notify}
            stampAudit={stampAudit}
            getLedgerStock={getLedgerStock}
            prefillPrescriptionId={billingPrefill?.prescriptionId}
          />
        )}
      </Route>
      <Route path="/invoices">
        {() => (
          <InvoicesView
            role={role}
            activeStore={activeStoreId}
            storesMeta={storesMeta}
            invoices={invoices}
            setInvoices={setInvoices}
            movements={movements}
            setMovements={setMovements}
            creditNotes={creditNotes}
            setCreditNotes={setCreditNotes}
            stampAudit={stampAudit}
            notify={notify}
          />
        )}
      </Route>
      <Route path="/gst-reports">
        {() => (
          <GstReportsView
            storesMeta={storesMeta}
            invoices={invoices}
            creditNotes={creditNotes}
            movements={movements}
          />
        )}
      </Route>
      <Route path="/imports" component={importPage} />
      <Route path="/products" component={productPage} />
      <Route path="/products/:id" component={productDetail} />
      <Route path="/suppliers" component={supplierPage} />
      <Route path="/suppliers/:id" component={supplierPage} />
      <Route path="/matching" component={matchingPage} />
      <Route path="/ask" component={askPage} />
      <Route path="/whatsapp" component={whatsappPage} />
      <Route path="/prescriptions" component={prescriptionPage} />
      <Route path="/admin" component={adminPage} />
      <Route path="/client-confirmation" component={clientPage} />
      <Route>
        <div className="empty-state">
          <CircleHelp size={28} />
          <strong>Screen not found</strong>
          <Button onClick={() => setLocation('/')}>Return to dashboard</Button>
        </div>
      </Route>
    </Switch>
  );

  return (
    <div className="app-shell">
      <aside className="sidebar">
        <div className="brand">
          <div className="brand-mark">S</div>
          <div>
            <div className="brand-title">Sampada</div>
            <div className="brand-sub">Stock & GST Billing Prototype</div>
          </div>
        </div>

        <div className="nav-label">Workspace</div>
        <nav className="nav-list">
          {nav.map(({ name, path, Icon }) => {
            const blocked =
              (role === 'Manager' && (name === 'Prescriptions' || name === 'Billing')) ||
              (role === 'Staff' && ['Imports', 'Suppliers', 'Matching Queue', 'Ask', 'Admin', 'GST Reports'].includes(name));
            return blocked ? (
              <span key={path} className="nav-item" title="Not available for this role" aria-disabled="true" style={{ opacity: 0.45, cursor: 'not-allowed' }}>
                <Icon />
                <span>{name}</span>
              </span>
            ) : (
              <Link
                href={path}
                key={path}
                className={`nav-item ${location === path || (path !== '/' && location.startsWith(`${path}/`)) ? 'active' : ''}`}
                data-testid={`nav-${name.toLowerCase().replaceAll(' ', '-')}`}
              >
                <Icon />
                <span>{name}</span>
                {name === 'Matching Queue' && matches.filter(x => x.status === 'Pending').length > 0 && (
                  <span className="nav-count">{matches.filter(x => x.status === 'Pending').length}</span>
                )}
                {name === 'Store Stock' && reconciliations.filter(x => x.status === 'Pending').length > 0 && (
                  <span className="nav-count" style={{ background: '#f56565' }}>{reconciliations.filter(x => x.status === 'Pending').length}</span>
                )}
              </Link>
            );
          })}
        </nav>

        <div className="side-bottom">
          <div className="side-live"><i className="dot" />Stock Ledger Active</div>
          <div>All sales deduct batch inventory immediately.</div>
        </div>
      </aside>

      <main className="main">
        <header className="topbar">
          <div className="global-search">
            <Search size={15} />
            <input
              aria-label="Search products"
              placeholder="Search products…"
              value={search}
              onChange={e => setSearch(e.target.value)}
              onKeyDown={e => {
                if (e.key === 'Enter') {
                  setFilter(search);
                  setLocation('/products');
                }
              }}
            />
            {search && (
              <button onClick={() => setSearch('')} aria-label="Clear search" style={{ border: 0, background: 'transparent', color: '#788' }}>
                ×
              </button>
            )}
          </div>

          <div className="top-spacer" />

          {/* Store Switcher */}
          <div style={{ display: 'flex', alignItems: 'center', gap: 6, marginRight: 12 }}>
            <Store size={14} color="#526268" />
            <label className="small-note" htmlFor="store-switch">Store:</label>
            <select
              id="store-switch"
              className="control"
              style={{ padding: '4px 8px', fontSize: 12, height: 32 }}
              value={activeStoreId}
              onChange={e => setActiveStoreId(e.target.value)}
            >
              <option value="all">All Stores (Consolidated)</option>
              {storesMeta.map(s => (
                <option key={s.id} value={s.id}>{s.name}</option>
              ))}
            </select>
          </div>

          <label className="small-note" htmlFor="role-switch">Role:</label>
          <select
            id="role-switch"
            className="role-select"
            data-testid="role-switcher"
            value={role}
            onChange={e => setRole(e.target.value as Role)}
          >
            {(['Admin', 'Store Owner', 'Manager', 'Pharmacist', 'Staff'] as Role[]).map(x => (
              <option key={x}>{x}</option>
            ))}
          </select>
        </header>

        <div className="prototype-strip">
          <span>
            <Zap size={13} />
            <b>Stock Ledger & GST Billing Prototype:</b> Sample data only · <b>Illustrative GST logic: to be verified by CA</b> · No live external integrations
          </span>
          <span>
            User: <b>{userName}</b> ({role}) <ShieldCheck size={13} />
          </span>
        </div>

        <div className="page screen-enter" key={location}>
          {activeRoute}
        </div>
      </main>

      <button className="feedback-fab" onClick={() => viewFeedback(title)} data-testid="client-feedback-button">
        <MessageCircle size={16} />Client feedback
      </button>

      {feedbackOpen && (
        <div className="drawer-shade" onMouseDown={e => { if (e.target === e.currentTarget) setFeedbackOpen(false); }}>
          <aside className="drawer" role="dialog" aria-label="Client feedback">
            <div className="drawer-head">
              <div>
                <div className="eyebrow">CLIENT REVIEW</div>
                <h2>Feedback on {feedbackScreen}</h2>
              </div>
              <button className="close-button" onClick={() => setFeedbackOpen(false)} aria-label="Close feedback">×</button>
            </div>
            <div className="notice info">Your feedback stays in this browser session and can be downloaded as JSON.</div>
            <div className="feedback-form section-gap">
              <div className="pill-group">
                <button className={feedbackStatus === 'Looks right' ? 'selected' : ''} onClick={() => setFeedbackStatus('Looks right')}>Looks right</button>
                <button className={feedbackStatus === 'Needs change' ? 'selected' : ''} onClick={() => setFeedbackStatus('Needs change')}>Needs change</button>
              </div>
              <Field label="Comment">
                <textarea rows={5} value={feedbackText} onChange={e => setFeedbackText(e.target.value)} placeholder="What should stay or change on this screen?" />
              </Field>
              <Button kind="primary" onClick={saveFeedback}><Check size={14} />Save feedback</Button>
            </div>
            <Panel title="Saved this session" sub={`${feedback.length} notes`}>
              {feedback.length ? (
                <div className="data-list">
                  {feedback.slice().reverse().map((f, i) => (
                    <div className="list-row" key={i}>
                      <div>
                        <div className="row-primary">{f.screen}</div>
                        <div className="row-secondary">{f.comment || 'No comment added'}</div>
                      </div>
                      <Status>{f.status}</Status>
                    </div>
                  ))}
                </div>
              ) : (
                <div className="empty-state">
                  <MessageCircle size={22} />
                  <strong>No feedback yet</strong>
                  <p>Choose a decision and save a note.</p>
                </div>
              )}
            </Panel>
            <div className="modal-actions">
              <Button onClick={exportFeedback}><ArrowDownToLine size={14} />Download feedback (.json)</Button>
            </div>
          </aside>
        </div>
      )}

      {modal && modal.body !== null && modal.title !== 'Print preview' ? (
        <Modal title={modal.title} close={() => setModal(null)}>{modal.body}</Modal>
      ) : null}

      {toast && (
        <div role="status" style={{ position: 'fixed', zIndex: 90, bottom: 80, left: '50%', transform: 'translateX(-50%)', background: '#213b40', color: '#fff', padding: '11px 18px', borderRadius: 8, fontSize: 13, fontWeight: 500, boxShadow: '0 5px 22px #142c3060', border: '1px solid #3c5d64' }}>
          {toast}
        </div>
      )}
    </div>
  );
}

function MatchResolve({ m, products, onChoose, onCreate, onSkip }: { m: any; products: Product[]; onChoose: (id: string) => void; onCreate: (name: string) => void; onSkip: () => void; }) {
  const [chosen, setChosen] = useState(m.candidates.find((x: string) => x) || '');
  const [newName, setNewName] = useState(m.rawText);
  const normalize = (value: string) => value.toLowerCase().replace(/[^a-z0-9]/g, '');
  const raw = normalize(m.rawText);
  const compare = (value: string) => {
    const token = normalize(value).replace(/(mcg|mg|iu|g)$/, '');
    return token && raw.includes(token) ? 'Text match' : 'Not stated in source';
  };
  return (
    <>
      <p>Incoming text: <b>{m.rawText}</b></p>
      <p className="small-note">Candidate confidence is a local similarity hint only. Check strength and pack against the source before choosing; nothing maps automatically.</p>
      <div className="data-list">
        {m.candidates.filter((id: string) => id && products.some((p: Product) => p.id === id)).map((id: string, i: number) => {
          const p = products.find((x: Product) => x.id === id)!;
          const score = Math.max(58, 94 - i * 11);
          const strength = compare(p.strength);
          const pack = compare(p.pack);
          return (
            <label className="panel" key={id} style={{ display: 'grid', gridTemplateColumns: '24px 1fr', gap: 10, padding: 13, marginBottom: 9, cursor: 'pointer', background: chosen === id ? '#f2f7f2' : '#fff' }}>
              <input type="radio" name="candidate" checked={chosen === id} onChange={() => setChosen(id)} />
              <div>
                <div style={{ display: 'flex', justifyContent: 'space-between', gap: 8 }}>
                  <span className="row-primary">{p.name}</span>
                  <span className="badge blue">Similarity hint {score}%</span>
                </div>
                <div className="row-secondary">Master record · {p.group} · {p.id}</div>
                <div style={{ display: 'flex', gap: 7, marginTop: 9, flexWrap: 'wrap' }}>
                  <span className={`badge ${strength === 'Text match' ? '' : 'warn'}`}>Strength · {p.strength} · {strength}</span>
                  <span className={`badge ${pack === 'Text match' ? '' : 'warn'}`}>Pack · {p.pack} · {pack}</span>
                </div>
              </div>
            </label>
          );
        })}
      </div>
      <Field label="New product name">
        <input value={newName} onChange={e => setNewName(e.target.value)} />
      </Field>
      <div className="modal-actions">
        <Button onClick={onSkip}>Skip</Button>
        <Button onClick={() => onCreate(newName)} disabled={!newName.trim()}>Create new product</Button>
        <Button kind="primary" disabled={!chosen} onClick={() => onChoose(chosen)}>Select existing product</Button>
      </div>
    </>
  );
}

export default App;