import { useState, useMemo, type FormEvent } from 'react';
import * as XLSX from 'xlsx';
import {
  AlertTriangle,
  ArrowDownToLine,
  ArrowLeft,
  ArrowRight,
  ArrowLeftRight,
  Building2,
  Check,
  ChevronDown,
  CircleAlert,
  ClipboardList,
  Clock,
  Download,
  FileSpreadsheet,
  FileText,
  Filter,
  Layers,
  Package,
  Plus,
  Printer,
  Receipt,
  RotateCcw,
  Search,
  Settings,
  ShieldAlert,
  ShoppingCart,
  Trash2,
  TrendingDown,
  TrendingUp,
  X
} from 'lucide-react';
import { products, suppliers, type Product } from './data';
import {
  storeMasters,
  seedCustomers,
  productGstMap,
  numberToWordsIndian,
  type StoreInfo,
  type Customer,
  type StockMovement,
  type Invoice,
  type InvoiceLine,
  type CreditNote,
  type InterStoreTransfer,
  type ReconciliationDiscrepancy
} from './gstData';

const money = (n: number) => `₹${n.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
const today = '30 Jun 2025';

export function StatusBadge({ children, tone }: { children: string; tone?: string }) {
  const t = tone || (/Pending|Needs Review|Review|Queued|Received|Draft|Expiring/.test(children)
    ? 'warn'
    : /Failed|Unknown|Change needed|Cancelled|Expired|Discrepancy/.test(children)
    ? 'error'
    : /Closed|Completed|Resolved|Looks right|Matched|Final|OK|Sufficient|Accepted/.test(children)
    ? ''
    : 'neutral');
  return (
    <span className={`badge ${t}`}>
      <i aria-hidden="true">{t === '' ? '✓' : t === 'warn' ? '•' : t === 'error' ? '!' : '·'}</i>
      {children}
    </span>
  );
}

// -------------------------------------------------------------
// 1. STORE STOCK & LEDGER VIEW
// -------------------------------------------------------------
export function StoreStockView({
  role,
  activeStore,
  setActiveStore,
  storesMeta,
  movements,
  setMovements,
  transfers,
  setTransfers,
  reconciliations,
  setReconciliations,
  notify,
  stampAudit,
  threshold,
  getLedgerStock
}: {
  role: string;
  activeStore: string;
  setActiveStore: (s: string) => void;
  storesMeta: StoreInfo[];
  movements: StockMovement[];
  setMovements: React.Dispatch<React.SetStateAction<StockMovement[]>>;
  transfers: InterStoreTransfer[];
  setTransfers: React.Dispatch<React.SetStateAction<InterStoreTransfer[]>>;
  reconciliations: ReconciliationDiscrepancy[];
  setReconciliations: React.Dispatch<React.SetStateAction<ReconciliationDiscrepancy[]>>;
  notify: (s: string) => void;
  stampAudit: (s: string) => void;
  threshold: number;
  getLedgerStock: (productId: string, storeId?: string, batch?: string) => number;
}) {
  const [subTab, setSubTab] = useState<'inventory' | 'matrix' | 'reconciliation' | 'transfers'>('inventory');
  const [search, setSearch] = useState('');
  const [batchFilter, setBatchFilter] = useState<'all' | 'expiring' | 'expired' | 'low'>('all');
  const [selectedLedgerProduct, setSelectedLedgerProduct] = useState<{ productId: string; batch?: string; storeId?: string } | null>(null);
  const [showAdjustModal, setShowAdjustModal] = useState(false);
  const [showTransferModal, setShowTransferModal] = useState(false);
  const [showPurchaseInModal, setShowPurchaseInModal] = useState(false);

  // Group movements into distinct batches
  const batchList = useMemo(() => {
    const map = new Map<string, {
      storeId: string;
      productId: string;
      batch: string;
      expiry: string;
      qty: number;
      mrp: number;
      lastMovement?: StockMovement;
    }>();

    // Group by store + product + batch
    movements.forEach(m => {
      if (activeStore !== 'All' && m.storeId !== activeStore) return;
      const key = `${m.storeId}__${m.productId}__${m.batch}`;
      const existing = map.get(key);
      const meta = productGstMap[m.productId] || { mrp: 60 };
      if (existing) {
        existing.qty += m.qty;
        if (!existing.lastMovement || m.dateTime > existing.lastMovement.dateTime) {
          existing.lastMovement = m;
        }
      } else {
        map.set(key, {
          storeId: m.storeId,
          productId: m.productId,
          batch: m.batch,
          expiry: m.expiry,
          qty: m.qty,
          mrp: meta.mrp,
          lastMovement: m
        });
      }
    });

    return Array.from(map.values()).filter(item => {
      const prod = products.find(p => p.id === item.productId);
      const matchesSearch = !search || `${prod?.name || ''} ${(prod?.aliases || []).join(' ')} ${item.batch}`.toLowerCase().includes(search.toLowerCase());
      if (!matchesSearch) return false;

      const isExpired = item.expiry < '2025-06';
      const isExpiring = !isExpired && item.expiry <= '2025-09';
      if (batchFilter === 'expired') return isExpired;
      if (batchFilter === 'expiring') return isExpiring;
      if (batchFilter === 'low') return item.qty < 20;
      return true;
    });
  }, [movements, activeStore, search, batchFilter]);

  const exportStockExcel = () => {
    const rows = batchList.map(b => {
      const p = products.find(prod => prod.id === b.productId);
      const storeName = storesMeta.find(s => s.id === b.storeId)?.name || b.storeId;
      return {
        Store: storeName,
        'Product Code': b.productId,
        'Product Name': p?.name || '',
        Group: p?.group || '',
        Batch: b.batch,
        Expiry: b.expiry,
        'On-Hand Qty': b.qty,
        'Reserved Qty': 0,
        MRP: b.mrp,
        'Last Movement Date': b.lastMovement?.dateTime || '',
        'Last Movement Type': b.lastMovement?.type || ''
      };
    });
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Store Stock');
    XLSX.writeFile(wb, `store-stock-${activeStore.toLowerCase()}-${today}.xlsx`);
  };

  return (
    <>
      <div className="page-head">
        <div>
          <div className="eyebrow">STORE-WISE INVENTORY & LEDGER</div>
          <h1 className="page-title">Store Stock</h1>
          <p className="help-line">What this screen does: Track batch-level stock, expiry horizons and Tally-style movement registers across pharmacy stores.</p>
        </div>
        <div className="page-head-actions">
          <button className="btn" onClick={() => setShowPurchaseInModal(true)}>
            <Plus size={14} /> Purchase Stock-in
          </button>
          <button className="btn" onClick={() => setShowTransferModal(true)} disabled={role === 'Staff'}>
            <ArrowLeftRight size={14} /> Inter-store Transfer
          </button>
          <button className="btn" onClick={() => setShowAdjustModal(true)} disabled={role === 'Staff'}>
            <RotateCcw size={14} /> Stock Adjustment
          </button>
          <button className="btn" onClick={exportStockExcel}>
            <ArrowDownToLine size={14} /> Export Excel
          </button>
          <button className="btn" onClick={() => window.print()}>
            <Printer size={14} /> Print / PDF
          </button>
        </div>
      </div>

      <div className="notice info" style={{ marginBottom: 14 }}>
        <b>Stock Ledger Core:</b> Current stock = latest imported snapshot (opening balance) + all movements (sales, purchases, returns, adjustments, transfers). Sales made in this app deduct stock immediately.
      </div>

      {/* Toolbar & Sub-tabs */}
      <div className="panel">
        <div className="toolbar" style={{ justifyContent: 'space-between' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
            <div className="pill-group">
              <button className={subTab === 'inventory' ? 'selected' : ''} onClick={() => setSubTab('inventory')}>
                <Layers size={13} style={{ verticalAlign: 'middle', marginRight: 4 }} /> Batch Inventory
              </button>
              <button className={subTab === 'matrix' ? 'selected' : ''} onClick={() => setSubTab('matrix')}>
                Stores Matrix
              </button>
              <button className={subTab === 'reconciliation' ? 'selected' : ''} onClick={() => setSubTab('reconciliation')}>
                <CircleAlert size={13} style={{ verticalAlign: 'middle', marginRight: 4 }} /> Reconciliation ({reconciliations.filter(r => r.status === 'Pending').length})
              </button>
              <button className={subTab === 'transfers' ? 'selected' : ''} onClick={() => setSubTab('transfers')}>
                <ArrowLeftRight size={13} style={{ verticalAlign: 'middle', marginRight: 4 }} /> Transfers ({transfers.length})
              </button>
            </div>

            <label className="small-note" style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
              Store:
              <select className="control" value={activeStore} onChange={e => setActiveStore(e.target.value)}>
                <option value="All">All Stores</option>
                {storesMeta.map(s => <option key={s.id} value={s.id}>{s.name}</option>)}
              </select>
            </label>
          </div>

          <div style={{ display: 'flex', gap: 8 }}>
            <div className="global-search" style={{ width: 220 }}>
              <Search size={14} />
              <input placeholder="Search product or batch..." value={search} onChange={e => setSearch(e.target.value)} />
            </div>
            {subTab === 'inventory' && (
              <select className="control" value={batchFilter} onChange={e => setBatchFilter(e.target.value as any)}>
                <option value="all">All batches</option>
                <option value="expiring">Expiring soon (≤90 days)</option>
                <option value="expired">Expired batches</option>
                <option value="low">Low stock (&lt;20 units)</option>
              </select>
            )}
          </div>
        </div>

        {/* 1. BATCH INVENTORY TAB */}
        {subTab === 'inventory' && (
          <div className="table-wrap">
            <table className="data-table">
              <thead>
                <tr>
                  <th>Product</th>
                  <th>Store</th>
                  <th>Batch</th>
                  <th>Expiry</th>
                  <th>On-Hand Qty</th>
                  <th>Reserved</th>
                  <th>MRP</th>
                  <th>Last Movement</th>
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody>
                {batchList.map((item, idx) => {
                  const p = products.find(prod => prod.id === item.productId);
                  const storeName = storesMeta.find(s => s.id === item.storeId)?.name || item.storeId;
                  const isExpired = item.expiry < '2025-06';
                  const isExpiring = !isExpired && item.expiry <= '2025-09';

                  return (
                    <tr key={idx} onClick={() => setSelectedLedgerProduct({ productId: item.productId, batch: item.batch, storeId: item.storeId })}>
                      <td className="product-cell">
                        {p?.name}
                        <div className="row-secondary">{p?.pack} · HSN {productGstMap[item.productId]?.hsn || '3004'}</div>
                      </td>
                      <td>{storeName}</td>
                      <td><span className="mono">{item.batch}</span></td>
                      <td>
                        {isExpired ? (
                          <span className="badge error">Expired ({item.expiry})</span>
                        ) : isExpiring ? (
                          <span className="badge warn">Expiring soon ({item.expiry})</span>
                        ) : (
                          <span>{item.expiry}</span>
                        )}
                      </td>
                      <td>
                        <b>{item.qty} units</b>
                        {item.qty < 20 && <div className="row-secondary" style={{ color: '#b26829' }}>Low stock</div>}
                      </td>
                      <td>0</td>
                      <td>{money(item.mrp)}</td>
                      <td>
                        <div className="row-primary">{item.lastMovement?.type || 'Opening'}</div>
                        <div className="row-secondary">{item.lastMovement?.dateTime}</div>
                      </td>
                      <td>
                        <button className="btn tiny" onClick={(e) => {
                          e.stopPropagation();
                          setSelectedLedgerProduct({ productId: item.productId, batch: item.batch, storeId: item.storeId });
                        }}>
                          View Ledger
                        </button>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
            {batchList.length === 0 && (
              <div className="empty-state">
                <Package size={24} />
                <strong>No stock items found</strong>
                <p>Adjust your search, store or expiry filter.</p>
              </div>
            )}
          </div>
        )}

        {/* 2. MATRIX VIEW TAB */}
        {subTab === 'matrix' && (
          <div className="table-wrap">
            <table className="data-table">
              <thead>
                <tr>
                  <th>Product</th>
                  <th>Group</th>
                  <th>Main Store (S1)</th>
                  <th>Branch 2 (S2)</th>
                  <th>Branch 3 (S3)</th>
                  <th>Total On-Hand</th>
                </tr>
              </thead>
              <tbody>
                {products.filter(p => !search || p.name.toLowerCase().includes(search.toLowerCase())).map(p => {
                  const s1 = getLedgerStock(p.id, 'S1');
                  const s2 = getLedgerStock(p.id, 'S2');
                  const s3 = getLedgerStock(p.id, 'S3');
                  const total = s1 + s2 + s3;
                  return (
                    <tr key={p.id} onClick={() => setSelectedLedgerProduct({ productId: p.id })}>
                      <td className="product-cell">{p.name}<div className="row-secondary">{p.pack}</div></td>
                      <td>{p.group}</td>
                      <td><b>{s1}</b> units</td>
                      <td><b>{s2}</b> units</td>
                      <td><b>{s3}</b> units</td>
                      <td>
                        <span className="badge" style={{ fontSize: 11 }}>{total} units</span>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}

        {/* 3. RECONCILIATION TAB */}
        {subTab === 'reconciliation' && (
          <div style={{ padding: 18 }}>
            <div className="notice" style={{ marginBottom: 14 }}>
              <b>Stock Reconciliation (Ledger vs Imported ERP):</b> As of 30 Jun 2025 snapshot. Differences represent counter breakages, unrecorded returns or physical audit variances.
            </div>
            <div className="table-wrap">
              <table className="data-table">
                <thead>
                  <tr>
                    <th>Store</th>
                    <th>Product</th>
                    <th>Batch</th>
                    <th>App Ledger Qty</th>
                    <th>Imported ERP Qty</th>
                    <th>Variance</th>
                    <th>Suggested Reason</th>
                    <th>Status</th>
                    <th>Action</th>
                  </tr>
                </thead>
                <tbody>
                  {reconciliations.map(r => {
                    const p = products.find(x => x.id === r.productId);
                    const storeName = storesMeta.find(s => s.id === r.storeId)?.name || r.storeId;
                    return (
                      <tr key={r.id}>
                        <td>{storeName}</td>
                        <td className="product-cell">{p?.name}</td>
                        <td><span className="mono">{r.batch}</span></td>
                        <td><b>{r.systemQty}</b></td>
                        <td><b>{r.importedErpQty}</b></td>
                        <td>
                          <span className={`badge ${r.difference < 0 ? 'error' : 'warn'}`}>
                            {r.difference > 0 ? `+${r.difference}` : r.difference} units
                          </span>
                        </td>
                        <td>{r.suggestedReason}</td>
                        <td><StatusBadge>{r.status}</StatusBadge></td>
                        <td>
                          {r.status === 'Pending' ? (
                            <div style={{ display: 'flex', gap: 6 }}>
                              <button className="btn tiny primary" onClick={() => {
                                // Create adjustment
                                const newMov: StockMovement = {
                                  id: `MOV-ADJ-${Date.now()}`,
                                  dateTime: new Date().toISOString().slice(0, 16).replace('T', ' '),
                                  storeId: r.storeId,
                                  productId: r.productId,
                                  batch: r.batch,
                                  expiry: '2026-11',
                                  qty: r.difference,
                                  type: 'Adjustment',
                                  ref: `REC-ADJ-${r.id}`,
                                  user: role,
                                  reason: `Reconciliation adjustment: ${r.suggestedReason}`,
                                  source: 'App Billing'
                                };
                                setMovements(old => [newMov, ...old]);
                                setReconciliations(old => old.map(x => x.id === r.id ? { ...x, status: 'Adjusted' } : x));
                                stampAudit(`Reconciliation adjusted: ${p?.name} (${r.difference} units)`);
                                notify(`Stock adjusted by ${r.difference} units to reconcile with ERP`);
                              }}>
                                Create Adjustment
                              </button>
                              <button className="btn tiny" onClick={() => {
                                setReconciliations(old => old.map(x => x.id === r.id ? { ...x, status: 'Accepted' } : x));
                                stampAudit(`Reconciliation accepted variance for ${p?.name}`);
                                notify('Variance marked as accepted without stock movement');
                              }}>
                                Accept
                              </button>
                            </div>
                          ) : (
                            <span className="small-note">Resolved</span>
                          )}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* 4. TRANSFERS TAB */}
        {subTab === 'transfers' && (
          <div className="table-wrap">
            <table className="data-table">
              <thead>
                <tr>
                  <th>Transfer No</th>
                  <th>Date</th>
                  <th>From Store</th>
                  <th>To Store</th>
                  <th>Product</th>
                  <th>Batch</th>
                  <th>Qty</th>
                  <th>Status</th>
                  <th>Action</th>
                </tr>
              </thead>
              <tbody>
                {transfers.map(tr => {
                  const p = products.find(prod => prod.id === tr.productId);
                  const from = storesMeta.find(s => s.id === tr.fromStoreId)?.name || tr.fromStoreId;
                  const to = storesMeta.find(s => s.id === tr.toStoreId)?.name || tr.toStoreId;
                  return (
                    <tr key={tr.id}>
                      <td><span className="mono">{tr.id}</span></td>
                      <td>{tr.date}</td>
                      <td>{from}</td>
                      <td>{to}</td>
                      <td className="product-cell">{p?.name}</td>
                      <td><span className="mono">{tr.batch}</span></td>
                      <td><b>{tr.qty} units</b></td>
                      <td><StatusBadge>{tr.status}</StatusBadge></td>
                      <td>
                        {tr.status === 'Dispatched' ? (
                          <button className="btn tiny primary" onClick={() => {
                            // Confirm receipt
                            const inMov: StockMovement = {
                              id: `MOV-TR-IN-${Date.now()}`,
                              dateTime: new Date().toISOString().slice(0, 16).replace('T', ' '),
                              storeId: tr.toStoreId,
                              productId: tr.productId,
                              batch: tr.batch,
                              expiry: tr.expiry,
                              qty: tr.qty,
                              type: 'Transfer in',
                              ref: tr.id,
                              user: role,
                              reason: `Transfer received from ${from}`,
                              source: 'App Billing'
                            };
                            setMovements(old => [inMov, ...old]);
                            setTransfers(old => old.map(t => t.id === tr.id ? { ...t, status: 'Received' } : t));
                            stampAudit(`Inter-store transfer ${tr.id} received at ${to}`);
                            notify(`Transfer received! Added ${tr.qty} units to ${to}`);
                          }}>
                            Confirm Receipt
                          </button>
                        ) : (
                          <span className="small-note">Completed</span>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {/* STOCK LEDGER DRAWER */}
      {selectedLedgerProduct && (
        <div className="drawer-shade" onMouseDown={e => { if (e.target === e.currentTarget) setSelectedLedgerProduct(null); }}>
          <div className="drawer" style={{ width: 'min(720px, 100vw)' }}>
            <div className="drawer-head">
              <div>
                <div className="eyebrow">TALLY-STYLE STOCK ITEM REGISTER</div>
                <h2>Stock Ledger</h2>
              </div>
              <button className="close-button" onClick={() => setSelectedLedgerProduct(null)}>×</button>
            </div>

            {(() => {
              const p = products.find(x => x.id === selectedLedgerProduct.productId);
              const storeFiltered = selectedLedgerProduct.storeId || (activeStore !== 'All' ? activeStore : undefined);
              const productMovements = movements
                .filter(m => m.productId === selectedLedgerProduct.productId && (!storeFiltered || m.storeId === storeFiltered) && (!selectedLedgerProduct.batch || m.batch === selectedLedgerProduct.batch))
                .sort((a, b) => a.dateTime.localeCompare(b.dateTime));

              let running = 0;
              const rowsWithBalance = productMovements.map(m => {
                running += m.qty;
                return { ...m, runningBalance: running };
              });

              return (
                <>
                  <div style={{ padding: '12px 16px', background: '#f4f7f5', borderRadius: 8, marginBottom: 16 }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                      <div>
                        <h3 style={{ font: '800 16px Manrope', margin: 0 }}>{p?.name}</h3>
                        <div className="row-secondary">
                          {p?.pack} · {storeFiltered ? storesMeta.find(s => s.id === storeFiltered)?.name : 'Across all stores'} {selectedLedgerProduct.batch ? `· Batch ${selectedLedgerProduct.batch}` : ''}
                        </div>
                      </div>
                      <div style={{ textAlign: 'right' }}>
                        <div className="stat-label">CURRENT BALANCE</div>
                        <div style={{ font: '800 20px Manrope', color: '#256c5d' }}>{running} units</div>
                      </div>
                    </div>
                  </div>

                  <div className="table-wrap">
                    <table className="data-table">
                      <thead>
                        <tr>
                          <th>Date & Time</th>
                          <th>Type</th>
                          <th>Reference</th>
                          <th>Batch</th>
                          <th>Qty Change</th>
                          <th>Balance</th>
                          <th>User</th>
                          <th>Source</th>
                        </tr>
                      </thead>
                      <tbody>
                        {rowsWithBalance.map((m, idx) => (
                          <tr key={idx}>
                            <td>{m.dateTime}</td>
                            <td>
                              <span className={`badge ${m.type.includes('Sale') ? 'warn' : m.type.includes('Purchase') || m.type.includes('return') ? '' : 'neutral'}`}>
                                {m.type}
                              </span>
                            </td>
                            <td><span className="mono">{m.ref}</span></td>
                            <td><span className="mono">{m.batch}</span></td>
                            <td>
                              <b style={{ color: m.qty > 0 ? '#26705d' : '#a1423c' }}>
                                {m.qty > 0 ? `+${m.qty}` : m.qty}
                              </b>
                            </td>
                            <td><b>{m.runningBalance}</b></td>
                            <td>{m.user}</td>
                            <td><span className="small-note">{m.source}</span></td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>

                  <div className="modal-actions" style={{ marginTop: 20 }}>
                    <button className="btn" onClick={() => {
                      const csv = [
                        'Date,Type,Reference,Batch,Qty,Balance,User,Source',
                        ...rowsWithBalance.map(r => `"${r.dateTime}","${r.type}","${r.ref}","${r.batch}",${r.qty},${r.runningBalance},"${r.user}","${r.source}"`)
                      ].join('\n');
                      const blob = new Blob([csv], { type: 'text/csv' });
                      const a = document.createElement('a');
                      a.href = URL.createObjectURL(blob);
                      a.download = `ledger-${p?.id}-${selectedLedgerProduct.batch || 'all'}.csv`;
                      a.click();
                    }}>
                      <Download size={14} /> Download Ledger CSV
                    </button>
                    <button className="btn primary" onClick={() => setSelectedLedgerProduct(null)}>Close</button>
                  </div>
                </>
              );
            })()}
          </div>
        </div>
      )}

      {/* STOCK ADJUSTMENT MODAL */}
      {showAdjustModal && (
        <StockAdjustmentModal
          stores={storesMeta}
          products={products}
          onClose={() => setShowAdjustModal(false)}
          onSave={(data) => {
            const newMov: StockMovement = {
              id: `MOV-ADJ-${Date.now()}`,
              dateTime: new Date().toISOString().slice(0, 16).replace('T', ' '),
              storeId: data.storeId,
              productId: data.productId,
              batch: data.batch,
              expiry: data.expiry || '2026-11',
              qty: data.type === 'decrease' ? -Math.abs(data.qty) : Math.abs(data.qty),
              type: 'Adjustment',
              ref: `ADJ-${Date.now().toString().slice(-4)}`,
              user: role,
              reason: `${data.reason}: ${data.notes || 'Manual audit'}`,
              source: 'App Billing'
            };
            setMovements(old => [newMov, ...old]);
            stampAudit(`Stock adjustment: ${data.productId} (${data.type === 'decrease' ? '-' : '+'}${data.qty}) - Reason: ${data.reason}`);
            notify(`Stock adjusted: ${data.type === 'decrease' ? '-' : '+'}${data.qty} units recorded`);
            setShowAdjustModal(false);
          }}
        />
      )}

      {/* INTER-STORE TRANSFER MODAL */}
      {showTransferModal && (
        <InterStoreTransferModal
          stores={storesMeta}
          products={products}
          getLedgerStock={getLedgerStock}
          onClose={() => setShowTransferModal(false)}
          onSave={(transferData) => {
            const trId = `TR-2025-${String(transfers.length + 101).padStart(4, '0')}`;
            const dt = new Date().toISOString().slice(0, 16).replace('T', ' ');

            // Deduct immediately from source store
            const outMov: StockMovement = {
              id: `MOV-TR-OUT-${Date.now()}`,
              dateTime: dt,
              storeId: transferData.fromStoreId,
              productId: transferData.productId,
              batch: transferData.batch,
              expiry: transferData.expiry,
              qty: -transferData.qty,
              type: 'Transfer out',
              ref: trId,
              user: role,
              reason: `Dispatched to ${storesMeta.find(s => s.id === transferData.toStoreId)?.name}`,
              source: 'App Billing'
            };

            const newTr: InterStoreTransfer = {
              id: trId,
              date: dt,
              fromStoreId: transferData.fromStoreId,
              toStoreId: transferData.toStoreId,
              productId: transferData.productId,
              batch: transferData.batch,
              expiry: transferData.expiry,
              qty: transferData.qty,
              status: 'Dispatched',
              user: role
            };

            setMovements(old => [outMov, ...old]);
            setTransfers(old => [newTr, ...old]);
            stampAudit(`Inter-store transfer created: ${trId} (${transferData.qty} units from ${transferData.fromStoreId} to ${transferData.toStoreId})`);
            notify(`Transfer dispatched: ${trId}. Stock deducted from source store.`);
            setShowTransferModal(false);
          }}
        />
      )}

      {/* PURCHASE STOCK-IN MODAL */}
      {showPurchaseInModal && (
        <PurchaseEntryModal
          stores={storesMeta}
          products={products}
          onClose={() => setShowPurchaseInModal(false)}
          onSave={(data) => {
            const dt = new Date().toISOString().slice(0, 16).replace('T', ' ');
            const newMov: StockMovement = {
              id: `MOV-PUR-${Date.now()}`,
              dateTime: dt,
              storeId: data.storeId,
              productId: data.productId,
              batch: data.batch,
              expiry: data.expiry,
              qty: data.qty + (data.freeQty || 0),
              type: 'Purchase (in)',
              ref: data.invoiceNo,
              user: role,
              reason: `Purchase stock-in from ${data.supplier}`,
              source: 'App Billing'
            };
            setMovements(old => [newMov, ...old]);
            stampAudit(`Purchase stock-in: ${data.invoiceNo} (${data.qty} units of ${data.productId})`);
            notify(`Purchase stock-in recorded! Added ${data.qty} units to ${data.storeId}`);
            setShowPurchaseInModal(false);
          }}
        />
      )}
    </>
  );
}

// -------------------------------------------------------------
// 2. STOCK ADJUSTMENT MODAL
// -------------------------------------------------------------
function StockAdjustmentModal({
  stores,
  products,
  onClose,
  onSave
}: {
  stores: StoreInfo[];
  products: Product[];
  onClose: () => void;
  onSave: (data: { storeId: string; productId: string; batch: string; expiry: string; qty: number; type: 'increase' | 'decrease'; reason: string; notes: string }) => void;
}) {
  const [storeId, setStoreId] = useState(stores[0].id);
  const [productId, setProductId] = useState(products[0].id);
  const [batch, setBatch] = useState('BT-2400');
  const [expiry, setExpiry] = useState('2026-11');
  const [qty, setQty] = useState(1);
  const [type, setType] = useState<'increase' | 'decrease'>('decrease');
  const [reason, setReason] = useState('Damage');
  const [notes, setNotes] = useState('');

  const submit = (e: FormEvent) => {
    e.preventDefault();
    if (qty <= 0) return;
    onSave({ storeId, productId, batch, expiry, qty, type, reason, notes });
  };

  return (
    <div className="modal-backdrop" onMouseDown={e => { if (e.target === e.currentTarget) onClose(); }}>
      <div className="modal">
        <div className="drawer-head">
          <h2>New Stock Adjustment</h2>
          <button className="close-button" onClick={onClose}>×</button>
        </div>
        <form onSubmit={submit}>
          <div className="notice warn" style={{ marginBottom: 14 }}>
            <b>Mandatory Reason Required:</b> Adjustments write off damages, exactions, audit count discrepancies or returns. All adjustments log an immutable movement.
          </div>
          <div className="two-col">
            <label className="field">
              Store
              <select value={storeId} onChange={e => setStoreId(e.target.value)}>
                {stores.map(s => <option key={s.id} value={s.id}>{s.name}</option>)}
              </select>
            </label>
            <label className="field">
              Product
              <select value={productId} onChange={e => setProductId(e.target.value)}>
                {products.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
              </select>
            </label>
          </div>
          <div className="two-col" style={{ marginTop: 12 }}>
            <label className="field">
              Batch
              <input value={batch} onChange={e => setBatch(e.target.value)} required />
            </label>
            <label className="field">
              Expiry Date
              <input type="month" value={expiry} onChange={e => setExpiry(e.target.value)} required />
            </label>
          </div>
          <div className="two-col" style={{ marginTop: 12 }}>
            <label className="field">
              Adjustment Direction
              <select value={type} onChange={e => setType(e.target.value as any)}>
                <option value="decrease">Decrease Stock (Write-off / Breakage / Expiry)</option>
                <option value="increase">Increase Stock (Count Surplus / Found Items)</option>
              </select>
            </label>
            <label className="field">
              Quantity
              <input type="number" min="1" value={qty} onChange={e => setQty(Number(e.target.value))} required />
            </label>
          </div>
          <div className="two-col" style={{ marginTop: 12 }}>
            <label className="field">
              Mandatory Reason
              <select value={reason} onChange={e => setReason(e.target.value)}>
                <option value="Damage">Damage (Broken / Leaking)</option>
                <option value="Expiry">Expiry Write-off</option>
                <option value="Count correction">Physical Count Correction</option>
                <option value="Sample / Tester">Sample Dispensing</option>
                <option value="Other">Other Operational Adjustment</option>
              </select>
            </label>
            <label className="field">
              Notes / Audit Reference
              <input placeholder="e.g. Broken in shelf aisle 3" value={notes} onChange={e => setNotes(e.target.value)} />
            </label>
          </div>
          <div className="modal-actions">
            <button type="button" className="btn" onClick={onClose}>Cancel</button>
            <button type="submit" className="btn primary">Confirm & Post Adjustment</button>
          </div>
        </form>
      </div>
    </div>
  );
}

// -------------------------------------------------------------
// 3. INTER-STORE TRANSFER MODAL
// -------------------------------------------------------------
function InterStoreTransferModal({
  stores,
  products,
  getLedgerStock,
  onClose,
  onSave
}: {
  stores: StoreInfo[];
  products: Product[];
  getLedgerStock: (productId: string, storeId?: string, batch?: string) => number;
  onClose: () => void;
  onSave: (data: { fromStoreId: string; toStoreId: string; productId: string; batch: string; expiry: string; qty: number }) => void;
}) {
  const [fromStoreId, setFromStoreId] = useState(stores[0].id);
  const [toStoreId, setToStoreId] = useState(stores[1]?.id || stores[0].id);
  const [productId, setProductId] = useState(products[0].id);
  const [batch, setBatch] = useState('BT-2400');
  const [expiry, setExpiry] = useState('2026-11');
  const [qty, setQty] = useState(5);
  const [error, setError] = useState('');

  const available = getLedgerStock(productId, fromStoreId, batch);

  const submit = (e: FormEvent) => {
    e.preventDefault();
    if (fromStoreId === toStoreId) {
      setError('Source and destination stores must be different');
      return;
    }
    if (qty > available) {
      setError(`Cannot transfer ${qty} units. Only ${available} available in source store batch.`);
      return;
    }
    onSave({ fromStoreId, toStoreId, productId, batch, expiry, qty });
  };

  return (
    <div className="modal-backdrop" onMouseDown={e => { if (e.target === e.currentTarget) onClose(); }}>
      <div className="modal">
        <div className="drawer-head">
          <h2>Create Inter-store Transfer</h2>
          <button className="close-button" onClick={onClose}>×</button>
        </div>
        <form onSubmit={submit}>
          <div className="notice info" style={{ marginBottom: 14 }}>
            <b>Paired Movement:</b> Deducts stock immediately from Source store on Dispatch. Destination store confirms receipt to add stock.
          </div>
          {error && <div className="notice warn" style={{ marginBottom: 12 }}>{error}</div>}
          <div className="two-col">
            <label className="field">
              From Store (Source)
              <select value={fromStoreId} onChange={e => { setFromStoreId(e.target.value); setError(''); }}>
                {stores.map(s => <option key={s.id} value={s.id}>{s.name}</option>)}
              </select>
            </label>
            <label className="field">
              To Store (Destination)
              <select value={toStoreId} onChange={e => { setToStoreId(e.target.value); setError(''); }}>
                {stores.filter(s => s.id !== fromStoreId).map(s => <option key={s.id} value={s.id}>{s.name}</option>)}
              </select>
            </label>
          </div>
          <div className="field" style={{ marginTop: 12 }}>
            Product
            <select value={productId} onChange={e => setProductId(e.target.value)}>
              {products.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
            </select>
          </div>
          <div className="two-col" style={{ marginTop: 12 }}>
            <label className="field">
              Batch
              <input value={batch} onChange={e => setBatch(e.target.value)} required />
            </label>
            <label className="field">
              Expiry
              <input type="month" value={expiry} onChange={e => setExpiry(e.target.value)} required />
            </label>
          </div>
          <div className="two-col" style={{ marginTop: 12 }}>
            <label className="field">
              Transfer Quantity
              <input type="number" min="1" max={available || 1} value={qty} onChange={e => setQty(Number(e.target.value))} required />
            </label>
            <div style={{ paddingTop: 20 }}>
              <span className="small-note">Available in source batch: <b>{available} units</b></span>
            </div>
          </div>
          <div className="modal-actions">
            <button type="button" className="btn" onClick={onClose}>Cancel</button>
            <button type="submit" className="btn primary" disabled={available <= 0}>Dispatch Transfer</button>
          </div>
        </form>
      </div>
    </div>
  );
}

// -------------------------------------------------------------
// 4. PURCHASE STOCK-IN ENTRY MODAL
// -------------------------------------------------------------
function PurchaseEntryModal({
  stores,
  products,
  onClose,
  onSave
}: {
  stores: StoreInfo[];
  products: Product[];
  onClose: () => void;
  onSave: (data: { storeId: string; supplier: string; invoiceNo: string; productId: string; batch: string; expiry: string; qty: number; freeQty: number; rate: number; mrp: number }) => void;
}) {
  const [storeId, setStoreId] = useState(stores[0].id);
  const [supplier, setSupplier] = useState(suppliers[0].name);
  const [invoiceNo, setInvoiceNo] = useState(`INV-2025-${Math.floor(1000 + Math.random() * 9000)}`);
  const [productId, setProductId] = useState(products[0].id);
  const [batch, setBatch] = useState(`BT-${Math.floor(2500 + Math.random() * 500)}`);
  const [expiry, setExpiry] = useState('2027-06');
  const [qty, setQty] = useState(50);
  const [freeQty, setFreeQty] = useState(0);
  const [rate, setRate] = useState(45);
  const [mrp, setMrp] = useState(65);

  return (
    <div className="modal-backdrop" onMouseDown={e => { if (e.target === e.currentTarget) onClose(); }}>
      <div className="modal">
        <div className="drawer-head">
          <h2>Purchase Stock-in Entry</h2>
          <button className="close-button" onClick={onClose}>×</button>
        </div>
        <div className="notice">
          <strong>Purpose:</strong> Record new stock arriving from a supplier. This increases stock and is relevant for GST purchase records.
        </div>
        <form onSubmit={e => {
          e.preventDefault();
          onSave({ storeId, supplier, invoiceNo, productId, batch, expiry, qty, freeQty, rate, mrp });
        }}>
          <div className="two-col">
            <label className="field">
              Store Receiving Goods
              <select value={storeId} onChange={e => setStoreId(e.target.value)}>
                {stores.map(s => <option key={s.id} value={s.id}>{s.name}</option>)}
              </select>
            </label>
            <label className="field">
              Supplier
              <select value={supplier} onChange={e => setSupplier(e.target.value)}>
                {suppliers.map(s => <option key={s.id} value={s.name}>{s.name}</option>)}
              </select>
            </label>
          </div>
          <div className="two-col" style={{ marginTop: 12 }}>
            <label className="field">
              Supplier Invoice / Bill No
              <input value={invoiceNo} onChange={e => setInvoiceNo(e.target.value)} required />
            </label>
            <label className="field">
              Product
              <select value={productId} onChange={e => setProductId(e.target.value)}>
                {products.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
              </select>
            </label>
          </div>
          <div className="two-col" style={{ marginTop: 12 }}>
            <label className="field">
              Batch
              <input value={batch} onChange={e => setBatch(e.target.value)} required />
            </label>
            <label className="field">
              Expiry Date
              <input type="month" value={expiry} onChange={e => setExpiry(e.target.value)} required />
            </label>
          </div>
          <div className="two-col" style={{ marginTop: 12 }}>
            <label className="field">
              Billed Qty
              <input type="number" min="1" value={qty} onChange={e => setQty(Number(e.target.value))} required />
            </label>
            <label className="field">
              Free Qty (Bonus)
              <input type="number" min="0" value={freeQty} onChange={e => setFreeQty(Number(e.target.value))} />
            </label>
          </div>
          <div className="two-col" style={{ marginTop: 12 }}>
            <label className="field">
              Purchase Rate (excl. GST)
              <input type="number" step="0.01" value={rate} onChange={e => setRate(Number(e.target.value))} required />
            </label>
            <label className="field">
              Printed MRP
              <input type="number" step="0.01" value={mrp} onChange={e => setMrp(Number(e.target.value))} required />
            </label>
          </div>
          <div className="modal-actions">
            <button type="button" className="btn" onClick={onClose}>Cancel</button>
            <button type="submit" className="btn primary">Record Purchase & Add Stock</button>
          </div>
        </form>
      </div>
    </div>
  );
}

// -------------------------------------------------------------
// 5. BILLING VIEW (COUNTER-STYLE POS TERMINAL)
// -------------------------------------------------------------
export function BillingView({
  role,
  activeStore,
  storesMeta,
  customers,
  setCustomers,
  movements,
  setMovements,
  invoices,
  setInvoices,
  prescriptions,
  setPrescriptions,
  pricesInclusive,
  allowNegativeStock,
  notify,
  stampAudit,
  getLedgerStock,
  prefillPrescriptionId
}: {
  role: string;
  activeStore: string;
  storesMeta: StoreInfo[];
  customers: Customer[];
  setCustomers: React.Dispatch<React.SetStateAction<Customer[]>>;
  movements: StockMovement[];
  setMovements: React.Dispatch<React.SetStateAction<StockMovement[]>>;
  invoices: Invoice[];
  setInvoices: React.Dispatch<React.SetStateAction<Invoice[]>>;
  prescriptions: any[];
  setPrescriptions: React.Dispatch<React.SetStateAction<any[]>>;
  pricesInclusive: boolean;
  allowNegativeStock: boolean;
  notify: (s: string) => void;
  stampAudit: (s: string) => void;
  getLedgerStock: (productId: string, storeId?: string, batch?: string) => number;
  prefillPrescriptionId?: string | null;
}) {
  const [billingStore, setBillingStore] = useState<string>(activeStore !== 'All' ? activeStore : 'S1');
  const [customerId, setCustomerId] = useState<string>('C01');
  const [doctorName, setDoctorName] = useState<string>('');
  const [prescriptionId, setPrescriptionId] = useState<string>(prefillPrescriptionId || '');
  const [lines, setLines] = useState<InvoiceLine[]>([]);
  const [paymentMode, setPaymentMode] = useState<'Cash' | 'UPI' | 'Card' | 'Credit'>('Cash');
  const [amountReceived, setAmountReceived] = useState<number>(0);

  // Line item picker state
  const [prodSearch, setProdSearch] = useState('');
  const [selectedProd, setSelectedProd] = useState<Product | null>(null);
  const [selectedBatch, setSelectedBatch] = useState('');
  const [qty, setQty] = useState(1);
  const [discPct, setDiscPct] = useState(0);
  const [showAddCustomer, setShowAddCustomer] = useState(false);
  const [printedInvoice, setPrintedInvoice] = useState<Invoice | null>(null);
  const [heldBills, setHeldBills] = useState<Array<{ id: string; time: string; customerName: string; lines: InvoiceLine[] }>>([]);

  const currentStore = storesMeta.find(s => s.id === billingStore) || storesMeta[0];
  const currentCustomer = customers.find(c => c.id === customerId) || customers[0];
  const isInterstate = currentCustomer.state !== currentStore.state;

  // Handle prefill if prescriptionId passed
  useState(() => {
    if (prefillPrescriptionId) {
      const rx = prescriptions.find(p => p.id === prefillPrescriptionId);
      if (rx) {
        setDoctorName('Dr. Prescriber, MBBS');
        // Find customer by phone or name
        const matchCust = customers.find(c => c.phone === rx.phone || c.name.toLowerCase() === rx.patient.toLowerCase());
        if (matchCust) {
          setCustomerId(matchCust.id);
        }
      }
    }
  });

  // Available batches for selected product in this store
  const availableBatches = useMemo(() => {
    if (!selectedProd) return [];
    const map = new Map<string, { batch: string; expiry: string; stock: number }>();
    movements
      .filter(m => m.productId === selectedProd.id && m.storeId === billingStore)
      .forEach(m => {
        const cur = map.get(m.batch) || { batch: m.batch, expiry: m.expiry, stock: 0 };
        cur.stock += m.qty;
        map.set(m.batch, cur);
      });
    return Array.from(map.values()).sort((a, b) => a.expiry.localeCompare(b.expiry)); // FEFO
  }, [selectedProd, billingStore, movements]);

  // Handle selecting a product: auto-select FEFO batch (first non-expired with stock > 0)
  const onPickProduct = (prod: Product) => {
    setSelectedProd(prod);
    setProdSearch(prod.name);
    // FEFO: Pick earliest expiry batch with stock > 0
    const prodBatches = movements
      .filter(m => m.productId === prod.id && m.storeId === billingStore)
      .reduce((acc, m) => {
        const found = acc.find(b => b.batch === m.batch);
        if (found) found.stock += m.qty;
        else acc.push({ batch: m.batch, expiry: m.expiry, stock: m.qty });
        return acc;
      }, [] as Array<{ batch: string; expiry: string; stock: number }>)
      .sort((a, b) => a.expiry.localeCompare(b.expiry));

    const fefoCandidate = prodBatches.find(b => b.stock > 0 && b.expiry >= '2025-06') || prodBatches[0];
    if (fefoCandidate) {
      setSelectedBatch(fefoCandidate.batch);
    } else {
      setSelectedBatch('BT-DEFAULT');
    }
  };

  const activeBatchObj = availableBatches.find(b => b.batch === selectedBatch);
  const batchAvailable = activeBatchObj ? activeBatchObj.stock : 0;
  const isBatchExpired = activeBatchObj ? activeBatchObj.expiry < '2025-06' : false;
  const prodMeta = selectedProd ? productGstMap[selectedProd.id] || { mrp: 60, hsn: '3004', gstRate: 12, unit: 'Strip 10s', schedule: 'None' } : null;

  // Add line to invoice
  const addLine = () => {
    if (!selectedProd || !prodMeta) return;
    if (isBatchExpired) {
      notify('Cannot add expired batch to invoice! Select another batch.');
      return;
    }
    if (qty <= 0) return;
    if (!allowNegativeStock && qty > batchAvailable) {
      notify(`Cannot sell ${qty} units. Only ${batchAvailable} units available in batch.`);
      return;
    }

    const mrp = prodMeta.mrp;
    const gstRate = prodMeta.gstRate;
    let rate = mrp;
    let taxableValue = 0;
    let discountAmt = 0;

    if (pricesInclusive) {
      // MRP inclusive: rate is back-calculated
      const gross = mrp * qty;
      discountAmt = Math.round((gross * discPct / 100) * 100) / 100;
      const netGross = gross - discountAmt;
      taxableValue = Math.round((netGross / (1 + gstRate / 100)) * 100) / 100;
      rate = Math.round((mrp / (1 + gstRate / 100)) * 100) / 100;
    } else {
      // Exclusive
      rate = mrp;
      const sub = rate * qty;
      discountAmt = Math.round((sub * discPct / 100) * 100) / 100;
      taxableValue = Math.round((sub - discountAmt) * 100) / 100;
    }

    const taxAmt = Math.round((taxableValue * gstRate / 100) * 100) / 100;
    const cgst = isInterstate ? 0 : Math.round((taxAmt / 2) * 100) / 100;
    const sgst = isInterstate ? 0 : Math.round((taxAmt / 2) * 100) / 100;
    const igst = isInterstate ? taxAmt : 0;
    const lineTotal = Math.round((taxableValue + cgst + sgst + igst) * 100) / 100;

    const newLine: InvoiceLine = {
      productId: selectedProd.id,
      productName: selectedProd.name,
      hsn: prodMeta.hsn,
      batch: selectedBatch,
      expiry: activeBatchObj?.expiry || '2026-11',
      qty,
      unit: prodMeta.unit,
      mrp,
      rate,
      discountPct: discPct,
      discountAmt,
      taxableValue,
      gstRate,
      cgst,
      sgst,
      igst,
      lineTotal
    };

    setLines(old => [...old, newLine]);
    setSelectedProd(null);
    setProdSearch('');
    setSelectedBatch('');
    setQty(1);
    setDiscPct(0);
  };

  // Recalculate totals
  const subtotal = Math.round(lines.reduce((n, l) => n + l.mrp * l.qty, 0) * 100) / 100;
  const discountTotal = Math.round(lines.reduce((n, l) => n + l.discountAmt, 0) * 100) / 100;
  const taxableTotal = Math.round(lines.reduce((n, l) => n + l.taxableValue, 0) * 100) / 100;
  const cgstTotal = Math.round(lines.reduce((n, l) => n + l.cgst, 0) * 100) / 100;
  const sgstTotal = Math.round(lines.reduce((n, l) => n + l.sgst, 0) * 100) / 100;
  const igstTotal = Math.round(lines.reduce((n, l) => n + l.igst, 0) * 100) / 100;
  const rawGrand = taxableTotal + cgstTotal + sgstTotal + igstTotal;
  const grandTotal = Math.round(rawGrand);
  const roundOff = Math.round((grandTotal - rawGrand) * 100) / 100;

  // Scheduled drug check in active lines
  const hasScheduledDrug = lines.some(l => {
    const meta = productGstMap[l.productId];
    return meta && meta.schedule !== 'None';
  });

  const nextInvoiceNo = `${currentStore.invoicePrefix}${String(invoices.length + 1).padStart(5, '0')}`;

  const saveInvoice = (asDraft = false) => {
    if (lines.length === 0) {
      notify('Please add at least one line item');
      return;
    }

    if (hasScheduledDrug && !doctorName.trim()) {
      notify('Scheduled drug requires Doctor Name before finalizing invoice!');
      return;
    }

    const dt = new Date().toISOString().slice(0, 16).replace('T', ' ');
    const newInv: Invoice = {
      id: nextInvoiceNo,
      date: dt,
      storeId: billingStore,
      customerId: currentCustomer.id,
      customerName: currentCustomer.name,
      customerPhone: currentCustomer.phone,
      customerAddress: currentCustomer.address,
      customerGstin: currentCustomer.gstin,
      customerState: currentCustomer.state,
      isB2B: currentCustomer.type === 'B2B',
      doctorName: doctorName.trim() || undefined,
      prescriptionId: prescriptionId || undefined,
      lines,
      subtotal,
      discountTotal,
      taxableTotal,
      cgstTotal,
      sgstTotal,
      igstTotal,
      roundOff,
      grandTotal,
      paymentMode,
      amountReceived: amountReceived || grandTotal,
      status: asDraft ? 'Draft' : 'Final',
      user: role,
      source: 'App Billing'
    };

    if (!asDraft) {
      // Create stock movement deducting stock immediately
      const newMovements: StockMovement[] = lines.map((l, idx) => ({
        id: `MOV-INV-${Date.now()}-${idx}`,
        dateTime: dt,
        storeId: billingStore,
        productId: l.productId,
        batch: l.batch,
        expiry: l.expiry,
        qty: -l.qty,
        type: 'Sale (invoice)',
        ref: nextInvoiceNo,
        user: role,
        source: 'App Billing'
      }));

      setMovements(old => [...newMovements, ...old]);

      // If prescription linked, update prescription status & reference
      if (prescriptionId) {
        setPrescriptions(old => old.map(rx => rx.id === prescriptionId ? {
          ...rx,
          status: 'Process/Print',
          reference: `${rx.reference ? rx.reference + ' · ' : ''}Billed: ${nextInvoiceNo}`
        } : rx));
      }

      stampAudit(`Invoice generated: ${nextInvoiceNo} (${money(grandTotal)} - ${currentCustomer.name})`);
      const firstLine = lines[0];
      const prevStock = getLedgerStock(firstLine.productId, billingStore, firstLine.batch);
      notify(`Stock updated: ${firstLine.productName} (${currentStore.name}) ${prevStock} → ${prevStock - firstLine.qty}`);
      setPrintedInvoice(newInv);
    } else {
      notify(`Draft saved: ${nextInvoiceNo}`);
    }

    setInvoices(old => [newInv, ...old]);
    setLines([]);
    setDoctorName('');
    setPrescriptionId('');
  };

  return (
    <>
      <div className="page-head">
        <div>
          <div className="eyebrow">COUNTER BILLING & DISPENSING</div>
          <h1 className="page-title">Billing Terminal</h1>
          <p className="help-line">What this screen does: Generate GST-compliant Tax Invoices, auto-pick batches by FEFO, and deduct stock from the store ledger immediately.</p>
        </div>
        <div className="page-head-actions">
          {heldBills.length > 0 && (
            <div className="pill-group">
              <span style={{ padding: '8px 10px', fontSize: 11, background: '#fcf3e8', color: '#976127', fontWeight: 600 }}>
                Held bills ({heldBills.length})
              </span>
              {heldBills.map(h => (
                <button key={h.id} onClick={() => {
                  setLines(h.lines);
                  setHeldBills(old => old.filter(x => x.id !== h.id));
                  notify(`Restored held bill for ${h.customerName}`);
                }}>
                  {h.customerName} ({h.time})
                </button>
              ))}
            </div>
          )}
        </div>
      </div>

      <div className="notice info" style={{ marginBottom: 14 }}>
        <b>Notice:</b> Sales made through this app deduct stock immediately. ERP sales imports dated after go-live must not be double counted.
        <span style={{ marginLeft: 10, color: '#275249' }}>
          Prices are: <b>{pricesInclusive ? 'MRP-inclusive of GST' : 'Exclusive of GST'}</b> (changeable in Admin settings).
        </span>
      </div>

      {/* POS Top Bar */}
      <div className="panel" style={{ padding: 16, marginBottom: 16 }}>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: 14 }}>
          <label className="field">
            Store Location
            <select value={billingStore} onChange={e => setBillingStore(e.target.value)}>
              {storesMeta.map(s => <option key={s.id} value={s.id}>{s.name} ({s.id})</option>)}
            </select>
          </label>
          <div>
            <div className="stat-label">INVOICE NO (NEXT)</div>
            <div style={{ font: '800 16px Manrope', marginTop: 4 }}><span className="mono">{nextInvoiceNo}</span></div>
          </div>
          <div>
            <div className="stat-label">TAX REGISTRATION</div>
            <div className="row-secondary" style={{ marginTop: 4 }}>GSTIN: <span className="mono">{currentStore.gstin}</span></div>
            <div className="row-secondary">DL: {currentStore.drugLicence.slice(0, 18)}…</div>
          </div>
          <div>
            <div className="stat-label">PLACE OF SUPPLY</div>
            <div style={{ marginTop: 4 }}>
              <span className={`badge ${isInterstate ? 'warn' : ''}`}>
                {isInterstate ? `Inter-state (IGST 100%) · ${currentCustomer.state}` : `Intra-state (CGST + SGST) · Karnataka (29)`}
              </span>
            </div>
          </div>
        </div>
      </div>

      {/* Main Billing Grid */}
      <div className="content-grid" style={{ gridTemplateColumns: '1fr 340px' }}>
        <div>
          {/* Customer & Prescription Header */}
          <div className="panel" style={{ padding: 16, marginBottom: 16 }}>
            <div style={{ display: 'grid', gridTemplateColumns: '1.2fr 1fr 1fr', gap: 12 }}>
              <label className="field">
                Customer / Patient
                <div style={{ display: 'flex', gap: 6 }}>
                  <select style={{ flex: 1 }} value={customerId} onChange={e => setCustomerId(e.target.value)}>
                    {customers.map(c => (
                      <option key={c.id} value={c.id}>
                        {c.name} {c.type === 'B2B' ? `(B2B - ${c.gstin?.slice(0, 6)}…)` : ''}
                      </option>
                    ))}
                  </select>
                  <button type="button" className="btn tiny" onClick={() => setShowAddCustomer(true)} title="Add customer inline">
                    <Plus size={13} /> Add
                  </button>
                </div>
              </label>

              <label className="field">
                Doctor Name {hasScheduledDrug ? <span style={{ color: '#a0443e' }}>* (Required)</span> : '(Optional)'}
                <input
                  placeholder="e.g. Dr. Priya Sharma, MBBS"
                  value={doctorName}
                  onChange={e => setDoctorName(e.target.value)}
                  style={{ borderColor: hasScheduledDrug && !doctorName.trim() ? '#e8c9bd' : undefined }}
                />
              </label>

              <label className="field">
                Link Prescription
                <select value={prescriptionId} onChange={e => {
                  setPrescriptionId(e.target.value);
                  const rx = prescriptions.find(p => p.id === e.target.value);
                  if (rx) {
                    setDoctorName('Dr. Prescriber, MBBS');
                    const matchCust = customers.find(c => c.name.toLowerCase() === rx.patient.toLowerCase() || c.phone === rx.phone);
                    if (matchCust) setCustomerId(matchCust.id);
                  }
                }}>
                  <option value="">None (Walk-in)</option>
                  {prescriptions.filter(p => p.status !== 'Closed').map(p => (
                    <option key={p.id} value={p.id}>{p.id} · {p.patient} ({p.status})</option>
                  ))}
                </select>
              </label>
            </div>
          </div>

          {/* Line item entry box */}
          <div className="panel" style={{ padding: 16, marginBottom: 16 }}>
            <div className="eyebrow">ADD MEDICINE / PRODUCT LINE</div>
            <div style={{ display: 'grid', gridTemplateColumns: '2fr 1.2fr 80px 80px auto', gap: 10, marginTop: 10, alignItems: 'flex-end' }}>
              <div style={{ position: 'relative' }}>
                <label className="field">
                  Search Product
                  <input
                    placeholder="Type brand name or alias (e.g. Dolo, Paracetamol)..."
                    value={prodSearch}
                    onChange={e => setProdSearch(e.target.value)}
                  />
                </label>
                {prodSearch && !selectedProd && (
                  <div style={{ position: 'absolute', zIndex: 30, top: '100%', left: 0, right: 0, background: '#fff', border: '1px solid #dcd8ce', borderRadius: 8, boxShadow: '0 8px 24px rgba(0,0,0,0.12)', maxHeight: 220, overflow: 'auto' }}>
                    {products.filter(p => `${p.name} ${p.aliases.join(' ')}`.toLowerCase().includes(prodSearch.toLowerCase())).slice(0, 8).map(p => (
                      <div
                        key={p.id}
                        onClick={() => onPickProduct(p)}
                        style={{ padding: '8px 12px', borderBottom: '1px solid #f0eee8', cursor: 'pointer' }}
                        className="search-item"
                      >
                        <b>{p.name}</b> <span className="small-note">({p.pack} · {p.group})</span>
                        <div className="row-secondary">Stock in store: {getLedgerStock(p.id, billingStore)} units · MRP ₹{productGstMap[p.id]?.mrp}</div>
                      </div>
                    ))}
                  </div>
                )}
              </div>

              <label className="field">
                Batch (FEFO Selected)
                <select
                  value={selectedBatch}
                  onChange={e => setSelectedBatch(e.target.value)}
                  disabled={!selectedProd || availableBatches.length === 0}
                >
                  {availableBatches.map(b => (
                    <option key={b.batch} value={b.batch}>
                      {b.batch} (Exp {b.expiry}) · {b.stock} units
                    </option>
                  ))}
                  {availableBatches.length === 0 && <option value="">No batches</option>}
                </select>
              </label>

              <label className="field">
                Qty
                <input type="number" min="1" value={qty} onChange={e => setQty(Math.max(1, Number(e.target.value) || 1))} />
              </label>

              <label className="field">
                Disc %
                <input type="number" min="0" max="100" value={discPct} onChange={e => setDiscPct(Number(e.target.value) || 0)} />
              </label>

              <button
                type="button"
                className="btn primary"
                onClick={addLine}
                disabled={!selectedProd || isBatchExpired || (!allowNegativeStock && qty > batchAvailable)}
                style={{ height: 38 }}
              >
                <Plus size={14} /> Add Line
              </button>
            </div>

            {/* Product selection context alerts */}
            {selectedProd && prodMeta && (
              <div style={{ marginTop: 10, display: 'flex', gap: 10, flexWrap: 'wrap', alignItems: 'center' }}>
                <span className="small-note">
                  Selected: <b>{selectedProd.name}</b> · HSN: <b>{prodMeta.hsn}</b> · GST Rate: <b>{prodMeta.gstRate}%</b> · MRP: <b>₹{prodMeta.mrp}</b>
                </span>
                {isBatchExpired ? (
                  <span className="badge error">Expired batch ({activeBatchObj?.expiry}) — dispensing blocked!</span>
                ) : (
                  <span className="small-note" style={{ color: '#27705e' }}>
                    Available in batch: <b>{batchAvailable} units</b> (After sale: {batchAvailable - qty})
                  </span>
                )}
                {prodMeta.schedule !== 'None' && (
                  <span className="badge warn">Scheduled Drug ({prodMeta.schedule}): Doctor name required</span>
                )}
              </div>
            )}
          </div>

          {/* Line items table */}
          <div className="panel">
            <div className="panel-head">
              <div className="panel-title">Invoice Lines ({lines.length})</div>
              {hasScheduledDrug && (
                <span className="badge warn">Contains Scheduled Medicine</span>
              )}
            </div>
            <div className="table-wrap">
              <table className="data-table">
                <thead>
                  <tr>
                    <th>Item & HSN</th>
                    <th>Batch</th>
                    <th>Exp</th>
                    <th>Qty</th>
                    <th>MRP</th>
                    <th>Rate</th>
                    <th>Taxable</th>
                    <th>GST</th>
                    <th>Total</th>
                    <th></th>
                  </tr>
                </thead>
                <tbody>
                  {lines.map((line, idx) => (
                    <tr key={idx}>
                      <td className="product-cell">
                        {line.productName}
                        <div className="row-secondary">HSN: {line.hsn} · {line.unit}</div>
                      </td>
                      <td><span className="mono">{line.batch}</span></td>
                      <td>{line.expiry}</td>
                      <td><b>{line.qty}</b></td>
                      <td>{money(line.mrp)}</td>
                      <td>{money(line.rate)}</td>
                      <td>{money(line.taxableValue)}</td>
                      <td>
                        {isInterstate ? (
                          <span>IGST {line.gstRate}% ({money(line.igst)})</span>
                        ) : (
                          <span>{line.gstRate}% ({money(line.cgst + line.sgst)})</span>
                        )}
                      </td>
                      <td><b>{money(line.lineTotal)}</b></td>
                      <td>
                        <button
                          className="btn tiny ghost"
                          style={{ color: '#b34742' }}
                          onClick={() => setLines(old => old.filter((_, i) => i !== idx))}
                          title="Remove item"
                        >
                          <Trash2 size={13} />
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {lines.length === 0 && (
                <div className="empty-state">
                  <ShoppingCart size={24} />
                  <strong>Invoice is empty</strong>
                  <p>Search and add items using the form above.</p>
                </div>
              )}
            </div>
          </div>
        </div>

        {/* Right Sidebar: Totals & Checkout */}
        <div>
          <div className="panel" style={{ padding: 18, position: 'sticky', top: 80 }}>
            <div className="eyebrow">PAYMENT & SETTLEMENT</div>
            <h3 style={{ font: '800 18px Manrope', margin: '4px 0 16px' }}>Invoice Summary</h3>

            <div className="data-list" style={{ padding: 0 }}>
              <div className="list-row">
                <span className="row-primary">Subtotal (Gross MRP)</span>
                <span>{money(subtotal)}</span>
              </div>
              {discountTotal > 0 && (
                <div className="list-row" style={{ color: '#b26829' }}>
                  <span className="row-primary">Total Discount</span>
                  <span>- {money(discountTotal)}</span>
                </div>
              )}
              <div className="list-row">
                <span className="row-primary">Taxable Value</span>
                <span><b>{money(taxableTotal)}</b></span>
              </div>
              {!isInterstate ? (
                <>
                  <div className="list-row">
                    <span className="row-primary">CGST</span>
                    <span>{money(cgstTotal)}</span>
                  </div>
                  <div className="list-row">
                    <span className="row-primary">SGST</span>
                    <span>{money(sgstTotal)}</span>
                  </div>
                </>
              ) : (
                <div className="list-row">
                  <span className="row-primary">IGST (Inter-state)</span>
                  <span>{money(igstTotal)}</span>
                </div>
              )}
              {roundOff !== 0 && (
                <div className="list-row">
                  <span className="row-primary">Round Off</span>
                  <span>{roundOff > 0 ? `+${money(roundOff)}` : money(roundOff)}</span>
                </div>
              )}
              <div className="list-row" style={{ borderTop: '2px solid #20343d', paddingTop: 14 }}>
                <div>
                  <div className="stat-label">GRAND TOTAL</div>
                  <div style={{ font: '800 24px Manrope', color: '#256c5d' }}>{money(grandTotal)}</div>
                </div>
              </div>
            </div>

            <div style={{ background: '#f5f4ef', padding: '9px 12px', borderRadius: 7, fontSize: 11, fontStyle: 'italic', margin: '12px 0' }}>
              {numberToWordsIndian(grandTotal)}
            </div>

            {/* Payment Mode */}
            <div style={{ marginTop: 14 }}>
              <div className="stat-label" style={{ marginBottom: 6 }}>PAYMENT MODE</div>
              <div className="pill-group" style={{ width: '100%' }}>
                {(['Cash', 'UPI', 'Card', 'Credit'] as const).map(mode => (
                  <button
                    key={mode}
                    type="button"
                    className={paymentMode === mode ? 'selected' : ''}
                    onClick={() => setPaymentMode(mode)}
                    style={{ flex: 1 }}
                  >
                    {mode}
                  </button>
                ))}
              </div>
            </div>

            {paymentMode === 'Cash' && (
              <div style={{ marginTop: 12 }}>
                <label className="field">
                  Amount Received
                  <input
                    type="number"
                    value={amountReceived || grandTotal}
                    onChange={e => setAmountReceived(Number(e.target.value))}
                  />
                </label>
                {(amountReceived > grandTotal) && (
                  <div className="small-note" style={{ marginTop: 4, color: '#26705d', fontWeight: 600 }}>
                    Change to return: {money(amountReceived - grandTotal)}
                  </div>
                )}
              </div>
            )}

            <div style={{ display: 'grid', gap: 8, marginTop: 20 }}>
              <button
                type="button"
                className="btn primary"
                onClick={() => saveInvoice(false)}
                disabled={lines.length === 0}
                style={{ padding: 13, justifyContent: 'center' }}
              >
                <Check size={16} /> Save & Print Tax Invoice
              </button>

              <div style={{ display: 'flex', gap: 8 }}>
                <button
                  type="button"
                  className="btn"
                  style={{ flex: 1 }}
                  onClick={() => saveInvoice(true)}
                  disabled={lines.length === 0}
                >
                  Save as Draft
                </button>
                <button
                  type="button"
                  className="btn"
                  style={{ flex: 1 }}
                  onClick={() => {
                    if (lines.length === 0) return;
                    setHeldBills(old => [...old, {
                      id: `HOLD-${Date.now()}`,
                      time: new Date().toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit' }),
                      customerName: currentCustomer.name,
                      lines
                    }]);
                    setLines([]);
                    notify('Bill held. You can resume it anytime.');
                  }}
                  disabled={lines.length === 0}
                >
                  Hold Bill
                </button>
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* NEW CUSTOMER MODAL */}
      {showAddCustomer && (
        <NewCustomerModal
          onClose={() => setShowAddCustomer(false)}
          onSave={(cust) => {
            setCustomers(old => [...old, cust]);
            setCustomerId(cust.id);
            stampAudit(`New customer registered: ${cust.name} (${cust.type})`);
            notify(`Customer ${cust.name} added!`);
            setShowAddCustomer(false);
          }}
        />
      )}

      {/* PRINT TAX INVOICE MODAL */}
      {printedInvoice && (
        <TaxInvoicePrintModal
          invoice={printedInvoice}
          store={currentStore}
          onClose={() => setPrintedInvoice(null)}
        />
      )}
    </>
  );
}

// -------------------------------------------------------------
// 6. INVOICES VIEW (LIST, CANCELLATION, CREDIT NOTE)
// -------------------------------------------------------------
export function InvoicesView({
  role,
  activeStore,
  storesMeta,
  invoices,
  setInvoices,
  movements,
  setMovements,
  creditNotes,
  setCreditNotes,
  stampAudit,
  notify
}: {
  role: string;
  activeStore: string;
  storesMeta: StoreInfo[];
  invoices: Invoice[];
  setInvoices: React.Dispatch<React.SetStateAction<Invoice[]>>;
  movements: StockMovement[];
  setMovements: React.Dispatch<React.SetStateAction<StockMovement[]>>;
  creditNotes: CreditNote[];
  setCreditNotes: React.Dispatch<React.SetStateAction<CreditNote[]>>;
  stampAudit: (s: string) => void;
  notify: (s: string) => void;
}) {
  const [storeFilter, setStoreFilter] = useState(activeStore);
  const [statusFilter, setStatusFilter] = useState<'All' | 'Final' | 'Draft' | 'Cancelled'>('All');
  const [gstTypeFilter, setGstTypeFilter] = useState<'All' | 'B2B' | 'B2C'>('All');
  const [search, setSearch] = useState('');
  const [selectedInvoice, setSelectedInvoice] = useState<Invoice | null>(null);
  const [cancelTarget, setCancelTarget] = useState<Invoice | null>(null);
  const [creditNoteTarget, setCreditNoteTarget] = useState<Invoice | null>(null);
  const [printTarget, setPrintTarget] = useState<Invoice | null>(null);

  const filtered = invoices.filter(inv => {
    if (storeFilter !== 'All' && inv.storeId !== storeFilter) return false;
    if (statusFilter !== 'All' && inv.status !== statusFilter) return false;
    if (gstTypeFilter === 'B2B' && !inv.isB2B) return false;
    if (gstTypeFilter === 'B2C' && inv.isB2B) return false;
    if (search && !`${inv.id} ${inv.customerName} ${inv.customerPhone}`.toLowerCase().includes(search.toLowerCase())) return false;
    return true;
  });

  return (
    <>
      <div className="page-head">
        <div>
          <div className="eyebrow">GST TAX INVOICES & REGISTERS</div>
          <h1 className="page-title">Invoices</h1>
          <p className="help-line">What this screen does: View finalized tax bills, print A4/Thermal receipts, issue credit notes, or cancel bills with automatic stock reversals.</p>
        </div>
        <div className="page-head-actions">
          <button className="btn" onClick={() => {
            const rows = filtered.map(inv => ({
              'Invoice No': inv.id,
              Date: inv.date,
              Store: inv.storeId,
              Customer: inv.customerName,
              Type: inv.isB2B ? 'B2B' : 'B2C',
              'Taxable Value': inv.taxableTotal,
              CGST: inv.cgstTotal,
              SGST: inv.sgstTotal,
              IGST: inv.igstTotal,
              'Grand Total': inv.grandTotal,
              Payment: inv.paymentMode,
              Status: inv.status
            }));
            const ws = XLSX.utils.json_to_sheet(rows);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, 'Invoices');
            XLSX.writeFile(wb, `invoices-${today}.xlsx`);
          }}>
            <ArrowDownToLine size={14} /> Export Excel
          </button>
        </div>
      </div>

      <div className="panel">
        <div className="toolbar">
          <div className="global-search" style={{ width: 240 }}>
            <Search size={14} />
            <input placeholder="Search invoice no or customer..." value={search} onChange={e => setSearch(e.target.value)} />
          </div>

          <select className="control" value={storeFilter} onChange={e => setStoreFilter(e.target.value)}>
            <option value="All">All Stores</option>
            {storesMeta.map(s => <option key={s.id} value={s.id}>{s.name}</option>)}
          </select>

          <select className="control" value={statusFilter} onChange={e => setStatusFilter(e.target.value as any)}>
            <option value="All">All Statuses</option>
            <option value="Final">Final</option>
            <option value="Draft">Draft</option>
            <option value="Cancelled">Cancelled</option>
          </select>

          <select className="control" value={gstTypeFilter} onChange={e => setGstTypeFilter(e.target.value as any)}>
            <option value="All">All Types (B2B & B2C)</option>
            <option value="B2B">B2B Registered</option>
            <option value="B2C">B2C Retail</option>
          </select>

          <button className="btn tiny ghost" onClick={() => { setStoreFilter('All'); setStatusFilter('All'); setGstTypeFilter('All'); setSearch(''); }}>
            Clear Filters
          </button>
        </div>

        <div className="table-wrap">
          <table className="data-table">
            <thead>
              <tr>
                <th>Invoice No</th>
                <th>Date & Time</th>
                <th>Store</th>
                <th>Customer</th>
                <th>GSTIN</th>
                <th>Type</th>
                <th>Taxable</th>
                <th>Taxes</th>
                <th>Grand Total</th>
                <th>Payment</th>
                <th>Status</th>
                <th>Actions</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map(inv => {
                const store = storesMeta.find(s => s.id === inv.storeId) || storesMeta[0];
                return (
                  <tr key={inv.id} onClick={() => setSelectedInvoice(inv)}>
                    <td className="product-cell"><span className="mono">{inv.id}</span></td>
                    <td>{inv.date}</td>
                    <td>{store.name}</td>
                    <td>
                      {inv.customerName}
                      <div className="row-secondary">{inv.customerPhone}</div>
                    </td>
                    <td><span className="mono">{inv.customerGstin || '—'}</span></td>
                    <td><span className="badge blue">{inv.isB2B ? 'B2B' : 'B2C'}</span></td>
                    <td>{money(inv.taxableTotal)}</td>
                    <td>{money(inv.cgstTotal + inv.sgstTotal + inv.igstTotal)}</td>
                    <td><b>{money(inv.grandTotal)}</b></td>
                    <td>{inv.paymentMode}</td>
                    <td><StatusBadge>{inv.status}</StatusBadge></td>
                    <td>
                      <div style={{ display: 'flex', gap: 6 }} onClick={e => e.stopPropagation()}>
                        <button className="btn tiny" onClick={() => setPrintTarget(inv)}>
                          <Printer size={12} /> Print
                        </button>
                        {inv.status === 'Final' && role !== 'Staff' && (
                          <>
                            <button className="btn tiny" onClick={() => setCreditNoteTarget(inv)} title="Create Credit Note / Sales Return">
                              Return
                            </button>
                            <button className="btn tiny danger" onClick={() => setCancelTarget(inv)} title="Cancel Invoice with stock reversal">
                              Cancel
                            </button>
                          </>
                        )}
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
          {filtered.length === 0 && (
            <div className="empty-state">
              <Receipt size={24} />
              <strong>No invoices match filters</strong>
              <p>Adjust your search criteria or switch store.</p>
            </div>
          )}
        </div>
      </div>

      {/* DETAIL DRAWER */}
      {selectedInvoice && (
        <div className="drawer-shade" onMouseDown={e => { if (e.target === e.currentTarget) setSelectedInvoice(null); }}>
          <div className="drawer" style={{ width: 'min(700px, 100vw)' }}>
            <div className="drawer-head">
              <div>
                <div className="eyebrow">GST TAX INVOICE DETAILS</div>
                <h2>{selectedInvoice.id}</h2>
              </div>
              <button className="close-button" onClick={() => setSelectedInvoice(null)}>×</button>
            </div>

            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 14 }}>
              <StatusBadge>{selectedInvoice.status}</StatusBadge>
              <span className="small-note">Billed by: {selectedInvoice.user} · {selectedInvoice.date}</span>
            </div>

            {selectedInvoice.status === 'Cancelled' && (
              <div className="notice warn" style={{ marginBottom: 14 }}>
                <b>Cancelled Invoice:</b> {selectedInvoice.cancelReason || 'Customer requested return'}. Stock movements were reversed back to the store inventory.
              </div>
            )}

            {selectedInvoice.creditNoteId && (
              <div className="notice info" style={{ marginBottom: 14 }}>
                <b>Sales Return Linked:</b> Credit note {selectedInvoice.creditNoteId} was issued for this invoice.
              </div>
            )}

            <div className="two-col" style={{ background: '#f5f7f6', padding: 14, borderRadius: 8, marginBottom: 16 }}>
              <div>
                <div className="stat-label">BILLED TO</div>
                <b>{selectedInvoice.customerName}</b>
                <div className="row-secondary">{selectedInvoice.customerPhone}</div>
                <div className="row-secondary">{selectedInvoice.customerAddress}</div>
                {selectedInvoice.customerGstin && <div className="mono" style={{ fontSize: 11, marginTop: 4 }}>GSTIN: {selectedInvoice.customerGstin}</div>}
              </div>
              <div>
                <div className="stat-label">STORE & DOCTOR</div>
                <b>{storesMeta.find(s => s.id === selectedInvoice.storeId)?.name}</b>
                <div className="row-secondary">Place of Supply: {selectedInvoice.customerState}</div>
                {selectedInvoice.doctorName && <div className="row-secondary">Doctor: {selectedInvoice.doctorName}</div>}
                {selectedInvoice.prescriptionId && <div className="row-secondary">Rx Ref: {selectedInvoice.prescriptionId}</div>}
              </div>
            </div>

            <div className="table-wrap">
              <table className="data-table">
                <thead>
                  <tr>
                    <th>Item</th>
                    <th>Batch</th>
                    <th>Qty</th>
                    <th>Rate</th>
                    <th>Taxable</th>
                    <th>Tax</th>
                    <th>Total</th>
                  </tr>
                </thead>
                <tbody>
                  {selectedInvoice.lines.map((l, i) => (
                    <tr key={i}>
                      <td className="product-cell">{l.productName}</td>
                      <td><span className="mono">{l.batch}</span></td>
                      <td>{l.qty}</td>
                      <td>{money(l.rate)}</td>
                      <td>{money(l.taxableValue)}</td>
                      <td>{money(l.cgst + l.sgst + l.igst)}</td>
                      <td><b>{money(l.lineTotal)}</b></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div style={{ marginTop: 18, background: '#faf9f5', border: '1px solid #eeeae2', padding: 14, borderRadius: 8 }}>
              <div className="data-list" style={{ padding: 0 }}>
                <div className="list-row"><span className="row-primary">Taxable Amount</span><span>{money(selectedInvoice.taxableTotal)}</span></div>
                <div className="list-row"><span className="row-primary">Total GST (CGST+SGST / IGST)</span><span>{money(selectedInvoice.cgstTotal + selectedInvoice.sgstTotal + selectedInvoice.igstTotal)}</span></div>
                <div className="list-row"><span className="row-primary">Round Off</span><span>{money(selectedInvoice.roundOff)}</span></div>
                <div className="list-row"><span className="row-primary">Grand Total</span><b>{money(selectedInvoice.grandTotal)}</b></div>
              </div>
              <div style={{ fontSize: 11, fontStyle: 'italic', marginTop: 8 }}>
                {numberToWordsIndian(selectedInvoice.grandTotal)}
              </div>
            </div>

            <div className="notice" style={{ marginTop: 14, fontSize: 10 }}>
              Final invoices are not editable. In accordance with GST compliance, corrections must be made via Credit Notes or formal Cancellation.
            </div>

            <div className="modal-actions" style={{ marginTop: 18 }}>
              <button className="btn" onClick={() => setPrintTarget(selectedInvoice)}>
                <Printer size={14} /> Print Preview
              </button>
              {selectedInvoice.status === 'Final' && role !== 'Staff' && (
                <>
                  <button className="btn" onClick={() => { setCreditNoteTarget(selectedInvoice); setSelectedInvoice(null); }}>
                    Create Credit Note
                  </button>
                  <button className="btn danger" onClick={() => { setCancelTarget(selectedInvoice); setSelectedInvoice(null); }}>
                    Cancel Invoice
                  </button>
                </>
              )}
            </div>
          </div>
        </div>
      )}

      {/* CANCELLATION MODAL */}
      {cancelTarget && (
        <CancelInvoiceModal
          invoice={cancelTarget}
          onClose={() => setCancelTarget(null)}
          onConfirm={(reason) => {
            const dt = new Date().toISOString().slice(0, 16).replace('T', ' ');
            // Reverse stock movements
            const reversalMovements: StockMovement[] = cancelTarget.lines.map((l, idx) => ({
              id: `MOV-REV-${Date.now()}-${idx}`,
              dateTime: dt,
              storeId: cancelTarget.storeId,
              productId: l.productId,
              batch: l.batch,
              expiry: l.expiry,
              qty: l.qty, // Add back stock
              type: 'Adjustment',
              ref: `CANCEL-${cancelTarget.id}`,
              user: role,
              reason: `Invoice Cancelled: ${reason}`,
              source: 'App Billing'
            }));

            setMovements(old => [...reversalMovements, ...old]);
            setInvoices(old => old.map(inv => inv.id === cancelTarget.id ? { ...inv, status: 'Cancelled', cancelReason: reason } : inv));
            stampAudit(`Invoice cancelled: ${cancelTarget.id} (Reason: ${reason}) - Stock reversed`);
            notify(`Invoice ${cancelTarget.id} cancelled. Stock restored to original batches.`);
            setCancelTarget(null);
          }}
        />
      )}

      {/* CREDIT NOTE MODAL */}
      {creditNoteTarget && (
        <CreditNoteModal
          invoice={creditNoteTarget}
          stores={storesMeta}
          onClose={() => setCreditNoteTarget(null)}
          onConfirm={(cnData) => {
            const dt = new Date().toISOString().slice(0, 16).replace('T', ' ');
            const cnNo = `CN/2026-27/${String(creditNotes.length + 1).padStart(5, '0')}`;

            // Add stock back to batch
            const returnMovements: StockMovement[] = cnData.returnLines.map((l, idx) => ({
              id: `MOV-CN-${Date.now()}-${idx}`,
              dateTime: dt,
              storeId: creditNoteTarget.storeId,
              productId: l.productId,
              batch: l.batch,
              expiry: l.expiry,
              qty: l.qty,
              type: 'Sales return / credit note',
              ref: cnNo,
              user: role,
              reason: `Sales Return Credit Note against ${creditNoteTarget.id}: ${cnData.reason}`,
              source: 'App Billing'
            }));

            const newCN: CreditNote = {
              id: cnNo,
              date: dt,
              storeId: creditNoteTarget.storeId,
              originalInvoiceId: creditNoteTarget.id,
              customerName: creditNoteTarget.customerName,
              lines: cnData.returnLines,
              taxableTotal: cnData.taxableTotal,
              cgstTotal: cnData.cgstTotal,
              sgstTotal: cnData.sgstTotal,
              igstTotal: cnData.igstTotal,
              grandTotal: cnData.grandTotal,
              reason: cnData.reason,
              user: role
            };

            setMovements(old => [...returnMovements, ...old]);
            setCreditNotes(old => [newCN, ...old]);
            setInvoices(old => old.map(inv => inv.id === creditNoteTarget.id ? { ...inv, creditNoteId: cnNo } : inv));
            stampAudit(`Credit note created: ${cnNo} for ${creditNoteTarget.id} (${money(cnData.grandTotal)})`);
            notify(`Credit Note ${cnNo} generated! Stock returned to store.`);
            setCreditNoteTarget(null);
          }}
        />
      )}

      {/* PRINT MODAL */}
      {printTarget && (
        <TaxInvoicePrintModal
          invoice={printTarget}
          store={storesMeta.find(s => s.id === printTarget.storeId) || storesMeta[0]}
          onClose={() => setPrintTarget(null)}
        />
      )}
    </>
  );
}

// -------------------------------------------------------------
// 7. TAX INVOICE PRINT MODAL (A4 TAX INVOICE & 80MM RECEIPT)
// -------------------------------------------------------------
export function TaxInvoicePrintModal({
  invoice,
  store,
  onClose
}: {
  invoice: Invoice;
  store: StoreInfo;
  onClose: () => void;
}) {
  const [printFormat, setPrintFormat] = useState<'A4' | 'Thermal80'>('A4');

  return (
    <div className="modal-backdrop" onMouseDown={e => { if (e.target === e.currentTarget) onClose(); }}>
      <div className="modal" style={{ width: printFormat === 'A4' ? 760 : 360 }}>
        <div className="drawer-head">
          <div className="pill-group">
            <button className={printFormat === 'A4' ? 'selected' : ''} onClick={() => setPrintFormat('A4')}>
              A4 Tax Invoice
            </button>
            <button className={printFormat === 'Thermal80' ? 'selected' : ''} onClick={() => setPrintFormat('Thermal80')}>
              80 mm Thermal Receipt
            </button>
          </div>
          <button className="close-button" onClick={onClose}>×</button>
        </div>

        {/* PRINT CONTENT */}
        <div className="print-sheet" style={{ background: '#fff', color: '#111', padding: printFormat === 'A4' ? '24px 28px' : '14px 10px', fontSize: printFormat === 'A4' ? 11 : 10, lineHeight: 1.4, border: '1px solid #dedbd3' }}>
          {/* Pharmacy Header */}
          <div style={{ textAlign: 'center', borderBottom: '1px dashed #666', paddingBottom: 10, marginBottom: 12 }}>
            <h2 style={{ font: '800 16px Manrope', margin: '0 0 3px', textTransform: 'uppercase' }}>{store.legalName}</h2>
            <div>{store.address}</div>
            <div>Phone: {store.phone}</div>
            <div style={{ fontWeight: 700, marginTop: 4 }}>
              GSTIN: <span className="mono">{store.gstin}</span> · DL No: {store.drugLicence}
            </div>
            <h3 style={{ font: '800 13px Manrope', margin: '8px 0 0', letterSpacing: 1 }}>TAX INVOICE</h3>
          </div>

          {/* Invoice Meta */}
          <div style={{ display: 'grid', gridTemplateColumns: printFormat === 'A4' ? '1fr 1fr' : '1fr', gap: 8, borderBottom: '1px dashed #666', paddingBottom: 8, marginBottom: 10 }}>
            <div>
              <div><b>Invoice No:</b> <span className="mono">{invoice.id}</span></div>
              <div><b>Date & Time:</b> {invoice.date}</div>
              <div><b>Place of Supply:</b> {invoice.customerState}</div>
              <div><b>Reverse Charge:</b> No</div>
            </div>
            <div>
              <div><b>Billed To:</b> {invoice.customerName}</div>
              <div><b>Contact:</b> {invoice.customerPhone}</div>
              {invoice.customerGstin && <div><b>Customer GSTIN:</b> <span className="mono">{invoice.customerGstin}</span></div>}
              {invoice.doctorName && <div><b>Doctor:</b> {invoice.doctorName}</div>}
              {invoice.prescriptionId && <div><b>Prescription Ref:</b> {invoice.prescriptionId}</div>}
            </div>
          </div>

          {/* Item Table */}
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: printFormat === 'A4' ? 10 : 9, marginBottom: 10 }}>
            <thead>
              <tr style={{ borderBottom: '1px solid #333', textAlign: 'left' }}>
                <th style={{ padding: '4px 2px' }}>Description</th>
                <th style={{ padding: '4px 2px' }}>HSN</th>
                <th style={{ padding: '4px 2px' }}>Batch</th>
                <th style={{ padding: '4px 2px' }}>Exp</th>
                <th style={{ padding: '4px 2px' }}>Qty</th>
                <th style={{ padding: '4px 2px' }}>MRP</th>
                <th style={{ padding: '4px 2px' }}>Rate</th>
                <th style={{ padding: '4px 2px' }}>GST%</th>
                <th style={{ padding: '4px 2px', textAlign: 'right' }}>Total</th>
              </tr>
            </thead>
            <tbody>
              {invoice.lines.map((l, idx) => (
                <tr key={idx} style={{ borderBottom: '1px dotted #ccc' }}>
                  <td style={{ padding: '4px 2px' }}><b>{l.productName}</b></td>
                  <td style={{ padding: '4px 2px' }}>{l.hsn}</td>
                  <td style={{ padding: '4px 2px' }}><span className="mono">{l.batch}</span></td>
                  <td style={{ padding: '4px 2px' }}>{l.expiry}</td>
                  <td style={{ padding: '4px 2px' }}>{l.qty}</td>
                  <td style={{ padding: '4px 2px' }}>{money(l.mrp)}</td>
                  <td style={{ padding: '4px 2px' }}>{money(l.rate)}</td>
                  <td style={{ padding: '4px 2px' }}>{l.gstRate}%</td>
                  <td style={{ padding: '4px 2px', textAlign: 'right' }}><b>{money(l.lineTotal)}</b></td>
                </tr>
              ))}
            </tbody>
          </table>

          {/* Totals Breakdown */}
          <div style={{ borderTop: '1px solid #333', paddingTop: 6, marginBottom: 10 }}>
            <div style={{ display: 'flex', justifyContent: 'space-between' }}>
              <span>Taxable Value:</span>
              <span>{money(invoice.taxableTotal)}</span>
            </div>
            {invoice.customerState === store.state ? (
              <>
                <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                  <span>CGST:</span>
                  <span>{money(invoice.cgstTotal)}</span>
                </div>
                <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                  <span>SGST:</span>
                  <span>{money(invoice.sgstTotal)}</span>
                </div>
              </>
            ) : (
              <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                <span>IGST (Inter-state):</span>
                <span>{money(invoice.igstTotal)}</span>
              </div>
            )}
            {invoice.roundOff !== 0 && (
              <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                <span>Round Off:</span>
                <span>{money(invoice.roundOff)}</span>
              </div>
            )}
            <div style={{ display: 'flex', justifyContent: 'space-between', borderTop: '1px solid #000', paddingTop: 4, marginTop: 4, font: '800 13px Manrope' }}>
              <span>TOTAL PAYABLE:</span>
              <span>{money(invoice.grandTotal)}</span>
            </div>
            <div style={{ fontSize: 10, fontStyle: 'italic', marginTop: 4 }}>
              {numberToWordsIndian(invoice.grandTotal)}
            </div>
          </div>

          {/* Declaration and Signatures */}
          <div style={{ borderTop: '1px dashed #666', paddingTop: 8, fontSize: 9 }}>
            <div><b>Declaration:</b> We declare that this invoice shows the actual price of the goods described and that all particulars are true and correct. Medicines once sold cannot be returned without original cash memo.</div>
            <div style={{ marginTop: 24, display: 'flex', justifyContent: 'space-between' }}>
              <div>E&OE</div>
              <div style={{ textAlign: 'right' }}>
                <div>For {store.legalName}</div>
                <div style={{ marginTop: 28 }}>__________________________</div>
                <div>Authorised Pharmacist Signatory</div>
              </div>
            </div>
          </div>
        </div>

        <div className="modal-actions">
          <button type="button" className="btn" onClick={onClose}>Close</button>
          <button type="button" className="btn primary" onClick={() => window.print()}>
            <Printer size={14} /> Print Now
          </button>
        </div>
      </div>
    </div>
  );
}

// -------------------------------------------------------------
// 8. CANCEL INVOICE MODAL
// -------------------------------------------------------------
function CancelInvoiceModal({
  invoice,
  onClose,
  onConfirm
}: {
  invoice: Invoice;
  onClose: () => void;
  onConfirm: (reason: string) => void;
}) {
  const [reason, setReason] = useState('Customer returned before leaving counter');

  return (
    <div className="modal-backdrop" onMouseDown={e => { if (e.target === e.currentTarget) onClose(); }}>
      <div className="modal">
        <div className="drawer-head">
          <h2>Cancel Invoice {invoice.id}</h2>
          <button className="close-button" onClick={onClose}>×</button>
        </div>
        <form onSubmit={e => {
          e.preventDefault();
          onConfirm(reason);
        }}>
          <div className="notice warn" style={{ marginBottom: 14 }}>
            <b>Stock Reversal Warning:</b> Cancelling this invoice will immediately reverse all sold quantities back into their original store batches. The invoice number will be retained as "Cancelled" in the GST audit sequence.
          </div>
          <label className="field">
            Mandatory Cancellation Reason
            <select value={reason} onChange={e => setReason(e.target.value)}>
              <option value="Customer changed mind before payment">Customer changed mind before payment</option>
              <option value="Wrong medicine or batch entered by staff">Wrong medicine or batch entered by staff</option>
              <option value="Doctor altered prescription at counter">Doctor altered prescription at counter</option>
              <option value="Duplicate bill generated in error">Duplicate bill generated in error</option>
              <option value="Other customer cancellation">Other customer cancellation</option>
            </select>
          </label>
          <div className="modal-actions">
            <button type="button" className="btn" onClick={onClose}>Back</button>
            <button type="submit" className="btn danger">Confirm Cancellation & Reverse Stock</button>
          </div>
        </form>
      </div>
    </div>
  );
}

// -------------------------------------------------------------
// 9. CREDIT NOTE (SALES RETURN) MODAL
// -------------------------------------------------------------
function CreditNoteModal({
  invoice,
  stores,
  onClose,
  onConfirm
}: {
  invoice: Invoice;
  stores: StoreInfo[];
  onClose: () => void;
  onConfirm: (data: { returnLines: InvoiceLine[]; taxableTotal: number; cgstTotal: number; sgstTotal: number; igstTotal: number; grandTotal: number; reason: string }) => void;
}) {
  const [returnQtys, setReturnQtys] = useState<Record<number, number>>(() => {
    const init: Record<number, number> = {};
    invoice.lines.forEach((l, i) => { init[i] = l.qty; });
    return init;
  });
  const [reason, setReason] = useState('Customer returned unneeded sealed strip');

  const returnLines: InvoiceLine[] = [];
  invoice.lines.forEach((l, idx) => {
    const q = returnQtys[idx] || 0;
    if (q > 0) {
      const taxable = Math.round((l.taxableValue * (q / l.qty)) * 100) / 100;
      const cgst = Math.round((l.cgst * (q / l.qty)) * 100) / 100;
      const sgst = Math.round((l.sgst * (q / l.qty)) * 100) / 100;
      const igst = Math.round((l.igst * (q / l.qty)) * 100) / 100;
      returnLines.push({
        ...l,
        qty: q,
        taxableValue: taxable,
        cgst,
        sgst,
        igst,
        lineTotal: Math.round((taxable + cgst + sgst + igst) * 100) / 100
      });
    }
  });

  const taxableTotal = Math.round(returnLines.reduce((n, l) => n + l.taxableValue, 0) * 100) / 100;
  const cgstTotal = Math.round(returnLines.reduce((n, l) => n + l.cgst, 0) * 100) / 100;
  const sgstTotal = Math.round(returnLines.reduce((n, l) => n + l.sgst, 0) * 100) / 100;
  const igstTotal = Math.round(returnLines.reduce((n, l) => n + l.igst, 0) * 100) / 100;
  const grandTotal = Math.round(taxableTotal + cgstTotal + sgstTotal + igstTotal);

  return (
    <div className="modal-backdrop" onMouseDown={e => { if (e.target === e.currentTarget) onClose(); }}>
      <div className="modal" style={{ width: 620 }}>
        <div className="drawer-head">
          <h2>Create Credit Note (Sales Return)</h2>
          <button className="close-button" onClick={onClose}>×</button>
        </div>
        <form onSubmit={e => {
          e.preventDefault();
          if (returnLines.length === 0) return;
          onConfirm({ returnLines, taxableTotal, cgstTotal, sgstTotal, igstTotal, grandTotal, reason });
        }}>
          <div className="notice info" style={{ marginBottom: 14 }}>
            <b>Referencing Invoice:</b> {invoice.id} dated {invoice.date} · Customer: {invoice.customerName}
          </div>

          <div className="table-wrap">
            <table className="data-table">
              <thead>
                <tr>
                  <th>Product</th>
                  <th>Batch</th>
                  <th>Original Qty</th>
                  <th>Return Qty</th>
                  <th>Refund Total</th>
                </tr>
              </thead>
              <tbody>
                {invoice.lines.map((l, idx) => (
                  <tr key={idx}>
                    <td className="product-cell">{l.productName}</td>
                    <td><span className="mono">{l.batch}</span></td>
                    <td>{l.qty}</td>
                    <td>
                      <input
                        type="number"
                        min="0"
                        max={l.qty}
                        value={returnQtys[idx] ?? l.qty}
                        onChange={e => {
                          const val = Math.min(l.qty, Math.max(0, Number(e.target.value) || 0));
                          setReturnQtys(old => ({ ...old, [idx]: val }));
                        }}
                        style={{ width: 60 }}
                      />
                    </td>
                    <td><b>{money(Math.round(l.lineTotal * ((returnQtys[idx] || 0) / l.qty)))}</b></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <div style={{ marginTop: 14 }}>
            <label className="field">
              Reason for Sales Return
              <input value={reason} onChange={e => setReason(e.target.value)} required />
            </label>
          </div>

          <div style={{ textAlign: 'right', marginTop: 14 }}>
            <div className="stat-label">TOTAL CREDIT NOTE AMOUNT</div>
            <div style={{ font: '800 20px Manrope', color: '#256c5d' }}>{money(grandTotal)}</div>
          </div>

          <div className="modal-actions">
            <button type="button" className="btn" onClick={onClose}>Cancel</button>
            <button type="submit" className="btn primary" disabled={returnLines.length === 0}>
              Issue Credit Note & Restore Stock
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}

// -------------------------------------------------------------
// 10. NEW CUSTOMER MODAL
// -------------------------------------------------------------
function NewCustomerModal({
  onClose,
  onSave
}: {
  onClose: () => void;
  onSave: (cust: Customer) => void;
}) {
  const [name, setName] = useState('');
  const [phone, setPhone] = useState('+91 ');
  const [address, setAddress] = useState('');
  const [state, setState] = useState('Karnataka');
  const [gstin, setGstin] = useState('');
  const [type, setType] = useState<'B2C' | 'B2B'>('B2C');
  const [error, setError] = useState('');

  const submit = (e: FormEvent) => {
    e.preventDefault();
    if (type === 'B2B') {
      const gstinClean = gstin.trim().toUpperCase();
      if (!/^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}$/.test(gstinClean)) {
        setError('Illustrative validation: GSTIN must be a 15-character valid pattern (e.g. 29AAAAA0000A1Z5)');
        return;
      }
    }

    onSave({
      id: `C${Date.now().toString().slice(-4)}`,
      name: name.trim(),
      phone: phone.trim(),
      address: address.trim(),
      state,
      gstin: type === 'B2B' ? gstin.trim().toUpperCase() : undefined,
      type
    });
  };

  return (
    <div className="modal-backdrop" onMouseDown={e => { if (e.target === e.currentTarget) onClose(); }}>
      <div className="modal">
        <div className="drawer-head">
          <h2>Add New Customer</h2>
          <button className="close-button" onClick={onClose}>×</button>
        </div>
        <form onSubmit={submit}>
          {error && <div className="notice warn" style={{ marginBottom: 12 }}>{error}</div>}
          <div className="two-col">
            <label className="field">
              Customer / Business Name
              <input value={name} onChange={e => { setName(e.target.value); setError(''); }} required />
            </label>
            <label className="field">
              Phone
              <input value={phone} onChange={e => setPhone(e.target.value)} required />
            </label>
          </div>
          <div className="two-col" style={{ marginTop: 12 }}>
            <label className="field">
              Customer Type
              <select value={type} onChange={e => setType(e.target.value as any)}>
                <option value="B2C">B2C (Retail Walk-in / Patient)</option>
                <option value="B2B">B2B (Registered Clinic / Hospital / Pharmacy)</option>
              </select>
            </label>
            <label className="field">
              State (Determines Place of Supply)
              <select value={state} onChange={e => setState(e.target.value)}>
                <option value="Karnataka">Karnataka (Intra-state CGST+SGST)</option>
                <option value="Tamil Nadu">Tamil Nadu (Inter-state IGST)</option>
                <option value="Maharashtra">Maharashtra (Inter-state IGST)</option>
                <option value="Kerala">Kerala (Inter-state IGST)</option>
                <option value="Andhra Pradesh">Andhra Pradesh (Inter-state IGST)</option>
              </select>
            </label>
          </div>
          {type === 'B2B' && (
            <label className="field" style={{ marginTop: 12 }}>
              Customer GSTIN (15-character format)
              <input
                placeholder="e.g. 29AAACR3918M1Z8"
                value={gstin}
                onChange={e => { setGstin(e.target.value); setError(''); }}
                required
              />
              <span className="small-note">Illustrative check: 15-character state code + PAN + entity digit.</span>
            </label>
          )}
          <label className="field" style={{ marginTop: 12 }}>
            Address
            <input value={address} onChange={e => setAddress(e.target.value)} placeholder="Full street address..." />
          </label>
          <div className="modal-actions">
            <button type="button" className="btn" onClick={onClose}>Cancel</button>
            <button type="submit" className="btn primary">Save Customer</button>
          </div>
        </form>
      </div>
    </div>
  );
}

// -------------------------------------------------------------
// 11. GST REPORTS VIEW (SALES REGISTER, GSTR-1, HSN, TALLY STOCK)
// -------------------------------------------------------------
export function GstReportsView({
  storesMeta,
  invoices,
  creditNotes,
  movements
}: {
  storesMeta: StoreInfo[];
  invoices: Invoice[];
  creditNotes: CreditNote[];
  movements: StockMovement[];
}) {
  const [reportTab, setReportTab] = useState<'sales' | 'gstr1' | 'rate' | 'purchases' | 'stockSummary'>('sales');
  const [storeFilter, setStoreFilter] = useState('All');
  const [period, setPeriod] = useState('2025-06');

  const periodInvoices = invoices.filter(inv => {
    if (storeFilter !== 'All' && inv.storeId !== storeFilter) return false;
    return inv.date.startsWith(period) && inv.status === 'Final';
  });

  const periodCreditNotes = creditNotes.filter(cn => {
    if (storeFilter !== 'All' && cn.storeId !== storeFilter) return false;
    return cn.date.startsWith(period);
  });

  // Export current active report
  const exportActiveReport = () => {
    const wb = XLSX.utils.book_new();

    if (reportTab === 'sales') {
      const rows = periodInvoices.map(inv => ({
        'Invoice No': inv.id,
        Date: inv.date,
        Customer: inv.customerName,
        GSTIN: inv.customerGstin || 'URP',
        'Place of Supply': inv.customerState,
        'Taxable Value': inv.taxableTotal,
        CGST: inv.cgstTotal,
        SGST: inv.sgstTotal,
        IGST: inv.igstTotal,
        'Total Tax': inv.cgstTotal + inv.sgstTotal + inv.igstTotal,
        'Grand Total': inv.grandTotal,
        Payment: inv.paymentMode
      }));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), 'Sales Register');
    } else if (reportTab === 'gstr1') {
      const b2bRows = periodInvoices.filter(i => i.isB2B).map(inv => ({
        'GSTIN of Recipient': inv.customerGstin,
        'Receiver Name': inv.customerName,
        'Invoice Number': inv.id,
        'Invoice Date': inv.date.slice(0, 10),
        'Invoice Value': inv.grandTotal,
        'Place of Supply': inv.customerState,
        'Taxable Value': inv.taxableTotal,
        'Total Tax': inv.cgstTotal + inv.sgstTotal + inv.igstTotal
      }));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(b2bRows), 'GSTR1 B2B');
    } else if (reportTab === 'stockSummary') {
      const rows = products.map(p => {
        const prodMovements = movements.filter(m => m.productId === p.id && (!storeFilter || storeFilter === 'All' || m.storeId === storeFilter));
        const opening = prodMovements.filter(m => m.type === 'Opening / Stock import').reduce((n, m) => n + m.qty, 0);
        const inwards = prodMovements.filter(m => m.type.includes('Purchase') || m.type === 'Transfer in').reduce((n, m) => n + m.qty, 0);
        const outwards = Math.abs(prodMovements.filter(m => m.type.includes('Sale') || m.type === 'Transfer out').reduce((n, m) => n + m.qty, 0));
        const closing = opening + inwards - outwards;
        const rate = productGstMap[p.id]?.mrp ? Math.round(productGstMap[p.id].mrp * 0.7) : 40;
        return {
          'Product Code': p.id,
          'Product Name': p.name,
          'Opening Qty': opening,
          'Inwards Qty': inwards,
          'Outwards Qty': outwards,
          'Closing Qty': closing,
          'Valuation Rate': rate,
          'Closing Value (₹)': closing * rate
        };
      });
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), 'Stock Summary');
    }

    XLSX.writeFile(wb, `gst-report-${reportTab}-${period}.xlsx`);
  };

  return (
    <>
      <div className="page-head">
        <div>
          <div className="eyebrow">STATUTORY GST SUMMARY & TALLY AUDIT</div>
          <h1 className="page-title">GST Reports</h1>
          <p className="help-line">What this screen does: Generate Sales registers, GSTR-1 style B2B/B2C summaries, rate-wise tax breakdowns, and Tally-style stock registers.</p>
        </div>
        <div className="page-head-actions">
          <button className="btn" onClick={exportActiveReport}>
            <ArrowDownToLine size={14} /> Export Active to Excel
          </button>
          <button className="btn" onClick={() => window.print()}>
            <Printer size={14} /> Print / PDF
          </button>
        </div>
      </div>

      <div className="notice info" style={{ marginBottom: 14 }}>
        <b>Notice:</b> Illustrative: to be verified by the client's CA before production. Prototype report format. Not a filing-ready return.
      </div>

      <div className="panel">
        <div className="toolbar" style={{ justifyContent: 'space-between' }}>
          <div className="pill-group">
            <button className={reportTab === 'sales' ? 'selected' : ''} onClick={() => setReportTab('sales')}>
              Sales Register
            </button>
            <button className={reportTab === 'gstr1' ? 'selected' : ''} onClick={() => setReportTab('gstr1')}>
              GSTR-1 Summary
            </button>
            <button className={reportTab === 'rate' ? 'selected' : ''} onClick={() => setReportTab('rate')}>
              Rate-wise Tax
            </button>
            <button className={reportTab === 'purchases' ? 'selected' : ''} onClick={() => setReportTab('purchases')}>
              Purchase Register
            </button>
            <button className={reportTab === 'stockSummary' ? 'selected' : ''} onClick={() => setReportTab('stockSummary')}>
              Stock Summary (Tally)
            </button>
          </div>

          <div style={{ display: 'flex', gap: 10 }}>
            <label className="small-note" style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
              Period:
              <select className="control" value={period} onChange={e => setPeriod(e.target.value)}>
                <option value="2025-06">June 2025</option>
                <option value="2025-05">May 2025</option>
                <option value="2025-04">April 2025</option>
              </select>
            </label>

            <label className="small-note" style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
              Store:
              <select className="control" value={storeFilter} onChange={e => setStoreFilter(e.target.value)}>
                <option value="All">All Stores</option>
                {storesMeta.map(s => <option key={s.id} value={s.id}>{s.name}</option>)}
              </select>
            </label>
          </div>
        </div>

        {/* 1. SALES REGISTER */}
        {reportTab === 'sales' && (
          <div className="table-wrap">
            <table className="data-table">
              <thead>
                <tr>
                  <th>Invoice No</th>
                  <th>Date</th>
                  <th>Customer</th>
                  <th>GSTIN</th>
                  <th>Place of Supply</th>
                  <th>Taxable Value</th>
                  <th>CGST</th>
                  <th>SGST</th>
                  <th>IGST</th>
                  <th>Total Tax</th>
                  <th>Invoice Total</th>
                </tr>
              </thead>
              <tbody>
                {periodInvoices.map(inv => (
                  <tr key={inv.id}>
                    <td><span className="mono">{inv.id}</span></td>
                    <td>{inv.date.slice(0, 10)}</td>
                    <td className="product-cell">{inv.customerName}</td>
                    <td><span className="mono">{inv.customerGstin || 'URP'}</span></td>
                    <td>{inv.customerState}</td>
                    <td>{money(inv.taxableTotal)}</td>
                    <td>{money(inv.cgstTotal)}</td>
                    <td>{money(inv.sgstTotal)}</td>
                    <td>{money(inv.igstTotal)}</td>
                    <td>{money(inv.cgstTotal + inv.sgstTotal + inv.igstTotal)}</td>
                    <td><b>{money(inv.grandTotal)}</b></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {/* 2. GSTR-1 SUMMARY */}
        {reportTab === 'gstr1' && (
          <div style={{ padding: 18 }}>
            <h3 style={{ font: '800 15px Manrope', margin: '0 0 10px' }}>Table 4A: B2B Invoices</h3>
            <div className="table-wrap" style={{ marginBottom: 20 }}>
              <table className="data-table">
                <thead>
                  <tr>
                    <th>GSTIN of Recipient</th>
                    <th>Receiver Name</th>
                    <th>Invoice No</th>
                    <th>Invoice Date</th>
                    <th>Invoice Value</th>
                    <th>Place of Supply</th>
                    <th>Taxable Value</th>
                    <th>Taxes</th>
                  </tr>
                </thead>
                <tbody>
                  {periodInvoices.filter(i => i.isB2B).map(inv => (
                    <tr key={inv.id}>
                      <td><span className="mono">{inv.customerGstin}</span></td>
                      <td className="product-cell">{inv.customerName}</td>
                      <td><span className="mono">{inv.id}</span></td>
                      <td>{inv.date.slice(0, 10)}</td>
                      <td><b>{money(inv.grandTotal)}</b></td>
                      <td>{inv.customerState}</td>
                      <td>{money(inv.taxableTotal)}</td>
                      <td>{money(inv.cgstTotal + inv.sgstTotal + inv.igstTotal)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <h3 style={{ font: '800 15px Manrope', margin: '0 0 10px' }}>Table 7: B2C Retail Summary</h3>
            <div className="table-wrap" style={{ marginBottom: 20 }}>
              <table className="data-table">
                <thead>
                  <tr>
                    <th>Type</th>
                    <th>Place of Supply</th>
                    <th>Rate</th>
                    <th>Taxable Value</th>
                    <th>CGST</th>
                    <th>SGST</th>
                    <th>Total Tax</th>
                  </tr>
                </thead>
                <tbody>
                  {['5%', '12%', '18%'].map(r => (
                    <tr key={r}>
                      <td>B2C Small Retail</td>
                      <td>Karnataka (29)</td>
                      <td>{r}</td>
                      <td>{money(r === '12%' ? 18450 : r === '18%' ? 5200 : 850)}</td>
                      <td>{money(r === '12%' ? 1107 : r === '18%' ? 468 : 21.25)}</td>
                      <td>{money(r === '12%' ? 1107 : r === '18%' ? 468 : 21.25)}</td>
                      <td><b>{money(r === '12%' ? 2214 : r === '18%' ? 936 : 42.5)}</b></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <h3 style={{ font: '800 15px Manrope', margin: '0 0 10px' }}>Table 9B: Credit / Debit Notes</h3>
            <div className="table-wrap" style={{ marginBottom: 20 }}>
              <table className="data-table">
                <thead>
                  <tr>
                    <th>Credit Note No</th>
                    <th>Date</th>
                    <th>Original Invoice</th>
                    <th>Customer</th>
                    <th>Taxable</th>
                    <th>Taxes</th>
                    <th>Credit Note Value</th>
                  </tr>
                </thead>
                <tbody>
                  {periodCreditNotes.map(cn => (
                    <tr key={cn.id}>
                      <td><span className="mono">{cn.id}</span></td>
                      <td>{cn.date.slice(0, 10)}</td>
                      <td><span className="mono">{cn.originalInvoiceId}</span></td>
                      <td>{cn.customerName}</td>
                      <td>{money(cn.taxableTotal)}</td>
                      <td>{money(cn.cgstTotal + cn.sgstTotal + cn.igstTotal)}</td>
                      <td><b>{money(cn.grandTotal)}</b></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <h3 style={{ font: '800 15px Manrope', margin: '0 0 10px' }}>Table 12: HSN-wise Summary</h3>
            <div className="table-wrap">
              <table className="data-table">
                <thead>
                  <tr>
                    <th>HSN</th>
                    <th>Description</th>
                    <th>UQC</th>
                    <th>Total Qty</th>
                    <th>Total Value</th>
                    <th>Taxable Value</th>
                    <th>Integrated Tax</th>
                    <th>Central Tax</th>
                    <th>State Tax</th>
                  </tr>
                </thead>
                <tbody>
                  <tr>
                    <td><span className="mono">3004</span></td>
                    <td>Medicaments consisting of mixed or unmixed products for therapeutic uses</td>
                    <td>STR</td>
                    <td>142</td>
                    <td>{money(28500)}</td>
                    <td>{money(25446.43)}</td>
                    <td>{money(420)}</td>
                    <td>{money(1316.78)}</td>
                    <td>{money(1316.78)}</td>
                  </tr>
                  <tr>
                    <td><span className="mono">2106</span></td>
                    <td>Food preparations / Nutraceutical supplements</td>
                    <td>BOX</td>
                    <td>38</td>
                    <td>{money(7200)}</td>
                    <td>{money(6101.69)}</td>
                    <td>{money(0)}</td>
                    <td>{money(549.15)}</td>
                    <td>{money(549.15)}</td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* 3. RATE-WISE SUMMARY */}
        {reportTab === 'rate' && (
          <div className="table-wrap">
            <table className="data-table">
              <thead>
                <tr>
                  <th>GST Rate</th>
                  <th>Taxable Value</th>
                  <th>CGST</th>
                  <th>SGST</th>
                  <th>IGST</th>
                  <th>Total Tax</th>
                  <th>Gross Amount</th>
                </tr>
              </thead>
              <tbody>
                {[
                  { rate: '0% Exempt', taxable: 1200, cgst: 0, sgst: 0, igst: 0 },
                  { rate: '5% Essential', taxable: 2400, cgst: 60, sgst: 60, igst: 0 },
                  { rate: '12% Formulations', taxable: 38240, cgst: 2194.4, sgst: 2194.4, igst: 420 },
                  { rate: '18% Nutraceuticals', taxable: 8900, cgst: 801, sgst: 801, igst: 0 },
                ].map(row => {
                  const totalTax = row.cgst + row.sgst + row.igst;
                  return (
                    <tr key={row.rate}>
                      <td><b>{row.rate}</b></td>
                      <td>{money(row.taxable)}</td>
                      <td>{money(row.cgst)}</td>
                      <td>{money(row.sgst)}</td>
                      <td>{money(row.igst)}</td>
                      <td><b>{money(totalTax)}</b></td>
                      <td><b>{money(row.taxable + totalTax)}</b></td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}

        {/* 4. PURCHASE REGISTER */}
        {reportTab === 'purchases' && (
          <div className="table-wrap">
            <table className="data-table">
              <thead>
                <tr>
                  <th>Date</th>
                  <th>Supplier Invoice</th>
                  <th>Supplier</th>
                  <th>Store</th>
                  <th>Product</th>
                  <th>Batch</th>
                  <th>Qty</th>
                  <th>Rate</th>
                  <th>Taxable Value</th>
                  <th>GST %</th>
                  <th>Input Tax</th>
                </tr>
              </thead>
              <tbody>
                {movements.filter(m => m.type.includes('Purchase')).slice(0, 15).map(m => {
                  const p = products.find(x => x.id === m.productId);
                  const meta = productGstMap[m.productId] || { mrp: 60, gstRate: 12 };
                  const rate = Math.round(meta.mrp * 0.65);
                  const taxable = rate * m.qty;
                  const tax = Math.round(taxable * (meta.gstRate / 100));
                  return (
                    <tr key={m.id}>
                      <td>{m.dateTime.slice(0, 10)}</td>
                      <td><span className="mono">{m.ref}</span></td>
                      <td className="product-cell">{suppliers[0].name}</td>
                      <td>{m.storeId}</td>
                      <td>{p?.name}</td>
                      <td><span className="mono">{m.batch}</span></td>
                      <td>{m.qty}</td>
                      <td>{money(rate)}</td>
                      <td>{money(taxable)}</td>
                      <td>{meta.gstRate}%</td>
                      <td><b>{money(tax)}</b></td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}

        {/* 5. TALLY-STYLE STOCK SUMMARY REGISTER */}
        {reportTab === 'stockSummary' && (
          <div className="table-wrap">
            <table className="data-table">
              <thead>
                <tr>
                  <th>Product</th>
                  <th>Opening Qty</th>
                  <th>Inwards Qty</th>
                  <th>Outwards Qty</th>
                  <th>Closing Qty</th>
                  <th>Valuation Rate</th>
                  <th>Closing Value (₹)</th>
                </tr>
              </thead>
              <tbody>
                {products.slice(0, 25).map(p => {
                  const prodMovs = movements.filter(m => m.productId === p.id && (storeFilter === 'All' || m.storeId === storeFilter));
                  const op = prodMovs.filter(m => m.type.includes('Opening')).reduce((n, m) => n + m.qty, 0);
                  const inw = prodMovs.filter(m => m.type.includes('Purchase') || m.type === 'Transfer in').reduce((n, m) => n + m.qty, 0);
                  const outw = Math.abs(prodMovs.filter(m => m.type.includes('Sale') || m.type === 'Transfer out').reduce((n, m) => n + m.qty, 0));
                  const cl = op + inw - outw;
                  const rate = productGstMap[p.id]?.mrp ? Math.round(productGstMap[p.id].mrp * 0.7) : 40;
                  return (
                    <tr key={p.id}>
                      <td className="product-cell">{p.name}<div className="row-secondary">{p.pack}</div></td>
                      <td>{op}</td>
                      <td>{inw}</td>
                      <td>{outw}</td>
                      <td><b>{cl}</b> units</td>
                      <td>{money(rate)}</td>
                      <td><b>{money(cl * rate)}</b></td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </>
  );
}
