/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import {
  ResponsiveContainer,
  AreaChart,
  Area,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  Cell
} from 'recharts';
import {
  Settings,
  MenuItem,
  Category,
  PaymentMethod,
  TableItem,
  StaffUser,
  CartItem,
  Order
} from './types';
import {
  DEFAULT_SETTINGS,
  DEFAULT_PAYMENT_METHODS,
  DEFAULT_CATEGORIES,
  DEFAULT_MENU,
  DEFAULT_USERS,
  DEFAULT_TABLES,
  announceOrderCall,
  getFallbackImage,
  generateDynamicQris,
  copyToClipboardText,
  formatCurrencyInput,
  parseIndonesianNumber,
  generateQrDataUrl
} from './utils';
import {
  initializeFirestoreData,
  subscribeSettings,
  subscribeCategories,
  subscribeMenu,
  subscribePaymentMethods,
  subscribeTables,
  subscribeStaff,
  subscribeOrders,
  createOrder,
  updateOrderStatus,
  deleteOrder,
  saveSettings as persistSettings,
  saveMenuItem,
  deleteMenuItem,
  batchSaveMenuItems,
  saveCategory,
  deleteCategory,
  saveTable,
  deleteTable,
  saveStaffUser,
  deleteStaffUser,
  savePaymentMethod,
  deletePaymentMethod
} from './dbService';
import { testConnection } from './firebase';

export default function App() {
  const [isOffline, setIsOffline] = useState(!navigator.onLine);
  const [firebaseConnected, setFirebaseConnected] = useState<boolean>(false);
  const [isSyncing, setIsSyncing] = useState<boolean>(false);

  // Active view: 'landing' | 'customer' | 'admin_login' | 'admin'
  const [view, setView] = useState<'landing' | 'customer' | 'admin_login' | 'admin'>('landing');
  const [adminTab, setAdminTab] = useState<'pos' | 'orders' | 'analytics' | 'menu' | 'tables' | 'staff' | 'settings'>('pos');
  const [currentUser, setCurrentUser] = useState<StaffUser | null>(null);
  const [analyticsDays, setAnalyticsDays] = useState<number>(7);
  const [orderFilterStatus, setOrderFilterStatus] = useState<string>('all');

  // Firestore Realtime Data
  const [settings, setSettings] = useState<Settings>(DEFAULT_SETTINGS);
  const [categories, setCategories] = useState<Category[]>(DEFAULT_CATEGORIES);
  const [menuItems, setMenuItems] = useState<MenuItem[]>(DEFAULT_MENU);
  const [paymentMethods, setPaymentMethods] = useState<PaymentMethod[]>(DEFAULT_PAYMENT_METHODS);
  const [tables, setTables] = useState<TableItem[]>(DEFAULT_TABLES);
  const [staffUsers, setStaffUsers] = useState<StaffUser[]>(DEFAULT_USERS);
  const [orders, setOrders] = useState<Order[]>([]);

  // Customer State
  const [customerTable, setCustomerTable] = useState('');
  const [customerOrderType, setCustomerOrderType] = useState<'dine-in' | 'takeaway'>('dine-in');
  const [customerName, setCustomerName] = useState('');
  const [customerCart, setCustomerCart] = useState<CartItem[]>([]);
  const [customerActiveCat, setCustomerActiveCat] = useState('Semua');
  const [showCart, setShowCart] = useState(false);
  const [checkoutData, setCheckoutData] = useState({ paymentMethod: '', notes: '' });
  const [activeOrder, setActiveOrder] = useState<Order | null>(null);
  const [showOrderTrackerModal, setShowOrderTrackerModal] = useState<boolean>(false);
  const [showTableSelectModal, setShowTableSelectModal] = useState<boolean>(false);
  const [isSubmittingOrder, setIsSubmittingOrder] = useState<boolean>(false);
  const [isSubmittingPos, setIsSubmittingPos] = useState<boolean>(false);
  const lastOrderSubmissionRef = useRef<{ time: number; cartSignature: string }>({ time: 0, cartSignature: '' });

  // Staff Login Form
  const [loginForm, setLoginForm] = useState({ username: '', pin: '' });

  // POS State
  const [posCart, setPosCart] = useState<CartItem[]>([]);
  const [posActiveCat, setPosActiveCat] = useState('Semua');
  const [posTableNo, setPosTableNo] = useState('');
  const [posCustomerName, setPosCustomerName] = useState('');
  const [posPaymentMethod, setPosPaymentMethod] = useState('');
  const [posOrderType, setPosOrderType] = useState<'dine-in' | 'takeaway'>('dine-in');
  const [posCashGiven, setPosCashGiven] = useState('');

  // Receipt Modal
  const [receiptOrder, setReceiptOrder] = useState<Order | null>(null);

  // Modals
  const [showMenuModal, setShowMenuModal] = useState(false);
  const [editingMenu, setEditingMenu] = useState<MenuItem | null>(null);
  const [menuForm, setMenuForm] = useState({ name: '', price: '', category: 'Makanan Utama', image: '', description: '' });
  const [adminMenuSearch, setAdminMenuSearch] = useState('');
  const [adminMenuCatFilter, setAdminMenuCatFilter] = useState('Semua');

  const [showStaffModal, setShowStaffModal] = useState(false);
  const [editingStaff, setEditingStaff] = useState<StaffUser | null>(null);
  const [staffForm, setStaffForm] = useState({ username: '', pin: '', name: '', role: 'cashier' as 'cashier' | 'admin' });
  const [visibleStaffPins, setVisibleStaffPins] = useState<{ [key: string]: boolean }>({});

  const [showPaymentModal, setShowPaymentModal] = useState(false);
  const [editingPayment, setEditingPayment] = useState<PaymentMethod | null>(null);
  const [paymentForm, setPaymentForm] = useState({ name: '', type: 'cash' as 'cash' | 'qris' | 'transfer', isActive: true, instructions: '' });

  const [showCategoryModal, setShowCategoryModal] = useState(false);
  const [categoryForm, setCategoryForm] = useState({ name: '' });

  const [showCheckStatusModal, setShowCheckStatusModal] = useState(false);
  const [checkOrderIdInput, setCheckOrderIdInput] = useState('');

  // Custom Alert Modal
  const [modalConfig, setModalConfig] = useState<{
    isOpen: boolean;
    title: string;
    message: string;
    onConfirm?: (() => void) | null;
  }>({ isOpen: false, title: '', message: '', onConfirm: null });

  const showAlert = (title: string, message: string, onConfirm: (() => void) | null = null) => {
    setModalConfig({ isOpen: true, title, message, onConfirm });
  };

  const closeAlert = () => {
    setModalConfig({ isOpen: false, title: '', message: '', onConfirm: null });
  };

  // Initialize and subscribe to Firestore
  useEffect(() => {
    const handleOffline = () => setIsOffline(true);
    const handleOnline = () => setIsOffline(false);
    window.addEventListener('offline', handleOffline);
    window.addEventListener('online', handleOnline);

    // URL parameter detection for table QR scan: ?table=X
    const urlParams = new URLSearchParams(window.location.search);
    const tableParam = urlParams.get('table');
    if (tableParam) {
      setCustomerTable(tableParam);
      setView('customer');
    }

    // Connect and seed Firestore
    testConnection().then((connected) => {
      setFirebaseConnected(connected);
      if (connected) {
        setIsSyncing(true);
        initializeFirestoreData().finally(() => setIsSyncing(false));
      }
    });

    // Subscriptions to live Firestore
    const unsubSettings = subscribeSettings((data) => setSettings(data));
    const unsubCategories = subscribeCategories((data) => setCategories(data));
    const unsubMenu = subscribeMenu((data) => setMenuItems(data));
    const unsubPayments = subscribePaymentMethods((data) => setPaymentMethods(data));
    const unsubTables = subscribeTables((data) => setTables(data));
    const unsubStaff = subscribeStaff((data) => setStaffUsers(data));
    const unsubOrders = subscribeOrders((data) => setOrders(data));

    return () => {
      window.removeEventListener('offline', handleOffline);
      window.removeEventListener('online', handleOnline);
      unsubSettings();
      unsubCategories();
      unsubMenu();
      unsubPayments();
      unsubTables();
      unsubStaff();
      unsubOrders();
    };
  }, []);

  // Sync background styling if configured
  useEffect(() => {
    if (settings.appBackground) {
      const bg = settings.appBackground.trim();
      if (bg.startsWith('http') || bg.startsWith('data:image')) {
        document.body.style.backgroundImage = `url('${bg}')`;
        document.body.style.backgroundSize = 'cover';
        document.body.style.backgroundPosition = 'center';
        document.body.style.backgroundAttachment = 'fixed';
      } else {
        document.body.style.backgroundColor = bg || '#f4f4f5';
        document.body.style.backgroundImage = 'none';
      }
    }
  }, [settings.appBackground]);

  // Price calculations
  const calculateSubtotal = (cart: CartItem[]) =>
    cart.reduce((sum, item) => sum + item.price * item.quantity, 0);
  const calculateTax = (subtotal: number) => subtotal * (settings.taxRate / 100);
  const calculateService = (subtotal: number) => subtotal * (settings.serviceCharge / 100);
  const calculateTotal = (cart: CartItem[]) => {
    const sub = calculateSubtotal(cart);
    return sub + calculateTax(sub) + calculateService(sub);
  };

  const generateOrderId = () => {
    const prefix = (settings.orderPrefix || 'ORD').trim().toUpperCase();
    return `${prefix}-${Math.floor(1000 + Math.random() * 9000)}`;
  };

  // Staff Login
  const handleLoginSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const found = staffUsers.find(
      (u) => u.username === loginForm.username && u.pin === loginForm.pin
    );
    if (found) {
      setCurrentUser(found);
      setView('admin');
      setAdminTab('pos');
      setLoginForm({ username: '', pin: '' });
    } else {
      showAlert('Gagal Masuk', 'Username atau PIN 6-digit salah. Silakan coba lagi.');
    }
  };

  const handleLogout = () => {
    setCurrentUser(null);
    setView('landing');
  };

  // Customer Order Operations
  const handleCustomerLogin = (e: React.FormEvent) => {
    e.preventDefault();
    if (customerTable.trim() || customerOrderType === 'takeaway') {
      setView('customer');
      setShowTableSelectModal(false);
    }
  };

  const startOrdering = (type: 'dine-in' | 'takeaway', tableNum = '') => {
    setCustomerOrderType(type);
    if (type === 'takeaway') {
      setCustomerTable('Takeaway');
      setView('customer');
      setShowTableSelectModal(false);
    } else {
      const selected = tableNum.trim() || customerTable.trim();
      if (selected) {
        setCustomerTable(selected);
        setView('customer');
        setShowTableSelectModal(false);
      } else {
        setShowTableSelectModal(true);
      }
    }
  };

  const addToCustomerCart = (item: MenuItem) => {
    setCustomerCart((prev) => {
      const exists = prev.find((i) => i.id === item.id);
      if (exists) {
        return prev.map((i) =>
          i.id === item.id ? { ...i, quantity: i.quantity + 1 } : i
        );
      }
      return [...prev, { ...item, quantity: 1 }];
    });
  };

  const updateCustomerQuantity = (id: string | number, delta: number) => {
    setCustomerCart((prev) =>
      prev
        .map((item) => {
          if (item.id === id) {
            const newQ = item.quantity + delta;
            return newQ > 0 ? { ...item, quantity: newQ } : item;
          }
          return item;
        })
        .filter((item) => item.quantity > 0)
    );
  };

  const handleCustomerCheckout = async () => {
    if (isSubmittingOrder) return;

    if (!customerName.trim()) {
      showAlert('Nama Diperlukan', 'Mohon masukkan nama pemesan sebelum mengirim pesanan.');
      return;
    }
    if (!checkoutData.paymentMethod || customerCart.length === 0) return;

    // Cegah Pesanan Ganda (Double submit protection)
    const cartSignature = customerCart
      .map((c) => `${c.id}:${c.quantity}`)
      .sort()
      .join('|') + `|${customerTable}|${customerName.trim()}`;
    const now = Date.now();
    if (
      lastOrderSubmissionRef.current.cartSignature === cartSignature &&
      now - lastOrderSubmissionRef.current.time < 6000
    ) {
      showAlert(
        'Pesanan Sudah Terkirim',
        'Pesanan yang sama baru saja terkirim ke dapur. Mohon tunggu proses pesanan Anda agar tidak terjadi duplikasi.'
      );
      return;
    }

    setIsSubmittingOrder(true);
    const sub = calculateSubtotal(customerCart);
    const newOrder: Order = {
      id: generateOrderId(),
      customerName: customerName.trim(),
      table: customerOrderType === 'takeaway' ? 'Takeaway' : customerTable,
      orderType: customerOrderType,
      items: [...customerCart],
      subtotal: sub,
      tax: calculateTax(sub),
      service: calculateService(sub),
      total: calculateTotal(customerCart),
      paymentMethod: checkoutData.paymentMethod,
      notes: checkoutData.notes,
      status: 'unverified',
      time: new Date().toISOString(),
      source: 'QR',
      cashGiven: 0,
      changeAmount: 0
    };

    try {
      await createOrder(newOrder);
      lastOrderSubmissionRef.current = { time: Date.now(), cartSignature };
      setActiveOrder(newOrder);
      setReceiptOrder(newOrder);
      setCustomerCart([]);
      setShowCart(false);
      // Customer remains on menu page; alert confirms submission
      showAlert(
        'Pesanan Berhasil Dikirim!',
        `Pesanan ${newOrder.id} telah masuk ke sistem restoran. Anda tetap berada di halaman menu untuk menambah pesanan jika diperlukan.`
      );
    } catch (err) {
      showAlert('Gagal Mengirim', 'Terjadi kendala saat mengirim pesanan ke server Firestore.');
    } finally {
      setIsSubmittingOrder(false);
    }
  };

  // POS Operations
  const addToPosCart = (item: MenuItem) => {
    setPosCart((prev) => {
      const exists = prev.find((i) => i.id === item.id);
      if (exists) {
        return prev.map((i) =>
          i.id === item.id ? { ...i, quantity: i.quantity + 1 } : i
        );
      }
      return [...prev, { ...item, quantity: 1 }];
    });
  };

  const updatePosQuantity = (id: string | number, delta: number) => {
    setPosCart((prev) =>
      prev
        .map((item) => {
          if (item.id === id) {
            const newQ = item.quantity + delta;
            return newQ > 0 ? { ...item, quantity: newQ } : item;
          }
          return item;
        })
        .filter((item) => item.quantity > 0)
    );
  };

  const handlePosCheckout = async () => {
    if (isSubmittingPos) return;
    if (!posPaymentMethod || posCart.length === 0) return;
    const totalAmt = calculateTotal(posCart);
    const selectedPm = paymentMethods.find((p) => p.id === posPaymentMethod);
    const isCash = selectedPm?.type === 'cash';

    let cashGivenNum = 0;
    let changeNum = 0;

    if (isCash) {
      cashGivenNum = parseIndonesianNumber(posCashGiven);
      if (cashGivenNum < totalAmt) {
        showAlert('Pembayaran Kurang', 'Jumlah uang tunai yang diterima kurang dari total tagihan.');
        return;
      }
      changeNum = cashGivenNum - totalAmt;
    }

    setIsSubmittingPos(true);
    const sub = calculateSubtotal(posCart);
    const newOrder: Order = {
      id: generateOrderId(),
      customerName: posCustomerName.trim() || 'Kasir / Walk-in',
      table: posTableNo || 'Takeaway',
      orderType: posOrderType,
      items: [...posCart],
      subtotal: sub,
      tax: calculateTax(sub),
      service: calculateService(sub),
      total: totalAmt,
      paymentMethod: posPaymentMethod,
      notes: '',
      status: 'completed',
      time: new Date().toISOString(),
      source: 'KASIR',
      cashGiven: isCash ? cashGivenNum : 0,
      changeAmount: isCash ? changeNum : 0
    };

    try {
      await createOrder(newOrder);
      setPosCart([]);
      setPosTableNo('');
      setPosCustomerName('');
      setPosPaymentMethod('');
      setPosCashGiven('');
      setReceiptOrder(newOrder);
    } catch (err) {
      showAlert('Gagal Menyimpan', 'Terjadi kendala saat menyimpan transaksi di Firestore.');
    } finally {
      setIsSubmittingPos(false);
    }
  };

  // Menu Admin Operations
  const handleSaveMenu = async (e: React.FormEvent) => {
    e.preventDefault();
    const item: MenuItem = {
      id: editingMenu ? editingMenu.id : String(Date.now()),
      name: menuForm.name,
      price: Number(menuForm.price),
      category: menuForm.category,
      image: menuForm.image || 'https://images.unsplash.com/photo-1546069901-ba9599a7e63c?w=500&auto=format&fit=crop&q=80',
      description: menuForm.description
    };

    try {
      await saveMenuItem(item);
      setShowMenuModal(false);
      showAlert('Berhasil', 'Menu berhasil disimpan langsung ke database Firebase!');
    } catch (err) {
      showAlert('Gagal', 'Gagal menyimpan menu ke Firestore.');
    }
  };

  const handleDeleteMenu = (id: string | number) => {
    showAlert('Hapus Menu', 'Apakah Anda yakin ingin menghapus menu ini dari database?', async () => {
      try {
        await deleteMenuItem(id);
      } catch (err) {
        showAlert('Gagal', 'Gagal menghapus menu dari database.');
      }
    });
  };

  const openMenuModal = (item: MenuItem | null = null) => {
    if (item) {
      setEditingMenu(item);
      setMenuForm({
        name: item.name,
        price: String(item.price),
        category: item.category,
        image: item.image,
        description: item.description
      });
    } else {
      setEditingMenu(null);
      setMenuForm({
        name: '',
        price: '',
        category: categories.filter((c) => c.id !== 'Semua')[0]?.id || 'Makanan Utama',
        image: '',
        description: ''
      });
    }
    setShowMenuModal(true);
  };

  // Order Operations: Cancel, Delete & XLSX Export (Batal & Hapus Khusus Admin)
  const handleCancelOrder = (orderId: string) => {
    if (currentUser?.role !== 'admin') {
      showAlert('Akses Terbatas', 'Hanya pengguna dengan peran Administrator yang memiliki wewenang untuk membatalkan pesanan.');
      return;
    }
    showAlert(
      'Batalkan Pesanan',
      `Apakah Anda yakin ingin membatalkan pesanan #${orderId}? Status pesanan akan diubah menjadi 'Dibatalkan'.`,
      async () => {
        try {
          await updateOrderStatus(orderId, 'cancelled');
        } catch (err) {
          showAlert('Gagal', 'Terjadi kesalahan saat membatalkan pesanan.');
        }
      }
    );
  };

  const handleDeleteOrder = (orderId: string) => {
    if (currentUser?.role !== 'admin') {
      showAlert('Akses Terbatas', 'Hanya pengguna dengan peran Administrator yang memiliki wewenang untuk menghapus transaksi secara permanen.');
      return;
    }
    showAlert(
      'Hapus Transaksi Permanen',
      `PERINGATAN: Transaksi #${orderId} akan dihapus secara permanen dari database Firebase Firestore. Data yang dihapus tidak dapat dipulihkan. Lanjutkan?`,
      async () => {
        try {
          await deleteOrder(orderId);
        } catch (err) {
          showAlert('Gagal', 'Terjadi kesalahan saat menghapus transaksi.');
        }
      }
    );
  };

  const exportOrdersToXLSX = () => {
    if (orders.length === 0) {
      showAlert('Informasi', 'Belum ada transaksi untuk diekspor ke file Excel.');
      return;
    }

    const dataToExport = orders.map((o) => {
      const itemsDetail = o.items.map((i) => `${i.quantity}x ${i.name} (@Rp ${i.price.toLocaleString('id-ID')})`).join('; ');
      const paymentName = paymentMethods.find((p) => p.id === o.paymentMethod)?.name || o.paymentMethod;
      const statusMap: Record<string, string> = {
        unverified: 'Menunggu Verifikasi',
        processing: 'Sedang Disiapkan / Dimasak',
        completed: 'Selesai',
        cancelled: 'Dibatalkan'
      };

      return {
        'No Pesanan': o.id,
        'Waktu Transaksi': new Date(o.time).toLocaleString('id-ID'),
        'Nama Pelanggan': o.customerName || 'Pelanggan',
        'Meja': o.table || '-',
        'Tipe Pesanan': o.orderType === 'takeaway' ? 'Bungkus (Takeaway)' : 'Makan di Tempat (Dine-in)',
        'Menu Dipesan': itemsDetail,
        'Total Item': o.items.reduce((sum, item) => sum + item.quantity, 0),
        'Subtotal (Rp)': o.subtotal,
        'Pajak (Rp)': o.tax,
        'Service (Rp)': o.service,
        'Total Tagihan (Rp)': o.total,
        'Metode Pembayaran': paymentName,
        'Status Pesanan': statusMap[o.status] || o.status,
        'Sumber Pesanan': o.source || 'POS / QR Pelanggan',
        'Catatan Pelanggan': o.notes || '-'
      };
    });

    const worksheet = XLSX.utils.json_to_sheet(dataToExport);
    worksheet['!cols'] = [
      { wch: 15 }, // No Pesanan
      { wch: 20 }, // Waktu Transaksi
      { wch: 20 }, // Nama Pelanggan
      { wch: 10 }, // Meja
      { wch: 24 }, // Tipe Pesanan
      { wch: 42 }, // Menu Dipesan
      { wch: 12 }, // Total Item
      { wch: 15 }, // Subtotal
      { wch: 12 }, // Pajak
      { wch: 12 }, // Service
      { wch: 18 }, // Total Tagihan
      { wch: 20 }, // Metode Pembayaran
      { wch: 26 }, // Status Pesanan
      { wch: 22 }, // Sumber Pesanan
      { wch: 28 }, // Catatan Pelanggan
    ];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Daftar Transaksi');
    XLSX.writeFile(workbook, `Laporan-Transaksi-${(settings.restoName || 'Resto').replace(/\s+/g, '-')}-${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  // Excel Export & Import Menu
  const exportMenuToXLSX = () => {
    const dataToExport = menuItems.map((m) => ({
      ID: m.id,
      Nama_Menu: m.name,
      Harga: m.price,
      Kategori: m.category,
      Deskripsi: m.description,
      URL_Gambar: m.image
    }));

    const worksheet = XLSX.utils.json_to_sheet(dataToExport);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Katalog Menu');
    XLSX.writeFile(workbook, `Katalog-Menu-RestoQR-${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  const importMenuFromXLSX = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = async (event) => {
      try {
        const data = new Uint8Array(event.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const json = XLSX.utils.sheet_to_json(worksheet);

        if (json && json.length > 0) {
          const imported: MenuItem[] = (json as Array<Record<string, unknown>>).map((row, idx) => ({
            id: String(row.ID || Date.now() + idx),
            name: String(row.Nama_Menu || row.name || 'Menu Baru'),
            price: Number(row.Harga || row.price) || 0,
            category: String(row.Kategori || row.category || 'Makanan Utama'),
            description: String(row.Deskripsi || row.description || ''),
            image: String(row.URL_Gambar || row.image || 'https://images.unsplash.com/photo-1546069901-ba9599a7e63c?w=500&auto=format&fit=crop&q=80')
          }));

          await batchSaveMenuItems(imported);
          showAlert('Berhasil', `Berhasil mengimpor dan menyinkronkan ${imported.length} menu ke database Firebase!`);
        } else {
          showAlert('Gagal', 'File Excel kosong atau format tidak sesuai.');
        }
      } catch (err) {
        console.error(err);
        showAlert('Gagal', 'Terjadi kesalahan saat membaca file Excel.');
      }
      e.target.value = '';
    };
    reader.readAsArrayBuffer(file);
  };

  // Staff Management
  const handleSaveStaff = async (e: React.FormEvent) => {
    e.preventDefault();
    const staff: StaffUser = {
      id: editingStaff ? editingStaff.id : String(Date.now()),
      name: staffForm.name,
      username: staffForm.username,
      pin: staffForm.pin,
      role: staffForm.role
    };

    try {
      await saveStaffUser(staff);
      setShowStaffModal(false);
      showAlert('Berhasil', 'Data staff berhasil disimpan ke Firebase!');
    } catch (err) {
      showAlert('Gagal', 'Gagal menyimpan data staff.');
    }
  };

  const handleDeleteStaff = (id: string | number) => {
    if (staffUsers.length <= 1) {
      showAlert('Peringatan', 'Restoran harus memiliki setidaknya 1 akun staff.');
      return;
    }
    showAlert('Hapus Staff', 'Apakah Anda yakin ingin menghapus akun staff ini?', async () => {
      try {
        await deleteStaffUser(id);
      } catch (err) {
        showAlert('Gagal', 'Gagal menghapus data staff.');
      }
    });
  };

  const openStaffModal = (staff: StaffUser | null = null) => {
    if (staff) {
      setEditingStaff(staff);
      setStaffForm({
        name: staff.name,
        username: staff.username,
        pin: staff.pin,
        role: staff.role
      });
    } else {
      setEditingStaff(null);
      setStaffForm({
        name: '',
        username: '',
        pin: '',
        role: 'cashier'
      });
    }
    setShowStaffModal(true);
  };

  // Payment Methods Management
  const handleSavePayment = async (e: React.FormEvent) => {
    e.preventDefault();
    const pm: PaymentMethod = {
      id: editingPayment ? editingPayment.id : `pay_${Date.now()}`,
      name: paymentForm.name,
      type: paymentForm.type,
      isActive: paymentForm.isActive,
      instructions: paymentForm.instructions
    };

    try {
      await savePaymentMethod(pm);
      setShowPaymentModal(false);
      showAlert('Berhasil', 'Metode pembayaran tersimpan di Firebase!');
    } catch (err) {
      showAlert('Gagal', 'Gagal menyimpan metode pembayaran.');
    }
  };

  const handleDeletePayment = (id: string) => {
    if (paymentMethods.length <= 1) {
      showAlert('Peringatan', 'Restoran harus memiliki setidaknya 1 metode pembayaran.');
      return;
    }
    showAlert('Hapus Metode', 'Hapus metode pembayaran ini?', async () => {
      try {
        await deletePaymentMethod(id);
      } catch (err) {
        showAlert('Gagal', 'Gagal menghapus metode pembayaran.');
      }
    });
  };

  const openPaymentModal = (pm: PaymentMethod | null = null) => {
    if (pm) {
      setEditingPayment(pm);
      setPaymentForm({
        name: pm.name,
        type: pm.type,
        isActive: pm.isActive,
        instructions: pm.instructions
      });
    } else {
      setEditingPayment(null);
      setPaymentForm({
        name: '',
        type: 'cash',
        isActive: true,
        instructions: ''
      });
    }
    setShowPaymentModal(true);
  };

  // Tables Management
  const handleAddTable = async () => {
    const name = prompt('Masukkan Nomor/Nama Meja baru:');
    if (name && name.trim()) {
      const id = `T_${Date.now()}`;
      const link = `${window.location.origin}${window.location.pathname}?table=${encodeURIComponent(name.trim())}`;
      try {
        await saveTable({ id, name: name.trim(), link });
        showAlert('Berhasil', `Meja ${name} berhasil ditambahkan dan disimpan di Firebase!`);
      } catch (err) {
        showAlert('Gagal', 'Gagal menyimpan meja baru.');
      }
    }
  };

  const handleDeleteTable = (id: string) => {
    showAlert('Hapus Meja', 'Apakah Anda yakin ingin menghapus meja ini?', async () => {
      try {
        await deleteTable(id);
      } catch (err) {
        showAlert('Gagal', 'Gagal menghapus meja.');
      }
    });
  };

  const downloadTableQrPng = async (table: TableItem) => {
    const tableLink = table.link || `${window.location.origin}${window.location.pathname}?table=${encodeURIComponent(table.name)}`;
    const qrDataUrl = await generateQrDataUrl(tableLink, 300);

    const canvas = document.createElement('canvas');
    canvas.width = 440;
    canvas.height = 500;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    // Card background
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    // Border
    ctx.strokeStyle = '#e4e4e7';
    ctx.lineWidth = 3;
    ctx.strokeRect(12, 12, canvas.width - 24, canvas.height - 24);

    // Restaurant Title
    ctx.fillStyle = '#0f172a';
    ctx.font = 'bold 24px Inter, sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText(settings.restoName, canvas.width / 2, 60);

    // Table Badge
    const badgeW = 160;
    const badgeH = 40;
    const bx = (canvas.width - badgeW) / 2;
    const by = 80;
    ctx.fillStyle = '#dc2626';
    ctx.beginPath();
    ctx.roundRect(bx, by, badgeW, badgeH, 20);
    ctx.fill();

    ctx.fillStyle = '#ffffff';
    ctx.font = 'bold 18px Inter, sans-serif';
    ctx.fillText(table.name, canvas.width / 2, by + 26);

    // Instructions
    ctx.fillStyle = '#64748b';
    ctx.font = '14px Inter, sans-serif';
    ctx.fillText('Scan QR ini untuk melihat menu & memesan', canvas.width / 2, 150);

    // Draw QR code image
    const img = new Image();
    img.onload = () => {
      ctx.drawImage(img, (canvas.width - 260) / 2, 170, 260, 260);

      ctx.fillStyle = '#94a3b8';
      ctx.font = '11px monospace';
      ctx.fillText(`ID: ${table.id} • Realtime Resto QR`, canvas.width / 2, 465);

      const downloadUrl = canvas.toDataURL('image/png');
      const a = document.createElement('a');
      a.href = downloadUrl;
      a.download = `QR-${settings.restoName.replace(/\s+/g, '')}-${table.name}.png`;
      a.click();
    };
    img.src = qrDataUrl;
  };

  // Category Management
  const handleAddCategory = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!categoryForm.name.trim()) return;
    const name = categoryForm.name.trim();
    if (categories.some((c) => c.name.toLowerCase() === name.toLowerCase())) {
      showAlert('Peringatan', 'Kategori ini sudah ada.');
      return;
    }
    try {
      await saveCategory({ id: name, name });
      setCategoryForm({ name: '' });
      setShowCategoryModal(false);
      showAlert('Berhasil', 'Kategori berhasil ditambahkan!');
    } catch (err) {
      showAlert('Gagal', 'Gagal menambahkan kategori.');
    }
  };

  const handleDeleteCategory = (catId: string) => {
    if (catId === 'Semua') return;
    if (categories.length <= 2) {
      showAlert('Peringatan', 'Harus ada setidaknya 1 kategori selain "Semua".');
      return;
    }
    showAlert('Hapus Kategori', `Hapus kategori "${catId}"? Menu dalam kategori ini akan otomatis dipindahkan ke kategori lain.`, async () => {
      try {
        const fallbackCat = categories.find((c) => c.id !== 'Semua' && c.id !== catId)?.id || 'Makanan Utama';
        const updatedMenu = menuItems.map((m) => (m.category === catId ? { ...m, category: fallbackCat } : m));
        await batchSaveMenuItems(updatedMenu);
        await deleteCategory(catId);
      } catch (err) {
        showAlert('Gagal', 'Gagal menghapus kategori.');
      }
    });
  };

  // Settings Save
  const handleSaveSettings = async (newSettings: Settings) => {
    try {
      await persistSettings(newSettings);
      showAlert('Berhasil', 'Pengaturan restoran berhasil disimpan ke Firebase!');
    } catch (err) {
      showAlert('Gagal', 'Gagal menyimpan pengaturan ke database.');
    }
  };

  // Printing & Image Struk
  const printReceipt = async (order: Order) => {
    const paymentObj = paymentMethods.find((p) => p.id === order.paymentMethod);
    const paymentName = paymentObj?.name || order.paymentMethod;
    const isCash = paymentObj?.type === 'cash';
    const isQris = paymentObj?.type === 'qris';

    let qrisDataUrl = '';
    let txQrDataUrl = '';

    if (isQris) {
      const dynQris = generateDynamicQris(settings.staticQris, Math.round(order.total));
      qrisDataUrl = await generateQrDataUrl(dynQris, 140);
    }
    if (isCash) {
      txQrDataUrl = await generateQrDataUrl(order.id, 120);
    }

    const printWindow = window.open('', '_blank');
    if (!printWindow) {
      showAlert('Peringatan', 'Pop-up terblokir oleh browser. Izinkan pop-up untuk mencetak struk.');
      return;
    }

    printWindow.document.write(`
      <!DOCTYPE html>
      <html>
      <head>
        <title>Struk Pembayaran - ${order.id}</title>
        <style>
          body { font-family: 'Courier New', Courier, monospace; width: ${settings.receiptPaperWidth || '300px'}; margin: 0 auto; padding: 6px; font-size: 11px; color: #000; line-height: 1.25; }
          .center { text-align: center; }
          .bold { font-weight: bold; }
          .line { border-bottom: 1px dashed #000; margin: 5px 0; }
          .flex { display: flex; justify-content: space-between; }
          .qr-box { text-align: center; margin: 6px 0; }
          .qr-box img { width: 110px; height: 110px; display: inline-block; }
          .inst { font-size: 9px; margin-top: 3px; background: #f4f4f5; padding: 4px; border-radius: 4px; }
        </style>
      </head>
      <body>
        <div class="center bold" style="font-size: 16px;">${settings.restoName}</div>
        ${settings.receiptHeaderNote ? `<div class="center" style="font-size: 10px;">${settings.receiptHeaderNote}</div>` : ''}
        ${settings.receiptShowAddress && settings.restoAddress ? `<div class="center" style="font-size: 9px;">${settings.restoAddress}</div>` : ''}
        ${settings.receiptShowPhone && settings.restoPhone ? `<div class="center" style="font-size: 9px;">Telp: ${settings.restoPhone}</div>` : ''}
        <div class="line"></div>
        <div>ID Pesanan : <b>${order.id}</b></div>
        <div>Pemesan    : ${order.customerName || '-'}</div>
        <div>Waktu      : ${new Date(order.time).toLocaleString('id-ID')}</div>
        <div>Meja/Tipe  : ${order.table} (${order.orderType})</div>
        ${settings.receiptShowCashier ? `<div>Kasir/Sumber: ${currentUser?.name || order.source}</div>` : ''}
        <div class="line"></div>
        <div>
          ${order.items
            .map(
              (i) => `
            <div>${i.quantity}x ${i.name}</div>
            <div class="flex">
              <span>  @ Rp ${i.price.toLocaleString('id-ID')}</span>
              <span>Rp ${(i.price * i.quantity).toLocaleString('id-ID')}</span>
            </div>
          `
            )
            .join('')}
        </div>
        <div class="line"></div>
        <div class="flex"><span>Subtotal:</span><span>Rp ${order.subtotal.toLocaleString('id-ID')}</span></div>
        ${settings.taxRate > 0 ? `<div class="flex"><span>Pajak (${settings.taxRate}%):</span><span>Rp ${order.tax.toLocaleString('id-ID')}</span></div>` : ''}
        ${settings.serviceCharge > 0 ? `<div class="flex"><span>Service (${settings.serviceCharge}%):</span><span>Rp ${order.service.toLocaleString('id-ID')}</span></div>` : ''}
        <div class="line"></div>
        <div class="flex bold" style="font-size: 14px;"><span>TOTAL:</span><span>Rp ${order.total.toLocaleString('id-ID')}</span></div>

        ${isCash ? `
          <div class="flex" style="margin-top: 4px;"><span>Tunai:</span><span>Rp ${(order.cashGiven || 0).toLocaleString('id-ID')}</span></div>
          <div class="flex"><span>Kembalian:</span><span>Rp ${(order.changeAmount || 0).toLocaleString('id-ID')}</span></div>
          <div class="line"></div>
          <div class="center bold" style="font-size: 10px;">QR CODE TRANSAKSI</div>
          <div class="qr-box"><img src="${txQrDataUrl}" alt="QR Transaksi" /></div>
          <div class="center" style="font-size: 8px;">Scan ID: ${order.id}</div>
        ` : ''}

        ${isQris ? `
          <div class="line"></div>
          <div class="center bold" style="font-size: 10px;">QRIS PEMBAYARAN</div>
          <div class="qr-box"><img src="${qrisDataUrl}" alt="QRIS" /></div>
          <div class="center" style="font-size: 8px;">Scan untuk konfirmasi pembayaran</div>
        ` : ''}

        <div class="line"></div>
        <div class="center" style="margin-top: 3px;">Pembayaran: ${paymentName}</div>
        <div class="center" style="margin-top: 6px; font-size: 10px;">${settings.receiptFooterNote || 'Terima Kasih Atas Kunjungan Anda!'}</div>
        <script>
          window.onload = function() { window.print(); window.close(); }
        </script>
      </body>
      </html>
    `);
    printWindow.document.close();
  };

  const downloadReceiptImage = async (order: Order) => {
    const paymentObj = paymentMethods.find((p) => p.id === order.paymentMethod);
    const paymentName = paymentObj?.name || order.paymentMethod;
    const isCash = paymentObj?.type === 'cash';
    const isQris = paymentObj?.type === 'qris';

    let h = 280 + order.items.length * 18;
    if (settings.taxRate > 0) h += 18;
    if (settings.serviceCharge > 0) h += 18;
    if (isCash) h += 160;
    if (isQris) h += 170;
    if (settings.receiptShowAddress && settings.restoAddress) h += 18;
    if (settings.receiptShowPhone && settings.restoPhone) h += 18;

    const canvas = document.createElement('canvas');
    canvas.width = 380;
    canvas.height = h;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    ctx.fillStyle = '#000000';
    ctx.font = 'bold 16px Courier New, monospace';
    ctx.textAlign = 'center';
    ctx.fillText(settings.restoName, canvas.width / 2, 24);

    let currY = 40;
    if (settings.receiptHeaderNote) {
      ctx.font = '10px Courier New, monospace';
      ctx.fillText(settings.receiptHeaderNote, canvas.width / 2, currY);
      currY += 16;
    }
    if (settings.receiptShowAddress && settings.restoAddress) {
      ctx.font = '10px Courier New, monospace';
      ctx.fillText(settings.restoAddress, canvas.width / 2, currY);
      currY += 16;
    }
    if (settings.receiptShowPhone && settings.restoPhone) {
      ctx.font = '10px Courier New, monospace';
      ctx.fillText(`Telp: ${settings.restoPhone}`, canvas.width / 2, currY);
      currY += 16;
    }

    ctx.textAlign = 'left';
    ctx.font = '11px Courier New, monospace';
    let y = currY + 6;
    ctx.fillText(`ID Pesanan : ${order.id}`, 16, y); y += 16;
    ctx.fillText(`Pemesan    : ${order.customerName || '-'}`, 16, y); y += 16;
    ctx.fillText(`Waktu      : ${new Date(order.time).toLocaleString('id-ID')}`, 16, y); y += 16;
    ctx.fillText(`Meja/Tipe  : ${order.table} (${order.orderType})`, 16, y); y += 18;

    ctx.beginPath();
    ctx.setLineDash([4, 2]);
    ctx.moveTo(16, y); ctx.lineTo(364, y);
    ctx.stroke();
    ctx.setLineDash([]);
    y += 18;

    order.items.forEach((item) => {
      ctx.fillText(`${item.quantity}x ${item.name}`, 16, y);
      ctx.textAlign = 'right';
      ctx.fillText(`Rp ${(item.price * item.quantity).toLocaleString('id-ID')}`, 364, y);
      ctx.textAlign = 'left';
      y += 16;
    });

    ctx.beginPath();
    ctx.setLineDash([4, 2]);
    ctx.moveTo(16, y); ctx.lineTo(364, y);
    ctx.stroke();
    ctx.setLineDash([]);
    y += 18;

    ctx.fillText('Subtotal', 16, y);
    ctx.textAlign = 'right';
    ctx.fillText(`Rp ${order.subtotal.toLocaleString('id-ID')}`, 364, y);
    ctx.textAlign = 'left';
    y += 16;

    if (settings.taxRate > 0) {
      ctx.fillText(`Pajak (${settings.taxRate}%)`, 16, y);
      ctx.textAlign = 'right';
      ctx.fillText(`Rp ${order.tax.toLocaleString('id-ID')}`, 364, y);
      ctx.textAlign = 'left';
      y += 16;
    }

    if (settings.serviceCharge > 0) {
      ctx.fillText(`Service (${settings.serviceCharge}%)`, 16, y);
      ctx.textAlign = 'right';
      ctx.fillText(`Rp ${order.service.toLocaleString('id-ID')}`, 364, y);
      ctx.textAlign = 'left';
      y += 16;
    }

    ctx.beginPath();
    ctx.setLineDash([4, 2]);
    ctx.moveTo(16, y); ctx.lineTo(364, y);
    ctx.stroke();
    ctx.setLineDash([]);
    y += 20;

    ctx.font = 'bold 14px Courier New, monospace';
    ctx.fillText('TOTAL', 16, y);
    ctx.textAlign = 'right';
    ctx.fillText(`Rp ${order.total.toLocaleString('id-ID')}`, 364, y);
    ctx.textAlign = 'left';
    y += 20;

    if (isCash) {
      ctx.font = '11px Courier New, monospace';
      ctx.fillText('Tunai Diterima', 16, y);
      ctx.textAlign = 'right';
      ctx.fillText(`Rp ${(order.cashGiven || 0).toLocaleString('id-ID')}`, 364, y);
      ctx.textAlign = 'left';
      y += 16;

      ctx.fillText('Kembalian', 16, y);
      ctx.textAlign = 'right';
      ctx.fillText(`Rp ${(order.changeAmount || 0).toLocaleString('id-ID')}`, 364, y);
      ctx.textAlign = 'left';
      y += 18;
    }

    const finishDownload = () => {
      ctx.font = '11px Courier New, monospace';
      ctx.textAlign = 'center';
      ctx.fillText(`Pembayaran: ${paymentName}`, canvas.width / 2, y);
      y += 16;
      ctx.fillText(settings.receiptFooterNote || 'Terima Kasih Atas Kunjungan Anda!', canvas.width / 2, y);

      const downloadUrl = canvas.toDataURL('image/png');
      const a = document.createElement('a');
      a.href = downloadUrl;
      a.download = `Struk-${order.id}.png`;
      a.click();
    };

    if (isQris) {
      const dynQris = generateDynamicQris(settings.staticQris, Math.round(order.total));
      const qrDataUrl = await generateQrDataUrl(dynQris, 120);
      const img = new Image();
      img.onload = () => {
        ctx.beginPath();
        ctx.setLineDash([4, 2]);
        ctx.moveTo(16, y); ctx.lineTo(364, y);
        ctx.stroke();
        ctx.setLineDash([]);
        y += 16;

        ctx.font = 'bold 11px Courier New, monospace';
        ctx.textAlign = 'center';
        ctx.fillText('QRIS PEMBAYARAN', canvas.width / 2, y);
        y += 12;

        ctx.drawImage(img, (canvas.width - 100) / 2, y, 100, 100);
        y += 115;
        finishDownload();
      };
      img.src = qrDataUrl;
    } else if (isCash) {
      const txQrUrl = await generateQrDataUrl(order.id, 110);
      const img = new Image();
      img.onload = () => {
        ctx.beginPath();
        ctx.setLineDash([4, 2]);
        ctx.moveTo(16, y); ctx.lineTo(364, y);
        ctx.stroke();
        ctx.setLineDash([]);
        y += 16;

        ctx.font = 'bold 11px Courier New, monospace';
        ctx.textAlign = 'center';
        ctx.fillText('QR CODE TRANSAKSI', canvas.width / 2, y);
        y += 12;

        ctx.drawImage(img, (canvas.width - 100) / 2, y, 100, 100);
        y += 115;
        finishDownload();
      };
      img.src = txQrUrl;
    } else {
      finishDownload();
    }
  };

  return (
    <div className="min-h-screen text-xs sm:text-sm text-zinc-900 bg-zinc-100 flex flex-col font-sans">
      {/* Firebase & Network Banner */}
      <div className="w-full bg-zinc-900 text-zinc-300 text-xs py-1.5 px-4 flex flex-wrap items-center justify-between gap-2 shadow-sm border-b border-zinc-800">
        <div className="flex items-center space-x-2">
          <span className="flex h-2 w-2 relative">
            <span className={`animate-ping absolute inline-flex h-full w-full rounded-full ${firebaseConnected ? 'bg-emerald-400 opacity-75' : 'bg-amber-400 opacity-75'}`}></span>
            <span className={`relative inline-flex rounded-full h-2 w-2 ${firebaseConnected ? 'bg-emerald-500' : 'bg-amber-500'}`}></span>
          </span>
          <span className="font-semibold text-[11px] text-white">
            {firebaseConnected ? 'Firebase Realtime Cloud Database Aktif' : 'Menghubungkan ke Firebase...'}
          </span>
          {isSyncing && <span className="text-[10px] text-amber-300 animate-pulse font-mono">(Syncing...)</span>}
        </div>

        <div className="flex items-center space-x-3 text-[11px]">
          {isOffline ? (
            <span className="bg-red-500/20 text-red-400 px-2 py-0.5 rounded font-bold">Mode Offline (Cache)</span>
          ) : (
            <span className="text-zinc-400">Sinkronisasi Instan Antar Perangkat</span>
          )}
        </div>
      </div>

      {/* Alert / Confirmation Modal */}
      {modalConfig.isOpen && (
        <div className="fixed inset-0 bg-black/60 z-[70] flex items-center justify-center p-4 backdrop-blur-sm">
          <div className="bg-white w-full max-w-sm rounded-2xl p-5 shadow-2xl text-center border border-zinc-100">
            <div className="w-12 h-12 bg-rose-100 text-rose-600 rounded-full flex items-center justify-center mx-auto mb-3">
              <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
              </svg>
            </div>
            <h3 className="text-base font-bold text-zinc-900 mb-1.5">{modalConfig.title}</h3>
            <p className="text-zinc-600 text-xs mb-5">{modalConfig.message}</p>
            <div className="flex gap-2.5">
              {modalConfig.onConfirm && (
                <button
                  onClick={() => {
                    const cb = modalConfig.onConfirm;
                    closeAlert();
                    if (cb) cb();
                  }}
                  className="flex-1 bg-rose-600 text-white font-bold py-2.5 rounded-xl hover:bg-rose-700 shadow-md text-xs transition-colors"
                >
                  Ya
                </button>
              )}
              <button
                onClick={closeAlert}
                className="flex-1 bg-zinc-200 text-zinc-800 font-bold py-2.5 rounded-xl hover:bg-zinc-300 text-xs transition-colors"
              >
                {modalConfig.onConfirm ? 'Batal' : 'OK'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Check Order Status Modal */}
      {showCheckStatusModal && (
        <div className="fixed inset-0 bg-black/60 z-[60] flex items-center justify-center p-4 backdrop-blur-sm">
          <div className="bg-white w-full max-w-sm rounded-2xl p-5 shadow-2xl border border-zinc-100">
            <div className="flex justify-between items-center mb-3 border-b border-zinc-100 pb-2">
              <h3 className="text-base font-bold text-zinc-900">Cek Status Pesanan</h3>
              <button onClick={() => setShowCheckStatusModal(false)} className="text-zinc-400 hover:text-zinc-600 p-1">
                ✕
              </button>
            </div>
            <p className="text-xs text-zinc-500 mb-3">
              Masukkan ID Pesanan (contoh: ORD-1234) untuk memantau status pembuatan makanan secara langsung dari Firestore.
            </p>
            <div className="space-y-3">
              <div className="flex gap-2">
                <input
                  type="text"
                  placeholder="ORD-XXXX"
                  value={checkOrderIdInput}
                  onChange={(e) => setCheckOrderIdInput(e.target.value)}
                  className="flex-1 px-3 py-2.5 bg-zinc-50 border border-zinc-200 rounded-xl text-xs sm:text-sm font-mono font-bold uppercase outline-none focus:border-rose-500"
                />
                <button
                  onClick={async () => {
                    try {
                      if (navigator.clipboard && navigator.clipboard.readText) {
                        const text = await navigator.clipboard.readText();
                        if (text) setCheckOrderIdInput(text.trim());
                      }
                    } catch (err) {
                      showAlert('Info', 'Silakan tempel (Ctrl+V) langsung ke kolom input.');
                    }
                  }}
                  className="px-3 bg-zinc-100 hover:bg-zinc-200 text-zinc-700 font-bold rounded-xl text-xs flex-shrink-0"
                >
                  Paste
                </button>
              </div>
              <button
                onClick={() => {
                  const searchId = checkOrderIdInput.trim().toUpperCase();
                  const found = orders.find((o) => o.id.toUpperCase() === searchId);
                  if (found) {
                    setShowCheckStatusModal(false);
                    setActiveOrder(found);
                    setShowOrderTrackerModal(true);
                    setView('customer');
                  } else {
                    showAlert('Tidak Ditemukan', `Pesanan dengan ID ${searchId} tidak ditemukan.`);
                  }
                }}
                className="w-full bg-rose-600 hover:bg-rose-700 text-white font-bold py-2.5 rounded-xl text-xs shadow-md transition-colors"
              >
                Cek Status
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Receipt Modal */}
      {receiptOrder && (
        <div className="fixed inset-0 bg-black/60 z-50 flex items-center justify-center p-3 sm:p-4 backdrop-blur-sm">
          <div className="bg-white w-full max-w-sm rounded-2xl p-4 sm:p-5 shadow-2xl flex flex-col max-h-[90vh]">
            <div className="flex justify-between items-center mb-3 border-b border-zinc-100 pb-2">
              <h3 className="text-base font-bold text-zinc-900">Struk Pembayaran</h3>
              <button onClick={() => setReceiptOrder(null)} className="text-zinc-400 hover:text-zinc-600 p-1 font-bold">
                ✕
              </button>
            </div>

            <div className="flex-1 overflow-y-auto space-y-2.5 font-mono text-[11px] sm:text-xs bg-zinc-50 p-3.5 rounded-xl border border-zinc-200 text-zinc-800">
              <div className="text-center font-bold text-base text-zinc-900">{settings.restoName}</div>
              {settings.receiptHeaderNote && <div className="text-center text-[10px] text-zinc-500">{settings.receiptHeaderNote}</div>}
              {settings.receiptShowAddress && settings.restoAddress && <div className="text-center text-[10px] text-zinc-500">{settings.restoAddress}</div>}
              {settings.receiptShowPhone && settings.restoPhone && <div className="text-center text-[10px] text-zinc-500">Telp: {settings.restoPhone}</div>}
              <div className="border-b border-dashed border-zinc-300 pb-1"></div>
              <div className="flex justify-between items-center bg-white p-2 rounded border border-zinc-200">
                <span>ID: <b>{receiptOrder.id}</b></span>
                <button
                  onClick={async () => {
                    const ok = await copyToClipboardText(receiptOrder.id);
                    if (ok) showAlert('Berhasil', 'ID Pesanan berhasil disalin!');
                  }}
                  className="px-2 py-0.5 bg-rose-50 text-rose-700 rounded font-sans text-[10px] font-bold hover:bg-rose-100"
                >
                  Salin ID
                </button>
              </div>
              <div>Pemesan: <b>{receiptOrder.customerName || '-'}</b></div>
              <div>Waktu: {new Date(receiptOrder.time).toLocaleString('id-ID')}</div>
              <div>Meja/Tipe: {receiptOrder.table} ({receiptOrder.orderType})</div>
              {settings.receiptShowCashier && <div>Kasir: {currentUser?.name || receiptOrder.source}</div>}
              <div className="border-b border-dashed border-zinc-300 pb-1"></div>
              <div className="space-y-1">
                {receiptOrder.items.map((item, idx) => (
                  <div key={idx} className="flex justify-between">
                    <span>{item.quantity}x {item.name}</span>
                    <span>Rp {(item.price * item.quantity).toLocaleString('id-ID')}</span>
                  </div>
                ))}
              </div>
              <div className="border-b border-dashed border-zinc-300 pb-1"></div>
              <div className="flex justify-between"><span>Subtotal:</span><span>Rp {receiptOrder.subtotal.toLocaleString('id-ID')}</span></div>
              {settings.taxRate > 0 && <div className="flex justify-between"><span>Pajak ({settings.taxRate}%):</span><span>Rp {receiptOrder.tax.toLocaleString('id-ID')}</span></div>}
              {settings.serviceCharge > 0 && <div className="flex justify-between"><span>Service ({settings.serviceCharge}%):</span><span>Rp {receiptOrder.service.toLocaleString('id-ID')}</span></div>}
              <div className="border-b border-dashed border-zinc-300 pb-1"></div>
              <div className="flex justify-between font-bold text-sm text-zinc-900">
                <span>TOTAL:</span>
                <span className="text-rose-600">Rp {receiptOrder.total.toLocaleString('id-ID')}</span>
              </div>

              {paymentMethods.find((p) => p.id === receiptOrder.paymentMethod)?.type === 'cash' && (
                <div className="pt-1 space-y-0.5 border-t border-dashed border-zinc-300 mt-1">
                  <div className="flex justify-between"><span>Tunai Diterima:</span><span>Rp {(receiptOrder.cashGiven || 0).toLocaleString('id-ID')}</span></div>
                  <div className="flex justify-between font-bold"><span>Kembalian:</span><span>Rp ${(receiptOrder.changeAmount || 0).toLocaleString('id-ID')}</span></div>
                </div>
              )}

              <div className="border-b border-dashed border-zinc-300 pb-1 mt-1"></div>
              <div className="text-center text-[11px]">
                Pembayaran: {paymentMethods.find((p) => p.id === receiptOrder.paymentMethod)?.name || receiptOrder.paymentMethod}
              </div>
              <div className="text-center text-[10px] text-zinc-500 pt-0.5">
                {settings.receiptFooterNote || 'Terima Kasih Atas Kunjungan Anda!'}
              </div>
            </div>

            <div className="grid grid-cols-2 gap-2 mt-3">
              <button
                onClick={() => printReceipt(receiptOrder)}
                className="bg-rose-600 hover:bg-rose-700 text-white font-bold py-2.5 rounded-xl shadow-md text-xs flex items-center justify-center space-x-1.5 transition-colors"
              >
                <span>🖨️ Cetak Thermal</span>
              </button>
              <button
                onClick={() => downloadReceiptImage(receiptOrder)}
                className="bg-zinc-900 hover:bg-zinc-800 text-white font-bold py-2.5 rounded-xl shadow-md text-xs flex items-center justify-center space-x-1.5 transition-colors"
              >
                <span>💾 Simpan Gambar</span>
              </button>
            </div>
            <button
              onClick={() => setReceiptOrder(null)}
              className="mt-2 w-full bg-zinc-200 text-zinc-800 font-bold py-2 rounded-xl text-xs hover:bg-zinc-300 transition-colors"
            >
              Tutup
            </button>
          </div>
        </div>
      )}

      {/* Order Tracker Modal */}
      {showOrderTrackerModal && activeOrder && (
        <div className="fixed inset-0 bg-black/60 z-50 flex items-center justify-center p-4 backdrop-blur-sm">
          <div className="bg-white w-full max-w-md rounded-3xl shadow-2xl p-5 sm:p-6 text-center border border-zinc-200 flex flex-col max-h-[90vh]">
            <div className="flex justify-between items-center mb-3 border-b border-zinc-100 pb-2">
              <h3 className="text-base font-bold text-zinc-900">Status Pesanan Realtime</h3>
              <button
                onClick={() => setShowOrderTrackerModal(false)}
                className="text-zinc-400 hover:text-zinc-600 p-1 font-bold"
              >
                ✕
              </button>
            </div>
            {(() => {
              const live = orders.find((o) => o.id === activeOrder.id) || activeOrder;
              const status = live.status;
              return (
                <div className="overflow-y-auto flex-1">
                  <div className={`w-16 h-16 rounded-full flex items-center justify-center mx-auto mb-3 border-4 border-white shadow-md ${
                    status === 'completed' ? 'bg-emerald-100 text-emerald-600' :
                    status === 'processing' ? 'bg-blue-100 text-blue-600 animate-pulse' : 'bg-amber-100 text-amber-600'
                  }`}>
                    {status === 'completed' ? (
                      <svg className="w-8 h-8" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" /></svg>
                    ) : status === 'processing' ? (
                      <svg className="w-8 h-8" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z" /></svg>
                    ) : (
                      <svg className="w-8 h-8" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z" /></svg>
                    )}
                  </div>

                  <h4 className="text-base font-bold text-zinc-900 mb-0.5">
                    {status === 'completed' ? 'Pesanan Selesai / Siap Diambil!' :
                     status === 'processing' ? 'Sedang Dimasak di Dapur' : 'Menunggu Konfirmasi Kasir'}
                  </h4>
                  <p className="text-zinc-500 text-xs mb-3">
                    {live.orderType === 'takeaway' ? 'Tipe: Bawa Pulang' : `Meja: ${live.table}`} • ID: {live.id}
                  </p>

                  <div className="bg-zinc-50 rounded-2xl p-3.5 text-left border border-zinc-200 mb-3 text-xs">
                    <div className="flex justify-between items-center mb-2">
                      <span className="text-zinc-500">ID Pesanan:</span>
                      <div className="flex items-center space-x-1.5">
                        <span className="font-mono font-bold text-zinc-900">{live.id}</span>
                        <button
                          onClick={async () => {
                            const ok = await copyToClipboardText(live.id);
                            if (ok) showAlert('Berhasil', 'ID Pesanan disalin!');
                          }}
                          className="text-[10px] bg-white border border-zinc-200 px-2 py-0.5 rounded font-bold hover:bg-zinc-100"
                        >
                          Salin
                        </button>
                      </div>
                    </div>
                    <div className="space-y-1 border-t border-zinc-200 pt-2">
                      {live.items.map((it, idx) => (
                        <div key={idx} className="flex justify-between">
                          <span>{it.quantity}x {it.name}</span>
                          <span className="font-semibold">Rp {(it.price * it.quantity).toLocaleString('id-ID')}</span>
                        </div>
                      ))}
                    </div>
                    <div className="mt-2.5 pt-2 border-t border-zinc-200 flex justify-between font-bold text-sm">
                      <span>Total:</span>
                      <span className="text-rose-600">Rp {live.total.toLocaleString('id-ID')}</span>
                    </div>
                  </div>

                  <div className="flex gap-2">
                    <button
                      onClick={() => setReceiptOrder(live)}
                      className="flex-1 bg-rose-50 text-rose-700 font-bold py-2.5 rounded-xl hover:bg-rose-100 text-xs transition-colors"
                    >
                      Buka Struk Pembayaran
                    </button>
                    <button
                      onClick={() => setShowOrderTrackerModal(false)}
                      className="flex-1 bg-zinc-900 text-white font-bold py-2.5 rounded-xl text-xs hover:bg-zinc-800 transition-colors"
                    >
                      Kembali ke Menu
                    </button>
                  </div>
                </div>
              );
            })()}
          </div>
        </div>
      )}

      {/* Table & Order Type Selection Modal */}
      {showTableSelectModal && (
        <div className="fixed inset-0 bg-black/60 z-50 flex items-center justify-center p-4 backdrop-blur-sm">
          <div className="bg-white w-full max-w-sm rounded-3xl p-5 sm:p-6 shadow-2xl border border-zinc-100">
            <div className="flex justify-between items-center mb-3 border-b border-zinc-100 pb-2">
              <h3 className="text-base font-bold text-zinc-900">Pilih Meja Pemesanan</h3>
              <button
                onClick={() => setShowTableSelectModal(false)}
                className="text-zinc-400 hover:text-zinc-600 p-1 font-bold"
              >
                ✕
              </button>
            </div>
            <p className="text-xs text-zinc-500 mb-4">
              Silakan pilih opsi makan di tempat atau bawa pulang untuk membuka buku menu.
            </p>

            <form onSubmit={handleCustomerLogin} className="space-y-4">
              <div className="grid grid-cols-2 gap-2 bg-zinc-100 p-1 rounded-2xl border border-zinc-200">
                <button
                  type="button"
                  onClick={() => setCustomerOrderType('dine-in')}
                  className={`py-2 rounded-xl font-bold text-xs transition-all ${
                    customerOrderType === 'dine-in' ? 'bg-white text-rose-600 shadow-sm' : 'text-zinc-500 hover:text-zinc-700'
                  }`}
                >
                  Makan di Tempat
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setCustomerOrderType('takeaway');
                    setCustomerTable('Takeaway');
                  }}
                  className={`py-2 rounded-xl font-bold text-xs transition-all ${
                    customerOrderType === 'takeaway' ? 'bg-white text-rose-600 shadow-sm' : 'text-zinc-500 hover:text-zinc-700'
                  }`}
                >
                  Bawa Pulang
                </button>
              </div>

              {customerOrderType === 'dine-in' && (
                <div>
                  <label className="text-[11px] font-bold text-zinc-500 uppercase tracking-wider mb-1 block">
                    Nomor / Nama Meja
                  </label>
                  <input
                    type="text"
                    placeholder="Contoh: 1, 2, VIP"
                    required={customerOrderType === 'dine-in'}
                    value={customerTable}
                    onChange={(e) => setCustomerTable(e.target.value)}
                    className="w-full px-3.5 py-2.5 bg-zinc-50 border border-zinc-200 rounded-xl text-center text-base font-bold focus:border-rose-500 outline-none uppercase"
                  />
                  <div className="flex flex-wrap gap-1.5 mt-2">
                    {tables.map((t) => (
                      <button
                        key={t.id}
                        type="button"
                        onClick={() => setCustomerTable(t.name.replace(/\D/g, '') || t.name)}
                        className={`px-2.5 py-1 rounded-lg text-xs font-bold transition-colors ${
                          customerTable === (t.name.replace(/\D/g, '') || t.name)
                            ? 'bg-rose-600 text-white'
                            : 'bg-zinc-100 hover:bg-zinc-200 text-zinc-700'
                        }`}
                      >
                        {t.name}
                      </button>
                    ))}
                  </div>
                </div>
              )}

              <button
                type="submit"
                className="w-full bg-rose-600 hover:bg-rose-700 text-white font-bold py-3 rounded-xl shadow-md text-xs sm:text-sm transition-all"
              >
                Buka Menu Makanan
              </button>
            </form>
          </div>
        </div>
      )}

      {/* STORE LANDING PAGE (LANDING PAGE TOKO) */}
      {view === 'landing' && (
        <div className="min-h-screen bg-zinc-50 text-zinc-900 flex flex-col font-sans">
          {/* Store Navbar */}
          <header className="sticky top-0 z-40 bg-white/95 backdrop-blur-md border-b border-zinc-200 shadow-xs">
            <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
              <div className="flex items-center space-x-3">
                <div className="w-10 h-10 rounded-xl bg-gradient-to-tr from-rose-600 to-rose-500 text-white flex items-center justify-center font-black text-lg shadow-md shadow-rose-500/20">
                  🍽️
                </div>
                <div>
                  <h1 className="text-base sm:text-lg font-black text-zinc-900 leading-tight">
                    {settings.restoName}
                  </h1>
                  <p className="text-[10px] text-emerald-600 font-bold flex items-center gap-1">
                    <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse"></span>
                    Buka Hari Ini (10:00 - 22:00 WIB)
                  </p>
                </div>
              </div>

              {/* Desktop links */}
              <nav className="hidden md:flex items-center space-x-6 text-xs font-semibold text-zinc-600">
                <a href="#menu-populer" className="hover:text-rose-600 transition-colors">Menu Populer</a>
                <a href="#cara-pesan" className="hover:text-rose-600 transition-colors">Cara Pesan</a>
                <a href="#keunggulan" className="hover:text-rose-600 transition-colors">Keunggulan</a>
                <a href="#lokasi" className="hover:text-rose-600 transition-colors">Lokasi & Kontak</a>
              </nav>

              {/* Header Action Buttons */}
              <div className="flex items-center space-x-2">
                <button
                  onClick={() => setShowCheckStatusModal(true)}
                  className="px-3 py-1.5 bg-zinc-100 hover:bg-zinc-200 text-zinc-700 font-bold text-xs rounded-xl transition-colors flex items-center space-x-1"
                >
                  <span>🔍</span>
                  <span className="hidden sm:inline">Cek Status</span>
                </button>
                <button
                  onClick={() => startOrdering('dine-in')}
                  className="px-3.5 py-1.5 bg-rose-600 hover:bg-rose-700 text-white font-bold text-xs rounded-xl shadow-md shadow-rose-600/20 transition-all flex items-center space-x-1"
                >
                  <span>🍽️</span>
                  <span>Buka Menu</span>
                </button>
                <button
                  onClick={() => setView('admin_login')}
                  className="p-1.5 text-zinc-400 hover:text-rose-600 hover:bg-zinc-100 rounded-xl transition-colors text-xs font-bold"
                  title="Login Staff / Kasir"
                >
                  🔐
                </button>
              </div>
            </div>
          </header>

          {/* Hero Section */}
          <section className="relative overflow-hidden bg-gradient-to-b from-rose-50/70 via-white to-zinc-50 py-10 sm:py-16 border-b border-zinc-200">
            <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
              <div className="grid grid-cols-1 lg:grid-cols-12 gap-8 lg:gap-12 items-center">
                
                {/* Hero Left Content */}
                <div className="lg:col-span-7 space-y-5 text-center lg:text-left">
                  <div className="inline-flex items-center space-x-2 px-3 py-1 bg-rose-100 text-rose-800 rounded-full text-xs font-bold">
                    <span>⭐ 4.9 Rating (1.850+ Ulasan Puas)</span>
                    <span className="text-rose-300">•</span>
                    <span>100% Halal & Higienis</span>
                  </div>

                  <h1 className="text-3xl sm:text-4xl lg:text-5xl font-black text-zinc-950 tracking-tight leading-[1.15]">
                    Cita Rasa Autentik Nusantara, <span className="text-rose-600">Pesan Instan</span> dari Meja Anda
                  </h1>

                  <p className="text-zinc-600 text-xs sm:text-sm lg:text-base leading-relaxed max-w-2xl mx-auto lg:mx-0">
                    Nikmati aneka hidangan resep istimewa di <b>{settings.restoName}</b>. Tanpa perlu antre di kasir, cukup pilih nomor meja atau bungkus (takeaway) langsung dari ponsel Anda.
                  </p>

                  {/* Highlights Strip */}
                  <div className="grid grid-cols-3 gap-2.5 pt-2 max-w-lg mx-auto lg:mx-0">
                    <div className="bg-white p-3 rounded-2xl border border-zinc-200 shadow-xs text-center">
                      <span className="text-lg">⚡</span>
                      <p className="font-bold text-xs text-zinc-900 mt-0.5">Pesan Mandiri</p>
                      <p className="text-[10px] text-zinc-500">Scan QR langsung</p>
                    </div>
                    <div className="bg-white p-3 rounded-2xl border border-zinc-200 shadow-xs text-center">
                      <span className="text-lg">💳</span>
                      <p className="font-bold text-xs text-zinc-900 mt-0.5">QRIS & Tunai</p>
                      <p className="text-[10px] text-zinc-500">Dinamis & otomatis</p>
                    </div>
                    <div className="bg-white p-3 rounded-2xl border border-zinc-200 shadow-xs text-center">
                      <span className="text-lg">🍲</span>
                      <p className="font-bold text-xs text-zinc-900 mt-0.5">Selalu Segar</p>
                      <p className="text-[10px] text-zinc-500">Masak saat dipesan</p>
                    </div>
                  </div>
                </div>

                {/* Hero Right Quick Order Box */}
                <div className="lg:col-span-5">
                  <div className="bg-white rounded-3xl p-6 sm:p-7 shadow-xl border border-zinc-200 relative">
                    <div className="flex items-center justify-between mb-4 border-b border-zinc-100 pb-3">
                      <div>
                        <h3 className="font-black text-base text-zinc-900">Mulai Pemesanan</h3>
                        <p className="text-xs text-zinc-500">Pilih opsi makan di tempat atau bungkus</p>
                      </div>
                      <span className="px-2.5 py-1 bg-emerald-50 text-emerald-700 font-bold text-[10px] rounded-full border border-emerald-200">
                        Sistem Online
                      </span>
                    </div>

                    <form onSubmit={handleCustomerLogin} className="space-y-4">
                      <div>
                        <label className="text-[11px] font-bold text-zinc-500 uppercase tracking-wider mb-1.5 block">
                          Tipe Pesanan
                        </label>
                        <div className="grid grid-cols-2 gap-2 bg-zinc-100 p-1 rounded-2xl border border-zinc-200">
                          <button
                            type="button"
                            onClick={() => setCustomerOrderType('dine-in')}
                            className={`py-2.5 rounded-xl font-bold text-xs transition-all flex items-center justify-center space-x-1.5 ${
                              customerOrderType === 'dine-in'
                                ? 'bg-white text-rose-600 shadow-sm'
                                : 'text-zinc-600 hover:text-zinc-900'
                            }`}
                          >
                            <span>🍽️</span>
                            <span>Makan di Tempat</span>
                          </button>
                          <button
                            type="button"
                            onClick={() => {
                              setCustomerOrderType('takeaway');
                              setCustomerTable('Takeaway');
                            }}
                            className={`py-2.5 rounded-xl font-bold text-xs transition-all flex items-center justify-center space-x-1.5 ${
                              customerOrderType === 'takeaway'
                                ? 'bg-white text-rose-600 shadow-sm'
                                : 'text-zinc-600 hover:text-zinc-900'
                            }`}
                          >
                            <span>🥡</span>
                            <span>Bawa Pulang</span>
                          </button>
                        </div>
                      </div>

                      {customerOrderType === 'dine-in' && (
                        <div>
                          <div className="flex justify-between items-center mb-1.5">
                            <label className="text-[11px] font-bold text-zinc-500 uppercase tracking-wider block">
                              Pilih / Ketik Nomor Meja
                            </label>
                            <span className="text-[10px] text-zinc-400">Tersedia {tables.length} meja</span>
                          </div>

                          <div className="flex gap-2 mb-2">
                            <input
                              type="text"
                              placeholder="Nomor meja (cth: 1, 2, VIP)"
                              required={customerOrderType === 'dine-in'}
                              value={customerTable}
                              onChange={(e) => setCustomerTable(e.target.value)}
                              className="flex-1 px-3.5 py-2.5 bg-zinc-50 border border-zinc-200 rounded-xl font-bold text-sm focus:border-rose-500 outline-none uppercase text-center"
                            />
                          </div>

                          {/* Quick table chips */}
                          <div className="flex flex-wrap gap-1.5">
                            {tables.slice(0, 6).map((t) => (
                              <button
                                key={t.id}
                                type="button"
                                onClick={() => setCustomerTable(t.name.replace(/\D/g, '') || t.name)}
                                className={`px-2.5 py-1 rounded-lg text-xs font-bold transition-colors ${
                                  customerTable === (t.name.replace(/\D/g, '') || t.name)
                                    ? 'bg-rose-600 text-white'
                                    : 'bg-zinc-100 hover:bg-zinc-200 text-zinc-700'
                                }`}
                              >
                                {t.name}
                              </button>
                            ))}
                          </div>
                        </div>
                      )}

                      <button
                        type="submit"
                        className="w-full bg-rose-600 hover:bg-rose-700 text-white font-bold py-3.5 rounded-xl shadow-lg shadow-rose-600/30 transition-all text-xs sm:text-sm flex items-center justify-center space-x-2"
                      >
                        <span>Buka Buku Menu & Pesan</span>
                        <span>➔</span>
                      </button>
                    </form>

                    <div className="mt-4 pt-3 border-t border-zinc-100 flex items-center justify-between text-[11px] text-zinc-500">
                      <span>Sudah punya ID pesanan?</span>
                      <button
                        onClick={() => setShowCheckStatusModal(true)}
                        className="text-rose-600 font-bold hover:underline"
                      >
                        Cek Status Realtime
                      </button>
                    </div>
                  </div>
                </div>

              </div>
            </div>
          </section>

          {/* Featured Menu Section (Menu Populer) */}
          <section id="menu-populer" className="py-12 sm:py-16 bg-white border-b border-zinc-200">
            <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
              <div className="flex flex-col sm:flex-row justify-between items-start sm:items-end gap-3 mb-8">
                <div>
                  <span className="text-rose-600 font-bold text-xs uppercase tracking-wider block mb-1">
                    Favorit Pelanggan
                  </span>
                  <h2 className="text-2xl sm:text-3xl font-black text-zinc-950">
                    Menu Rekomendasi Hari Ini
                  </h2>
                </div>
                <button
                  onClick={() => startOrdering('dine-in')}
                  className="text-rose-600 font-bold text-xs sm:text-sm hover:underline flex items-center space-x-1"
                >
                  <span>Lihat Semua ({menuItems.length} Pilihan Menu)</span>
                  <span>➔</span>
                </button>
              </div>

              {/* Menu Cards Grid */}
              <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-4 gap-3 sm:gap-5">
                {menuItems.slice(0, 8).map((item) => (
                  <div
                    key={item.id}
                    className="bg-white rounded-2xl border border-zinc-200 overflow-hidden shadow-xs hover:shadow-md hover:border-rose-400 transition-all flex flex-col group"
                  >
                    <div className="aspect-square bg-zinc-100 overflow-hidden relative">
                      <img
                        src={item.image}
                        alt={item.name}
                        className="w-full h-full object-cover group-hover:scale-105 transition-transform duration-300"
                        onError={(e) => { (e.target as HTMLImageElement).src = getFallbackImage(item.name); }}
                      />
                      <span className="absolute top-2 left-2 bg-white/90 backdrop-blur-xs text-[10px] font-bold text-zinc-800 px-2 py-0.5 rounded-md shadow-xs">
                        {item.category}
                      </span>
                    </div>

                    <div className="p-3 sm:p-4 flex-1 flex flex-col">
                      <h4 className="font-bold text-xs sm:text-sm text-zinc-900 line-clamp-1 mb-1">
                        {item.name}
                      </h4>
                      <p className="text-[11px] text-zinc-500 line-clamp-2 mb-3">
                        {item.description}
                      </p>

                      <div className="mt-auto pt-2 border-t border-zinc-100 flex items-center justify-between">
                        <span className="font-black text-rose-600 text-xs sm:text-sm">
                          Rp {item.price.toLocaleString('id-ID')}
                        </span>
                        <button
                          onClick={() => {
                            addToCustomerCart(item);
                            startOrdering('dine-in');
                          }}
                          className="px-2.5 py-1.5 bg-rose-50 hover:bg-rose-600 text-rose-700 hover:text-white rounded-lg text-xs font-bold transition-colors"
                        >
                          + Pesan
                        </button>
                      </div>
                    </div>
                  </div>
                ))}
              </div>

              <div className="text-center mt-8">
                <button
                  onClick={() => startOrdering('dine-in')}
                  className="px-6 py-3 bg-zinc-900 hover:bg-rose-600 text-white font-bold rounded-xl text-xs sm:text-sm shadow-md transition-colors inline-flex items-center space-x-2"
                >
                  <span>Buka Katalog Menu Lengkap ({menuItems.length} Menu)</span>
                  <span>➔</span>
                </button>
              </div>
            </div>
          </section>

          {/* How It Works (Cara Pesan) */}
          <section id="cara-pesan" className="py-12 sm:py-16 bg-zinc-50 border-b border-zinc-200">
            <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 text-center">
              <span className="text-rose-600 font-bold text-xs uppercase tracking-wider block mb-1">
                Mudah & Praktis
              </span>
              <h2 className="text-2xl sm:text-3xl font-black text-zinc-950 mb-3">
                Cara Memesan di {settings.restoName}
              </h2>
              <p className="text-zinc-600 text-xs sm:text-sm max-w-xl mx-auto mb-10">
                Pesan hidangan favorit langsung dari genggaman ponsel Anda hanya dalam 3 langkah mudah.
              </p>

              <div className="grid grid-cols-1 md:grid-cols-3 gap-6 text-left">
                <div className="bg-white p-6 rounded-3xl border border-zinc-200 shadow-xs relative">
                  <span className="w-10 h-10 rounded-2xl bg-rose-100 text-rose-600 font-black text-base flex items-center justify-center mb-4">
                    1
                  </span>
                  <h3 className="font-bold text-sm sm:text-base text-zinc-900 mb-1.5">
                    Pilih Nomor Meja / Scan QR
                  </h3>
                  <p className="text-xs text-zinc-500 leading-relaxed">
                    Duduk di meja yang Anda sukai, lalu scan kode QR di meja atau tentukan nomor meja Anda secara langsung.
                  </p>
                </div>

                <div className="bg-white p-6 rounded-3xl border border-zinc-200 shadow-xs relative">
                  <span className="w-10 h-10 rounded-2xl bg-rose-100 text-rose-600 font-black text-base flex items-center justify-center mb-4">
                    2
                  </span>
                  <h3 className="font-bold text-sm sm:text-base text-zinc-900 mb-1.5">
                    Pilih Menu & Atur Catatan
                  </h3>
                  <p className="text-xs text-zinc-500 leading-relaxed">
                    Pilih makanan dan minuman favorit Anda, tambahkan catatan khusus (cth: tidak pedas, tanpa bawang).
                  </p>
                </div>

                <div className="bg-white p-6 rounded-3xl border border-zinc-200 shadow-xs relative">
                  <span className="w-10 h-10 rounded-2xl bg-rose-100 text-rose-600 font-black text-base flex items-center justify-center mb-4">
                    3
                  </span>
                  <h3 className="font-bold text-sm sm:text-base text-zinc-900 mb-1.5">
                    Bayar & Pantau Realtime
                  </h3>
                  <p className="text-xs text-zinc-500 leading-relaxed">
                    Pilih bayar via QRIS atau tunai di kasir. Pantau proses memasak di layar HP hingga pesanan siap saji.
                  </p>
                </div>
              </div>
            </div>
          </section>

          {/* Store Advantages (Keunggulan Toko) */}
          <section id="keunggulan" className="py-12 sm:py-16 bg-white border-b border-zinc-200">
            <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
              <div className="text-center mb-10">
                <span className="text-rose-600 font-bold text-xs uppercase tracking-wider block mb-1">
                  Komitmen Kami
                </span>
                <h2 className="text-2xl sm:text-3xl font-black text-zinc-950 mb-2">
                  Kenapa Pelanggan Memilih Kami?
                </h2>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                <div className="p-5 rounded-2xl bg-zinc-50 border border-zinc-200">
                  <span className="text-2xl">🌱</span>
                  <h4 className="font-bold text-xs sm:text-sm text-zinc-900 mt-2 mb-1">Bahan Baku Pilihan</h4>
                  <p className="text-[11px] text-zinc-500">Daging dan rempah segar dipasok setiap pagi demi menjaga kualitas rasa autentik.</p>
                </div>
                <div className="p-5 rounded-2xl bg-zinc-50 border border-zinc-200">
                  <span className="text-2xl">⚡</span>
                  <h4 className="font-bold text-xs sm:text-sm text-zinc-900 mt-2 mb-1">Penyajian Cepat</h4>
                  <p className="text-[11px] text-zinc-500">Pesanan dari HP langsung terkirim ke dapur secara realtime untuk langsung diproses.</p>
                </div>
                <div className="p-5 rounded-2xl bg-zinc-50 border border-zinc-200">
                  <span className="text-2xl">📢</span>
                  <h4 className="font-bold text-xs sm:text-sm text-zinc-900 mt-2 mb-1">Panggilan Suara</h4>
                  <p className="text-[11px] text-zinc-500">Notifikasi lonceng & suara ramah saat hidangan Anda telah siap disajikan.</p>
                </div>
                <div className="p-5 rounded-2xl bg-zinc-50 border border-zinc-200">
                  <span className="text-2xl">🧾</span>
                  <h4 className="font-bold text-xs sm:text-sm text-zinc-900 mt-2 mb-1">Struk Digital & Cetak</h4>
                  <p className="text-[11px] text-zinc-500">Simpan struk pembayaran sebagai bukti transaksi resmi atau unduh langsung di HP.</p>
                </div>
              </div>
            </div>
          </section>

          {/* Store Info, Location & Contact */}
          <section id="lokasi" className="py-12 sm:py-16 bg-zinc-900 text-white">
            <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
              <div className="grid grid-cols-1 lg:grid-cols-12 gap-8 items-center">
                <div className="lg:col-span-7 space-y-4">
                  <span className="text-rose-400 font-bold text-xs uppercase tracking-wider block">
                    Kunjungi Restoran Kami
                  </span>
                  <h2 className="text-2xl sm:text-3xl font-black text-white">
                    {settings.restoName}
                  </h2>
                  <p className="text-zinc-300 text-xs sm:text-sm max-w-xl leading-relaxed">
                    Suasana nyaman, hidangan higienis, dan pelayanan ramah menanti Anda dan keluarga.
                  </p>

                  <div className="space-y-3 pt-2 text-xs">
                    <div className="flex items-start space-x-3">
                      <span className="text-base text-rose-400">📍</span>
                      <div>
                        <p className="font-bold text-white">Alamat Restoran</p>
                        <p className="text-zinc-400">{settings.restoAddress || 'Jl. Sudirman No. 45, Jakarta'}</p>
                      </div>
                    </div>
                    <div className="flex items-start space-x-3">
                      <span className="text-base text-rose-400">📞</span>
                      <div>
                        <p className="font-bold text-white">Telepon / Reservasi</p>
                        <p className="text-zinc-400">{settings.restoPhone || '0812-3456-7890'}</p>
                      </div>
                    </div>
                    <div className="flex items-start space-x-3">
                      <span className="text-base text-rose-400">🕒</span>
                      <div>
                        <p className="font-bold text-white">Jam Operasional</p>
                        <p className="text-zinc-400">Senin - Minggu: 10:00 - 22:00 WIB</p>
                      </div>
                    </div>
                  </div>
                </div>

                {/* Right Call-To-Action Card */}
                <div className="lg:col-span-5 bg-zinc-800 p-6 sm:p-7 rounded-3xl border border-zinc-700 text-center space-y-4">
                  <div className="w-14 h-14 bg-rose-600/20 text-rose-500 rounded-2xl flex items-center justify-center mx-auto text-2xl">
                    🍲
                  </div>
                  <h3 className="font-bold text-base text-white">Sudah Berada di Restoran?</h3>
                  <p className="text-xs text-zinc-400">
                    Buka menu sekarang dan nikmati sajian hangat favorit Anda tanpa perlu antre di meja kasir.
                  </p>
                  <button
                    onClick={() => startOrdering('dine-in')}
                    className="w-full bg-rose-600 hover:bg-rose-700 text-white font-bold py-3.5 rounded-xl shadow-lg shadow-rose-600/30 transition-all text-xs sm:text-sm"
                  >
                    Buka Menu & Mulai Pesan
                  </button>
                  <button
                    onClick={() => setShowCheckStatusModal(true)}
                    className="w-full bg-zinc-700 hover:bg-zinc-600 text-zinc-200 font-bold py-2.5 rounded-xl text-xs transition-colors"
                  >
                    Cek Status Pesanan Saya
                  </button>
                </div>
              </div>
            </div>
          </section>

          {/* Footer */}
          <footer className="bg-zinc-950 text-zinc-500 py-6 text-xs border-t border-zinc-800">
            <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 flex flex-col sm:flex-row items-center justify-between gap-3">
              <div className="flex items-center space-x-2">
                <span className="font-bold text-zinc-300">{settings.restoName}</span>
                <span>•</span>
                <span>Resto QR Pro Cloud POS</span>
              </div>
              <div className="flex items-center space-x-4">
                <button
                  onClick={() => setShowCheckStatusModal(true)}
                  className="hover:text-zinc-300 transition-colors"
                >
                  Cek Pesanan
                </button>
                <button
                  onClick={() => setView('admin_login')}
                  className="hover:text-rose-400 transition-colors font-semibold"
                >
                  🔐 Login Kasir / Admin
                </button>
              </div>
            </div>
          </footer>
        </div>
      )}

      {/* ADMIN / CASHIER LOGIN */}
      {view === 'admin_login' && (
        <div className="min-h-screen bg-zinc-900 flex flex-col items-center justify-center p-4">
          <div className="w-full max-w-sm bg-white rounded-3xl p-6 sm:p-8 shadow-2xl border border-zinc-200">
            <div className="w-14 h-14 bg-rose-100 text-rose-600 rounded-2xl flex items-center justify-center mx-auto mb-4">
              <svg className="w-7 h-7" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 15v2m-6 4h12a2 2 0 002-2v-6a2 2 0 00-2-2H6a2 2 0 00-2 2v6a2 2 0 002 2zm10-10V7a4 4 0 00-8 0v4h8z" />
              </svg>
            </div>
            <h2 className="text-xl font-bold text-zinc-900 text-center mb-1">Login Staff Resto</h2>
            <p className="text-xs text-zinc-500 text-center mb-5">
              Masukan User <b className="font-bold text-zinc-700">Kasir/Admin</b> dan <b className="font-bold text-zinc-700">PIN</b> yang anda miliki
            </p>

            <form onSubmit={handleLoginSubmit} className="space-y-4">
              <div>
                <label className="text-[11px] font-bold text-zinc-500 uppercase tracking-wider mb-1 block">USER</label>
                <input
                  type="text"
                  required
                  placeholder="admin / kasir1"
                  value={loginForm.username}
                  onChange={(e) => setLoginForm({ ...loginForm, username: e.target.value })}
                  className="w-full px-3.5 py-2.5 bg-zinc-50 border border-zinc-200 rounded-xl text-zinc-900 font-medium text-xs sm:text-sm focus:outline-none focus:border-rose-500"
                />
              </div>
              <div>
                <label className="text-[11px] font-bold text-zinc-500 uppercase tracking-wider mb-1 block">PIN</label>
                <input
                  type="password"
                  maxLength={6}
                  required
                  placeholder="••••••"
                  value={loginForm.pin}
                  onChange={(e) => setLoginForm({ ...loginForm, pin: e.target.value })}
                  className="w-full px-3.5 py-2.5 bg-zinc-50 border border-zinc-200 rounded-xl text-zinc-900 font-medium text-center tracking-widest focus:outline-none focus:border-rose-500"
                />
              </div>
              <button
                type="submit"
                className="w-full bg-rose-600 hover:bg-rose-700 text-white font-bold py-3 rounded-xl transition-all shadow-md text-xs sm:text-sm"
              >
                Masuk Dashboard POS
              </button>
            </form>

            <button
              onClick={() => setView('landing')}
              className="w-full mt-3 bg-zinc-100 border border-zinc-200 text-zinc-600 font-bold py-2.5 rounded-xl hover:bg-zinc-200 text-xs transition-all"
            >
              Kembali ke Halaman Pelanggan
            </button>
          </div>
        </div>
      )}

      {/* CUSTOMER VIEW - SELALU TETAP DI HALAMAN MENU */}
      {view === 'customer' && (
        <div className="min-h-screen pb-24">
          <header className="bg-white px-4 py-3 shadow-sm sticky top-0 z-20 flex justify-between items-center border-b border-zinc-200">
            <div>
              <h1 className="text-base font-bold text-zinc-900">{settings.restoName}</h1>
              <span className="text-[10px] text-rose-600 font-bold bg-rose-50 inline-block px-2 py-0.5 rounded-md mt-0.5">
                {customerOrderType === 'takeaway' ? 'Takeaway' : `Meja: ${customerTable || '1'}`}
              </span>
            </div>
            <div className="flex items-center space-x-2">
              {activeOrder && (
                <button
                  onClick={() => setShowOrderTrackerModal(true)}
                  className="bg-rose-50 hover:bg-rose-100 text-rose-700 font-bold text-xs px-2.5 py-1.5 rounded-xl flex items-center space-x-1 border border-rose-200 transition-colors"
                >
                  <span>📋</span>
                  <span>Status Pesanan</span>
                </button>
              )}
              <button onClick={() => setView('landing')} className="text-zinc-500 hover:text-zinc-800 font-bold text-xs bg-zinc-100 px-3 py-1.5 rounded-xl">
                Ganti Meja
              </button>
            </div>
          </header>

          {/* Realtime Active Order Status Banner */}
          {activeOrder && (() => {
            const live = orders.find((o) => o.id === activeOrder.id) || activeOrder;
            const status = live.status;
            return (
              <div className="mx-3 sm:mx-5 mt-3 bg-gradient-to-r from-rose-50 to-amber-50 border border-rose-200 rounded-2xl p-3 shadow-sm flex items-center justify-between gap-3">
                <div className="flex items-center space-x-2.5 min-w-0">
                  <span className={`w-3 h-3 rounded-full flex-shrink-0 ${
                    status === 'completed' ? 'bg-emerald-500' :
                    status === 'processing' ? 'bg-blue-500 animate-pulse' : 'bg-amber-500'
                  }`} />
                  <div className="min-w-0">
                    <p className="font-bold text-xs text-zinc-900 truncate">
                      Pesanan <span className="font-mono text-rose-600 font-black">{live.id}</span>:{' '}
                      <span className={status === 'completed' ? 'text-emerald-700' : status === 'processing' ? 'text-blue-700' : 'text-amber-800'}>
                        {status === 'completed' ? 'Selesai / Siap Diambil!' :
                         status === 'processing' ? 'Sedang Dimasak di Dapur' : 'Menunggu Konfirmasi Kasir'}
                      </span>
                    </p>
                    <p className="text-[10px] text-zinc-500">
                      Total Rp {live.total.toLocaleString('id-ID')} • {live.items.length} Menu • Meja {live.table}
                    </p>
                  </div>
                </div>
                <div className="flex items-center space-x-1.5 flex-shrink-0">
                  <button
                    onClick={() => setShowOrderTrackerModal(true)}
                    className="px-2.5 py-1.5 bg-rose-600 text-white hover:bg-rose-700 rounded-xl text-xs font-bold transition-colors shadow-sm"
                  >
                    Detail / Struk
                  </button>
                  <button
                    onClick={() => setActiveOrder(null)}
                    className="p-1 text-zinc-400 hover:text-zinc-600 rounded-lg text-xs"
                    title="Tutup banner"
                  >
                    ✕
                  </button>
                </div>
              </div>
            );
          })()}

            {/* Category horizontal scroll */}
            <div className="bg-white px-3 py-2.5 shadow-sm border-b border-zinc-200 sticky top-[53px] z-10 flex space-x-2 overflow-x-auto">
              {categories.map((cat) => (
                <button
                  key={cat.id}
                  onClick={() => setCustomerActiveCat(cat.id)}
                  className={`px-3.5 py-1.5 rounded-full text-xs font-bold whitespace-nowrap transition-all ${customerActiveCat === cat.id ? 'bg-rose-600 text-white shadow' : 'bg-zinc-100 text-zinc-600 hover:bg-zinc-200'}`}
                >
                  {cat.name}
                </button>
              ))}
            </div>

            {/* Menu Grid */}
            <div className="p-3 sm:p-5 grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-5 gap-3">
              {menuItems
                .filter((item) => customerActiveCat === 'Semua' || item.category === customerActiveCat)
                .map((item) => {
                  const inCart = customerCart.find((c) => c.id === item.id);
                  return (
                    <div key={item.id} className="bg-white rounded-2xl p-2.5 shadow-sm border border-zinc-200 flex flex-col hover:border-rose-300 transition-all">
                      <div className="aspect-square rounded-xl overflow-hidden mb-2 bg-zinc-100">
                        <img
                          src={item.image}
                          alt={item.name}
                          className="w-full h-full object-cover"
                          onError={(e) => {
                            (e.target as HTMLImageElement).src = getFallbackImage(item.name);
                          }}
                        />
                      </div>
                      <h3 className="font-bold text-xs text-zinc-900 leading-tight mb-1">{item.name}</h3>
                      <p className="text-[11px] text-zinc-500 line-clamp-2 mb-2">{item.description}</p>
                      <div className="mt-auto flex items-center justify-between pt-1 border-t border-zinc-100">
                        <span className="font-bold text-rose-600 text-xs">Rp {item.price.toLocaleString('id-ID')}</span>
                        {inCart ? (
                          <div className="flex items-center space-x-1.5 bg-rose-50 rounded-lg p-0.5">
                            <button
                              onClick={() => updateCustomerQuantity(item.id, -1)}
                              className="w-5 h-5 flex items-center justify-center bg-white rounded text-rose-600 font-bold text-xs shadow-sm"
                            >
                              -
                            </button>
                            <span className="text-xs font-bold px-1">{inCart.quantity}</span>
                            <button
                              onClick={() => updateCustomerQuantity(item.id, 1)}
                              className="w-5 h-5 flex items-center justify-center bg-rose-600 rounded text-white font-bold text-xs shadow-sm"
                            >
                              +
                            </button>
                          </div>
                        ) : (
                          <button
                            onClick={() => addToCustomerCart(item)}
                            className="bg-zinc-900 hover:bg-rose-600 text-white p-1.5 rounded-lg transition-colors"
                          >
                            <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 4v16m8-8H4" />
                            </svg>
                          </button>
                        )}
                      </div>
                    </div>
                  );
                })}
            </div>

            {/* Floating Cart Button */}
            {customerCart.length > 0 && !showCart && (
              <div className="fixed bottom-4 left-0 right-0 px-4 z-30">
                <button
                  onClick={() => setShowCart(true)}
                  className="w-full max-w-md mx-auto bg-rose-600 text-white rounded-2xl p-4 shadow-xl flex justify-between items-center font-bold text-xs sm:text-sm"
                >
                  <div className="flex items-center space-x-2">
                    <span className="bg-white/20 px-2 py-0.5 rounded-full text-xs">
                      {customerCart.reduce((s, i) => s + i.quantity, 0)} Item
                    </span>
                    <span>Lihat Keranjang</span>
                  </div>
                  <span>Rp {calculateTotal(customerCart).toLocaleString('id-ID')}</span>
                </button>
              </div>
            )}

            {/* Cart Drawer */}
            {showCart && (
              <div className="fixed inset-0 bg-black/60 z-50 flex flex-col justify-end backdrop-blur-sm">
                <div className="bg-white w-full max-w-lg mx-auto rounded-t-3xl h-[85vh] flex flex-col">
                  <div className="p-4 border-b border-zinc-100 flex justify-between items-center">
                    <h2 className="font-bold text-base text-zinc-900">Keranjang Pesanan</h2>
                    <div className="flex items-center space-x-2">
                      <button
                        onClick={() => {
                          setCustomerCart([]);
                          setShowCart(false);
                        }}
                        className="text-red-500 text-xs font-bold px-2 py-1 hover:bg-red-50 rounded-lg"
                      >
                        Kosongkan
                      </button>
                      <button onClick={() => setShowCart(false)} className="text-zinc-500 p-1 font-bold">
                        ✕
                      </button>
                    </div>
                  </div>

                  <div className="flex-1 overflow-y-auto p-4 space-y-4">
                    <div className="space-y-2.5">
                      {customerCart.map((item) => (
                        <div key={item.id} className="flex gap-2.5 items-center bg-zinc-50 p-2.5 rounded-xl border border-zinc-100">
                          <img
                            src={item.image}
                            alt={item.name}
                            className="w-12 h-12 rounded-lg object-cover bg-zinc-200"
                            onError={(e) => { (e.target as HTMLImageElement).src = getFallbackImage(item.name); }}
                          />
                          <div className="flex-1 min-w-0">
                            <h4 className="font-bold text-xs text-zinc-900 truncate">{item.name}</h4>
                            <p className="text-rose-600 font-bold text-xs">Rp {(item.price * item.quantity).toLocaleString('id-ID')}</p>
                          </div>
                          <div className="flex items-center space-x-1.5 bg-white rounded-lg p-1 border border-zinc-200">
                            <button onClick={() => updateCustomerQuantity(item.id, -1)} className="w-5 h-5 flex items-center justify-center text-xs font-bold text-zinc-600">-</button>
                            <span className="text-xs font-bold w-4 text-center">{item.quantity}</span>
                            <button onClick={() => updateCustomerQuantity(item.id, 1)} className="w-5 h-5 flex items-center justify-center text-xs font-bold text-zinc-600">+</button>
                          </div>
                        </div>
                      ))}
                    </div>

                    <div>
                      <label className="text-xs font-bold text-zinc-600 block mb-1">Nama Pemesan *</label>
                      <input
                        type="text"
                        placeholder="Masukkan nama Anda..."
                        value={customerName}
                        onChange={(e) => setCustomerName(e.target.value)}
                        className="w-full p-2.5 bg-zinc-50 border border-zinc-200 rounded-xl text-xs font-bold focus:border-rose-500 outline-none"
                      />
                    </div>

                    <div>
                      <label className="text-xs font-bold text-zinc-600 block mb-1">Catatan Tambahan (Opsional)</label>
                      <input
                        type="text"
                        placeholder="Contoh: Sambal dipisah, tidak pakai daun bawang"
                        value={checkoutData.notes}
                        onChange={(e) => setCheckoutData({ ...checkoutData, notes: e.target.value })}
                        className="w-full p-2.5 bg-zinc-50 border border-zinc-200 rounded-xl text-xs focus:border-rose-500 outline-none"
                      />
                    </div>

                    <div>
                      <label className="text-xs font-bold text-zinc-600 block mb-1">Pilih Metode Pembayaran</label>
                      <div className="space-y-1.5">
                        {paymentMethods
                          .filter((p) => p.isActive)
                          .map((pm) => (
                            <label
                              key={pm.id}
                              className={`flex items-center p-3 rounded-xl border cursor-pointer transition-all ${checkoutData.paymentMethod === pm.id ? 'border-rose-500 bg-rose-50' : 'border-zinc-200 bg-white'}`}
                            >
                              <input
                                type="radio"
                                name="custPayment"
                                value={pm.id}
                                checked={checkoutData.paymentMethod === pm.id}
                                onChange={(e) => setCheckoutData({ ...checkoutData, paymentMethod: e.target.value })}
                                className="hidden"
                              />
                              <div className={`w-4 h-4 rounded-full border-2 flex items-center justify-center mr-3 ${checkoutData.paymentMethod === pm.id ? 'border-rose-600' : 'border-zinc-300'}`}>
                                {checkoutData.paymentMethod === pm.id && <div className="w-2 h-2 bg-rose-600 rounded-full" />}
                              </div>
                              <div className="flex-1">
                                <span className="font-bold text-xs text-zinc-900 block">{pm.name}</span>
                                {pm.instructions && <span className="text-[10px] text-zinc-500 block">{pm.instructions}</span>}
                              </div>
                            </label>
                          ))}
                      </div>
                    </div>

                    {/* Summary calculation */}
                    <div className="bg-zinc-50 p-3.5 rounded-xl space-y-1 text-xs border border-zinc-200">
                      <div className="flex justify-between text-zinc-500">
                        <span>Subtotal:</span>
                        <span className="font-semibold text-zinc-800">Rp {calculateSubtotal(customerCart).toLocaleString('id-ID')}</span>
                      </div>
                      {settings.taxRate > 0 && (
                        <div className="flex justify-between text-zinc-500">
                          <span>Pajak ({settings.taxRate}%):</span>
                          <span className="font-semibold text-zinc-800">Rp {calculateTax(calculateSubtotal(customerCart)).toLocaleString('id-ID')}</span>
                        </div>
                      )}
                      {settings.serviceCharge > 0 && (
                        <div className="flex justify-between text-zinc-500">
                          <span>Service Charge ({settings.serviceCharge}%):</span>
                          <span className="font-semibold text-zinc-800">Rp {calculateService(calculateSubtotal(customerCart)).toLocaleString('id-ID')}</span>
                        </div>
                      )}
                      <div className="flex justify-between font-bold text-sm text-zinc-900 pt-2 border-t border-zinc-200">
                        <span>Total Tagihan:</span>
                        <span className="text-rose-600">Rp {calculateTotal(customerCart).toLocaleString('id-ID')}</span>
                      </div>
                    </div>
                  </div>

                  <div className="p-4 border-t border-zinc-100 bg-white">
                    <button
                      onClick={handleCustomerCheckout}
                      disabled={!checkoutData.paymentMethod || !customerName.trim() || isSubmittingOrder || customerCart.length === 0}
                      className={`w-full py-3.5 rounded-xl font-bold transition-all shadow-md text-xs sm:text-sm flex items-center justify-center space-x-2 ${
                        checkoutData.paymentMethod && customerName.trim() && !isSubmittingOrder && customerCart.length > 0
                          ? 'bg-rose-600 text-white hover:bg-rose-700'
                          : 'bg-zinc-200 text-zinc-400 cursor-not-allowed'
                      }`}
                    >
                      {isSubmittingOrder ? (
                        <>
                          <span className="w-4 h-4 border-2 border-white/30 border-t-white rounded-full animate-spin inline-block"></span>
                          <span>Sedang Mengirim Pesanan...</span>
                        </>
                      ) : (
                        <span>Kirim Pesanan ke Restoran (Rp {calculateTotal(customerCart).toLocaleString('id-ID')})</span>
                      )}
                    </button>
                  </div>
                </div>
              </div>
            )}
          </div>
      )}

      {/* ADMIN & POS VIEW */}
      {view === 'admin' && (
        <div className="min-h-screen flex flex-col md:flex-row">
          {/* Sidebar */}
          <aside className="w-full md:w-56 bg-zinc-900 text-zinc-300 flex flex-col shadow-xl z-20 md:h-screen">
            <div className="p-4 border-b border-zinc-800">
              <h2 className="text-base font-bold text-white truncate">{settings.restoName}</h2>
              <div className="flex items-center space-x-1 mt-1">
                <span className="w-2 h-2 rounded-full bg-emerald-400"></span>
                <span className="text-[10px] text-zinc-400 uppercase font-mono tracking-wider">
                  {currentUser?.role === 'admin' ? 'Administrator' : 'Kasir Shift'}
                </span>
              </div>
            </div>

            <nav className="flex-1 p-2 space-y-1 flex flex-row md:flex-col overflow-x-auto">
              {[
                { id: 'pos', name: 'Kasir (POS)', icon: '🛒', adminOnly: false },
                { id: 'orders', name: 'Pesanan Masuk', icon: '📋', adminOnly: false },
                { id: 'analytics', name: 'Analitik', icon: '📊', adminOnly: false },
                { id: 'menu', name: 'Katalog Menu', icon: '🍽️', adminOnly: true },
                { id: 'tables', name: 'QR Meja', icon: '📱', adminOnly: true },
                { id: 'staff', name: 'Manajemen Staff', icon: '👥', adminOnly: true },
                { id: 'settings', name: 'Pengaturan Resto', icon: '⚙️', adminOnly: true }
              ].map((tab) => {
                const isLocked = tab.adminOnly && currentUser?.role !== 'admin';
                const unverifiedCount = orders.filter((o) => o.status === 'unverified').length;
                return (
                  <button
                    key={tab.id}
                    onClick={() => {
                      if (isLocked) {
                        showAlert('Akses Terbatas', 'Hanya administrator yang dapat membuka menu ini.');
                      } else {
                        setAdminTab(tab.id as typeof adminTab);
                      }
                    }}
                    className={`flex items-center space-x-2.5 px-3 py-2.5 rounded-xl transition-all w-full text-left flex-shrink-0 ${
                      adminTab === tab.id
                        ? 'bg-rose-600 text-white font-bold shadow-md'
                        : isLocked
                        ? 'opacity-40 hover:bg-zinc-800'
                        : 'hover:bg-zinc-800 hover:text-white'
                    }`}
                  >
                    <span>{tab.icon}</span>
                    <span className="text-xs flex-1 truncate">{tab.name}</span>
                    {tab.id === 'orders' && unverifiedCount > 0 && (
                      <span className="bg-amber-500 text-zinc-900 text-[10px] font-black px-1.5 py-0.5 rounded-full">
                        {unverifiedCount}
                      </span>
                    )}
                  </button>
                );
              })}
            </nav>

            <div className="p-3 border-t border-zinc-800">
              <button
                onClick={handleLogout}
                className="flex items-center space-x-2 w-full px-3 py-2 rounded-xl text-zinc-400 hover:text-white hover:bg-zinc-800 text-xs font-semibold"
              >
                <span>🚪</span>
                <span>Keluar ({currentUser?.name})</span>
              </button>
            </div>
          </aside>

          {/* Main Content Area */}
          <main className="flex-1 flex flex-col h-screen overflow-hidden">
            <header className="bg-white px-5 py-3 border-b border-zinc-200 flex items-center justify-between shadow-sm flex-shrink-0">
              <div>
                <h2 className="text-base font-bold text-zinc-900 capitalize">
                  {adminTab === 'pos' ? 'Mesin Kasir (POS)' :
                   adminTab === 'orders' ? 'Pesanan Masuk & Verifikasi Dapur' :
                   adminTab === 'analytics' ? 'Analitik Penjualan & Grafik Performa' :
                   adminTab === 'menu' ? 'Katalog Menu & Impor Excel' :
                   adminTab === 'tables' ? 'Manajemen QR Meja Pelanggan' :
                   adminTab === 'staff' ? 'Manajemen Staff & Akses PIN' : 'Pengaturan Resto & Struk'}
                </h2>
                <span className="text-[10px] text-emerald-600 font-semibold">● Terhubung Langsung ke Firestore Cloud</span>
              </div>
              <div className="flex items-center space-x-2">
                <span className="text-xs font-bold text-zinc-700 hidden sm:block">{currentUser?.name}</span>
                <span className="w-8 h-8 rounded-full bg-rose-100 text-rose-600 font-bold flex items-center justify-center text-xs">
                  {currentUser?.name.substring(0, 2).toUpperCase()}
                </span>
              </div>
            </header>

            <div className={`flex-1 p-4 ${adminTab === 'pos' ? 'overflow-hidden flex flex-col' : 'overflow-y-auto'}`}>
              {/* POS TAB */}
              {adminTab === 'pos' && (
                <div className="flex flex-col lg:flex-row gap-4 flex-1 h-[calc(100vh-140px)] min-h-0">
                  {/* Left: Menu Items Selector */}
                  <div className="flex-1 flex flex-col bg-white rounded-2xl border border-zinc-200 shadow-sm overflow-hidden min-h-0">
                    <div className="p-2.5 border-b border-zinc-200 bg-zinc-50 flex space-x-2 overflow-x-auto flex-shrink-0">
                      {categories.map((cat) => (
                        <button
                          key={cat.id}
                          onClick={() => setPosActiveCat(cat.id)}
                          className={`px-3 py-1.5 rounded-xl text-xs font-bold whitespace-nowrap transition-all ${posActiveCat === cat.id ? 'bg-zinc-900 text-white shadow' : 'bg-white text-zinc-600 border border-zinc-200 hover:bg-zinc-100'}`}
                        >
                          {cat.name}
                        </button>
                      ))}
                    </div>

                    <div className="flex-1 overflow-y-auto p-3 grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 xl:grid-cols-5 gap-3 content-start">
                      {menuItems
                        .filter((item) => posActiveCat === 'Semua' || item.category === posActiveCat)
                        .map((item) => (
                          <div
                            key={item.id}
                            onClick={() => addToPosCart(item)}
                            className="bg-white border border-zinc-200 rounded-xl p-2 cursor-pointer hover:border-rose-500 hover:shadow-md transition-all flex flex-col group"
                          >
                            <div className="aspect-square rounded-lg overflow-hidden mb-1.5 bg-zinc-100">
                              <img
                                src={item.image}
                                alt={item.name}
                                className="w-full h-full object-cover group-hover:scale-105 transition-transform"
                                onError={(e) => { (e.target as HTMLImageElement).src = getFallbackImage(item.name); }}
                              />
                            </div>
                            <h3 className="font-bold text-xs text-zinc-900 line-clamp-1">{item.name}</h3>
                            <p className="text-rose-600 font-bold text-xs mt-auto">Rp {item.price.toLocaleString('id-ID')}</p>
                          </div>
                        ))}
                    </div>
                  </div>

                  {/* Right: Cart & Checkout Form */}
                  <div className="w-full lg:w-96 flex flex-col bg-white rounded-2xl border border-zinc-200 shadow-sm overflow-hidden flex-shrink-0 h-full max-h-[50vh] lg:max-h-full">
                    <div className="p-3 border-b border-zinc-200 bg-zinc-50 flex justify-between items-center">
                      <h3 className="font-bold text-xs sm:text-sm text-zinc-900">Keranjang Kasir</h3>
                      {posCart.length > 0 && (
                        <button
                          onClick={() => {
                            setPosCart([]);
                            setPosCashGiven('');
                          }}
                          className="text-red-500 text-xs font-bold hover:underline"
                        >
                          Kosongkan
                        </button>
                      )}
                    </div>

                    <div className="flex-1 overflow-y-auto p-3 min-h-0 space-y-2">
                      {posCart.length === 0 ? (
                        <div className="h-full flex flex-col items-center justify-center text-zinc-400">
                          <p className="text-xs">Klik item menu di kiri untuk menambah transaksi</p>
                        </div>
                      ) : (
                        posCart.map((item) => (
                          <div key={item.id} className="flex gap-2 bg-zinc-50 p-2 rounded-xl border border-zinc-200 items-center">
                            <div className="flex-1 min-w-0">
                              <h4 className="text-xs font-bold text-zinc-900 truncate">{item.name}</h4>
                              <p className="text-rose-600 font-bold text-[11px]">Rp {(item.price * item.quantity).toLocaleString('id-ID')}</p>
                            </div>
                            <div className="flex items-center space-x-1 bg-white rounded-lg p-0.5 border border-zinc-200">
                              <button onClick={() => updatePosQuantity(item.id, -1)} className="w-5 h-5 flex items-center justify-center text-xs font-bold text-zinc-600">-</button>
                              <span className="text-xs font-bold w-4 text-center">{item.quantity}</span>
                              <button onClick={() => updatePosQuantity(item.id, 1)} className="w-5 h-5 flex items-center justify-center text-xs font-bold text-zinc-600">+</button>
                            </div>
                          </div>
                        ))
                      )}
                    </div>

                    {/* Totals & Payment Inputs */}
                    <div className="p-3 bg-zinc-50 border-t border-zinc-200 space-y-2">
                      {posCart.length > 0 && (
                        <div className="space-y-1.5">
                          <div className="space-y-1 text-xs">
                            <div className="flex justify-between text-zinc-500">
                              <span>Subtotal</span>
                              <span className="font-semibold text-zinc-800">Rp {calculateSubtotal(posCart).toLocaleString('id-ID')}</span>
                            </div>
                            {settings.taxRate > 0 && (
                              <div className="flex justify-between text-zinc-500">
                                <span>Pajak ({settings.taxRate}%)</span>
                                <span className="font-semibold text-zinc-800">Rp {calculateTax(calculateSubtotal(posCart)).toLocaleString('id-ID')}</span>
                              </div>
                            )}
                            {settings.serviceCharge > 0 && (
                              <div className="flex justify-between text-zinc-500">
                                <span>Service ({settings.serviceCharge}%)</span>
                                <span className="font-semibold text-zinc-800">Rp {calculateService(calculateSubtotal(posCart)).toLocaleString('id-ID')}</span>
                              </div>
                            )}
                            <div className="flex justify-between font-bold text-sm text-zinc-900 pt-1 border-t border-zinc-200">
                              <span>Total</span>
                              <span className="text-rose-600">Rp {calculateTotal(posCart).toLocaleString('id-ID')}</span>
                            </div>
                          </div>

                          <input
                            type="text"
                            placeholder="Nomor Meja / Takeaway"
                            value={posTableNo}
                            onChange={(e) => setPosTableNo(e.target.value)}
                            className="w-full px-2.5 py-1.5 rounded-lg border border-zinc-300 text-xs font-bold focus:border-rose-500 outline-none"
                          />
                          <input
                            type="text"
                            placeholder="Nama Pemesan *"
                            value={posCustomerName}
                            onChange={(e) => setPosCustomerName(e.target.value)}
                            className="w-full px-2.5 py-1.5 rounded-lg border border-zinc-300 text-xs font-bold focus:border-rose-500 outline-none"
                          />
                          <div className="flex gap-2">
                            <select
                              value={posOrderType}
                              onChange={(e) => setPosOrderType(e.target.value as 'dine-in' | 'takeaway')}
                              className="w-1/2 px-2 py-1.5 rounded-lg border border-zinc-300 text-xs font-bold bg-white"
                            >
                              <option value="dine-in">Dine In (Makan di Tempat)</option>
                              <option value="takeaway">Takeaway (Bungkus)</option>
                            </select>
                            <select
                              value={posPaymentMethod}
                              onChange={(e) => {
                                setPosPaymentMethod(e.target.value);
                                setPosCashGiven('');
                              }}
                              className="w-1/2 px-2 py-1.5 rounded-lg border border-zinc-300 text-xs font-bold bg-white"
                            >
                              <option value="">-- Metode Bayar --</option>
                              {paymentMethods
                                .filter((p) => p.isActive)
                                .map((pm) => (
                                  <option key={pm.id} value={pm.id}>{pm.name}</option>
                                ))}
                            </select>
                          </div>

                          {posPaymentMethod && paymentMethods.find((p) => p.id === posPaymentMethod)?.type === 'cash' && (
                            <div className="p-2 bg-rose-50 border border-rose-200 rounded-xl space-y-1">
                              <label className="text-[10px] font-bold text-rose-800 uppercase block">Uang Tunai Diterima (Rp)</label>
                              <input
                                type="text"
                                placeholder="Contoh: 100.000"
                                value={posCashGiven}
                                onChange={(e) => setPosCashGiven(formatCurrencyInput(e.target.value))}
                                className="w-full px-2 py-1 bg-white rounded border border-rose-300 text-xs font-bold text-zinc-900"
                              />
                              <div className="flex justify-between items-center text-xs font-bold text-rose-900">
                                <span>Kembalian:</span>
                                <span>
                                  Rp {Math.max(0, parseIndonesianNumber(posCashGiven) - calculateTotal(posCart)).toLocaleString('id-ID')}
                                </span>
                              </div>
                            </div>
                          )}
                        </div>
                      )}

                      <button
                        onClick={handlePosCheckout}
                        disabled={posCart.length === 0 || !posPaymentMethod || !posCustomerName.trim() || isSubmittingPos}
                        className={`w-full py-2.5 rounded-xl font-bold transition-all shadow-md text-xs sm:text-sm flex items-center justify-center space-x-2 ${
                          posCart.length > 0 && posPaymentMethod && posCustomerName.trim() && !isSubmittingPos
                            ? 'bg-rose-600 text-white hover:bg-rose-700'
                            : 'bg-zinc-200 text-zinc-400 cursor-not-allowed'
                        }`}
                      >
                        {isSubmittingPos ? (
                          <>
                            <span className="w-3.5 h-3.5 border-2 border-white/30 border-t-white rounded-full animate-spin inline-block"></span>
                            <span>Menyimpan Transaksi...</span>
                          </>
                        ) : (
                          <span>Bayar & Cetak Struk (Firebase Sync)</span>
                        )}
                      </button>
                    </div>
                  </div>
                </div>
              )}

              {/* ORDERS TAB */}
              {adminTab === 'orders' && (() => {
                const isAdmin = currentUser?.role === 'admin';

                return (
                  <div className="space-y-4">
                    <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3 bg-white p-4 rounded-2xl shadow-sm border border-zinc-200">
                      <div>
                        <div className="flex items-center space-x-2">
                          <h3 className="font-bold text-sm text-zinc-900">Pesanan Masuk & Transaksi Realtime</h3>
                          <span className="text-[10px] font-bold px-2 py-0.5 bg-rose-50 text-rose-700 border border-rose-200 rounded-full">
                            {orders.length} Total
                          </span>
                        </div>
                        <p className="text-xs text-zinc-500 mt-0.5">
                          {isAdmin
                            ? 'Administrator: Kelola verifikasi dapur, pembatalan pesanan, hapus transaksi, atau unduh laporan Excel.'
                            : 'Kasir Shift: Verifikasi pesanan dapur dan selesaikan transaksi kasir. Pembatalan & hapus transaksi hanya dapat dilakukan oleh Admin.'}
                        </p>
                      </div>

                      <div className="flex flex-wrap items-center gap-2 w-full sm:w-auto">
                        <button
                          onClick={exportOrdersToXLSX}
                          className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold px-3.5 py-2 rounded-xl text-xs transition-all flex items-center justify-center space-x-1.5 shadow-sm"
                          title="Ekspor seluruh daftar transaksi ke format Excel (.xlsx)"
                        >
                          <span>📥</span>
                          <span>Export XLSX</span>
                        </button>
                      </div>
                    </div>

                    {/* Status Filters */}
                    <div className="flex space-x-1.5 overflow-x-auto pb-1">
                      {[
                        { id: 'all', label: 'Semua Status', count: orders.length },
                        { id: 'unverified', label: '🔴 Menunggu Verifikasi', count: orders.filter((o) => o.status === 'unverified').length },
                        { id: 'processing', label: '🔵 Sedang Dimasak', count: orders.filter((o) => o.status === 'processing').length },
                        { id: 'completed', label: '🟢 Selesai', count: orders.filter((o) => o.status === 'completed').length },
                        { id: 'cancelled', label: '⛔ Dibatalkan', count: orders.filter((o) => o.status === 'cancelled').length }
                      ].map((tab) => (
                        <button
                          key={tab.id}
                          onClick={() => setOrderFilterStatus(tab.id)}
                          className={`px-3 py-1.5 rounded-xl text-xs font-bold whitespace-nowrap transition-all flex items-center space-x-1.5 ${
                            orderFilterStatus === tab.id
                              ? 'bg-zinc-900 text-white shadow-sm'
                              : 'bg-white text-zinc-600 border border-zinc-200 hover:bg-zinc-100'
                          }`}
                        >
                          <span>{tab.label}</span>
                          <span className={`text-[10px] px-1.5 py-0.2 rounded-full ${orderFilterStatus === tab.id ? 'bg-white/20 text-white' : 'bg-zinc-100 text-zinc-600'}`}>
                            {tab.count}
                          </span>
                        </button>
                      ))}
                    </div>

                    {/* Grouped Orders Table */}
                    {[
                      { key: 'unverified', title: '🔴 Menunggu Verifikasi Kasir (Scan QR Pelanggan)', badgeBg: 'bg-amber-100 text-amber-800' },
                      { key: 'processing', title: '🔵 Sedang Disiapkan / Dimasak di Dapur', badgeBg: 'bg-blue-100 text-blue-800' },
                      { key: 'completed', title: '🟢 Selesai / Siap Diambil', badgeBg: 'bg-emerald-100 text-emerald-800' },
                      { key: 'cancelled', title: '⛔ Pesanan Dibatalkan', badgeBg: 'bg-red-100 text-red-800' }
                    ]
                      .filter((grp) => orderFilterStatus === 'all' || orderFilterStatus === grp.key)
                      .map((grp) => {
                        const list = orders.filter((o) => o.status === grp.key);
                        if (list.length === 0) return null;
                        return (
                          <div key={grp.key} className="bg-white rounded-2xl border border-zinc-200 shadow-sm overflow-hidden">
                            <div className="p-3.5 bg-zinc-50 border-b border-zinc-200 flex justify-between items-center">
                              <h4 className="font-bold text-xs sm:text-sm text-zinc-900">{grp.title}</h4>
                              <span className={`text-[11px] font-bold px-2 py-0.5 rounded-full ${grp.badgeBg}`}>{list.length}</span>
                            </div>

                            <div className="overflow-x-auto">
                              <table className="w-full text-left border-collapse text-xs min-w-[700px]">
                                <thead>
                                  <tr className="border-b border-zinc-200 text-[11px] font-bold text-zinc-500 uppercase bg-zinc-50/50">
                                    <th className="p-3">ID & Waktu</th>
                                    <th className="p-3">Pemesan & Meja</th>
                                    <th className="p-3">Detail Menu</th>
                                    <th className="p-3">Total & Bayar</th>
                                    <th className="p-3 text-center">Aksi {isAdmin ? '& Kelola Transaksi' : 'Status'}</th>
                                  </tr>
                                </thead>
                                <tbody className="divide-y divide-zinc-100">
                                  {list.map((order) => (
                                    <tr key={order.id} className="hover:bg-zinc-50 transition-colors">
                                      <td className="p-3">
                                        <span className="font-mono font-bold text-zinc-900 block">{order.id}</span>
                                        <span className="text-[10px] text-zinc-400">
                                          {new Date(order.time).toLocaleTimeString('id-ID')} • {order.source}
                                        </span>
                                      </td>
                                      <td className="p-3">
                                        <span className="font-bold text-zinc-900 block">{order.customerName}</span>
                                        <span className="text-zinc-500 text-[11px]">{order.table} ({order.orderType === 'takeaway' ? 'Bungkus' : 'Dine In'})</span>
                                      </td>
                                      <td className="p-3">
                                        <div className="space-y-0.5 max-w-xs">
                                          {order.items.map((it, idx) => (
                                            <div key={idx} className="text-zinc-700">
                                              <b>{it.quantity}x</b> {it.name}
                                            </div>
                                          ))}
                                          {order.notes && <div className="text-[10px] text-amber-800 bg-amber-50 p-1 rounded">Catatan: {order.notes}</div>}
                                        </div>
                                      </td>
                                      <td className="p-3">
                                        <span className="font-bold text-rose-600 block">Rp {order.total.toLocaleString('id-ID')}</span>
                                        <span className="text-[10px] text-zinc-500">
                                          {paymentMethods.find((p) => p.id === order.paymentMethod)?.name || order.paymentMethod}
                                        </span>
                                      </td>
                                      <td className="p-3">
                                        <div className="flex flex-wrap gap-1.5 items-center justify-center">
                                          {order.status === 'unverified' && (
                                            <>
                                              <button
                                                onClick={() => updateOrderStatus(order.id, 'processing')}
                                                className="px-2.5 py-1 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg font-bold text-[11px] shadow-sm transition-colors"
                                                title="Verifikasi dan teruskan ke dapur"
                                              >
                                                ✓ Terima
                                              </button>
                                              {/* Batal Pesanan: Khusus Admin */}
                                              {isAdmin && (
                                                <button
                                                  onClick={() => handleCancelOrder(order.id)}
                                                  className="px-2.5 py-1 bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 rounded-lg font-bold text-[11px] transition-colors"
                                                  title="Batalkan pesanan ini (Khusus Admin)"
                                                >
                                                  ⛔ Batalkan
                                                </button>
                                              )}
                                            </>
                                          )}

                                          {order.status === 'processing' && (
                                            <>
                                              <button
                                                onClick={() => {
                                                  updateOrderStatus(order.id, 'completed');
                                                  setReceiptOrder({ ...order, status: 'completed' });
                                                }}
                                                className="px-2.5 py-1 bg-blue-600 hover:bg-blue-700 text-white rounded-lg font-bold text-[11px] shadow-sm transition-colors"
                                                title="Tandai pesanan selesai dan cetak struk"
                                              >
                                                ✓ Selesai
                                              </button>
                                              {/* Batal Pesanan: Khusus Admin */}
                                              {isAdmin && (
                                                <button
                                                  onClick={() => handleCancelOrder(order.id)}
                                                  className="px-2.5 py-1 bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 rounded-lg font-bold text-[11px] transition-colors"
                                                  title="Batalkan pesanan ini (Khusus Admin)"
                                                >
                                                  ⛔ Batalkan
                                                </button>
                                              )}
                                            </>
                                          )}

                                          {order.status === 'completed' && (
                                            <>
                                              <button
                                                onClick={() => {
                                                  announceOrderCall(
                                                    order.customerName || 'Pelanggan',
                                                    order.id,
                                                    settings.soundEnabled,
                                                    settings.voiceType || 'female',
                                                    settings.callTemplate,
                                                    order.table || 'Meja'
                                                  );
                                                }}
                                                className="px-2.5 py-1 bg-amber-500 hover:bg-amber-600 text-white rounded-lg font-bold text-[11px] shadow-sm flex items-center space-x-1"
                                                title="Panggil nama pelanggan lewat speaker"
                                              >
                                                <span>📢</span>
                                                <span>Panggil</span>
                                              </button>
                                              <button
                                                onClick={() => setReceiptOrder(order)}
                                                className="px-2 py-1 bg-zinc-100 hover:bg-zinc-200 text-zinc-700 rounded-lg text-[11px] font-semibold"
                                                title="Lihat / cetak struk"
                                              >
                                                Struk
                                              </button>
                                              {/* Batal Pesanan: Khusus Admin */}
                                              {isAdmin && (
                                                <button
                                                  onClick={() => handleCancelOrder(order.id)}
                                                  className="px-2 py-1 bg-zinc-50 hover:bg-zinc-100 text-zinc-600 border border-zinc-200 rounded-lg text-[11px] font-medium"
                                                  title="Ubah status menjadi dibatalkan (Khusus Admin)"
                                                >
                                                  Batal
                                                </button>
                                              )}
                                            </>
                                          )}

                                          {order.status === 'cancelled' && (
                                            isAdmin ? (
                                              <button
                                                onClick={() => updateOrderStatus(order.id, 'unverified')}
                                                className="px-2.5 py-1 bg-amber-50 hover:bg-amber-100 text-amber-800 border border-amber-200 rounded-lg font-bold text-[11px] transition-colors"
                                                title="Pulihkan pesanan ini ke daftar menunggu (Khusus Admin)"
                                              >
                                                ↺ Pulihkan
                                              </button>
                                            ) : (
                                              <span className="text-[11px] text-zinc-400 font-medium italic">
                                                Dibatalkan
                                              </span>
                                            )
                                          )}

                                          {/* Hapus Transaksi Permanen: Khusus Admin */}
                                          {isAdmin && (
                                            <button
                                              onClick={() => handleDeleteOrder(order.id)}
                                              className="px-2 py-1 bg-red-50 hover:bg-red-100 text-red-600 rounded-lg text-[11px] font-bold border border-red-200 transition-colors"
                                              title="Hapus transaksi ini secara permanen dari database (Khusus Admin)"
                                            >
                                              🗑️ Hapus
                                            </button>
                                          )}
                                        </div>
                                      </td>
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          </div>
                        );
                      })}

                    {orders.length === 0 && (
                      <div className="bg-white rounded-2xl p-8 border border-zinc-200 text-center text-zinc-400 space-y-2">
                        <span className="text-3xl block">📋</span>
                        <p className="text-xs font-semibold">Belum ada pesanan yang masuk.</p>
                        <p className="text-[11px] text-zinc-400">Pesanan dari kasir POS atau pemesanan mandiri QR pelanggan akan muncul di sini secara realtime.</p>
                      </div>
                    )}
                  </div>
                );
              })()}

              {/* ANALYTICS TAB */}
              {adminTab === 'analytics' && (() => {
                const now = new Date();
                const filteredOrders = orders.filter((o) => {
                  if (analyticsDays === 0) return true;
                  const orderDate = new Date(o.time);
                  const diffDays = (now.getTime() - orderDate.getTime()) / (1000 * 60 * 60 * 24);
                  return diffDays <= analyticsDays;
                });

                const completedOrders = filteredOrders.filter((o) => o.status === 'completed');
                const cancelledOrders = filteredOrders.filter((o) => o.status === 'cancelled');
                const unverifiedOrders = filteredOrders.filter((o) => o.status === 'unverified');
                const processingOrders = filteredOrders.filter((o) => o.status === 'processing');

                const totalRevenue = completedOrders.reduce((sum, o) => sum + o.total, 0);
                const totalCompleted = completedOrders.length;
                const aov = totalCompleted > 0 ? Math.round(totalRevenue / totalCompleted) : 0;

                // Daily revenue timeline aggregation
                const daysToGenerate = analyticsDays === 0 ? 14 : analyticsDays;
                const dailyRevenueMap: { [key: string]: { date: string; fullDate: string; revenue: number; count: number } } = {};

                for (let i = daysToGenerate - 1; i >= 0; i--) {
                  const d = new Date();
                  d.setDate(d.getDate() - i);
                  const key = d.toISOString().slice(0, 10);
                  const label = d.toLocaleDateString('id-ID', { day: 'numeric', month: 'short' });
                  dailyRevenueMap[key] = { date: label, fullDate: key, revenue: 0, count: 0 };
                }

                completedOrders.forEach((o) => {
                  const key = new Date(o.time).toISOString().slice(0, 10);
                  if (dailyRevenueMap[key]) {
                    dailyRevenueMap[key].revenue += o.total;
                    dailyRevenueMap[key].count += 1;
                  } else if (analyticsDays === 0) {
                    const label = new Date(o.time).toLocaleDateString('id-ID', { day: 'numeric', month: 'short' });
                    dailyRevenueMap[key] = { date: label, fullDate: key, revenue: o.total, count: 1 };
                  }
                });

                const dailyChartData = Object.values(dailyRevenueMap);

                // Top Selling Items aggregation
                const itemSalesMap: {
                  [name: string]: { name: string; category: string; quantity: number; revenue: number; price: number };
                } = {};

                completedOrders.forEach((o) => {
                  o.items.forEach((it) => {
                    if (!itemSalesMap[it.name]) {
                      itemSalesMap[it.name] = {
                        name: it.name,
                        category: it.category || 'Menu',
                        quantity: 0,
                        revenue: 0,
                        price: it.price
                      };
                    }
                    itemSalesMap[it.name].quantity += it.quantity;
                    itemSalesMap[it.name].revenue += it.price * it.quantity;
                  });
                });

                const topMenuItems = Object.values(itemSalesMap).sort((a, b) => b.quantity - a.quantity);
                const top1Dish = topMenuItems[0];

                const barChartData = topMenuItems.slice(0, 7).map((item) => ({
                  name: item.name.length > 14 ? item.name.substring(0, 14) + '…' : item.name,
                  fullName: item.name,
                  porsi: item.quantity,
                  omset: item.revenue
                }));

                // Payment Breakdown
                const paymentSummary: { [key: string]: { name: string; count: number; total: number } } = {};
                completedOrders.forEach((o) => {
                  const pm = paymentMethods.find((p) => p.id === o.paymentMethod);
                  const pmName = pm ? pm.name : (o.paymentMethod || 'Tunai');
                  if (!paymentSummary[pmName]) {
                    paymentSummary[pmName] = { name: pmName, count: 0, total: 0 };
                  }
                  paymentSummary[pmName].count += 1;
                  paymentSummary[pmName].total += o.total;
                });

                const dineInCount = completedOrders.filter((o) => o.orderType === 'dine-in').length;
                const takeawayCount = completedOrders.filter((o) => o.orderType === 'takeaway').length;

                return (
                  <div className="space-y-5 pb-6">
                    {/* Filter Period Header */}
                    <div className="bg-white p-4 rounded-2xl shadow-sm border border-zinc-200 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3">
                      <div>
                        <h3 className="font-bold text-sm text-zinc-900 flex items-center gap-2">
                          <span>📊</span>
                          <span>Analitik Penjualan & Performa Resto</span>
                        </h3>
                        <p className="text-xs text-zinc-500 mt-0.5">
                          Visualisasi pendapatan harian dan menu terlaris secara langsung dari transaksi kasir & QR.
                        </p>
                      </div>

                      <div className="flex items-center gap-1.5 bg-zinc-100 p-1 rounded-xl">
                        {[
                          { label: '7 Hari', days: 7 },
                          { label: '30 Hari', days: 30 },
                          { label: 'Semua Waktu', days: 0 }
                        ].map((btn) => (
                          <button
                            key={btn.days}
                            onClick={() => setAnalyticsDays(btn.days)}
                            className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${
                              analyticsDays === btn.days
                                ? 'bg-white text-zinc-900 shadow-sm'
                                : 'text-zinc-600 hover:text-zinc-900'
                            }`}
                          >
                            {btn.label}
                          </button>
                        ))}
                      </div>
                    </div>

                    {/* KPI Cards Grid */}
                    <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3.5">
                      <div className="bg-white p-4 rounded-2xl border border-zinc-200 shadow-sm relative overflow-hidden">
                        <div className="w-1.5 h-full bg-rose-500 absolute top-0 left-0" />
                        <span className="text-zinc-500 text-xs font-semibold block">Total Pendapatan (Omset)</span>
                        <h4 className="text-lg sm:text-xl font-black text-rose-600 mt-1">
                          Rp {totalRevenue.toLocaleString('id-ID')}
                        </h4>
                        <span className="text-[11px] text-zinc-400 mt-1 block">
                          Dari {totalCompleted} transaksi sukses
                        </span>
                      </div>

                      <div className="bg-white p-4 rounded-2xl border border-zinc-200 shadow-sm relative overflow-hidden">
                        <div className="w-1.5 h-full bg-emerald-500 absolute top-0 left-0" />
                        <span className="text-zinc-500 text-xs font-semibold block">Rata-Rata Transaksi (AOV)</span>
                        <h4 className="text-lg sm:text-xl font-black text-zinc-900 mt-1">
                          Rp {aov.toLocaleString('id-ID')}
                        </h4>
                        <span className="text-[11px] text-emerald-600 font-semibold mt-1 block">
                          Nilai rata-rata per struk
                        </span>
                      </div>

                      <div className="bg-white p-4 rounded-2xl border border-zinc-200 shadow-sm relative overflow-hidden">
                        <div className="w-1.5 h-full bg-blue-500 absolute top-0 left-0" />
                        <span className="text-zinc-500 text-xs font-semibold block">Menu Paling Laris #1</span>
                        <h4 className="text-base font-bold text-zinc-900 mt-1 truncate" title={top1Dish ? top1Dish.name : '-'}>
                          {top1Dish ? top1Dish.name : 'Belum Ada Data'}
                        </h4>
                        <span className="text-[11px] text-blue-600 font-semibold mt-1 block">
                          {top1Dish ? `${top1Dish.quantity} porsi terjual (Rp ${top1Dish.revenue.toLocaleString('id-ID')})` : '0 porsi'}
                        </span>
                      </div>

                      <div className="bg-white p-4 rounded-2xl border border-zinc-200 shadow-sm relative overflow-hidden">
                        <div className="w-1.5 h-full bg-amber-500 absolute top-0 left-0" />
                        <span className="text-zinc-500 text-xs font-semibold block">Aktivitas Pesanan</span>
                        <div className="flex items-baseline space-x-2 mt-1">
                          <span className="text-lg font-black text-zinc-900">{filteredOrders.length}</span>
                          <span className="text-xs text-zinc-500">Total Masuk</span>
                        </div>
                        <div className="flex items-center space-x-2 text-[10px] text-zinc-500 mt-1">
                          <span className="text-emerald-600 font-bold">{totalCompleted} Selesai</span>
                          <span>•</span>
                          <span className="text-amber-600 font-bold">{unverifiedOrders.length + processingOrders.length} Proses</span>
                          <span>•</span>
                          <span className="text-rose-600 font-bold">{cancelledOrders.length} Batal</span>
                        </div>
                      </div>
                    </div>

                    {/* CHARTS GRID */}
                    <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
                      {/* Grafik Pendapatan Harian */}
                      <div className="bg-white p-4 rounded-2xl border border-zinc-200 shadow-sm flex flex-col">
                        <div className="flex justify-between items-center mb-3">
                          <div>
                            <h4 className="font-bold text-sm text-zinc-900">Grafik Pendapatan Harian</h4>
                            <p className="text-[11px] text-zinc-500">Tren penjualan selesai per tanggal</p>
                          </div>
                          <span className="text-xs font-bold text-rose-600 bg-rose-50 px-2 py-0.5 rounded-lg border border-rose-200">
                            Rp {totalRevenue.toLocaleString('id-ID')}
                          </span>
                        </div>

                        <div className="w-full h-72 min-h-[280px]">
                          <ResponsiveContainer width="100%" height="100%">
                            <AreaChart data={dailyChartData} margin={{ top: 10, right: 10, left: -15, bottom: 0 }}>
                              <defs>
                                <linearGradient id="colorRevenue" x1="0" y1="0" x2="0" y2="1">
                                  <stop offset="5%" stopColor="#e11d48" stopOpacity={0.3} />
                                  <stop offset="95%" stopColor="#e11d48" stopOpacity={0.0} />
                                </linearGradient>
                              </defs>
                              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                              <XAxis dataKey="date" stroke="#94a3b8" fontSize={11} tickLine={false} />
                              <YAxis
                                stroke="#94a3b8"
                                fontSize={11}
                                tickLine={false}
                                axisLine={false}
                                tickFormatter={(v) =>
                                  v >= 1000000 ? `${(v / 1000000).toFixed(1)}jt` : v >= 1000 ? `${(v / 1000).toFixed(0)}k` : `${v}`
                                }
                              />
                              <Tooltip
                                content={({ active, payload, label }) => {
                                  if (active && payload && payload.length) {
                                    const val = payload[0].value as number;
                                    const count = payload[0].payload.count as number;
                                    return (
                                      <div className="bg-zinc-900 text-white p-2.5 rounded-xl text-xs shadow-xl border border-zinc-700">
                                        <p className="font-bold text-zinc-300">{label}</p>
                                        <p className="font-black text-rose-400 text-sm mt-0.5">Rp {val.toLocaleString('id-ID')}</p>
                                        <p className="text-[10px] text-zinc-400 mt-0.5">{count} Transaksi Selesai</p>
                                      </div>
                                    );
                                  }
                                  return null;
                                }}
                              />
                              <Area
                                type="monotone"
                                dataKey="revenue"
                                stroke="#e11d48"
                                strokeWidth={2.5}
                                fillOpacity={1}
                                fill="url(#colorRevenue)"
                              />
                            </AreaChart>
                          </ResponsiveContainer>
                        </div>
                      </div>

                      {/* Grafik Menu Terlaris */}
                      <div className="bg-white p-4 rounded-2xl border border-zinc-200 shadow-sm flex flex-col">
                        <div className="flex justify-between items-center mb-3">
                          <div>
                            <h4 className="font-bold text-sm text-zinc-900">Grafik Menu Terlaris (Porsi)</h4>
                            <p className="text-[11px] text-zinc-500">Top 7 hidangan yang paling banyak dipesan</p>
                          </div>
                          <span className="text-xs font-bold text-zinc-600 bg-zinc-100 px-2 py-0.5 rounded-lg border border-zinc-200">
                            {topMenuItems.length} Menu Terjual
                          </span>
                        </div>

                        {barChartData.length === 0 ? (
                          <div className="flex-1 flex flex-col items-center justify-center text-zinc-400 py-10">
                            <span className="text-3xl mb-1">🍽️</span>
                            <p className="text-xs">Belum ada menu yang terjual pada periode ini.</p>
                          </div>
                        ) : (
                          <div className="w-full h-72 min-h-[280px]">
                            <ResponsiveContainer width="100%" height="100%">
                              <BarChart data={barChartData} margin={{ top: 10, right: 10, left: -20, bottom: 25 }}>
                                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                                <XAxis
                                  dataKey="name"
                                  stroke="#94a3b8"
                                  fontSize={10}
                                  tickLine={false}
                                  interval={0}
                                  angle={-20}
                                  textAnchor="end"
                                />
                                <YAxis stroke="#94a3b8" fontSize={11} tickLine={false} axisLine={false} allowDecimals={false} />
                                <Tooltip
                                  content={({ active, payload }) => {
                                    if (active && payload && payload.length) {
                                      const data = payload[0].payload;
                                      return (
                                        <div className="bg-zinc-900 text-white p-2.5 rounded-xl text-xs shadow-xl border border-zinc-700">
                                          <p className="font-bold text-zinc-200">{data.fullName}</p>
                                          <p className="font-bold text-amber-400 mt-0.5">{data.porsi} Porsi Terjual</p>
                                          <p className="text-[10px] text-zinc-400">Total Omset: Rp {data.omset.toLocaleString('id-ID')}</p>
                                        </div>
                                      );
                                    }
                                    return null;
                                  }}
                                />
                                <Bar dataKey="porsi" fill="#e11d48" radius={[6, 6, 0, 0]}>
                                  {barChartData.map((_, index) => (
                                    <Cell
                                      key={`cell-${index}`}
                                      fill={index === 0 ? '#e11d48' : index === 1 ? '#f43f5e' : index === 2 ? '#fb7185' : '#fda4af'}
                                    />
                                  ))}
                                </Bar>
                              </BarChart>
                            </ResponsiveContainer>
                          </div>
                        )}
                      </div>
                    </div>

                    {/* Breakdown & Top Menu Ranking Table */}
                    <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
                      {/* Metode Pembayaran & Tipe Pesanan */}
                      <div className="space-y-4">
                        <div className="bg-white p-4 rounded-2xl border border-zinc-200 shadow-sm space-y-3">
                          <h4 className="font-bold text-xs sm:text-sm text-zinc-900">Distribusi Metode Pembayaran</h4>
                          <div className="space-y-2">
                            {Object.values(paymentSummary).length === 0 ? (
                              <p className="text-xs text-zinc-400">Belum ada transaksi.</p>
                            ) : (
                              Object.values(paymentSummary).map((pm) => {
                                const percent = totalRevenue > 0 ? Math.round((pm.total / totalRevenue) * 100) : 0;
                                return (
                                  <div key={pm.name} className="space-y-1">
                                    <div className="flex justify-between text-xs">
                                      <span className="font-bold text-zinc-800">{pm.name}</span>
                                      <span className="text-zinc-500 font-semibold">
                                        Rp {pm.total.toLocaleString('id-ID')} ({percent}%)
                                      </span>
                                    </div>
                                    <div className="w-full bg-zinc-100 rounded-full h-2 overflow-hidden">
                                      <div className="bg-rose-600 h-2 rounded-full" style={{ width: `${percent}%` }} />
                                    </div>
                                  </div>
                                );
                              })
                            )}
                          </div>
                        </div>

                        <div className="bg-white p-4 rounded-2xl border border-zinc-200 shadow-sm space-y-3">
                          <h4 className="font-bold text-xs sm:text-sm text-zinc-900">Tipe Layanan Pesanan</h4>
                          <div className="grid grid-cols-2 gap-2 text-center">
                            <div className="p-3 bg-zinc-50 rounded-xl border border-zinc-200">
                              <span className="text-lg block">🍽️</span>
                              <span className="text-xs font-bold text-zinc-800 block mt-1">Dine In</span>
                              <span className="text-xs text-rose-600 font-extrabold">{dineInCount} Pesanan</span>
                            </div>
                            <div className="p-3 bg-zinc-50 rounded-xl border border-zinc-200">
                              <span className="text-lg block">🛍️</span>
                              <span className="text-xs font-bold text-zinc-800 block mt-1">Takeaway</span>
                              <span className="text-xs text-rose-600 font-extrabold">{takeawayCount} Pesanan</span>
                            </div>
                          </div>
                        </div>
                      </div>

                      {/* Ranking Top Menu Items Table */}
                      <div className="lg:col-span-2 bg-white rounded-2xl border border-zinc-200 shadow-sm overflow-hidden flex flex-col">
                        <div className="p-4 bg-zinc-50 border-b border-zinc-200 flex justify-between items-center">
                          <div>
                            <h4 className="font-bold text-xs sm:text-sm text-zinc-900">Peringkat Menu Paling Laris</h4>
                            <p className="text-[11px] text-zinc-500">Urutan hidangan berdasarkan kuantitas porsi terjual</p>
                          </div>
                          <span className="text-[11px] font-bold px-2.5 py-1 bg-white rounded-lg border border-zinc-200 text-zinc-700">
                            Top 10 Menu
                          </span>
                        </div>

                        <div className="overflow-x-auto flex-1">
                          <table className="w-full text-left border-collapse text-xs">
                            <thead>
                              <tr className="border-b border-zinc-200 text-[11px] font-bold text-zinc-500 uppercase bg-zinc-50/50">
                                <th className="p-3 w-12 text-center">#</th>
                                <th className="p-3">Nama Menu</th>
                                <th className="p-3">Kategori</th>
                                <th className="p-3 text-right">Harga</th>
                                <th className="p-3 text-center">Terjual</th>
                                <th className="p-3 text-right">Total Omset</th>
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-zinc-100">
                              {topMenuItems.length === 0 ? (
                                <tr>
                                  <td colSpan={6} className="p-6 text-center text-zinc-400 text-xs">
                                    Belum ada data transaksi menu pada periode ini.
                                  </td>
                                </tr>
                              ) : (
                                topMenuItems.slice(0, 10).map((item, idx) => (
                                  <tr key={item.name} className="hover:bg-zinc-50 transition-colors">
                                    <td className="p-3 text-center">
                                      <span
                                        className={`inline-flex items-center justify-center w-5 h-5 rounded-full text-[10px] font-black ${
                                          idx === 0
                                            ? 'bg-amber-400 text-amber-950'
                                            : idx === 1
                                            ? 'bg-zinc-300 text-zinc-800'
                                            : idx === 2
                                            ? 'bg-amber-600/30 text-amber-900'
                                            : 'bg-zinc-100 text-zinc-600'
                                        }`}
                                      >
                                        {idx + 1}
                                      </span>
                                    </td>
                                    <td className="p-3 font-bold text-zinc-900">{item.name}</td>
                                    <td className="p-3">
                                      <span className="px-2 py-0.5 rounded-md text-[10px] font-bold uppercase bg-zinc-100 text-zinc-700">
                                        {item.category}
                                      </span>
                                    </td>
                                    <td className="p-3 text-right font-medium text-zinc-600">
                                      Rp {item.price.toLocaleString('id-ID')}
                                    </td>
                                    <td className="p-3 text-center font-bold text-rose-600">
                                      {item.quantity} porsi
                                    </td>
                                    <td className="p-3 text-right font-bold text-zinc-900">
                                      Rp {item.revenue.toLocaleString('id-ID')}
                                    </td>
                                  </tr>
                                ))
                              )}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    </div>
                  </div>
                );
              })()}

              {/* MENU CATALOG TAB */}
              {adminTab === 'menu' && (() => {
                const filteredAdminMenu = menuItems.filter((item) => {
                  const query = adminMenuSearch.toLowerCase().trim();
                  const matchesSearch = !query ||
                    item.name.toLowerCase().includes(query) ||
                    (item.description && item.description.toLowerCase().includes(query));
                  const matchesCat = adminMenuCatFilter === 'Semua' || item.category === adminMenuCatFilter;
                  return matchesSearch && matchesCat;
                });

                return (
                  <div className="space-y-4">
                    {/* Category Management */}
                    <div className="bg-white p-4 rounded-2xl shadow-sm border border-zinc-200 space-y-3">
                      <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-2">
                        <div>
                          <h3 className="font-bold text-sm text-zinc-900">Kategori Menu Restoran</h3>
                          <p className="text-xs text-zinc-500">Kelompokkan hidangan agar mudah ditemukan pelanggan dan kasir.</p>
                        </div>
                        <button
                          onClick={() => setShowCategoryModal(true)}
                          className="bg-zinc-900 hover:bg-zinc-800 text-white font-bold px-3 py-2 rounded-xl text-xs transition-all flex items-center space-x-1"
                        >
                          <span>+ Tambah Kategori</span>
                        </button>
                      </div>
                      <div className="flex flex-wrap gap-2 pt-1">
                        {categories.map((cat) => {
                          const count = cat.id === 'Semua' ? menuItems.length : menuItems.filter((m) => m.category === cat.id).length;
                          return (
                            <div
                              key={cat.id}
                              className={`flex items-center space-x-1.5 px-3 py-1.5 rounded-xl border text-xs font-bold transition-all cursor-pointer ${
                                adminMenuCatFilter === cat.id
                                  ? 'bg-rose-50 border-rose-300 text-rose-700'
                                  : 'bg-zinc-50 border-zinc-200 text-zinc-700 hover:bg-zinc-100'
                              }`}
                              onClick={() => setAdminMenuCatFilter(cat.id)}
                            >
                              <span>{cat.name}</span>
                              <span className="text-[10px] px-1.5 py-0.2 rounded-full bg-white/80 border border-zinc-200 text-zinc-500">
                                {count}
                              </span>
                              {cat.id !== 'Semua' && (
                                <button
                                  type="button"
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    handleDeleteCategory(cat.id);
                                  }}
                                  className="text-red-500 hover:text-red-700 ml-1 font-bold px-1"
                                  title="Hapus Kategori"
                                >
                                  ×
                                </button>
                              )}
                            </div>
                          );
                        })}
                      </div>
                    </div>

                    {/* Menu List Header & Controls */}
                    <div className="bg-white p-4 rounded-2xl shadow-sm border border-zinc-200 space-y-3">
                      <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3">
                        <div>
                          <div className="flex items-center space-x-2">
                            <h3 className="font-bold text-sm text-zinc-900">Katalog Menu</h3>
                            <span className="bg-zinc-100 text-zinc-700 font-bold text-xs px-2.5 py-0.5 rounded-full border border-zinc-200">
                              {filteredAdminMenu.length} / {menuItems.length} Produk
                            </span>
                          </div>
                          <p className="text-xs text-zinc-500 mt-0.5">
                            Data tersinkron otomatis ke cloud Firestore dan responsif di semua perangkat.
                          </p>
                        </div>
                        <div className="grid grid-cols-2 sm:flex sm:flex-wrap gap-2 w-full sm:w-auto">
                          <button
                            onClick={exportMenuToXLSX}
                            className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold px-3 py-2 rounded-xl text-xs transition-all flex items-center justify-center space-x-1"
                          >
                            <span>📥 Export</span>
                          </button>
                          <label className="bg-blue-600 hover:bg-blue-700 text-white font-bold px-3 py-2 rounded-xl text-xs transition-all flex items-center justify-center space-x-1 cursor-pointer">
                            <span>📤 Import</span>
                            <input type="file" accept=".xlsx, .xls" onChange={importMenuFromXLSX} className="hidden" />
                          </label>
                          <button
                            onClick={() => openMenuModal()}
                            className="col-span-2 sm:col-span-1 bg-rose-600 hover:bg-rose-700 text-white font-bold px-3.5 py-2 rounded-xl shadow-md text-xs transition-all flex items-center justify-center space-x-1"
                          >
                            <span>+ Tambah Menu</span>
                          </button>
                        </div>
                      </div>

                      {/* Search Bar & Quick Filter */}
                      <div className="flex flex-col sm:flex-row gap-2 pt-1 border-t border-zinc-100">
                        <div className="relative flex-1">
                          <input
                            type="text"
                            placeholder="Cari menu berdasarkan nama atau deskripsi..."
                            value={adminMenuSearch}
                            onChange={(e) => setAdminMenuSearch(e.target.value)}
                            className="w-full pl-8 pr-8 py-2 bg-zinc-50 border border-zinc-200 rounded-xl text-xs font-medium outline-none focus:border-rose-500 focus:bg-white transition-all"
                          />
                          <span className="absolute left-2.5 top-2.5 text-zinc-400 text-xs">🔍</span>
                          {adminMenuSearch && (
                            <button
                              onClick={() => setAdminMenuSearch('')}
                              className="absolute right-2.5 top-2 text-zinc-400 hover:text-zinc-600 text-xs font-bold"
                            >
                              ✕
                            </button>
                          )}
                        </div>
                        <select
                          value={adminMenuCatFilter}
                          onChange={(e) => setAdminMenuCatFilter(e.target.value)}
                          className="px-3 py-2 bg-zinc-50 border border-zinc-200 rounded-xl text-xs font-bold text-zinc-700 outline-none focus:border-rose-500 sm:w-48"
                        >
                          {categories.map((c) => (
                            <option key={c.id} value={c.id}>
                              {c.name}
                            </option>
                          ))}
                        </select>
                      </div>
                    </div>

                    {/* Empty State */}
                    {filteredAdminMenu.length === 0 && (
                      <div className="bg-white rounded-2xl p-8 text-center border border-dashed border-zinc-300">
                        <p className="text-3xl mb-2">🍽️</p>
                        <h4 className="font-bold text-zinc-800 text-sm">Tidak ada menu yang ditemukan</h4>
                        <p className="text-xs text-zinc-500 mt-1 max-w-sm mx-auto">
                          {adminMenuSearch || adminMenuCatFilter !== 'Semua'
                            ? 'Coba ubah kata kunci pencarian atau ganti filter kategori di atas.'
                            : 'Belum ada menu yang tersimpan. Klik tombol Tambah Menu untuk menambahkan hidangan.'}
                        </p>
                        <div className="mt-4 flex justify-center gap-2">
                          {(adminMenuSearch || adminMenuCatFilter !== 'Semua') && (
                            <button
                              onClick={() => {
                                setAdminMenuSearch('');
                                setAdminMenuCatFilter('Semua');
                              }}
                              className="px-3 py-1.5 bg-zinc-100 hover:bg-zinc-200 text-zinc-700 text-xs font-bold rounded-xl"
                            >
                              Reset Filter
                            </button>
                          )}
                          <button
                            onClick={() => openMenuModal()}
                            className="px-3.5 py-1.5 bg-rose-600 hover:bg-rose-700 text-white text-xs font-bold rounded-xl shadow"
                          >
                            + Tambah Menu Baru
                          </button>
                        </div>
                      </div>
                    )}

                    {/* MOBILE CARD VIEW: Visible on screens < 768px (Phones & small tablets) */}
                    <div className="block md:hidden space-y-3">
                      {filteredAdminMenu.map((item) => (
                        <div
                          key={item.id}
                          className="bg-white p-3.5 rounded-2xl shadow-sm border border-zinc-200 flex flex-col gap-3"
                        >
                          <div className="flex gap-3 items-start">
                            <img
                              src={item.image}
                              alt={item.name}
                              className="w-20 h-20 rounded-xl object-cover border border-zinc-200 shrink-0 bg-zinc-100"
                              onError={(e) => {
                                (e.target as HTMLImageElement).src = getFallbackImage(item.name);
                              }}
                            />
                            <div className="flex-1 min-w-0">
                              <div className="flex items-center gap-1.5 mb-1">
                                <span className="px-2 py-0.5 rounded-md text-[10px] font-bold uppercase bg-rose-50 text-rose-700 border border-rose-100">
                                  {item.category}
                                </span>
                              </div>
                              <h4 className="font-bold text-sm text-zinc-900 leading-tight truncate">{item.name}</h4>
                              <p className="text-xs text-rose-600 font-extrabold mt-1">
                                Rp {item.price.toLocaleString('id-ID')}
                              </p>
                              {item.description && (
                                <p className="text-[11px] text-zinc-500 line-clamp-2 mt-1 leading-snug">
                                  {item.description}
                                </p>
                              )}
                            </div>
                          </div>
                          <div className="flex items-center gap-2 pt-2 border-t border-zinc-100">
                            <button
                              onClick={() => openMenuModal(item)}
                              className="flex-1 py-2 bg-zinc-100 hover:bg-zinc-200 text-zinc-800 rounded-xl text-xs font-bold flex items-center justify-center space-x-1 transition-colors"
                            >
                              <span>✏️ Edit</span>
                            </button>
                            <button
                              onClick={() => handleDeleteMenu(item.id)}
                              className="flex-1 py-2 bg-red-50 hover:bg-red-100 text-red-600 rounded-xl text-xs font-bold flex items-center justify-center space-x-1 transition-colors"
                            >
                              <span>🗑️ Hapus</span>
                            </button>
                          </div>
                        </div>
                      ))}
                    </div>

                    {/* DESKTOP & TABLET TABLE VIEW: Visible on screens >= 768px (Laptops, Desktops, iPads) */}
                    <div className="hidden md:block bg-white rounded-2xl shadow-sm border border-zinc-200 overflow-hidden">
                      <div className="overflow-x-auto">
                        <table className="w-full text-left border-collapse text-xs sm:text-sm min-w-[650px]">
                          <thead>
                            <tr className="bg-zinc-50 border-b border-zinc-200 text-[11px] font-bold text-zinc-500 uppercase tracking-wider">
                              <th className="p-3.5">Foto</th>
                              <th className="p-3.5">Nama Menu</th>
                              <th className="p-3.5">Kategori</th>
                              <th className="p-3.5">Harga</th>
                              <th className="p-3.5">Deskripsi</th>
                              <th className="p-3.5 text-center">Aksi</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-zinc-100">
                            {filteredAdminMenu.map((item) => (
                              <tr key={item.id} className="hover:bg-zinc-50/80 transition-colors">
                                <td className="p-3.5">
                                  <img
                                    src={item.image}
                                    alt={item.name}
                                    className="w-12 h-12 rounded-xl object-cover border border-zinc-200 bg-zinc-100"
                                    onError={(e) => {
                                      (e.target as HTMLImageElement).src = getFallbackImage(item.name);
                                    }}
                                  />
                                </td>
                                <td className="p-3.5 font-bold text-zinc-900">{item.name}</td>
                                <td className="p-3.5">
                                  <span className="px-2.5 py-1 rounded-lg text-[10px] font-bold uppercase bg-zinc-100 text-zinc-700">
                                    {item.category}
                                  </span>
                                </td>
                                <td className="p-3.5 font-bold text-rose-600">
                                  Rp {item.price.toLocaleString('id-ID')}
                                </td>
                                <td className="p-3.5 text-zinc-500 max-w-xs truncate">{item.description || '-'}</td>
                                <td className="p-3.5 text-center space-x-2 whitespace-nowrap">
                                  <button
                                    onClick={() => openMenuModal(item)}
                                    className="px-3 py-1 bg-zinc-100 hover:bg-zinc-200 text-zinc-700 rounded-lg text-xs font-bold transition-colors"
                                  >
                                    Edit
                                  </button>
                                  <button
                                    onClick={() => handleDeleteMenu(item.id)}
                                    className="px-3 py-1 bg-red-50 hover:bg-red-100 text-red-600 rounded-lg text-xs font-bold transition-colors"
                                  >
                                    Hapus
                                  </button>
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  </div>
                );
              })()}

              {/* TABLES TAB */}
              {adminTab === 'tables' && (
                <div className="space-y-4">
                  <div className="flex justify-between items-center bg-white p-4 rounded-2xl shadow-sm border border-zinc-200">
                    <div>
                      <h3 className="font-bold text-sm text-zinc-900">Manajemen Meja & QR Pelanggan</h3>
                      <p className="text-xs text-zinc-500">Unduh gambar QR meja untuk dipajang di atas meja pelanggan.</p>
                    </div>
                    <button
                      onClick={handleAddTable}
                      className="bg-rose-600 hover:bg-rose-700 text-white font-bold px-4 py-2.5 rounded-xl shadow-md text-xs transition-all flex items-center space-x-1.5"
                    >
                      <span>+ Tambah Meja</span>
                    </button>
                  </div>

                  <div className="bg-white rounded-2xl shadow-sm border border-zinc-200 overflow-hidden">
                    <table className="w-full text-left border-collapse text-xs sm:text-sm">
                      <thead>
                        <tr className="bg-zinc-50 border-b border-zinc-200 text-[11px] font-bold text-zinc-500 uppercase">
                          <th className="p-3.5">Nama Meja</th>
                          <th className="p-3.5">Link Pemesanan</th>
                          <th className="p-3.5 text-center">Aksi QR Code</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-zinc-100">
                        {tables.map((tbl) => (
                          <tr key={tbl.id} className="hover:bg-zinc-50 transition-colors">
                            <td className="p-3.5 font-bold text-base text-zinc-900">{tbl.name}</td>
                            <td className="p-3.5 font-mono text-zinc-500 truncate max-w-xs">{tbl.link || `${window.location.origin}${window.location.pathname}?table=${tbl.name}`}</td>
                            <td className="p-3.5 text-center space-x-2">
                              <button
                                onClick={() => downloadTableQrPng(tbl)}
                                className="px-3.5 py-1.5 bg-rose-50 hover:bg-rose-100 text-rose-700 rounded-lg text-xs font-bold transition-colors"
                              >
                                Unduh
                              </button>
                              {tables.length > 1 && (
                                <button
                                  onClick={() => handleDeleteTable(tbl.id)}
                                  className="px-3 py-1.5 bg-red-50 hover:bg-red-100 text-red-600 rounded-lg text-xs font-bold transition-colors"
                                >
                                  Hapus
                                </button>
                              )}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {/* STAFF TAB */}
              {adminTab === 'staff' && (
                <div className="space-y-4">
                  <div className="flex justify-between items-center bg-white p-4 rounded-2xl shadow-sm border border-zinc-200">
                    <div>
                      <h3 className="font-bold text-sm text-zinc-900">Manajemen Akun Staff & Kasir</h3>
                      <p className="text-xs text-zinc-500">Atur akses login PIN 6-digit untuk kasir dan administrator.</p>
                    </div>
                    <button
                      onClick={() => openStaffModal()}
                      className="bg-rose-600 hover:bg-rose-700 text-white font-bold px-4 py-2.5 rounded-xl shadow-md text-xs transition-all flex items-center space-x-1.5"
                    >
                      <span>+ Tambah Staff</span>
                    </button>
                  </div>

                  <div className="bg-white rounded-2xl shadow-sm border border-zinc-200 overflow-hidden">
                    <table className="w-full text-left border-collapse text-xs sm:text-sm">
                      <thead>
                        <tr className="bg-zinc-50 border-b border-zinc-200 text-[11px] font-bold text-zinc-500 uppercase">
                          <th className="p-3.5">Nama Staff</th>
                          <th className="p-3.5">Username</th>
                          <th className="p-3.5">PIN (6-Digit)</th>
                          <th className="p-3.5">Role</th>
                          <th className="p-3.5 text-center">Aksi</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-zinc-100">
                        {staffUsers.map((staff) => {
                          const isPinVisible = visibleStaffPins[staff.id];
                          return (
                            <tr key={staff.id} className="hover:bg-zinc-50 transition-colors">
                              <td className="p-3.5 font-bold text-zinc-900">{staff.name}</td>
                              <td className="p-3.5 font-mono text-zinc-600">{staff.username}</td>
                              <td className="p-3.5 font-mono">
                                <div className="flex items-center space-x-2">
                                  <span className="bg-zinc-100 px-2.5 py-1 rounded-lg text-zinc-800 font-bold tracking-wider">
                                    {isPinVisible ? staff.pin : '••••••'}
                                  </span>
                                  <button
                                    onClick={() => setVisibleStaffPins((p) => ({ ...p, [staff.id]: !p[staff.id] }))}
                                    className="text-zinc-400 hover:text-zinc-600 text-xs"
                                  >
                                    {isPinVisible ? 'Sembunyi' : 'Lihat'}
                                  </button>
                                </div>
                              </td>
                              <td className="p-3.5">
                                <span className={`px-2.5 py-0.5 rounded-full text-[10px] font-bold uppercase ${staff.role === 'admin' ? 'bg-purple-100 text-purple-700' : 'bg-blue-100 text-blue-700'}`}>
                                  {staff.role}
                                </span>
                              </td>
                              <td className="p-3.5 text-center space-x-2">
                                <button
                                  onClick={() => openStaffModal(staff)}
                                  className="px-3 py-1 bg-zinc-100 hover:bg-zinc-200 text-zinc-700 rounded-lg text-xs font-bold"
                                >
                                  Edit
                                </button>
                                <button
                                  onClick={() => handleDeleteStaff(staff.id)}
                                  className="px-3 py-1 bg-red-50 hover:bg-red-100 text-red-600 rounded-lg text-xs font-bold"
                                >
                                  Hapus
                                </button>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {/* SETTINGS TAB */}
              {adminTab === 'settings' && (
                <div className="space-y-5 max-w-3xl mx-auto pb-10">
                  {/* Identity */}
                  <div className="bg-white rounded-2xl p-5 shadow-sm border border-zinc-200 space-y-4">
                    <h3 className="font-bold text-base text-zinc-900 border-b border-zinc-100 pb-2">Identitas Restoran & Pajak</h3>
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                      <div>
                        <label className="text-xs font-bold text-zinc-700 uppercase block mb-1">Nama Restoran</label>
                        <input
                          type="text"
                          value={settings.restoName}
                          onChange={(e) => setSettings({ ...settings, restoName: e.target.value })}
                          className="w-full px-3.5 py-2.5 bg-zinc-50 border border-zinc-200 rounded-xl text-xs font-bold outline-none focus:border-rose-500"
                        />
                      </div>
                      <div>
                        <label className="text-xs font-bold text-zinc-700 uppercase block mb-1">Nomor Telepon</label>
                        <input
                          type="text"
                          value={settings.restoPhone}
                          onChange={(e) => setSettings({ ...settings, restoPhone: e.target.value })}
                          className="w-full px-3.5 py-2.5 bg-zinc-50 border border-zinc-200 rounded-xl text-xs font-bold outline-none focus:border-rose-500"
                        />
                      </div>
                      <div className="sm:col-span-2">
                        <label className="text-xs font-bold text-zinc-700 uppercase block mb-1">Alamat Restoran</label>
                        <input
                          type="text"
                          value={settings.restoAddress}
                          onChange={(e) => setSettings({ ...settings, restoAddress: e.target.value })}
                          className="w-full px-3.5 py-2.5 bg-zinc-50 border border-zinc-200 rounded-xl text-xs font-bold outline-none focus:border-rose-500"
                        />
                      </div>
                      <div>
                        <label className="text-xs font-bold text-zinc-700 uppercase block mb-1">Pajak Restoran (%)</label>
                        <input
                          type="number"
                          value={settings.taxRate}
                          onChange={(e) => setSettings({ ...settings, taxRate: Number(e.target.value) })}
                          className="w-full px-3.5 py-2.5 bg-zinc-50 border border-zinc-200 rounded-xl text-xs font-bold outline-none focus:border-rose-500"
                        />
                      </div>
                      <div>
                        <label className="text-xs font-bold text-zinc-700 uppercase block mb-1">Service Charge (%)</label>
                        <input
                          type="number"
                          value={settings.serviceCharge}
                          onChange={(e) => setSettings({ ...settings, serviceCharge: Number(e.target.value) })}
                          className="w-full px-3.5 py-2.5 bg-zinc-50 border border-zinc-200 rounded-xl text-xs font-bold outline-none focus:border-rose-500"
                        />
                      </div>
                      <div>
                        <label className="text-xs font-bold text-zinc-700 uppercase block mb-1">Prefix / Kode Order (Cth: ORD)</label>
                        <input
                          type="text"
                          value={settings.orderPrefix || 'ORD'}
                          onChange={(e) => setSettings({ ...settings, orderPrefix: e.target.value })}
                          className="w-full px-3.5 py-2.5 bg-zinc-50 border border-zinc-200 rounded-xl text-xs font-bold uppercase outline-none focus:border-rose-500"
                        />
                      </div>
                    </div>
                  </div>

                  {/* Receipt & Notifications */}
                  <div className="bg-white rounded-2xl p-5 shadow-sm border border-zinc-200 space-y-4">
                    <h3 className="font-bold text-base text-zinc-900 border-b border-zinc-100 pb-2">Struk Pembayaran & Suara Panggilan</h3>
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                      <div>
                        <label className="text-xs font-bold text-zinc-700 uppercase block mb-1">Catatan Header Struk</label>
                        <input
                          type="text"
                          value={settings.receiptHeaderNote}
                          onChange={(e) => setSettings({ ...settings, receiptHeaderNote: e.target.value })}
                          className="w-full px-3.5 py-2.5 bg-zinc-50 border border-zinc-200 rounded-xl text-xs outline-none focus:border-rose-500"
                        />
                      </div>
                      <div>
                        <label className="text-xs font-bold text-zinc-700 uppercase block mb-1">Catatan Footer Struk</label>
                        <input
                          type="text"
                          value={settings.receiptFooterNote}
                          onChange={(e) => setSettings({ ...settings, receiptFooterNote: e.target.value })}
                          className="w-full px-3.5 py-2.5 bg-zinc-50 border border-zinc-200 rounded-xl text-xs outline-none focus:border-rose-500"
                        />
                      </div>
                      <div>
                        <label className="text-xs font-bold text-zinc-700 uppercase block mb-1">Lebar Kertas Struk</label>
                        <select
                          value={settings.receiptPaperWidth}
                          onChange={(e) => setSettings({ ...settings, receiptPaperWidth: e.target.value })}
                          className="w-full px-3.5 py-2.5 bg-zinc-50 border border-zinc-200 rounded-xl text-xs font-bold outline-none focus:border-rose-500 bg-white"
                        >
                          <option value="280px">Thermal 58mm (280px)</option>
                          <option value="300px">Thermal 80mm (300px)</option>
                          <option value="350px">Lebar Standar (350px)</option>
                        </select>
                      </div>
                      <div>
                        <label className="text-xs font-bold text-zinc-700 uppercase block mb-1">Suara Panggilan / Gender Voice</label>
                        <select
                          value={settings.voiceType || 'female'}
                          onChange={(e) => setSettings({ ...settings, voiceType: e.target.value as 'female' | 'male' })}
                          className="w-full px-3.5 py-2.5 bg-zinc-50 border border-zinc-200 rounded-xl text-xs font-bold outline-none focus:border-rose-500 bg-white"
                        >
                          <option value="female">Suara Wanita (Google Indonesian TTS)</option>
                          <option value="male">Suara Laki-laki (Deep Tone Voice)</option>
                        </select>
                      </div>
                      <div className="flex items-center justify-between p-3 bg-zinc-50 rounded-xl border border-zinc-200 sm:col-span-2">
                        <div>
                          <span className="font-bold text-xs block text-zinc-800">Suara Notifikasi Audio</span>
                          <span className="text-[11px] text-zinc-500">Bunyi chime dan pengumuman nama pelanggan saat pesanan siap</span>
                        </div>
                        <input
                          type="checkbox"
                          checked={settings.soundEnabled}
                          onChange={(e) => setSettings({ ...settings, soundEnabled: e.target.checked })}
                          className="w-5 h-5 accent-rose-600 rounded cursor-pointer"
                        />
                      </div>

                      {/* Pengaturan Kalimat Ucapan Notifikasi */}
                      <div className="sm:col-span-2 bg-rose-50/70 p-4 rounded-2xl border border-rose-200 space-y-2.5">
                        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                          <div>
                            <label className="text-xs font-bold text-rose-950 uppercase tracking-wider block">
                              Pengaturan Kalimat Ucapan Notifikasi
                            </label>
                            <p className="text-[11px] text-zinc-600">
                              Format kalimat yang diucapkan sistem suara saat pesanan pelanggan selesai.
                            </p>
                          </div>
                          <button
                            type="button"
                            onClick={() => {
                              announceOrderCall(
                                'Budi',
                                'ORD-1001',
                                settings.soundEnabled,
                                settings.voiceType || 'female',
                                settings.callTemplate,
                                'Meja 05'
                              );
                            }}
                            className="self-start sm:self-auto px-3 py-1.5 bg-amber-500 hover:bg-amber-600 text-white rounded-xl text-xs font-bold shadow-sm transition-colors flex items-center space-x-1.5 flex-shrink-0"
                          >
                            <span>🔊</span>
                            <span>Tes Suara Notifikasi</span>
                          </button>
                        </div>

                        <textarea
                          rows={2}
                          value={settings.callTemplate}
                          onChange={(e) => setSettings({ ...settings, callTemplate: e.target.value })}
                          placeholder="Pesanan atas nama {nama}, meja {meja}, dengan nomor {id}, sudah selesai dan siap diambil."
                          className="w-full px-3.5 py-2.5 bg-white border border-rose-200 rounded-xl text-xs font-medium text-zinc-900 outline-none focus:border-rose-500 shadow-xs"
                        />
                        
                        <div className="flex flex-wrap items-center justify-between gap-2 text-[11px] text-zinc-600">
                          <p>
                            Gunakan kode: <span className="font-mono font-bold text-rose-700 bg-white px-1.5 py-0.5 rounded border border-rose-200">{'{nama}'}</span> = Nama Pelanggan, <span className="font-mono font-bold text-rose-700 bg-white px-1.5 py-0.5 rounded border border-rose-200">{'{meja}'}</span> = Nomor Meja, <span className="font-mono font-bold text-rose-700 bg-white px-1.5 py-0.5 rounded border border-rose-200">{'{id}'}</span> = Nomor Pesanan.
                          </p>
                          <button
                            type="button"
                            onClick={() => setSettings({
                              ...settings,
                              callTemplate: "Pesanan atas nama {nama}, meja {meja}, dengan nomor {id}, sudah selesai dan siap diambil."
                            })}
                            className="text-rose-600 font-bold hover:underline"
                          >
                            Reset ke Default
                          </button>
                        </div>
                      </div>
                    </div>
                  </div>

                  {/* Payment & QRIS */}
                  <div className="bg-white rounded-2xl p-5 shadow-sm border border-zinc-200 space-y-4">
                    <div className="flex justify-between items-center border-b border-zinc-100 pb-2">
                      <h3 className="font-bold text-base text-zinc-900">Metode Pembayaran & Static QRIS</h3>
                      <button
                        onClick={() => openPaymentModal()}
                        className="px-3 py-1.5 bg-rose-600 text-white rounded-xl text-xs font-bold hover:bg-rose-700 shadow-sm"
                      >
                        + Tambah Metode
                      </button>
                    </div>
                    <div>
                      <label className="text-xs font-bold text-zinc-700 uppercase block mb-1">Static QRIS String</label>
                      <textarea
                        rows={2}
                        value={settings.staticQris}
                        onChange={(e) => setSettings({ ...settings, staticQris: e.target.value })}
                        className="w-full px-3.5 py-2.5 bg-zinc-50 border border-zinc-200 rounded-xl text-xs font-mono outline-none focus:border-rose-500"
                      />
                    </div>
                    <div className="space-y-2">
                      {paymentMethods.map((pm) => (
                        <div key={pm.id} className="flex justify-between items-center p-3 bg-zinc-50 rounded-xl border border-zinc-200">
                          <div>
                            <span className="font-bold text-xs text-zinc-800">{pm.name}</span>
                            <span className="text-[10px] text-zinc-500 block uppercase font-semibold">Tipe: {pm.type} • {pm.isActive ? 'Aktif' : 'Nonaktif'}</span>
                          </div>
                          <div className="flex space-x-1.5">
                            <button onClick={() => openPaymentModal(pm)} className="px-2.5 py-1 bg-white border border-zinc-200 text-zinc-700 rounded-lg text-xs font-bold">Edit</button>
                            <button onClick={() => handleDeletePayment(pm.id)} className="px-2.5 py-1 bg-red-50 text-red-600 rounded-lg text-xs font-bold">Hapus</button>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>

                  <button
                    onClick={() => handleSaveSettings(settings)}
                    className="w-full bg-rose-600 hover:bg-rose-700 text-white font-bold py-3.5 rounded-xl shadow-lg shadow-rose-500/30 transition-all text-sm"
                  >
                    Simpan Semua Pengaturan ke Firebase
                  </button>
                </div>
              )}
            </div>
          </main>
        </div>
      )}

      {/* MODAL: Menu Add/Edit */}
      {showMenuModal && (
        <div className="fixed inset-0 bg-black/60 z-50 flex items-center justify-center p-3 sm:p-4 backdrop-blur-sm">
          <div className="bg-white w-full max-w-md max-h-[92vh] overflow-y-auto rounded-2xl p-5 sm:p-6 shadow-2xl">
            <h3 className="text-base font-bold text-zinc-900 mb-4">{editingMenu ? 'Edit Menu' : 'Tambah Menu Baru'}</h3>
            <form onSubmit={handleSaveMenu} className="space-y-3.5">
              <div>
                <label className="text-xs font-bold text-zinc-600 block mb-1">Nama Menu *</label>
                <input
                  type="text"
                  required
                  value={menuForm.name}
                  onChange={(e) => setMenuForm({ ...menuForm, name: e.target.value })}
                  className="w-full p-2.5 bg-zinc-50 border border-zinc-200 rounded-xl text-xs font-bold outline-none focus:border-rose-500"
                />
              </div>
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="text-xs font-bold text-zinc-600 block mb-1">Harga (Rp) *</label>
                  <input
                    type="number"
                    required
                    value={menuForm.price}
                    onChange={(e) => setMenuForm({ ...menuForm, price: e.target.value })}
                    className="w-full p-2.5 bg-zinc-50 border border-zinc-200 rounded-xl text-xs font-bold outline-none focus:border-rose-500"
                  />
                </div>
                <div>
                  <label className="text-xs font-bold text-zinc-600 block mb-1">Kategori *</label>
                  <select
                    value={menuForm.category}
                    onChange={(e) => setMenuForm({ ...menuForm, category: e.target.value })}
                    className="w-full p-2.5 bg-zinc-50 border border-zinc-200 rounded-xl text-xs font-bold bg-white outline-none focus:border-rose-500"
                  >
                    {categories
                      .filter((c) => c.id !== 'Semua')
                      .map((c) => (
                        <option key={c.id} value={c.id}>{c.name}</option>
                      ))}
                  </select>
                </div>
              </div>
              <div>
                <label className="text-xs font-bold text-zinc-600 block mb-1">URL Gambar</label>
                <input
                  type="text"
                  placeholder="https://..."
                  value={menuForm.image}
                  onChange={(e) => setMenuForm({ ...menuForm, image: e.target.value })}
                  className="w-full p-2.5 bg-zinc-50 border border-zinc-200 rounded-xl text-xs outline-none focus:border-rose-500"
                />
              </div>
              <div>
                <label className="text-xs font-bold text-zinc-600 block mb-1">Deskripsi Hidangan</label>
                <textarea
                  rows={2}
                  value={menuForm.description}
                  onChange={(e) => setMenuForm({ ...menuForm, description: e.target.value })}
                  className="w-full p-2.5 bg-zinc-50 border border-zinc-200 rounded-xl text-xs outline-none focus:border-rose-500"
                />
              </div>
              <div className="flex gap-2 pt-2">
                <button type="submit" className="flex-1 bg-rose-600 text-white font-bold py-2.5 rounded-xl text-xs shadow hover:bg-rose-700">
                  Simpan ke Database
                </button>
                <button
                  type="button"
                  onClick={() => setShowMenuModal(false)}
                  className="flex-1 bg-zinc-200 text-zinc-800 font-bold py-2.5 rounded-xl text-xs hover:bg-zinc-300"
                >
                  Batal
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* MODAL: Staff Add/Edit */}
      {showStaffModal && (
        <div className="fixed inset-0 bg-black/60 z-50 flex items-center justify-center p-4 backdrop-blur-sm">
          <div className="bg-white w-full max-w-sm rounded-2xl p-5 shadow-2xl">
            <h3 className="text-base font-bold text-zinc-900 mb-4">{editingStaff ? 'Edit Staff' : 'Tambah Staff Baru'}</h3>
            <form onSubmit={handleSaveStaff} className="space-y-3.5">
              <div>
                <label className="text-xs font-bold text-zinc-600 block mb-1">Nama Lengkap *</label>
                <input
                  type="text"
                  required
                  value={staffForm.name}
                  onChange={(e) => setStaffForm({ ...staffForm, name: e.target.value })}
                  className="w-full p-2.5 bg-zinc-50 border border-zinc-200 rounded-xl text-xs font-bold outline-none focus:border-rose-500"
                />
              </div>
              <div>
                <label className="text-xs font-bold text-zinc-600 block mb-1">Username Login *</label>
                <input
                  type="text"
                  required
                  value={staffForm.username}
                  onChange={(e) => setStaffForm({ ...staffForm, username: e.target.value })}
                  className="w-full p-2.5 bg-zinc-50 border border-zinc-200 rounded-xl text-xs font-bold outline-none focus:border-rose-500"
                />
              </div>
              <div>
                <label className="text-xs font-bold text-zinc-600 block mb-1">PIN (6 Digit Angka) *</label>
                <input
                  type="password"
                  maxLength={6}
                  required
                  placeholder="••••••"
                  value={staffForm.pin}
                  onChange={(e) => setStaffForm({ ...staffForm, pin: e.target.value })}
                  className="w-full p-2.5 bg-zinc-50 border border-zinc-200 rounded-xl text-xs font-bold tracking-widest text-center outline-none focus:border-rose-500"
                />
              </div>
              <div>
                <label className="text-xs font-bold text-zinc-600 block mb-1">Role / Peran</label>
                <select
                  value={staffForm.role}
                  onChange={(e) => setStaffForm({ ...staffForm, role: e.target.value as 'cashier' | 'admin' })}
                  className="w-full p-2.5 bg-zinc-50 border border-zinc-200 rounded-xl text-xs font-bold bg-white outline-none focus:border-rose-500"
                >
                  <option value="cashier">Kasir (Hanya POS & Pesanan)</option>
                  <option value="admin">Administrator (Semua Akses & Pengaturan)</option>
                </select>
              </div>
              <div className="flex gap-2 pt-2">
                <button type="submit" className="flex-1 bg-rose-600 text-white font-bold py-2.5 rounded-xl text-xs shadow hover:bg-rose-700">
                  Simpan Staff
                </button>
                <button
                  type="button"
                  onClick={() => setShowStaffModal(false)}
                  className="flex-1 bg-zinc-200 text-zinc-800 font-bold py-2.5 rounded-xl text-xs hover:bg-zinc-300"
                >
                  Batal
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* MODAL: Payment Add/Edit */}
      {showPaymentModal && (
        <div className="fixed inset-0 bg-black/60 z-50 flex items-center justify-center p-4 backdrop-blur-sm">
          <div className="bg-white w-full max-w-sm rounded-2xl p-5 shadow-2xl">
            <h3 className="text-base font-bold text-zinc-900 mb-4">{editingPayment ? 'Edit Metode Pembayaran' : 'Tambah Metode Pembayaran'}</h3>
            <form onSubmit={handleSavePayment} className="space-y-3.5">
              <div>
                <label className="text-xs font-bold text-zinc-600 block mb-1">Nama Metode *</label>
                <input
                  type="text"
                  required
                  placeholder="Cth: QRIS ShopeePay"
                  value={paymentForm.name}
                  onChange={(e) => setPaymentForm({ ...paymentForm, name: e.target.value })}
                  className="w-full p-2.5 bg-zinc-50 border border-zinc-200 rounded-xl text-xs font-bold outline-none focus:border-rose-500"
                />
              </div>
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="text-xs font-bold text-zinc-600 block mb-1">Tipe</label>
                  <select
                    value={paymentForm.type}
                    onChange={(e) => setPaymentForm({ ...paymentForm, type: e.target.value as 'cash' | 'qris' | 'transfer' })}
                    className="w-full p-2.5 bg-zinc-50 border border-zinc-200 rounded-xl text-xs font-bold bg-white"
                  >
                    <option value="cash">Tunai (Cash)</option>
                    <option value="qris">QRIS</option>
                    <option value="transfer">Transfer Bank</option>
                  </select>
                </div>
                <div>
                  <label className="text-xs font-bold text-zinc-600 block mb-1">Status</label>
                  <select
                    value={paymentForm.isActive ? 'true' : 'false'}
                    onChange={(e) => setPaymentForm({ ...paymentForm, isActive: e.target.value === 'true' })}
                    className="w-full p-2.5 bg-zinc-50 border border-zinc-200 rounded-xl text-xs font-bold bg-white"
                  >
                    <option value="true">Aktif</option>
                    <option value="false">Nonaktif</option>
                  </select>
                </div>
              </div>
              <div>
                <label className="text-xs font-bold text-zinc-600 block mb-1">Instruksi Pembayaran</label>
                <textarea
                  rows={2}
                  placeholder="Cth: Silakan bayar tunai di meja kasir"
                  value={paymentForm.instructions}
                  onChange={(e) => setPaymentForm({ ...paymentForm, instructions: e.target.value })}
                  className="w-full p-2.5 bg-zinc-50 border border-zinc-200 rounded-xl text-xs outline-none focus:border-rose-500"
                />
              </div>
              <div className="flex gap-2 pt-2">
                <button type="submit" className="flex-1 bg-rose-600 text-white font-bold py-2.5 rounded-xl text-xs shadow hover:bg-rose-700">
                  Simpan
                </button>
                <button
                  type="button"
                  onClick={() => setShowPaymentModal(false)}
                  className="flex-1 bg-zinc-200 text-zinc-800 font-bold py-2.5 rounded-xl text-xs hover:bg-zinc-300"
                >
                  Batal
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* MODAL: Category Add */}
      {showCategoryModal && (
        <div className="fixed inset-0 bg-black/60 z-50 flex items-center justify-center p-4 backdrop-blur-sm">
          <div className="bg-white w-full max-w-xs rounded-2xl p-5 shadow-2xl">
            <h3 className="text-base font-bold text-zinc-900 mb-3">Tambah Kategori Baru</h3>
            <form onSubmit={handleAddCategory} className="space-y-3">
              <div>
                <label className="text-xs font-bold text-zinc-600 block mb-1">Nama Kategori *</label>
                <input
                  type="text"
                  required
                  placeholder="Cth: Paket Hemat"
                  value={categoryForm.name}
                  onChange={(e) => setCategoryForm({ name: e.target.value })}
                  className="w-full p-2.5 bg-zinc-50 border border-zinc-200 rounded-xl text-xs font-bold outline-none focus:border-rose-500"
                />
              </div>
              <div className="flex gap-2 pt-1">
                <button type="submit" className="flex-1 bg-rose-600 text-white font-bold py-2.5 rounded-xl text-xs shadow hover:bg-rose-700">
                  Tambah
                </button>
                <button
                  type="button"
                  onClick={() => setShowCategoryModal(false)}
                  className="flex-1 bg-zinc-200 text-zinc-800 font-bold py-2.5 rounded-xl text-xs hover:bg-zinc-300"
                >
                  Batal
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
      {/* MODAL: Custom Alert & Confirmation */}
      {modalConfig.isOpen && (
        <div className="fixed inset-0 bg-black/60 z-50 flex items-center justify-center p-4 backdrop-blur-sm animate-fade-in">
          <div className="bg-white w-full max-w-sm rounded-2xl p-5 shadow-2xl border border-zinc-100">
            <h3 className="text-base font-bold text-zinc-900 mb-2">{modalConfig.title}</h3>
            <p className="text-xs text-zinc-600 mb-4 leading-relaxed">{modalConfig.message}</p>
            <div className="flex gap-2">
              {modalConfig.onConfirm ? (
                <>
                  <button
                    type="button"
                    onClick={() => {
                      const cb = modalConfig.onConfirm;
                      closeAlert();
                      if (cb) cb();
                    }}
                    className="flex-1 bg-red-600 hover:bg-red-700 text-white font-bold py-2.5 rounded-xl text-xs shadow transition-colors"
                  >
                    Ya, Lanjutkan
                  </button>
                  <button
                    type="button"
                    onClick={closeAlert}
                    className="flex-1 bg-zinc-200 hover:bg-zinc-300 text-zinc-800 font-bold py-2.5 rounded-xl text-xs transition-colors"
                  >
                    Batal
                  </button>
                </>
              ) : (
                <button
                  type="button"
                  onClick={closeAlert}
                  className="w-full bg-rose-600 hover:bg-rose-700 text-white font-bold py-2.5 rounded-xl text-xs shadow transition-colors"
                >
                  Tutup
                </button>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
