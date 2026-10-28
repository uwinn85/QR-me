export interface Settings {
  restoName: string;
  taxRate: number;
  serviceCharge: number;
  orderPrefix: string;
  staticQris: string;
  appBackground: string;
  restoAddress: string;
  restoPhone: string;
  soundEnabled: boolean;
  voiceType: 'female' | 'male';
  callTemplate: string;
  receiptHeaderNote: string;
  receiptFooterNote: string;
  receiptShowAddress: boolean;
  receiptShowPhone: boolean;
  receiptShowCashier: boolean;
  receiptPaperWidth: string;
}

export interface MenuItem {
  id: string | number;
  name: string;
  price: number;
  category: string;
  image: string;
  description: string;
}

export interface Category {
  id: string;
  name: string;
}

export interface PaymentMethod {
  id: string;
  name: string;
  type: 'cash' | 'qris' | 'transfer';
  isActive: boolean;
  instructions: string;
}

export interface TableItem {
  id: string;
  name: string;
  link: string;
}

export interface StaffUser {
  id: string | number;
  username: string;
  pin: string;
  role: 'admin' | 'cashier';
  name: string;
}

export interface CartItem extends MenuItem {
  quantity: number;
}

export interface Order {
  id: string;
  customerName: string;
  table: string;
  orderType: 'dine-in' | 'takeaway';
  items: CartItem[];
  subtotal: number;
  tax: number;
  service: number;
  total: number;
  paymentMethod: string;
  notes?: string;
  status: 'unverified' | 'pending' | 'processing' | 'completed';
  time: string;
  source: string;
  cashGiven?: number;
  changeAmount?: number;
  createdAt?: string;
  updatedAt?: string;
}
