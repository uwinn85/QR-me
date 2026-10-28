import {
  collection,
  doc,
  getDoc,
  getDocs,
  setDoc,
  updateDoc,
  deleteDoc,
  onSnapshot,
  writeBatch
} from 'firebase/firestore';
import { db, handleFirestoreError, OperationType } from './firebase';
import {
  Settings,
  MenuItem,
  Category,
  PaymentMethod,
  TableItem,
  StaffUser,
  Order
} from './types';
import {
  DEFAULT_SETTINGS,
  DEFAULT_PAYMENT_METHODS,
  DEFAULT_CATEGORIES,
  DEFAULT_MENU,
  DEFAULT_USERS,
  DEFAULT_TABLES
} from './utils';

// Helper to seed initial data in Firestore if empty
export async function initializeFirestoreData(): Promise<void> {
  try {
    // 1. Settings
    const settingsRef = doc(db, 'settings', 'general');
    const settingsSnap = await getDoc(settingsRef);
    if (!settingsSnap.exists()) {
      await setDoc(settingsRef, DEFAULT_SETTINGS);
    }

    // 2. Categories
    const categoriesCol = collection(db, 'categories');
    const categoriesSnap = await getDocs(categoriesCol);
    if (categoriesSnap.empty) {
      const batch = writeBatch(db);
      for (const cat of DEFAULT_CATEGORIES) {
        batch.set(doc(db, 'categories', cat.id), cat);
      }
      await batch.commit();
    }

    // 3. Menu items
    const menuCol = collection(db, 'menu');
    const menuSnap = await getDocs(menuCol);
    if (menuSnap.empty) {
      const batch = writeBatch(db);
      for (const item of DEFAULT_MENU) {
        batch.set(doc(db, 'menu', String(item.id)), item);
      }
      await batch.commit();
    }

    // 4. Payment methods
    const paymentsCol = collection(db, 'paymentMethods');
    const paymentsSnap = await getDocs(paymentsCol);
    if (paymentsSnap.empty) {
      const batch = writeBatch(db);
      for (const pm of DEFAULT_PAYMENT_METHODS) {
        batch.set(doc(db, 'paymentMethods', pm.id), pm);
      }
      await batch.commit();
    }

    // 5. Tables
    const tablesCol = collection(db, 'tables');
    const tablesSnap = await getDocs(tablesCol);
    if (tablesSnap.empty) {
      const batch = writeBatch(db);
      for (const tbl of DEFAULT_TABLES) {
        const link = `${window.location.origin}${window.location.pathname}?table=${tbl.name.replace(/\D/g, '') || '1'}`;
        batch.set(doc(db, 'tables', tbl.id), { ...tbl, link });
      }
      await batch.commit();
    }

    // 6. Staff
    const staffCol = collection(db, 'staff');
    const staffSnap = await getDocs(staffCol);
    if (staffSnap.empty) {
      const batch = writeBatch(db);
      for (const user of DEFAULT_USERS) {
        batch.set(doc(db, 'staff', String(user.id)), user);
      }
      await batch.commit();
    }
  } catch (error) {
    console.warn('Initial seeding handled or offline:', error);
  }
}

// Subscriptions
export function subscribeSettings(callback: (settings: Settings) => void): () => void {
  const path = 'settings';
  return onSnapshot(
    doc(db, 'settings', 'general'),
    (snap) => {
      if (snap.exists()) {
        callback(snap.data() as Settings);
      } else {
        callback(DEFAULT_SETTINGS);
      }
    },
    (error) => {
      handleFirestoreError(error, OperationType.GET, path);
    }
  );
}

export function subscribeCategories(callback: (categories: Category[]) => void): () => void {
  const path = 'categories';
  return onSnapshot(
    collection(db, 'categories'),
    (snap) => {
      const list: Category[] = [];
      snap.forEach((d) => list.push(d.data() as Category));
      if (list.length > 0) {
        callback(list);
      }
    },
    (error) => {
      handleFirestoreError(error, OperationType.LIST, path);
    }
  );
}

export function subscribeMenu(callback: (menu: MenuItem[]) => void): () => void {
  const path = 'menu';
  return onSnapshot(
    collection(db, 'menu'),
    (snap) => {
      const list: MenuItem[] = [];
      snap.forEach((d) => list.push(d.data() as MenuItem));
      if (list.length > 0) {
        callback(list);
      }
    },
    (error) => {
      handleFirestoreError(error, OperationType.LIST, path);
    }
  );
}

export function subscribePaymentMethods(callback: (payments: PaymentMethod[]) => void): () => void {
  const path = 'paymentMethods';
  return onSnapshot(
    collection(db, 'paymentMethods'),
    (snap) => {
      const list: PaymentMethod[] = [];
      snap.forEach((d) => list.push(d.data() as PaymentMethod));
      if (list.length > 0) {
        callback(list);
      }
    },
    (error) => {
      handleFirestoreError(error, OperationType.LIST, path);
    }
  );
}

export function subscribeTables(callback: (tables: TableItem[]) => void): () => void {
  const path = 'tables';
  return onSnapshot(
    collection(db, 'tables'),
    (snap) => {
      const list: TableItem[] = [];
      snap.forEach((d) => list.push(d.data() as TableItem));
      if (list.length > 0) {
        callback(list);
      }
    },
    (error) => {
      handleFirestoreError(error, OperationType.LIST, path);
    }
  );
}

export function subscribeStaff(callback: (staff: StaffUser[]) => void): () => void {
  const path = 'staff';
  return onSnapshot(
    collection(db, 'staff'),
    (snap) => {
      const list: StaffUser[] = [];
      snap.forEach((d) => list.push(d.data() as StaffUser));
      if (list.length > 0) {
        callback(list);
      }
    },
    (error) => {
      handleFirestoreError(error, OperationType.LIST, path);
    }
  );
}

export function subscribeOrders(callback: (orders: Order[]) => void): () => void {
  const path = 'orders';
  return onSnapshot(
    collection(db, 'orders'),
    (snap) => {
      const list: Order[] = [];
      snap.forEach((d) => list.push(d.data() as Order));
      // Sort orders descending by time
      list.sort((a, b) => new Date(b.time).getTime() - new Date(a.time).getTime());
      callback(list);
    },
    (error) => {
      handleFirestoreError(error, OperationType.LIST, path);
    }
  );
}

// Mutations
export async function createOrder(order: Order): Promise<void> {
  const path = `orders/${order.id}`;
  try {
    await setDoc(doc(db, 'orders', order.id), order);
  } catch (error) {
    handleFirestoreError(error, OperationType.CREATE, path);
  }
}

export async function updateOrderStatus(orderId: string, status: Order['status']): Promise<void> {
  const path = `orders/${orderId}`;
  try {
    await updateDoc(doc(db, 'orders', orderId), { status, updatedAt: new Date().toISOString() });
  } catch (error) {
    handleFirestoreError(error, OperationType.UPDATE, path);
  }
}

export async function saveSettings(settings: Settings): Promise<void> {
  const path = 'settings/general';
  try {
    await setDoc(doc(db, 'settings', 'general'), settings);
  } catch (error) {
    handleFirestoreError(error, OperationType.WRITE, path);
  }
}

export async function saveMenuItem(item: MenuItem): Promise<void> {
  const path = `menu/${item.id}`;
  try {
    await setDoc(doc(db, 'menu', String(item.id)), item);
  } catch (error) {
    handleFirestoreError(error, OperationType.WRITE, path);
  }
}

export async function deleteMenuItem(id: string | number): Promise<void> {
  const path = `menu/${id}`;
  try {
    await deleteDoc(doc(db, 'menu', String(id)));
  } catch (error) {
    handleFirestoreError(error, OperationType.DELETE, path);
  }
}

export async function batchSaveMenuItems(items: MenuItem[]): Promise<void> {
  const path = 'menu';
  try {
    const batch = writeBatch(db);
    for (const item of items) {
      batch.set(doc(db, 'menu', String(item.id)), item);
    }
    await batch.commit();
  } catch (error) {
    handleFirestoreError(error, OperationType.WRITE, path);
  }
}

export async function saveCategory(cat: Category): Promise<void> {
  const path = `categories/${cat.id}`;
  try {
    await setDoc(doc(db, 'categories', cat.id), cat);
  } catch (error) {
    handleFirestoreError(error, OperationType.WRITE, path);
  }
}

export async function deleteCategory(id: string): Promise<void> {
  const path = `categories/${id}`;
  try {
    await deleteDoc(doc(db, 'categories', id));
  } catch (error) {
    handleFirestoreError(error, OperationType.DELETE, path);
  }
}

export async function saveTable(table: TableItem): Promise<void> {
  const path = `tables/${table.id}`;
  try {
    await setDoc(doc(db, 'tables', table.id), table);
  } catch (error) {
    handleFirestoreError(error, OperationType.WRITE, path);
  }
}

export async function deleteTable(id: string): Promise<void> {
  const path = `tables/${id}`;
  try {
    await deleteDoc(doc(db, 'tables', id));
  } catch (error) {
    handleFirestoreError(error, OperationType.DELETE, path);
  }
}

export async function saveStaffUser(staff: StaffUser): Promise<void> {
  const path = `staff/${staff.id}`;
  try {
    await setDoc(doc(db, 'staff', String(staff.id)), staff);
  } catch (error) {
    handleFirestoreError(error, OperationType.WRITE, path);
  }
}

export async function deleteStaffUser(id: string | number): Promise<void> {
  const path = `staff/${id}`;
  try {
    await deleteDoc(doc(db, 'staff', String(id)));
  } catch (error) {
    handleFirestoreError(error, OperationType.DELETE, path);
  }
}

export async function savePaymentMethod(pm: PaymentMethod): Promise<void> {
  const path = `paymentMethods/${pm.id}`;
  try {
    await setDoc(doc(db, 'paymentMethods', pm.id), pm);
  } catch (error) {
    handleFirestoreError(error, OperationType.WRITE, path);
  }
}

export async function deletePaymentMethod(id: string): Promise<void> {
  const path = `paymentMethods/${id}`;
  try {
    await deleteDoc(doc(db, 'paymentMethods', id));
  } catch (error) {
    handleFirestoreError(error, OperationType.DELETE, path);
  }
}
