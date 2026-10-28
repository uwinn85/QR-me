import QRCode from 'qrcode';
import { Settings, PaymentMethod, Category, MenuItem, StaffUser, TableItem } from './types';

export const DEFAULT_SETTINGS: Settings = {
  restoName: "Nusantara Resto",
  taxRate: 11,
  serviceCharge: 5,
  orderPrefix: "ORD",
  staticQris: "00020101021126580016ID.CO.QRIS.WWW01189360091430000000010303UMI51440014ID.A0107070303UMI021500000000000000303UMI5204581253033605802ID5915Nusantara Resto6007Jakarta6304",
  appBackground: "#f4f4f5",
  restoAddress: "Jl. Sudirman No. 45, Jakarta",
  restoPhone: "0812-3456-7890",
  soundEnabled: true,
  voiceType: "female",
  callTemplate: "Pesanan atas nama {nama}, meja {meja}, dengan nomor {id}, sudah selesai dan siap diambil.",
  receiptHeaderNote: "Struk Pembayaran Resmi",
  receiptFooterNote: "Terima Kasih Atas Kunjungan Anda!",
  receiptShowAddress: true,
  receiptShowPhone: true,
  receiptShowCashier: true,
  receiptPaperWidth: "300px"
};

export const DEFAULT_PAYMENT_METHODS: PaymentMethod[] = [
  { id: 'qris', name: 'QRIS (Gopay, OVO, Dana)', type: 'qris', isActive: true, instructions: 'Scan QRIS di bawah ini menggunakan aplikasi mobile banking atau e-wallet Anda.' },
  { id: 'cash', name: 'Tunai di Kasir', type: 'cash', isActive: true, instructions: 'Silakan tunjukkan QR Code transaksi atau ID Pesanan ke kasir untuk melakukan pembayaran tunai.' }
];

export const DEFAULT_CATEGORIES: Category[] = [
  { id: 'Semua', name: 'Semua Menu' },
  { id: 'Makanan Utama', name: 'Makanan Utama' },
  { id: 'Minuman', name: 'Minuman' },
  { id: 'Camilan', name: 'Camilan' },
  { id: 'Penutup', name: 'Penutup' }
];

export const DEFAULT_MENU: MenuItem[] = [
  { id: '1', name: 'Nasi Goreng Spesial', price: 45000, category: 'Makanan Utama', image: 'https://images.unsplash.com/photo-1603133872878-684f208fb84b?w=500&auto=format&fit=crop&q=80', description: 'Nasi goreng dengan telur, ayam suwir, dan udang renyah.' },
  { id: '2', name: 'Sop Buntut Bakar', price: 85000, category: 'Makanan Utama', image: 'https://images.unsplash.com/photo-1547928576-a4a33237cbc3?w=500&auto=format&fit=crop&q=80', description: 'Sop buntut sapi empuk pilihan dibakar dengan baluran kecap manis khas.' },
  { id: '3', name: 'Ayam Penyet Sambal Ijo', price: 38000, category: 'Makanan Utama', image: 'https://images.unsplash.com/photo-1626082927389-6cd097cdc6ec?w=500&auto=format&fit=crop&q=80', description: 'Ayam goreng rempah gurih dengan ulekan sambal ijo ekstra pedas.' },
  { id: '4', name: 'Es Teh Manis', price: 10000, category: 'Minuman', image: 'https://images.unsplash.com/photo-1556679343-c7306c1976bc?w=500&auto=format&fit=crop&q=80', description: 'Teh melati seduh segar dingin dengan manis pas.' },
  { id: '5', name: 'Es Kopi Susu Aren', price: 25000, category: 'Minuman', image: 'https://images.unsplash.com/photo-1517701550927-30cf4ba1dba5?w=500&auto=format&fit=crop&q=80', description: 'Espresso robusta & arabika, susu segar, dan gula aren organik.' },
  { id: '6', name: 'Tahu Cabe Garam', price: 22000, category: 'Camilan', image: 'https://images.unsplash.com/photo-1546069901-ba9599a7e63c?w=500&auto=format&fit=crop&q=80', description: 'Tahu krispi goreng gurih ditumis bawang putih wangi dan cabai rawit.' },
  { id: '7', name: 'Pisang Goreng Keju', price: 20000, category: 'Penutup', image: 'https://images.unsplash.com/photo-1579954115545-a95591f28bfc?w=500&auto=format&fit=crop&q=80', description: 'Pisang kepok goreng keemasan dengan parutan keju cheddar melimpah.' }
];

export const DEFAULT_USERS: StaffUser[] = [
  { id: '1', username: 'admin', pin: '123456', role: 'admin', name: 'Administrator Utama' },
  { id: '2', username: 'kasir1', pin: '111111', role: 'cashier', name: 'Kasir Shift 1' },
  { id: '3', username: 'kasir2', pin: '222222', role: 'cashier', name: 'Kasir Shift 2' }
];

export const DEFAULT_TABLES: TableItem[] = [
  { id: 'T1', name: 'Meja 1', link: '' },
  { id: 'T2', name: 'Meja 2', link: '' }
];

// Audio chime using Web Audio API
export const playOrderCallSound = () => {
  try {
    const AudioContextClass = window.AudioContext || (window as unknown as { webkitAudioContext: typeof AudioContext }).webkitAudioContext;
    const audioCtx = new AudioContextClass();
    if (audioCtx.state === 'suspended') {
      audioCtx.resume();
    }
    
    const playTone = (freq: number, startTime: number, duration: number) => {
      const osc = audioCtx.createOscillator();
      const gain = audioCtx.createGain();
      
      osc.type = 'sine';
      osc.frequency.setValueAtTime(freq, audioCtx.currentTime + startTime);
      
      gain.gain.setValueAtTime(0, audioCtx.currentTime + startTime);
      gain.gain.linearRampToValueAtTime(0.3, audioCtx.currentTime + startTime + 0.05);
      gain.gain.exponentialRampToValueAtTime(0.001, audioCtx.currentTime + startTime + duration);
      
      osc.connect(gain);
      gain.connect(audioCtx.destination);
      
      osc.start(audioCtx.currentTime + startTime);
      osc.stop(audioCtx.currentTime + startTime + duration);
    };

    playTone(587.33, 0, 0.35); // D5
    playTone(880, 0.22, 0.5);  // A5
  } catch (e) {
    console.warn("AudioContext not supported or blocked", e);
  }
};

let currentAudioPlayer: HTMLAudioElement | null = null;

export const announceOrderCall = (
  customerName: string,
  orderId: string,
  soundEnabled = true,
  voiceType: 'female' | 'male' = 'female',
  template = '',
  tableName = ''
) => {
  if (currentAudioPlayer) {
    currentAudioPlayer.pause();
    currentAudioPlayer.currentTime = 0;
    currentAudioPlayer = null;
  }
  if ('speechSynthesis' in window) {
    window.speechSynthesis.cancel();
  }

  const textTemplate = template || "Pesanan atas nama {nama}, meja {meja}, dengan nomor {id}, sudah selesai dan siap diambil.";
  const cleanName = (customerName || 'Pelanggan').replace(/[^a-zA-Z0-9\s]/g, '').trim();
  const cleanId = (orderId || '').replace(/[^a-zA-Z0-9-]/g, '').trim();
  const cleanTable = (tableName || 'Meja').replace(/[^a-zA-Z0-9\s-]/g, '').trim();

  const speechText = textTemplate
    .replace(/{nama}/gi, cleanName)
    .replace(/{id}/gi, cleanId)
    .replace(/{meja}/gi, cleanTable)
    .replace(/{nomor_meja}/gi, cleanTable)
    .replace(/{table}/gi, cleanTable);

  if (soundEnabled) {
    playOrderCallSound();
  }

  const delay = soundEnabled ? 650 : 50;

  if (voiceType === 'male' && 'speechSynthesis' in window) {
    setTimeout(() => {
      const utterance = new SpeechSynthesisUtterance(speechText);
      utterance.lang = 'id-ID';
      utterance.rate = 0.88;
      utterance.pitch = 0.65;
      
      const voices = window.speechSynthesis.getVoices();
      const idVoices = voices.filter(v => v.lang.toLowerCase().includes('id'));
      if (idVoices.length > 0) {
        const maleVoice = idVoices.find(v => v.name.toLowerCase().includes('male') || v.name.toLowerCase().includes('david') || v.name.toLowerCase().includes('andika')) || idVoices[0];
        utterance.voice = maleVoice;
      }
      window.speechSynthesis.speak(utterance);
    }, delay);
    return;
  }

  const encodedText = encodeURIComponent(speechText);
  const googleTtsUrl = `https://translate.google.com/translate_tts?ie=UTF-8&q=${encodedText}&tl=id&client=tw-ob`;

  const audio = new Audio(googleTtsUrl);
  currentAudioPlayer = audio;

  setTimeout(() => {
    audio.play().catch(err => {
      console.warn("Google TTS audio play failed, falling back to Web Speech API:", err);
      if ('speechSynthesis' in window) {
        const utterance = new SpeechSynthesisUtterance(speechText);
        utterance.lang = 'id-ID';
        utterance.rate = 0.95;
        utterance.pitch = 1.2;
        window.speechSynthesis.speak(utterance);
      }
    });
  }, delay);
};

export const getFallbackImage = (name: string) => {
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="300" height="300" viewBox="0 0 300 300"><rect fill="#f43f5e" width="300" height="300"/><text fill="#ffffff" font-family="sans-serif" font-size="20" font-weight="bold" x="50%" y="50%" dominant-baseline="middle" text-anchor="middle">${name.substring(0, 15)}</text></svg>`;
  return `data:image/svg+xml;base64,${btoa(svg)}`;
};

export function calculateCRC16(str: string): string {
  let crc = 0xFFFF;
  for (let c = 0; c < str.length; c++) {
    crc ^= str.charCodeAt(c) << 8;
    for (let i = 0; i < 8; i++) {
      if (crc & 0x8000) {
        crc = (crc << 1) ^ 0x1021;
      } else {
        crc = crc << 1;
      }
    }
  }
  let hex = (crc & 0xFFFF).toString(16).toUpperCase();
  while (hex.length < 4) hex = '0' + hex;
  return hex;
}

export function generateDynamicQris(staticQrisString: string, amount: number): string {
  if (!staticQrisString || typeof staticQrisString !== 'string') return '';
  let qris = staticQrisString.trim();
  if (qris.length < 10) return qris;

  if (qris.endsWith('6304') && qris.length >= 4) {
    qris = qris.slice(0, -4);
  } else {
    const idx63 = qris.lastIndexOf('6304');
    if (idx63 !== -1) {
      qris = qris.slice(0, idx63);
    }
  }

  if (qris.includes('010211')) {
    qris = qris.replace('010211', '010212');
  } else if (!qris.includes('010212')) {
    qris = '000201010212' + qris.replace('000201', '');
  }

  let cleanQris = '';
  let i = 0;
  while (i < qris.length) {
    const tag = qris.slice(i, i + 2);
    const lenStr = qris.slice(i + 2, i + 4);
    const len = parseInt(lenStr, 10);
    if (isNaN(len)) {
      cleanQris += qris.slice(i);
      break;
    }
    const value = qris.slice(i + 4, i + 4 + len);
    if (tag !== '54') {
      cleanQris += tag + lenStr + value;
    }
    i += 4 + len;
  }

  const amountStr = Math.round(amount).toString();
  const tag54Val = '54' + (amountStr.length < 10 ? '0' + amountStr.length : amountStr.length) + amountStr;

  const idx58 = cleanQris.indexOf('5802ID');
  const idx59 = cleanQris.indexOf('59');
  if (idx58 !== -1) {
    cleanQris = cleanQris.slice(0, idx58) + tag54Val + cleanQris.slice(idx58);
  } else if (idx59 !== -1) {
    cleanQris = cleanQris.slice(0, idx59) + tag54Val + cleanQris.slice(idx59);
  } else {
    cleanQris += tag54Val;
  }

  cleanQris += '6304';
  const crcVal = calculateCRC16(cleanQris);
  return cleanQris + crcVal;
}

export const copyToClipboardText = async (text: string): Promise<boolean> => {
  try {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      await navigator.clipboard.writeText(text);
      return true;
    }
  } catch (err) {
    console.warn("Clipboard API failed, using fallback", err);
  }
  try {
    const textarea = document.createElement("textarea");
    textarea.value = text;
    textarea.style.position = "fixed";
    textarea.style.opacity = "0";
    document.body.appendChild(textarea);
    textarea.focus();
    textarea.select();
    const successful = document.execCommand('copy');
    document.body.removeChild(textarea);
    return successful;
  } catch (err) {
    console.error("Fallback copy failed", err);
    return false;
  }
};

export const formatCurrencyInput = (val: string): string => {
  const numbers = val.replace(/[^0-9]/g, '');
  if (!numbers) return '';
  return Number(numbers).toLocaleString('id-ID');
};

export const parseIndonesianNumber = (val: string | number): number => {
  if (!val) return 0;
  const cleaned = val.toString().replace(/[^0-9]/g, '');
  return Number(cleaned) || 0;
};

export async function generateQrDataUrl(text: string, width = 200): Promise<string> {
  try {
    return await QRCode.toDataURL(text, {
      width,
      margin: 1,
      color: {
        dark: '#000000',
        light: '#ffffff'
      }
    });
  } catch (err) {
    console.error('QR code generation error:', err);
    return '';
  }
}
