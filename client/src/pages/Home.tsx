import { useEffect, useMemo, useState } from "react";
import {
  Archive,
  ArrowDownToLine,
  BarChart3,
  Bell,
  BookOpen,
  CalendarDays,
  Check,
  ChevronLeft,
  CircleDollarSign,
  ClipboardList,
  Coffee,
  Download,
  FileSpreadsheet,
  FileText,
  History,
  LayoutDashboard,
  Menu,
  MoreHorizontal,
  Plus,
  Printer,
  ReceiptText,
  Search,
  Settings2,
  ShieldCheck,
  ShoppingBag,
  Sparkles,
  Trash2,
  UserRound,
  UsersRound,
  Utensils,
  X,
} from "lucide-react";
import { toast } from "sonner";
import XLSX from "xlsx-js-style";

type Meal = { id: string; name: string; price: number; active: boolean };
type Employee = { id: string; name: string; active: boolean };
type OrderItem = { mealId: string; mealName: string; quantity: number; unitPrice: number };
type Order = {
  id: string;
  invoiceNo: string;
  createdAt: string;
  items: OrderItem[];
  participants: string[];
  total: number;
  note: string;
};
type AppState = { meals: Meal[]; employees: Employee[]; orders: Order[]; monthlyDiscounts: Record<string, Record<string, number>> };
type View = "overview" | "orders" | "meals" | "employees" | "reports" | "invoices";

type BeforeInstallPromptEvent = Event & {
  prompt: () => Promise<void>;
  userChoice: Promise<{ outcome: "accepted" | "dismissed" }>;
};

const STORAGE_KEY = "alrimi-cafeteria-pwa.v1";
const currency = new Intl.NumberFormat("ar-YE", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const dateFormatter = new Intl.DateTimeFormat("ar-YE", { day: "numeric", month: "short", year: "numeric" });
const timeFormatter = new Intl.DateTimeFormat("ar-YE", { hour: "2-digit", minute: "2-digit" });
const monthFormatter = new Intl.DateTimeFormat("ar-YE", { month: "long", year: "numeric" });

const defaultEmployees = [
  "أحمد محمد أحمد المختار", "أحمد يحيى أحمد الظيلمي", "أنور محمد حميد الجبرة", "أيمن محمد عبد الله القران", "بكيل أحمد علي ملطش", "جميل علي أحمد الزعري", "حسن عباس محمد شرحه", "خالد عثمان مهيوب الدبعي", "خالد يحيى صالح المشرة", "سلطان غالب محمد العصيمي", "سمير محمد عبده الحكيمي", "صالح علي حسين المثلي", "عادل أحمد أحمد عيسى", "عبدالله أحمد عبدالوهاب القباطي", "علي حسين حسين الرخومي", "علي حمود عبدالله سنبل", "علي يحيى عبده الربيعه", "عماد عياش أحمد الحريشي", "عمر عبدالله علي المقطري", "فواز حسن صالح سالم", "نبيل قايد عبدالكريم القدسي", "ياسين عبده فارع الحكيمي", "أحمد صالح علي السفياني", "علي أحمد علي ناقلة", "أنور صالح حسين الزريقي", "خالد محمد يحيى الخزان", "سمير أحمد محمد مثنى", "ماجد محمد حسن الشيباني", "محمد شمسان عبدالله عثمان", "محمد علي حسين الرخومي", "مصباح محمد علي البدوي", "خالد محمد أحمد الزهرات", "وائل محمد أحمد المطعمي", "علي صالح علي المخارشي", "أحمد محمد مهدي المطري", "محمد يحيى محمد نعمان", "هاني عبدالله عبدالله البريطي", "أبو بكر عبدالله ناجي سعيد", "عبدالله عبدالرزاق حيدر الصلوي", "حاشدي محمد صديق خلوفه", "علي الأشموري", "محمد المطري", "يوسف القبيسي", "شكري الجعفري", "صخر الكباب", "حازم الأشموري", "ياسر الأصبحي", "عبدالرحمن الريمي",
];

const seedState: AppState = {
  meals: [
    { id: "meal-breakfast", name: "فطور صباحي", price: 18, active: true },
    { id: "meal-lunch", name: "وجبة غداء", price: 28, active: true },
    { id: "meal-special", name: "وجبة خاصة", price: 35, active: true },
  ],
  employees: defaultEmployees.map((name, index) => ({ id: `emp-${index + 1}`, name, active: true })),
  orders: [],
  monthlyDiscounts: {},
};

const navItems: { id: View; label: string; icon: typeof LayoutDashboard }[] = [
  { id: "overview", label: "نظرة عامة", icon: LayoutDashboard },
  { id: "orders", label: "الطلبات اليومية", icon: ClipboardList },
  { id: "meals", label: "الوجبات والأسعار", icon: Utensils },
  { id: "employees", label: "الموظفون", icon: UsersRound },
  { id: "reports", label: "التقارير الشهرية", icon: BarChart3 },
  { id: "invoices", label: "الفواتير", icon: ReceiptText },
];

function makeId(prefix: string) {
  return `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}
function money(value: number) { return `${currency.format(value)} ر.ي`; }
function formatDate(value: string) { return dateFormatter.format(new Date(value)); }
function formatTime(value: string) { return timeFormatter.format(new Date(value)); }
function currentMonth() { const now = new Date(); return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`; }
function monthLabel(value: string) { return monthFormatter.format(new Date(`${value}-01T12:00:00`)); }
function loadState(): AppState {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return seedState;
    const parsed = JSON.parse(raw) as Partial<AppState>;
    return {
      meals: parsed.meals?.length ? parsed.meals : seedState.meals,
      employees: parsed.employees?.length ? parsed.employees : seedState.employees,
      orders: parsed.orders ?? [],
      monthlyDiscounts: parsed.monthlyDiscounts ?? {},
    };
  } catch { return seedState; }
}
function downloadFile(content: string, filename: string, type = "text/plain;charset=utf-8") {
  const blob = new Blob(["\ufeff", content], { type });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url; anchor.download = filename; anchor.click(); URL.revokeObjectURL(url);
}
function excelCellStyle(kind: "header" | "body" | "total" | "date") {
  const fill = kind === "header" ? "17201D" : kind === "total" ? "FFF0E8" : "FFFFFF";
  return {
    fill: { patternType: "solid", fgColor: { rgb: fill } },
    font: { name: "Arial", sz: kind === "header" ? 11 : 10, bold: kind === "header" || kind === "total", color: { rgb: kind === "header" ? "FFFFFF" : "17201D" } },
    alignment: { horizontal: "center", vertical: "center", wrapText: kind === "header" },
    border: { top: { style: "thin", color: { rgb: "D9E1DA" } }, bottom: { style: "thin", color: { rgb: "D9E1DA" } }, left: { style: "thin", color: { rgb: "D9E1DA" } }, right: { style: "thin", color: { rgb: "D9E1DA" } } },
    numFmt: kind === "date" ? "dd/mm/yyyy" : kind === "body" || kind === "total" ? "0.00" : "General",
  };
}
function setPrintReadySheet(ws: XLSX.WorkSheet, lastCell: string, firstColWidth = 30) {
  ws["!cols"] = [{ wch: firstColWidth }, ...Array.from({ length: 31 }, () => ({ wch: 12 }))];
  ws["!freeze"] = { xSplit: 1, ySplit: 1 };
  ws["!autofilter"] = { ref: `A1:${lastCell}` };
  ws["!pageSetup"] = { orientation: "landscape", fitToWidth: 1, fitToHeight: 0, paperSize: 9 };
  ws["!margins"] = { left: 0.25, right: 0.25, top: 0.45, bottom: 0.45, header: 0.2, footer: 0.2 };
  ws["!printHeader"] = "1:1";
}
function monthlyMatrix(month: string, orders: Order[], employees: Employee[]) {
  const [year, monthNumber] = month.split("-").map(Number);
  const daysInMonth = new Date(year, monthNumber, 0).getDate();
  const dates = Array.from({ length: daysInMonth }, (_, index) => new Date(year, monthNumber - 1, index + 1));
  const values = new Map<string, number>();
  orders.forEach((order) => {
    const day = Number(order.createdAt.slice(8, 10));
    const share = order.participants.length ? order.total / order.participants.length : 0;
    order.participants.forEach((employeeId) => values.set(`${employeeId}-${day}`, (values.get(`${employeeId}-${day}`) ?? 0) + share));
  });
  const rows = employees.map((employee) => {
    const daily = dates.map((_, index) => values.get(`${employee.id}-${index + 1}`) ?? 0);
    return { employee, daily, total: daily.reduce((sum, value) => sum + value, 0) };
  });
  return { dates, rows };
}
function createMonthlyWorkbook(month: string, orders: Order[], employees: Employee[]) {
  const { dates, rows } = monthlyMatrix(month, orders, employees);
  const header = ["الموظف", ...dates.map((date) => `${String(date.getDate()).padStart(2, "0")}/${String(date.getMonth() + 1).padStart(2, "0")}/${date.getFullYear()}`), "الاجمالي"];
  const data = [header, ...rows.map((row) => [row.employee.name, ...row.daily, row.total])];
  const ws = XLSX.utils.aoa_to_sheet(data);
  const lastColumn = dates.length + 2;
  const lastCell = `${XLSX.utils.encode_col(lastColumn - 1)}${data.length}`;
  data.forEach((row, rowIndex) => row.forEach((_, colIndex) => {
    const cell = ws[XLSX.utils.encode_cell({ r: rowIndex, c: colIndex })];
    if (!cell) return;
    cell.s = rowIndex === 0 ? excelCellStyle("header") : colIndex === 0 ? excelCellStyle("body") : colIndex === lastColumn - 1 ? excelCellStyle("total") : excelCellStyle("body");
    if (rowIndex === 0 && colIndex > 0 && colIndex < lastColumn - 1) { cell.s = excelCellStyle("header"); }
  }));
  setPrintReadySheet(ws, lastCell);
  ws["!rows"] = [{ hpt: 24 }, ...rows.map(() => ({ hpt: 20 }))];
  const detailHeader = ["رقم الفاتورة", "التاريخ", "الأصناف", "المشاركون", "الإجمالي", "نصيب الموظف"];
  const detailRows = orders.map((order) => [order.invoiceNo, formatDate(order.createdAt), order.items.map((item) => `${item.mealName} × ${item.quantity}`).join("، "), order.participants.map((id) => employees.find((employee) => employee.id === id)?.name ?? "").join("، "), order.total, order.participants.length ? order.total / order.participants.length : 0]);
  const details = XLSX.utils.aoa_to_sheet([detailHeader, ...detailRows]);
  [detailHeader, ...detailRows].forEach((row, rowIndex) => row.forEach((_, colIndex) => {
    const cell = details[XLSX.utils.encode_cell({ r: rowIndex, c: colIndex })];
    if (!cell) return;
    cell.s = rowIndex === 0 ? excelCellStyle("header") : colIndex === 1 ? excelCellStyle("date") : colIndex >= 4 ? excelCellStyle("body") : excelCellStyle("body");
  }));
  details["!cols"] = [{ wch: 24 }, { wch: 15 }, { wch: 30 }, { wch: 55 }, { wch: 15 }, { wch: 15 }];
  details["!pageSetup"] = { orientation: "landscape", fitToWidth: 1, fitToHeight: 0, paperSize: 9 };
  details["!autofilter"] = { ref: `A1:F${detailRows.length + 1}` };
  return { workbook: XLSX.utils.book_new(), ws, details };
}
function exportMonthlyExcel(month: string, orders: Order[], employees: Employee[]) {
  const { workbook, ws, details } = createMonthlyWorkbook(month, orders, employees);
  XLSX.utils.book_append_sheet(workbook, ws, `كشف ${month.replace("-", "_")}`);
  XLSX.utils.book_append_sheet(workbook, details, "تفاصيل الطلبات");
  XLSX.writeFile(workbook, `كشف-حساب-${month}.xlsx`);
}
function employeeDetails(month: string, employee: Employee, orders: Order[]) {
  return orders.filter((order) => order.participants.includes(employee.id)).flatMap((order) => {
    const shareFactor = order.participants.length || 1;
    return order.items.map((item) => [order.invoiceNo, formatDate(order.createdAt), item.mealName, item.quantity, item.unitPrice * item.quantity, (item.unitPrice * item.quantity) / shareFactor]);
  });
}
function exportEmployeeExcel(month: string, employee: Employee, orders: Order[], discount: number) {
  const detailHeader = ["رقم الفاتورة", "التاريخ", "الوجبة", "الكمية", "قيمة الوجبة", "نصيب الموظف"];
  const details = employeeDetails(month, employee, orders);
  const subtotal = details.reduce((sum, row) => sum + Number(row[5]), 0);
  const summary = [["اسم الموظف", employee.name], ["الشهر", monthLabel(month)], ["إجمالي نصيب الوجبات", subtotal], ["الخصم الشهري", discount], ["صافي المستحق", subtotal - discount]];
  const ws = XLSX.utils.aoa_to_sheet([["كشف حساب موظف مفصل"], ...summary, [], detailHeader, ...details]);
  ws["A1"].s = { ...excelCellStyle("header"), font: { ...excelCellStyle("header").font, sz: 15 } };
  ["A2", "A3", "A4", "A5", "A6"].forEach((cell) => { if (ws[cell]) ws[cell].s = excelCellStyle("body"); });
  ["B2", "B3"].forEach((cell) => { if (ws[cell]) ws[cell].s = excelCellStyle("body"); });
  ["B4", "B5", "B6"].forEach((cell) => { if (ws[cell]) ws[cell].s = excelCellStyle("total"); });
  const headerRow = 8;
  detailHeader.forEach((_, colIndex) => { const cell = ws[XLSX.utils.encode_cell({ r: headerRow - 1, c: colIndex })]; if (cell) cell.s = excelCellStyle("header"); });
  details.forEach((row, rowIndex) => row.forEach((_, colIndex) => { const cell = ws[XLSX.utils.encode_cell({ r: headerRow + rowIndex, c: colIndex })]; if (cell) cell.s = excelCellStyle(colIndex >= 4 ? "body" : "body"); }));
  ws["!cols"] = [{ wch: 24 }, { wch: 16 }, { wch: 28 }, { wch: 12 }, { wch: 16 }, { wch: 16 }];
  ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 5 } }];
  ws["!pageSetup"] = { orientation: "portrait", fitToWidth: 1, fitToHeight: 0, paperSize: 9 };
  ws["!autofilter"] = { ref: `A${headerRow}:F${headerRow + details.length}` };
  const workbook = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(workbook, ws, "كشف الموظف"); XLSX.writeFile(workbook, `كشف-${employee.name}-${month}.xlsx`);
}
function buildInvoice(order: Order, employees: Employee[]) {
  const names = order.participants.map((id) => employees.find((employee) => employee.id === id)?.name ?? "موظف");
  return [
    "نظام بوفيه يدكو", "فاتورة طلب وجبة", "-----------------------------", `رقم الفاتورة: ${order.invoiceNo}`,
    `التاريخ: ${formatDate(order.createdAt)} - ${formatTime(order.createdAt)}`, "", ...order.items.map((item) => `${item.mealName} × ${item.quantity} = ${money(item.unitPrice * item.quantity)}`),
    "-----------------------------", `الإجمالي: ${money(order.total)}`, `عدد المشاركين: ${names.length}`, "", "المشاركون:", ...names.map((name) => `- ${name}`), order.note ? `\nملاحظة: ${order.note}` : "", "", "شكرًا لكم",
  ].join("\n");
}

function EmptyState({ icon: Icon, title, description, action }: { icon: typeof Archive; title: string; description: string; action?: React.ReactNode }) {
  return <div className="empty-state"><div className="empty-icon"><Icon size={22} /></div><h3>{title}</h3><p>{description}</p>{action}</div>;
}

export default function Home() {
  const [state, setState] = useState<AppState>(() => loadState());
  const [view, setView] = useState<View>("overview");
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [query, setQuery] = useState("");
  const [installPrompt, setInstallPrompt] = useState<BeforeInstallPromptEvent | null>(null);
  const [showOrderModal, setShowOrderModal] = useState(false);
  const [showMealModal, setShowMealModal] = useState(false);
  const [showEmployeeModal, setShowEmployeeModal] = useState(false);
  const [draftItems, setDraftItems] = useState<OrderItem[]>([]);
  const [draftMealId, setDraftMealId] = useState("");
  const [draftQty, setDraftQty] = useState(1);
  const [draftNote, setDraftNote] = useState("");
  const [draftParticipants, setDraftParticipants] = useState<string[]>([]);
  const [mealName, setMealName] = useState("");
  const [mealPrice, setMealPrice] = useState("");
  const [employeeName, setEmployeeName] = useState("");
  const [reportMonth, setReportMonth] = useState(currentMonth());

  useEffect(() => { localStorage.setItem(STORAGE_KEY, JSON.stringify(state)); }, [state]);
  useEffect(() => {
    const onInstall = (event: Event) => { event.preventDefault(); setInstallPrompt(event as BeforeInstallPromptEvent); };
    window.addEventListener("beforeinstallprompt", onInstall);
    if ("serviceWorker" in navigator) navigator.serviceWorker.register("/sw.js").catch(() => undefined);
    return () => window.removeEventListener("beforeinstallprompt", onInstall);
  }, []);

  const activeMeals = state.meals.filter((meal) => meal.active);
  const activeEmployees = state.employees.filter((employee) => employee.active);
  const monthOrders = useMemo(() => state.orders.filter((order) => order.createdAt.slice(0, 7) === reportMonth), [reportMonth, state.orders]);
  const todayOrders = state.orders.filter((order) => order.createdAt.slice(0, 10) === new Date().toISOString().slice(0, 10));
  const todayTotal = todayOrders.reduce((sum, order) => sum + order.total, 0);
  const monthTotal = monthOrders.reduce((sum, order) => sum + order.total, 0);
  const filteredEmployees = activeEmployees.filter((employee) => employee.name.includes(query));
  const filteredMeals = activeMeals.filter((meal) => meal.name.includes(query));
  const recentOrders = [...state.orders].sort((a, b) => b.createdAt.localeCompare(a.createdAt)).slice(0, 6);
  const draftTotal = draftItems.reduce((sum, item) => sum + item.unitPrice * item.quantity, 0);

  const openView = (next: View) => { setView(next); setSidebarOpen(false); setQuery(""); };
  const updateState = (updater: (previous: AppState) => AppState) => setState((previous) => updater(previous));
  const notify = (message: string) => toast.success(message);

  const addMeal = () => {
    const price = Number(mealPrice);
    if (!mealName.trim() || !Number.isFinite(price) || price <= 0) { toast.error("أدخل اسم الوجبة والسعر بشكل صحيح"); return; }
    updateState((previous) => ({ ...previous, meals: [...previous.meals, { id: makeId("meal"), name: mealName.trim(), price, active: true }] }));
    setMealName(""); setMealPrice(""); setShowMealModal(false); notify("تمت إضافة الوجبة بنجاح");
  };
  const addEmployee = () => {
    if (!employeeName.trim()) { toast.error("اكتب اسم الموظف أولًا"); return; }
    updateState((previous) => ({ ...previous, employees: [...previous.employees, { id: makeId("emp"), name: employeeName.trim(), active: true }] }));
    setEmployeeName(""); setShowEmployeeModal(false); notify("تمت إضافة الموظف بنجاح");
  };
  const removeMeal = (id: string) => updateState((previous) => ({ ...previous, meals: previous.meals.filter((meal) => meal.id !== id) }));
  const removeEmployee = (id: string) => updateState((previous) => ({ ...previous, employees: previous.employees.filter((employee) => employee.id !== id) }));
  const addDraftItem = () => {
    const meal = activeMeals.find((item) => item.id === draftMealId);
    if (!meal || draftQty < 1) { toast.error("اختر الوجبة والكمية"); return; }
    setDraftItems((previous) => {
      const existing = previous.find((item) => item.mealId === meal.id);
      if (existing) return previous.map((item) => item.mealId === meal.id ? { ...item, quantity: item.quantity + draftQty } : item);
      return [...previous, { mealId: meal.id, mealName: meal.name, quantity: draftQty, unitPrice: meal.price }];
    });
    setDraftQty(1); notify("أضيفت الوجبة إلى الطلب");
  };
  const createOrder = () => {
    if (!draftItems.length || !draftParticipants.length) { toast.error("أضف وجبة واحدة واختر موظفًا واحدًا على الأقل"); return; }
    const now = new Date().toISOString();
    const dayCode = now.slice(0, 10).replaceAll("-", "");
    const dayCount = state.orders.filter((order) => order.createdAt.slice(0, 10) === now.slice(0, 10)).length + 1;
    const order: Order = { id: makeId("order"), invoiceNo: `INV-${dayCode}-${String(dayCount).padStart(3, "0")}`, createdAt: now, items: draftItems, participants: draftParticipants, total: draftTotal, note: draftNote.trim() };
    updateState((previous) => ({ ...previous, orders: [...previous.orders, order] }));
    setDraftItems([]); setDraftParticipants([]); setDraftNote(""); setShowOrderModal(false); notify(`تم تسجيل الطلب ${order.invoiceNo}`);
    setView("orders");
  };
  const exportOrders = () => {
    const details = state.orders.map((order) => [order.invoiceNo, formatDate(order.createdAt), order.items.map((item) => `${item.mealName} × ${item.quantity}`).join("، "), order.participants.map((id) => state.employees.find((employee) => employee.id === id)?.name ?? "").join("، "), order.total, order.participants.length ? order.total / order.participants.length : 0, order.note]);
    const header = ["رقم الفاتورة", "التاريخ", "الأصناف", "المشاركون", "الإجمالي", "نصيب الموظف", "الملاحظة"];
    const ws = XLSX.utils.aoa_to_sheet([header, ...details]);
    [header, ...details].forEach((row, rowIndex) => row.forEach((_, colIndex) => { const cell = ws[XLSX.utils.encode_cell({ r: rowIndex, c: colIndex })]; if (cell) cell.s = rowIndex === 0 ? excelCellStyle("header") : colIndex === 1 ? excelCellStyle("date") : colIndex >= 4 && colIndex <= 5 ? excelCellStyle("body") : excelCellStyle("body"); }));
    ws["!cols"] = [{ wch: 25 }, { wch: 16 }, { wch: 30 }, { wch: 55 }, { wch: 16 }, { wch: 16 }, { wch: 28 }];
    ws["!pageSetup"] = { orientation: "landscape", fitToWidth: 1, fitToHeight: 0, paperSize: 9 };
    ws["!autofilter"] = { ref: `A1:G${details.length + 1}` };
    const workbook = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(workbook, ws, "سجل الطلبات"); XLSX.writeFile(workbook, `سجل-الطلبات-${reportMonth}.xlsx`); notify("تم تصدير سجل الطلبات إلى Excel");
  };
  const exportReport = () => { exportMonthlyExcel(reportMonth, monthOrders, activeEmployees); notify("تم تصدير كشف الشهر إلى Excel"); };
  const setEmployeeDiscount = (employee: Employee) => {
    const current = state.monthlyDiscounts[reportMonth]?.[employee.id] ?? 0;
    const value = window.prompt(`أدخل خصم ${employee.name} لشهر ${monthLabel(reportMonth)} بالريال اليمني`, String(current));
    if (value === null) return;
    const discount = Number(value);
    if (!Number.isFinite(discount) || discount < 0) { toast.error("أدخل قيمة خصم صحيحة"); return; }
    updateState((previous) => ({ ...previous, monthlyDiscounts: { ...previous.monthlyDiscounts, [reportMonth]: { ...(previous.monthlyDiscounts[reportMonth] ?? {}), [employee.id]: discount } } }));
    notify("تم حفظ الخصم الشهري");
  };
  const exportEmployeeReport = (employee: Employee) => { exportEmployeeExcel(reportMonth, employee, monthOrders, state.monthlyDiscounts[reportMonth]?.[employee.id] ?? 0); notify(`تم تصدير كشف ${employee.name}`); };
  const printMonthlyReport = () => {
    const { dates, rows } = monthlyMatrix(reportMonth, monthOrders, activeEmployees);
    const tableRows = rows.map((row) => `<tr><th>${row.employee.name}</th>${row.daily.map((value) => `<td>${value ? value.toFixed(2) : ""}</td>`).join("")}<th>${row.total.toFixed(2)}</th></tr>`).join("");
    const headings = dates.map((date) => `<th>${date.getDate()}</th>`).join("");
    const printWindow = window.open("", "_blank", "width=1200,height=800");
    if (!printWindow) return;
    printWindow.document.write(`<html dir="rtl"><head><title>كشف حساب ${monthLabel(reportMonth)}</title><style>@page{size:landscape;margin:8mm}body{font-family:Arial,Tahoma,sans-serif;color:#17201d}h1{text-align:center;margin:0 0 4px;font-size:20px}p{text-align:center;margin:0 0 14px;color:#66756b;font-size:12px}table{border-collapse:collapse;width:100%;font-size:9px;table-layout:fixed;direction:rtl}th,td{border:1px solid #b8c3ba;text-align:center;padding:5px;white-space:nowrap}thead th{background:#17201d;color:#fff}tbody th:first-child{width:170px;text-align:right;background:#fff0e8}tbody th:last-child{background:#fff0e8}</style></head><body><h1>نظام بوفيه يدكو</h1><p>كشف حساب الموظفين — ${monthLabel(reportMonth)} — المبالغ بالريال اليمني</p><table><thead><tr><th>الموظف</th>${headings}<th>الاجمالي</th></tr></thead><tbody>${tableRows}</tbody></table><script>window.print();</script></body></html>`);
    printWindow.document.close();
  };
  const printInvoice = (order: Order) => {
    const invoice = buildInvoice(order, state.employees).replaceAll("\n", "<br />");
    const printWindow = window.open("", "_blank", "width=760,height=700");
    if (!printWindow) return;
    printWindow.document.write(`<html dir="rtl"><head><title>${order.invoiceNo}</title><style>body{font-family:Arial,sans-serif;padding:36px;color:#17201d;line-height:1.9}h1{color:#e56b2f} .line{border-top:1px solid #ddd;margin:18px 0}</style></head><body><h1>نظام بوفيه يدكو</h1><div class="line"></div>${invoice}<script>window.print();</script></body></html>`);
    printWindow.document.close();
  };
  const downloadInvoice = (order: Order) => {
    const invoice = buildInvoice(order, state.employees).replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll("\n", "<br />");
    const html = `<html dir="rtl" lang="ar"><head><meta charset="utf-8"><title>${order.invoiceNo}</title><style>body{font-family:Arial,Tahoma,sans-serif;max-width:760px;margin:40px auto;padding:24px;color:#17201d;line-height:2;border:1px solid #d9e1da;border-radius:16px}h1{color:#e56b2f;font-size:24px}.line{border-top:1px solid #d9e1da;margin:18px 0}@media print{body{margin:0;border:0}}</style></head><body><h1>نظام بوفيه يدكو</h1><div class="line"></div>${invoice}</body></html>`;
    downloadFile(html, `${order.invoiceNo}.html`, "text/html;charset=utf-8");
    notify("تم تحميل الفاتورة بصيغة قابلة للطباعة");
  };
  const clearData = () => { if (!window.confirm("سيتم حذف الطلبات فقط من هذا الجهاز. هل تريد المتابعة؟")) return; updateState((previous) => ({ ...previous, orders: [] })); notify("تم مسح سجل الطلبات"); };
  const installApp = async () => { if (!installPrompt) { toast.info("استخدم قائمة المتصفح ثم اختر تثبيت التطبيق"); return; } await installPrompt.prompt(); setInstallPrompt(null); };

  return <div className="app-shell" dir="rtl">
    <aside className={`sidebar ${sidebarOpen ? "is-open" : ""}`}>
      <div className="brand"><div className="brand-mark"><Coffee size={22} /></div><div><strong>يدكو</strong><span>نظام البوفيه</span></div></div>
      <div className="sidebar-label">القائمة الرئيسية</div>
      <nav>{navItems.map(({ id, label, icon: Icon }) => <button key={id} className={`nav-item ${view === id ? "active" : ""}`} onClick={() => openView(id)}><Icon size={18} /><span>{label}</span>{id === "orders" && todayOrders.length > 0 && <b className="nav-count">{todayOrders.length}</b>}</button>)}</nav>
      <div className="sidebar-spacer" />
      <div className="install-card"><div className="install-spark"><Sparkles size={17} /></div><strong>ثبّت التطبيق</strong><p>وصول أسرع من شاشة جهازك دون فتح المتصفح.</p><button onClick={installApp}><Download size={15} /> تثبيت الآن</button></div>
      <div className="sidebar-foot"><ShieldCheck size={14} /> بياناتك محفوظة على هذا الجهاز</div>
    </aside>
    {sidebarOpen && <button className="sidebar-overlay" aria-label="إغلاق القائمة" onClick={() => setSidebarOpen(false)} />}
    <main className="main-content">
      <header className="topbar"><div className="topbar-start"><button className="icon-btn menu-btn" onClick={() => setSidebarOpen(true)} aria-label="فتح القائمة"><Menu size={21} /></button><div><div className="eyebrow">مساحة الإدارة / {navItems.find((item) => item.id === view)?.label}</div><h1>{view === "overview" ? "صباح الخير، عبدالرحمن" : navItems.find((item) => item.id === view)?.label}</h1></div></div><div className="topbar-actions"><div className="live-status"><i /> يعمل محليًا</div><button className="icon-btn" aria-label="الإشعارات"><Bell size={18} /></button><div className="user-avatar">ع</div></div></header>
      <div className="page-body">
        {view === "overview" && <OverviewView todayOrders={todayOrders} todayTotal={todayTotal} monthTotal={monthTotal} activeEmployees={activeEmployees} activeMeals={activeMeals} recentOrders={recentOrders} onNewOrder={() => setShowOrderModal(true)} onViewOrders={() => openView("orders")} onViewMeals={() => openView("meals")} onPrint={printInvoice} state={state} />}
        {view === "orders" && <OrdersView orders={[...state.orders].sort((a, b) => b.createdAt.localeCompare(a.createdAt))} employees={state.employees} query={query} setQuery={setQuery} onNewOrder={() => setShowOrderModal(true)} onExport={exportOrders} onPrint={printInvoice} onClear={clearData} />}
        {view === "meals" && <MealsView meals={filteredMeals} query={query} setQuery={setQuery} onAdd={() => setShowMealModal(true)} onRemove={removeMeal} />}
        {view === "employees" && <EmployeesView employees={filteredEmployees} query={query} setQuery={setQuery} onAdd={() => setShowEmployeeModal(true)} onRemove={removeEmployee} />}
        {view === "reports" && <ReportsView month={reportMonth} setMonth={setReportMonth} orders={monthOrders} employees={activeEmployees} total={monthTotal} discounts={state.monthlyDiscounts[reportMonth] ?? {}} onExport={exportReport} onPrint={printMonthlyReport} onSetDiscount={setEmployeeDiscount} onExportEmployee={exportEmployeeReport} />}
        {view === "invoices" && <InvoicesView orders={[...state.orders].sort((a, b) => b.createdAt.localeCompare(a.createdAt))} employees={state.employees} onPrint={printInvoice} onDownload={downloadInvoice} />}
      </div>
      <footer className="app-footer"><span>نظام بوفيه يدكو</span><span>تصميم وبرمجة المهندس / عبدالرحمن الريمي</span><span>آخر حفظ تلقائي</span></footer>
    </main>

    {showOrderModal && <Modal title="تسجيل طلب جديد" icon={<ShoppingBag size={18} />} onClose={() => setShowOrderModal(false)} wide><div className="modal-grid"><div><div className="field-label">أصناف الطلب</div><div className="add-item-row"><select value={draftMealId} onChange={(event) => setDraftMealId(event.target.value)}><option value="">اختر الوجبة</option>{activeMeals.map((meal) => <option key={meal.id} value={meal.id}>{meal.name} — {money(meal.price)}</option>)}</select><input type="number" min="1" value={draftQty} onChange={(event) => setDraftQty(Number(event.target.value))} /><button className="secondary-btn compact" onClick={addDraftItem}><Plus size={16} /> إضافة</button></div><div className="draft-list">{draftItems.length ? draftItems.map((item) => <div className="draft-row" key={item.mealId}><span>{item.mealName}</span><span>{item.quantity} × {money(item.unitPrice)}</span><button className="remove-btn" onClick={() => setDraftItems((items) => items.filter((draft) => draft.mealId !== item.mealId))}><X size={15} /></button></div>) : <div className="inline-empty">لم تتم إضافة أصناف بعد</div>}</div><div className="draft-total"><span>إجمالي الطلب</span><strong>{money(draftTotal)}</strong></div><label className="field-label" htmlFor="note">ملاحظة الطلب <span>اختياري</span></label><textarea id="note" value={draftNote} onChange={(event) => setDraftNote(event.target.value)} placeholder="أضف أي ملاحظة للمطبخ..." /></div><div><div className="field-label">الموظفون المشاركون <span>اختر واحدًا أو أكثر</span></div><div className="employee-picker">{activeEmployees.map((employee) => <label key={employee.id} className={`check-row ${draftParticipants.includes(employee.id) ? "checked" : ""}`}><input type="checkbox" checked={draftParticipants.includes(employee.id)} onChange={() => setDraftParticipants((ids) => ids.includes(employee.id) ? ids.filter((id) => id !== employee.id) : [...ids, employee.id])} /><span>{employee.name}</span>{draftParticipants.includes(employee.id) && <Check size={15} />}</label>)}</div><div className="selected-count">تم اختيار {draftParticipants.length} موظف</div></div></div><div className="modal-actions"><button className="ghost-btn" onClick={() => setShowOrderModal(false)}>إلغاء</button><button className="primary-btn" onClick={createOrder}><Check size={17} /> حفظ وتسجيل الطلب</button></div></Modal>}
    {showMealModal && <Modal title="إضافة وجبة" icon={<Utensils size={18} />} onClose={() => setShowMealModal(false)}><label className="field-label">اسم الوجبة</label><input autoFocus value={mealName} onChange={(event) => setMealName(event.target.value)} placeholder="مثال: وجبة متوسطة" /><label className="field-label">السعر <span>بالريال اليمني</span></label><input type="number" min="0" step="0.5" value={mealPrice} onChange={(event) => setMealPrice(event.target.value)} placeholder="0.00" /><div className="modal-actions"><button className="ghost-btn" onClick={() => setShowMealModal(false)}>إلغاء</button><button className="primary-btn" onClick={addMeal}><Plus size={17} /> إضافة الوجبة</button></div></Modal>}
    {showEmployeeModal && <Modal title="إضافة موظف" icon={<UserRound size={18} />} onClose={() => setShowEmployeeModal(false)}><label className="field-label">اسم الموظف</label><input autoFocus value={employeeName} onChange={(event) => setEmployeeName(event.target.value)} placeholder="اكتب اسم الموظف" /><div className="modal-actions"><button className="ghost-btn" onClick={() => setShowEmployeeModal(false)}>إلغاء</button><button className="primary-btn" onClick={addEmployee}><Plus size={17} /> إضافة الموظف</button></div></Modal>}
  </div>;
}

function OverviewView({ todayOrders, todayTotal, monthTotal, activeEmployees, activeMeals, recentOrders, onNewOrder, onViewOrders, onViewMeals, onPrint, state }: { todayOrders: Order[]; todayTotal: number; monthTotal: number; activeEmployees: Employee[]; activeMeals: Meal[]; recentOrders: Order[]; onNewOrder: () => void; onViewOrders: () => void; onViewMeals: () => void; onPrint: (order: Order) => void; state: AppState }) {
  return <div className="view-stack"><section className="hero-panel"><div><div className="hero-kicker"><Sparkles size={14} /> لوحة التشغيل اليومية</div><h2>كل ما تحتاجه لإدارة<br /><em>بوفيه يدكو</em> في مكان واحد.</h2><p>سجّل الطلبات، وزّع التكلفة على الموظفين، واستخرج التقارير خلال ثوانٍ.</p><button className="hero-btn" onClick={onNewOrder}><Plus size={17} /> تسجيل طلب جديد <ChevronLeft size={16} /></button></div><div className="hero-graphic"><div className="hero-orbit orbit-one" /><div className="hero-orbit orbit-two" /><div className="hero-plate"><Utensils size={58} /></div><div className="floating-stat"><span>طلبات اليوم</span><strong>{todayOrders.length}</strong><small>طلب نشط</small></div></div></section><section className="stats-grid"><StatCard icon={<ShoppingBag />} label="طلبات اليوم" value={String(todayOrders.length)} detail={todayOrders.length ? "يتم تحديثها لحظيًا" : "ابدأ بتسجيل أول طلب"} tone="orange" /><StatCard icon={<CircleDollarSign />} label="إجمالي اليوم" value={money(todayTotal)} detail="قيمة الطلبات المسجلة" tone="teal" /><StatCard icon={<UsersRound />} label="الموظفون" value={String(activeEmployees.length)} detail="موظف في القائمة" tone="purple" /><StatCard icon={<BarChart3 />} label="إجمالي الشهر" value={money(monthTotal)} detail="من بداية الشهر الحالي" tone="blue" /></section><div className="content-grid"><section className="panel recent-panel"><PanelHeading icon={<History size={17} />} title="آخر الطلبات" action={<button className="text-btn" onClick={onViewOrders}>عرض الكل <ChevronLeft size={15} /></button>} />{recentOrders.length ? <div className="order-table">{recentOrders.map((order) => <OrderRow key={order.id} order={order} employees={state.employees} onPrint={onPrint} />)}</div> : <EmptyState icon={ClipboardList} title="لا توجد طلبات حتى الآن" description="ابدأ بتسجيل طلب وجبة جديد ليظهر هنا." action={<button className="secondary-btn" onClick={onNewOrder}><Plus size={15} /> تسجيل طلب</button>} />}</section><section className="panel quick-panel"><PanelHeading icon={<Sparkles size={17} />} title="اختصارات سريعة" /><div className="quick-actions"><button onClick={onNewOrder}><div className="quick-icon orange"><Plus size={19} /></div><span><strong>طلب وجبة جديد</strong><small>تسجيل وتوزيع التكلفة</small></span><ChevronLeft size={17} /></button><button onClick={onViewMeals}><div className="quick-icon teal"><Utensils size={19} /></div><span><strong>{activeMeals.length} وجبات نشطة</strong><small>متاحة للاختيار اليوم</small></span><ChevronLeft size={17} /></button><div className="tip-card"><div><BookOpen size={18} /><strong>نصيحة اليوم</strong></div><p>حافظ على تحديث أسعار الوجبات قبل بداية كل شهر لتكون التقارير دقيقة.</p></div></div></section></div></div>;
}
function StatCard({ icon, label, value, detail, tone }: { icon: React.ReactNode; label: string; value: string; detail: string; tone: string }) { return <div className="stat-card"><div className={`stat-icon ${tone}`}>{icon}</div><div className="stat-copy"><span>{label}</span><strong>{value}</strong><small>{detail}</small></div><MoreHorizontal className="stat-more" size={18} /></div>; }
function PanelHeading({ icon, title, action }: { icon: React.ReactNode; title: string; action?: React.ReactNode }) { return <div className="panel-heading"><div><span className="heading-icon">{icon}</span><h3>{title}</h3></div>{action}</div>; }
function OrderRow({ order, employees, onPrint }: { order: Order; employees: Employee[]; onPrint: (order: Order) => void }) { return <div className="order-row"><div className="order-symbol"><ReceiptText size={17} /></div><div className="order-main"><strong>{order.invoiceNo}</strong><span>{order.items.map((item) => `${item.mealName} × ${item.quantity}`).join("، ")}</span></div><div className="order-meta"><strong>{money(order.total)}</strong><span>{formatTime(order.createdAt)} · {order.participants.length} موظفين</span></div><button className="row-action" onClick={() => onPrint(order)} aria-label="طباعة الفاتورة"><Printer size={16} /></button></div>; }

function OrdersView({ orders, employees, query, setQuery, onNewOrder, onExport, onPrint, onClear }: { orders: Order[]; employees: Employee[]; query: string; setQuery: (value: string) => void; onNewOrder: () => void; onExport: () => void; onPrint: (order: Order) => void; onClear: () => void }) { return <div className="view-stack"><PageIntro eyebrow="سجل التشغيل" title="الطلبات اليومية" description="سجل واضح لكل طلب، مع توزيع التكلفة والفاتورة جاهزة للطباعة." actions={<><button className="secondary-btn" onClick={onExport}><FileSpreadsheet size={16} /> تصدير Excel</button><button className="primary-btn" onClick={onNewOrder}><Plus size={17} /> طلب جديد</button></>} /><section className="panel"><div className="toolbar"><div className="search-box"><Search size={17} /><input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="ابحث برقم الفاتورة أو الملاحظة..." /></div><div className="toolbar-note"><CalendarDays size={15} /> {orders.length} طلب مسجل</div></div>{orders.length ? <div className="wide-table"><div className="table-head"><span>الفاتورة</span><span>التاريخ</span><span>الأصناف</span><span>المشاركون</span><span>الإجمالي</span><span>إجراء</span></div>{orders.filter((order) => `${order.invoiceNo} ${order.note}`.includes(query)).map((order) => <div className="table-row" key={order.id}><strong>{order.invoiceNo}</strong><span>{formatDate(order.createdAt)}<small>{formatTime(order.createdAt)}</small></span><span>{order.items.map((item) => <small key={item.mealId}>{item.mealName} × {item.quantity}</small>)}</span><span><b className="avatar-stack">{order.participants.slice(0, 3).map((id) => <i key={id}>{(employees.find((employee) => employee.id === id)?.name ?? "م").charAt(0)}</i>)}{order.participants.length > 3 && <i>+{order.participants.length - 3}</i>}</b><small>{order.participants.length} موظف</small></span><strong>{money(order.total)}</strong><button className="row-action" onClick={() => onPrint(order)}><Printer size={16} /></button></div>)}</div> : <EmptyState icon={ClipboardList} title="لا توجد طلبات مسجلة" description="سجّل أول طلب ليظهر في السجل." action={<button className="primary-btn" onClick={onNewOrder}><Plus size={16} /> تسجيل طلب</button>} />}<div className="danger-zone"><button className="ghost-btn danger" onClick={onClear}><Trash2 size={15} /> مسح سجل الطلبات</button></div></section></div>; }

function MealsView({ meals, query, setQuery, onAdd, onRemove }: { meals: Meal[]; query: string; setQuery: (value: string) => void; onAdd: () => void; onRemove: (id: string) => void }) { return <div className="view-stack"><PageIntro eyebrow="إعدادات المطبخ" title="الوجبات والأسعار" description="أدر قائمة الوجبات والأسعار التي تظهر عند تسجيل الطلبات." actions={<button className="primary-btn" onClick={onAdd}><Plus size={17} /> إضافة وجبة</button>} /><section className="panel"><div className="toolbar"><div className="search-box"><Search size={17} /><input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="ابحث عن وجبة..." /></div><span className="pill"><Utensils size={14} /> {meals.length} وجبات نشطة</span></div><div className="entity-grid">{meals.map((meal) => <div className="entity-card" key={meal.id}><div className="entity-icon orange"><Utensils size={20} /></div><div className="entity-copy"><strong>{meal.name}</strong><span>متاحة للطلب</span></div><div className="entity-price">{money(meal.price)}</div><button className="icon-btn subtle" onClick={() => onRemove(meal.id)} aria-label="حذف الوجبة"><Trash2 size={16} /></button></div>)}</div>{!meals.length && <EmptyState icon={Utensils} title="لا توجد وجبات مطابقة" description="أضف وجبة جديدة أو جرّب كلمة بحث أخرى." />}</section></div>; }
function EmployeesView({ employees, query, setQuery, onAdd, onRemove }: { employees: Employee[]; query: string; setQuery: (value: string) => void; onAdd: () => void; onRemove: (id: string) => void }) { return <div className="view-stack"><PageIntro eyebrow="دليل الفريق" title="الموظفون" description="قائمة الموظفين المشاركين في توزيع تكاليف الوجبات." actions={<button className="primary-btn" onClick={onAdd}><Plus size={17} /> إضافة موظف</button>} /><section className="panel"><div className="toolbar"><div className="search-box"><Search size={17} /><input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="ابحث باسم الموظف..." /></div><span className="pill"><UsersRound size={14} /> {employees.length} موظف</span></div><div className="people-grid">{employees.map((employee, index) => <div className="person-card" key={employee.id}><div className={`person-avatar avatar-${index % 5}`}>{employee.name.charAt(0)}</div><div className="entity-copy"><strong>{employee.name}</strong><span><i className="online-dot" /> نشط في النظام</span></div><button className="icon-btn subtle" onClick={() => onRemove(employee.id)} aria-label="حذف الموظف"><Trash2 size={16} /></button></div>)}</div>{!employees.length && <EmptyState icon={UsersRound} title="لا يوجد موظفون مطابقون" description="أضف موظفًا جديدًا أو غيّر كلمة البحث." />}</section></div>; }
function ReportsView({ month, setMonth, orders, employees, total, discounts, onExport, onPrint, onSetDiscount, onExportEmployee }: { month: string; setMonth: (value: string) => void; orders: Order[]; employees: Employee[]; total: number; discounts: Record<string, number>; onExport: () => void; onPrint: () => void; onSetDiscount: (employee: Employee) => void; onExportEmployee: (employee: Employee) => void }) { const rows = employees.map((employee) => { const employeeOrders = orders.filter((order) => order.participants.includes(employee.id)); const subtotal = employeeOrders.reduce((sum, order) => sum + order.total / order.participants.length, 0); const discount = discounts[employee.id] ?? 0; return { employee, count: employeeOrders.length, subtotal, discount, total: subtotal - discount }; }).filter((row) => row.count || row.discount); return <div className="view-stack"><PageIntro eyebrow="التحليل المالي" title="التقارير الشهرية" description="كشف توزيع التكلفة لكل موظف حسب الشهر المختار." actions={<><button className="ghost-btn" onClick={onPrint}><Printer size={16} /> طباعة الكشف</button><button className="secondary-btn" onClick={onExport}><FileSpreadsheet size={16} /> تصدير كشف الشهر</button></>} /><section className="report-hero"><div><span>الفترة الحالية</span><strong>{monthLabel(month)}</strong><small>توزيع عادل حسب عدد المشاركين في كل طلب</small></div><div className="report-total"><span>إجمالي الفترة</span><strong>{money(total)}</strong></div><div className="report-filter"><label htmlFor="report-month">اختيار الشهر</label><input id="report-month" type="month" value={month} onChange={(event) => setMonth(event.target.value)} /></div></section><section className="panel"><PanelHeading icon={<BarChart3 size={17} />} title="كشف الموظفين" action={<span className="pill">{rows.length} موظف في الكشف</span>} />{rows.length ? <div className="wide-table report-table"><div className="table-head"><span>الموظف</span><span>الطلبات</span><span>الإجمالي</span><span>الخصم</span><span>الصافي</span><span>إجراء</span></div>{rows.map((row) => <div className="table-row" key={row.employee.id}><span className="person-inline"><i>{row.employee.name.charAt(0)}</i><strong>{row.employee.name}</strong></span><strong>{row.count}</strong><span>{money(row.subtotal)}</span><span className="discount-value">{money(row.discount)}</span><strong>{money(row.total)}</strong><span className="report-actions"><button className="row-action" onClick={() => onSetDiscount(row.employee)} title="إضافة أو تعديل الخصم"><Settings2 size={15} /></button><button className="row-action" onClick={() => onExportEmployee(row.employee)} title="تصدير كشف الموظف"><Download size={15} /></button></span></div>)}</div> : <EmptyState icon={BarChart3} title="لا توجد بيانات لهذا الشهر" description="سجّل طلبات أو أضف خصمًا ليظهر الموظف في الكشف." />}</section></div>; }
function InvoicesView({ orders, employees, onPrint, onDownload }: { orders: Order[]; employees: Employee[]; onPrint: (order: Order) => void; onDownload: (order: Order) => void }) { return <div className="view-stack"><PageIntro eyebrow="المستندات" title="الفواتير" description="الفواتير النصية الناتجة من الطلبات، جاهزة للتحميل أو الطباعة." /><section className="panel invoice-grid">{orders.length ? orders.map((order) => <div className="invoice-card" key={order.id}><div className="invoice-card-top"><div className="invoice-file"><FileText size={20} /></div><span className="status-badge"><Check size={13} /> محفوظة</span></div><strong>{order.invoiceNo}</strong><span>{formatDate(order.createdAt)} · {money(order.total)}</span><div className="invoice-actions"><button className="secondary-btn" onClick={() => onDownload(order)}><ArrowDownToLine size={15} /> تحميل</button><button className="ghost-btn" onClick={() => onPrint(order)}><Printer size={15} /> طباعة</button></div></div>) : <EmptyState icon={ReceiptText} title="لا توجد فواتير بعد" description="بعد تسجيل أول طلب ستظهر فاتورته هنا." />}</section></div>; }
function PageIntro({ eyebrow, title, description, actions }: { eyebrow: string; title: string; description: string; actions?: React.ReactNode }) { return <div className="page-intro"><div><div className="eyebrow accent">{eyebrow}</div><h2>{title}</h2><p>{description}</p></div><div className="intro-actions">{actions}</div></div>; }
function Modal({ title, icon, onClose, children, wide = false }: { title: string; icon: React.ReactNode; onClose: () => void; children: React.ReactNode; wide?: boolean }) { return <div className="modal-backdrop" onMouseDown={(event) => { if (event.target === event.currentTarget) onClose(); }}><div className={`modal-card ${wide ? "wide" : ""}`}><div className="modal-header"><div><span className="heading-icon">{icon}</span><h3>{title}</h3></div><button className="icon-btn" onClick={onClose} aria-label="إغلاق"><X size={18} /></button></div>{children}</div></div>; }

export { STORAGE_KEY };
