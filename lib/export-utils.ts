import * as XLSX from "xlsx";
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import { formatDateIndia, formatDateTime } from "./utils";

export interface ExportApplicationItem {
  id: string;
  referenceNumber: string;
  employeeName: string;
  employeeEmail: string;
  employeeIdNumber?: string | null;
  departmentSnapshot: string;
  leaveTypeSnapshot: string;
  startDate: string;
  endDate: string;
  totalDays: number;
  isHalfDay?: boolean;
  status: string;
  reportingManagerName?: string | null;
  handoverPersonName?: string | null;
  reason?: string | null;
  submittedAt: string;
  decisionBy?: string | null;
}

export interface PDFExportOptions {
  activeFilters?: string[];
  totalRecords?: number;
}

export function exportApplicationsToExcel(
  applications: ExportApplicationItem[],
  filename?: string
) {
  const rows = applications.map((app, index) => ({
    "S.No": index + 1,
    "Reference Number": app.referenceNumber,
    "Employee Name": app.employeeName,
    "Employee Email": app.employeeEmail,
    "Employee ID": app.employeeIdNumber || "N/A",
    "Department": app.departmentSnapshot,
    "Leave Type": app.leaveTypeSnapshot,
    "Start Date": formatDateIndia(app.startDate),
    "End Date": formatDateIndia(app.endDate),
    "Duration (Days)": app.totalDays,
    "Half Day": app.isHalfDay ? "Yes" : "No",
    "Status": app.status,
    "Decision By": app.decisionBy || "—",
    "Reporting Manager": app.reportingManagerName || "N/A",
    "Handover Person": app.handoverPersonName || "N/A",
    "Reason": app.reason || "N/A",
    "Submitted Date": formatDateTime(app.submittedAt),
  }));

  const worksheet = XLSX.utils.json_to_sheet(rows);

  worksheet["!cols"] = [
    { wch: 6 },
    { wch: 22 },
    { wch: 22 },
    { wch: 26 },
    { wch: 14 },
    { wch: 18 },
    { wch: 18 },
    { wch: 14 },
    { wch: 14 },
    { wch: 15 },
    { wch: 10 },
    { wch: 12 },
    { wch: 18 },
    { wch: 20 },
    { wch: 20 },
    { wch: 32 },
    { wch: 22 },
  ];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Leave Applications");

  const dateStr = new Date().toISOString().slice(0, 10);
  const finalFilename = filename || `leave-applications-report-${dateStr}.xlsx`;
  XLSX.writeFile(workbook, finalFilename);
}

export function exportApplicationsToPDF(
  applications: ExportApplicationItem[],
  options?: PDFExportOptions
) {
  const doc = new jsPDF({
    orientation: "landscape",
    unit: "mm",
    format: "a4",
  });

  const pageWidth = doc.internal.pageSize.getWidth();

  // Header Title
  doc.setFontSize(16);
  doc.setTextColor(17, 24, 39);
  doc.text("VRD Exports — Leave Applications Report", 14, 15);

  // Metadata Subtitle
  doc.setFontSize(9);
  doc.setTextColor(107, 114, 128);
  const count = options?.totalRecords ?? applications.length;
  const metaText = `Generated: ${formatDateTime(new Date().toISOString())}  |  Total Applications: ${count}`;
  doc.text(metaText, 14, 21);

  // Active filters summary
  const filterText = options?.activeFilters?.length
    ? `Filters: ${options.activeFilters.join("  •  ")}`
    : "";
  if (filterText) {
    doc.setFontSize(8);
    doc.setTextColor(156, 163, 175);
    doc.text(filterText, 14, 26);
  }

  const tableData = applications.map((app, index) => [
    String(index + 1),
    app.referenceNumber,
    `${app.employeeName}\n${app.employeeEmail}`,
    app.departmentSnapshot,
    app.leaveTypeSnapshot,
    `${formatDateIndia(app.startDate)}\nto ${formatDateIndia(app.endDate)}`,
    String(app.totalDays),
    app.status,
    app.decisionBy || "—",
    formatDateTime(app.submittedAt),
  ]);

  autoTable(doc, {
    startY: filterText ? 30 : 25,
    head: [[
      "#",
      "Reference",
      "Employee",
      "Department",
      "Leave Type",
      "Dates",
      "Days",
      "Status",
      "Decision By",
      "Submitted",
    ]],
    body: tableData,
    theme: "striped",
    headStyles: {
      fillColor: [17, 24, 39],
      textColor: [255, 255, 255],
      fontSize: 8,
      fontStyle: "bold",
    },
    styles: {
      fontSize: 7.5,
      cellPadding: 2.5,
      textColor: [55, 65, 81],
      overflow: "linebreak",
    },
    alternateRowStyles: {
      fillColor: [249, 250, 251],
    },
    columnStyles: {
      0: { cellWidth: 8, halign: "center" },
      1: { cellWidth: 32, fontStyle: "bold" },
      2: { cellWidth: 46 },
      3: { cellWidth: 30 },
      4: { cellWidth: 28 },
      5: { cellWidth: 28 },
      6: { cellWidth: 14, halign: "center" },
      7: { cellWidth: 22, halign: "center" },
      8: { cellWidth: 28 },
      9: { cellWidth: 32 },
    },
    didDrawPage: (data) => {
      const pageCount = (doc.internal as unknown as { getNumberOfPages: () => number }).getNumberOfPages();
      doc.setFontSize(8);
      doc.setTextColor(156, 163, 175);
      doc.text(
        `Page ${data.pageNumber} of ${pageCount}`,
        pageWidth - 25,
        doc.internal.pageSize.getHeight() - 8
      );
    },
  });

  const dateStr = new Date().toISOString().slice(0, 10);
  doc.save(`leave-applications-report-${dateStr}.pdf`);
}
