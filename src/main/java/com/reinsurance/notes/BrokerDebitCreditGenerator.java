package com.reinsurance.notes;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;

public class BrokerDebitCreditGenerator {

    public static void main(String[] args) {

        // ✅ Detect base directory (works for both IntelliJ and JAR/BAT)
        String basePath = System.getProperty("user.dir");

        // Default locations for packaged use
        String excelFilePath = basePath + File.separator + "resources" + File.separator + "DebitNoteCalculations.xlsx";
        String templatePath = basePath + File.separator + "resources" + File.separator + "DebitNoteTemplate.docx";
        String outputFolder = basePath + File.separator + "resources" + File.separator + "output" + File.separator;

        // ✅ Fallback for IntelliJ execution
        File excelFile = new File(excelFilePath);
        if (!excelFile.exists()) {
            excelFilePath = basePath + File.separator + "src" + File.separator + "main" + File.separator + "resources" + File.separator + "DebitNoteCalculations.xlsx";
            templatePath = basePath + File.separator + "src" + File.separator + "main" + File.separator + "resources" + File.separator + "DebitNoteTemplate.docx";
            outputFolder = basePath + File.separator + "src" + File.separator + "main" + File.separator + "resources" + File.separator + "output" + File.separator;
        }

        System.out.println("==============================================");
        System.out.println("   Reinsurance Debit Note Generator v2.1");
        System.out.println("==============================================");
        System.out.println("Excel File: " + excelFilePath);
        System.out.println("Template File: " + templatePath);
        System.out.println("Output Folder: " + outputFolder);
        System.out.println("Processing data...\n");

        try (FileInputStream fis = new FileInputStream(excelFilePath);
             Workbook wb = new XSSFWorkbook(fis)) {

            Sheet sheet = wb.getSheetAt(0);
            File outDir = new File(outputFolder);
            if (!outDir.exists()) outDir.mkdirs();

            DateTimeFormatter df = DateTimeFormatter.ofPattern("dd-MMM-yyyy");
            int processedCol = 21; // ✅ last column (Processed)

            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null || isRowEmpty(row)) {
                    System.out.println("⚠️ Skipping blank row " + r);
                    continue;
                }

                String processedFlag = getString(row, processedCol);
                if (processedFlag.equalsIgnoreCase("Yes") || processedFlag.equalsIgnoreCase("Processed")) {
                    System.out.println("⏩ Skipping row " + r + " (already processed)");
                    continue;
                }

                // --- Input Columns ---
                String debitNoteNo = getString(row, 0);
                if (debitNoteNo == null || debitNoteNo.isEmpty()) {
                    debitNoteNo = "DN-" + String.format("%03d", r);
                }

                String docDate = getString(row, 1);
                if (docDate.isEmpty()) docDate = LocalDate.now().format(df);

                String interest = getString(row, 2);
                String insured = getString(row, 3);
                String reinsurer = getString(row, 4);
                String period = getString(row, 5);

                double SI = getDouble(row, 6);
                double cedentRate = getDouble(row, 7);
                double reinsRate = getDouble(row, 8);
                double share = getDouble(row, 9);
                double brokerage = getDouble(row, 10);
                double cedingCommPct = getDouble(row, 15); // 🆕 Ceding Commission %

                if (SI == 0 || cedentRate == 0 || share == 0) {
                    System.out.println("⚠️ Skipping incomplete row " + r);
                    continue;
                }

                // --- Calculations ---
                // Cedent side
                double grossPremiumCedent = SI * (cedentRate / 100);
                double sharePremiumCedent = grossPremiumCedent * (share / 100);

                // Reinsurer side (use Reinsurance Rate if provided)
                double effectiveReinsRate = (reinsRate > 0) ? reinsRate : cedentRate;
                double grossPremiumReins = SI * (effectiveReinsRate / 100);
                double sharePremiumReins = grossPremiumReins * (share / 100);

                // Common deductions
                double cedingCommissionAmt = sharePremiumCedent * (cedingCommPct / 100);
                double grossBrokerage = sharePremiumReins * (brokerage / 100);
                double netBrokerage = grossBrokerage / 2.0;

                // Correct accounting logic
                double netPremiumFromYou = sharePremiumCedent - cedingCommissionAmt; // Debit (Cedent)
                double netPremiumToYou = sharePremiumReins - netBrokerage - cedingCommissionAmt; // Credit (Reinsurer)

                // --- Write Calculated Values to Excel ---
                setNumeric(row, 11, grossPremiumCedent);
                setNumeric(row, 12, sharePremiumCedent);
                setNumeric(row, 13, grossPremiumReins);
                setNumeric(row, 14, sharePremiumReins);
                setNumeric(row, 15, cedingCommPct);
                setNumeric(row, 16, cedingCommissionAmt);
                setNumeric(row, 17, grossBrokerage);
                setNumeric(row, 18, netBrokerage);
                setNumeric(row, 19, netPremiumFromYou);
                setNumeric(row, 20, netPremiumToYou);
                setString(row, 21, "Yes");

                // --- Generate Debit Note Word ---
                String safeFileName = debitNoteNo.replaceAll("[^a-zA-Z0-9-_]", "_");
                generateDebitNote(
                        templatePath,
                        outputFolder + safeFileName + ".docx",
                        debitNoteNo,
                        docDate,
                        interest,
                        insured,
                        reinsurer,
                        period,
                        SI,
                        cedentRate,
                        grossPremiumCedent,
                        share,
                        sharePremiumCedent,
                        netPremiumFromYou // Debit side
                );

                System.out.println("✅ Processed and generated Debit Note for: " + debitNoteNo);
            }

            // Save updated Excel
            try (FileOutputStream fos = new FileOutputStream(excelFilePath)) {
                wb.write(fos);
            }

            System.out.println("\n✅ All Debit Notes Processed and Excel Updated Successfully.");
            openOutputFolder(outputFolder);

        } catch (FileNotFoundException e) {
            System.err.println("❌ File not found! Please check folder structure:");
            System.err.println("Expected Excel file at: " + excelFilePath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // --- Utility: Check if row is empty ---
    private static boolean isRowEmpty(Row row) {
        if (row == null) return true;
        for (int c = 0; c <= 10; c++) {
            Cell cell = row.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
            if (cell != null && cell.getCellType() != CellType.BLANK) {
                if (cell.getCellType() == CellType.STRING && !cell.getStringCellValue().trim().isEmpty())
                    return false;
                if (cell.getCellType() == CellType.NUMERIC && cell.getNumericCellValue() != 0)
                    return false;
            }
        }
        return true;
    }

    // --- Excel Helpers ---
    private static String getString(Row row, int idx) {
        Cell c = row.getCell(idx, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        if (c == null) return "";
        if (c.getCellType() == CellType.STRING) return c.getStringCellValue().trim();
        if (c.getCellType() == CellType.NUMERIC) return String.valueOf(c.getNumericCellValue());
        return "";
    }

    private static double getDouble(Row row, int idx) {
        Cell c = row.getCell(idx, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        if (c == null) return 0.0;
        if (c.getCellType() == CellType.NUMERIC) return c.getNumericCellValue();
        if (c.getCellType() == CellType.STRING) {
            try {
                return Double.parseDouble(c.getStringCellValue().replace(",", "").trim());
            } catch (Exception e) {
                return 0.0;
            }
        }
        return 0.0;
    }

    private static void setNumeric(Row row, int idx, double val) {
        if (idx > 21) return;
        Cell c = row.getCell(idx);
        if (c == null) c = row.createCell(idx, CellType.NUMERIC);
        c.setCellValue(val);
    }

    private static void setString(Row row, int idx, String val) {
        if (idx > 21) return;
        Cell c = row.getCell(idx);
        if (c == null) c = row.createCell(idx, CellType.STRING);
        c.setCellValue(val);
    }

    // --- Word Template Generator ---
    private static void setCellText(XWPFTable table, int rowIdx, int colIdx, String text) {
        XWPFTableRow row = table.getRow(rowIdx);
        if (row != null && row.getCell(colIdx) != null) {
            row.getCell(colIdx).removeParagraph(0);
            row.getCell(colIdx).setText(text);
        }
    }

    private static void generateDebitNote(
            String templatePath, String outputPath,
            String debitNoteNo, String documentDate, String interest,
            String insured, String reinsurer, String period,
            double sumInsured, double rate, double facPremiumFull,
            double share, double sharePremium, double netPremiumFromYou)
            throws IOException, InvalidFormatException {

        try (FileInputStream fis = new FileInputStream(templatePath);
             XWPFDocument doc = new XWPFDocument(fis)) {

            XWPFTable table = doc.getTables().get(0);

            setCellText(table, 0, 2, debitNoteNo);
            setCellText(table, 1, 2, documentDate);
            setCellText(table, 3, 2, interest);
            setCellText(table, 4, 2, insured);
            setCellText(table, 5, 2, reinsurer);
            setCellText(table, 6, 2, period);
            setCellText(table, 7, 2, "USD " + format(sumInsured));
            setCellText(table, 8, 2, String.format("%.2f%%", rate));
            setCellText(table, 9, 2, "USD " + format(facPremiumFull));
            setCellText(table, 10, 2, String.format("%.2f%% of 100%%", share));
            setCellText(table, 11, 2, "USD " + format(sharePremium));
            setCellText(table, 12, 2, "USD " + format(netPremiumFromYou));

            File outFile = new File(outputPath);
            outFile.getParentFile().mkdirs();
            try (FileOutputStream fos = new FileOutputStream(outFile)) {
                doc.write(fos);
            }
        }
    }

    private static String format(double val) {
        return String.format("%,.2f", val);
    }

    // --- Auto open folder ---
    private static void openOutputFolder(String outputFolder) {
        try {
            File folder = new File(outputFolder);
            if (folder.exists()) {
                new ProcessBuilder("explorer.exe", folder.getAbsolutePath()).start();
            }
        } catch (IOException e) {
            System.err.println("⚠️ Unable to open output folder automatically.");
        }
    }
}
